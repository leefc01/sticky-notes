// Tests notes.js against a small mock of the PowerPoint JavaScript API.
// Run: node --test test/
const test = require("node:test");
const assert = require("node:assert");

let nextId = 1;
function mkShape(name, tags = {}) {
  const shape = {
    id: "sh" + nextId++, name, visible: true, deleted: false, left: 0, top: 0,
    _tags: { ...tags },
    tags: {
      load() {},
      get items() { return Object.entries(shape._tags).map(([key, value]) => ({ key, value })); },
      add(k, v) { shape._tags[k.toUpperCase()] = v; },
      delete(k) { delete shape._tags[k.toUpperCase()]; },
    },
    fill: { setSolidColor(c) { shape.fillColor = c; } },
    lineFormat: {},
    textFrame: { textRange: { text: "", font: {}, paragraphFormat: {}, load() {},
      getSubstring(start, length) {
        const tr = shape.textFrame.textRange;
        return {
          setSelected() {
            if (global.__deck.textSelectFails) global.__deck.pendingFail = true;
            else global.__deck.textSelection = { shape: shape.id, start, length };
          },
          set text(v) { tr.text = tr.text.slice(0, start) + v + tr.text.slice(start + length); },
        };
      } } },
    delete() { shape.deleted = true; },
    load() {},
  };
  return shape;
}

function mkDeck(spec) {
  // spec: array per slide of shapes
  const slides = spec.map((shapes, i) => {
    const slide = { id: "s" + i, _shapes: shapes, setSelectedShapes(ids) { deck.shapeSelection = ids; } };
    slide.shapes = {
      load() {},
      get items() { return slide._shapes.filter((s) => !s.deleted); },
      addGeometricShape(type, o) {
        const s = mkShape("tmp"); Object.assign(s, o); s.type = type; slide._shapes.push(s); return s;
      },
    };
    return slide;
  });
  const deck = { slides, selectedSlide: 0, selectedShapes: [], syncs: 0, apiLevel: "1.10" };
  global.__deck = deck;
  const presentation = {
    slides: { load() {}, get items() { return slides; } },
    getSelectedSlides() { return { load() {}, items: [slides[deck.selectedSlide]] }; },
    setSelectedSlides(ids) { deck.selectedSlide = slides.findIndex((s) => s.id === ids[0]); },
    getSelectedShapes() { return { load() {}, items: deck.selectedShapes }; },
  };
  global.PowerPoint = {
    run: (fn) => Promise.resolve().then(() => fn({ presentation, sync: () => {
      deck.syncs++;
      if (deck.pendingFail) { deck.pendingFail = false; return Promise.reject(new Error("not supported")); }
      return Promise.resolve();
    } })),
    GeometricShapeType: { rectangle: "Rectangle" },
    TextVerticalAlignment: { top: "Top" },
    ParagraphHorizontalAlignment: { left: "Left" },
  };
  global.Office = {
    context: { requirements: { isSetSupported: (n, v) => n === "PowerPointApi" && +v.split(".")[1] <= +deck.apiLevel.split(".")[1] } },
  };
  return deck;
}

const note = () => mkShape("SGNote_x", { SGNOTE: "1" });
const notesOn = (deck, i) => deck.slides[i]._shapes.filter((s) => !s.deleted && s._tags.SGNOTE === "1");
const store = {};
global.localStorage = { getItem: (k) => (k in store ? store[k] : null), setItem: (k, v) => { store[k] = String(v); } };
const api = require("../notes.js");
function resetPrefs(p = {}) { delete store["sgnotes.prefs.v1"]; api._reloadPrefs(); if (Object.keys(p).length) api.savePrefs(p); }

test("Add note: flush top-right, tagged, styled; second note cascades", async () => {
  resetPrefs();
  const deck = mkDeck([[], [mkShape("Title 1")]]);
  deck.selectedSlide = 1;
  await api.addNote();
  await api.addNote();
  const ns = notesOn(deck, 1);
  assert.strictEqual(ns.length, 2);
  assert.deepStrictEqual([ns[0].left, ns[0].top, ns[0].width, ns[0].height], [816, 0, 144, 108]);
  assert.deepStrictEqual([ns[1].left, ns[1].top], [798, 18]);
  assert.match(ns[0].name, /^SGNote_/);
  assert.strictEqual(ns[0].fillColor, "#FFE97A");
  assert.strictEqual(ns[0].textFrame.textRange.text, "Note");
  assert.strictEqual(ns[0].altTextTitle, "Sticky note");
  assert.strictEqual(api.ui.lastStatus, "2 notes on 1 slide");
});

test("Shapes that aren't notes are ignored (no tag, wrong name)", async () => {
  const deck = mkDeck([[mkShape("SGNote_fake"), mkShape("Box", { SGNOTE: "1" }), note()]]);
  const s = await PowerPoint.run((ctx) => api.scan(ctx));
  assert.strictEqual(s.notes.length, 1);
  await api.deleteNotes();
  assert.strictEqual(deck.slides[0]._shapes.filter((x) => x.deleted).length, 1);
});

test("Next / Previous skip empty slides and wrap", async () => {
  const deck = mkDeck([[], [note()], [], [note()], []]);
  deck.selectedSlide = 0;
  await api.next(); assert.strictEqual(deck.selectedSlide, 1);
  await api.next(); assert.strictEqual(deck.selectedSlide, 3);
  await api.next(); assert.strictEqual(deck.selectedSlide, 1, "wraps to first");
  await api.previous(); assert.strictEqual(deck.selectedSlide, 3, "wraps to last");
  deck.selectedSlide = 4;
  await api.previous(); assert.strictEqual(deck.selectedSlide, 3);
});

test("Next with no notes shows a message and stays put", async () => {
  const deck = mkDeck([[], []]);
  deck.selectedSlide = 1;
  await api.next();
  assert.strictEqual(deck.selectedSlide, 1);
  assert.strictEqual(api.ui.lastMsg, "No notes in this deck");
});

test("Delete: selected note only; otherwise all on slide; message when none", async () => {
  const a = note(), b = note(), title = mkShape("Title");
  const deck = mkDeck([[a, b, title], [note()]]);
  deck.selectedShapes = [{ id: a.id }];
  await api.deleteNotes();
  assert.deepStrictEqual([a.deleted, b.deleted, title.deleted], [true, false, false]);

  deck.selectedShapes = [{ id: title.id }]; // a non-note selected → all notes on the slide
  await api.deleteNotes();
  assert.deepStrictEqual([b.deleted, title.deleted], [true, false]);
  assert.strictEqual(notesOn(deck, 1).length, 1, "other slides untouched");

  deck.selectedShapes = [];
  await api.deleteNotes();
  assert.strictEqual(api.ui.lastMsg, "No notes on this slide");
});

test("Delete all: asks with counts, cancel keeps, confirm removes only notes", async () => {
  const title = mkShape("Title");
  const deck = mkDeck([[note(), note(), title], [], [note()]]);
  await api.askDeleteAll();
  assert.strictEqual(api._pending(), true);
  api.cancelDeleteAll();
  await api.confirmDeleteAll(); // no-op after cancel
  assert.strictEqual(notesOn(deck, 0).length, 2);

  await api.askDeleteAll();
  await api.confirmDeleteAll();
  assert.strictEqual(notesOn(deck, 0).length + notesOn(deck, 2).length, 0);
  assert.strictEqual(title.deleted, false);
  assert.strictEqual(api.ui.lastMsg, "Deleted 3 notes");
});

test("Hide all / Show all; new notes after Hide all are visible", async () => {
  const n1 = note(), n2 = note(), other = mkShape("Pic");
  const deck = mkDeck([[n1, other], [n2]]);
  await api.hideAll();
  assert.deepStrictEqual([n1.visible, n2.visible, other.visible], [false, false, true]);
  deck.selectedSlide = 1;
  await api.addNote();
  assert.strictEqual(notesOn(deck, 1).at(-1).visible, true);
  await api.showAll();
  assert.deepStrictEqual([n1.visible, n2.visible], [true, true]);
});

test("Hide all on an older build: message, nothing changes", async () => {
  const n1 = note();
  const deck = mkDeck([[n1]]);
  deck.apiLevel = "1.8";
  await api.hideAll();
  assert.strictEqual(n1.visible, true);
  assert.match(api.ui.lastMsg, /Update Office/);
});

test("Large deck: each command uses a handful of round-trips", async () => {
  const spec = Array.from({ length: 150 }, (_, i) =>
    [mkShape("Title"), mkShape("Body"), ...(i % 3 === 0 ? [note()] : [])]);
  const deck = mkDeck(spec);
  for (const [name, fn] of [["next", api.next], ["hide", api.hideAll], ["show", api.showAll], ["add", api.addNote], ["delete", api.deleteNotes]]) {
    deck.syncs = 0;
    await fn();
    assert.ok(deck.syncs <= 6, `${name} used ${deck.syncs} syncs`);
  }
});

test("Position setting: top left and top center, with cascade", async () => {
  resetPrefs({ position: "tl" });
  let deck = mkDeck([[]]);
  await api.addNote(); await api.addNote();
  assert.deepStrictEqual(notesOn(deck, 0).map((s) => [s.left, s.top]), [[0, 0], [18, 18]]);
  resetPrefs({ position: "tc" });
  deck = mkDeck([[]]);
  await api.addNote(); await api.addNote();
  assert.deepStrictEqual(notesOn(deck, 0).map((s) => [s.left, s.top]), [[408, 0], [408, 18]]);
});

test("Top right uses the real slide width (4:3 deck)", () => {
  assert.deepStrictEqual(api._notePosition("tr", 720, 0), { left: 576, top: 0 });
});

test("Color setting applies to new notes only, and is tagged", async () => {
  resetPrefs();
  const deck = mkDeck([[]]);
  await api.addNote();
  api.savePrefs({ color: "blue" });
  await api.addNote();
  const [a, b] = notesOn(deck, 0);
  assert.deepStrictEqual([a.fillColor, b.fillColor], ["#FFE97A", "#BFE3FF"]);
  assert.strictEqual(b.lineFormat.color, "#5FA8E0");
  assert.strictEqual(b._tags.SGNOTE_COLOR, "blue");
});

test("Apply color to all notes recolors every note, nothing else", async () => {
  resetPrefs({ color: "green" });
  const other = mkShape("Pic");
  const deck = mkDeck([[note(), other], [note()]]);
  await api.recolorAll();
  const all = [...notesOn(deck, 0), ...notesOn(deck, 1)];
  assert.ok(all.every((s) => s.fillColor === "#C8F0B8" && s._tags.SGNOTE_COLOR === "green"));
  assert.strictEqual(other.fillColor, undefined);
  assert.strictEqual(api.ui.lastMsg, "Recolored 2 notes green");
});

test("Prefix: text starts with prefix, only the placeholder word is selected", async () => {
  resetPrefs({ prefix: "FL:" });
  const deck = mkDeck([[]]);
  await api.addNote();
  const n = notesOn(deck, 0)[0];
  assert.strictEqual(n.textFrame.textRange.text, "FL: Note");
  assert.deepStrictEqual(deck.textSelection, { shape: n.id, start: 4, length: 4 });
});

test("No prefix: whole placeholder selected for typing", async () => {
  resetPrefs();
  const deck = mkDeck([[]]);
  await api.addNote();
  const n = notesOn(deck, 0)[0];
  assert.strictEqual(n.textFrame.textRange.text, "Note");
  assert.deepStrictEqual(deck.textSelection, { shape: n.id, start: 0, length: 4 });
});

test("If text selection fails, the note is still added and selected as a shape", async () => {
  resetPrefs();
  const deck = mkDeck([[]]);
  deck.textSelectFails = true;
  await api.addNote();
  const n = notesOn(deck, 0)[0];
  assert.ok(n, "note added");
  assert.deepStrictEqual(deck.shapeSelection, [n.id]);
  assert.strictEqual(api.ui.lastMsg, "", "no error shown");
});

test("Preferences persist and bad stored values fall back to defaults", () => {
  resetPrefs({ position: "tc", color: "pink", prefix: "ABCDEFGHIJKLMNOPQRSTUVWXYZ" });
  assert.deepStrictEqual(api._reloadPrefs(), { position: "tc", color: "pink", prefix: "ABCDEFGHIJKLMNOPQRST", stampOn: false, who: "", skipDone: true });
  store["sgnotes.prefs.v1"] = JSON.stringify({ position: "bottom", color: "red", prefix: 5 });
  assert.deepStrictEqual(api._reloadPrefs(), { position: "tr", color: "yellow", prefix: "", stampOn: false, who: "", skipDone: true });
  store["sgnotes.prefs.v1"] = "not json";
  assert.deepStrictEqual(api._reloadPrefs(), { position: "tr", color: "yellow", prefix: "", stampOn: false, who: "", skipDone: true });
});

test("Changing position: a lone note at the new spot isn't offset", async () => {
  resetPrefs({ position: "tl" });
  const deck = mkDeck([[]]);
  await api.addNote();                       // top left, slot 0
  api.savePrefs({ position: "tr" });
  await api.addNote();                       // top right should be slot 0, not 1
  await api.addNote();                       // then cascade at top right
  assert.deepStrictEqual(notesOn(deck, 0).map((s) => [s.left, s.top]), [[0, 0], [816, 0], [798, 18]]);
});

test("A note dragged away frees its slot; gaps are filled first", () => {
  const w = 960;
  assert.strictEqual(api._freeSlot("tr", w, [{ left: 500, top: 300 }]), 0, "moved note doesn't count");
  assert.strictEqual(api._freeSlot("tr", w, [{ left: 798, top: 18 }]), 0, "slot 0 free even if slot 1 taken");
  assert.strictEqual(api._freeSlot("tr", w, [{ left: 816, top: 0 }, { left: 798, top: 18 }]), 2);
});

test("Settings order and labels are title case, Top Right default", () => {
  const html = require("fs").readFileSync(require("path").join(__dirname, "..", "taskpane.html"), "utf8");
  const order = [...html.matchAll(/data-position="(\w+)">([^<]+)</g)].map((m) => [m[1], m[2]]);
  assert.deepStrictEqual(order, [["tl", "Top Left"], ["tc", "Top Center"], ["tr", "Top Right"]]);
  resetPrefs();
  assert.strictEqual(api.prefs().position, "tr");
});

test("Pane version matches the manifest version", () => {
  const src = require("fs").readFileSync(require("path").join(__dirname, "..", "tools", "build.py"), "utf8");
  const manifestVersion = src.match(/VERSION = "([\d.]+)"/)[1];
  assert.strictEqual(manifestVersion, api.version + ".0");
});

// ---------- v1.3: Done ----------

const tagged = (text, extra = {}) => { const s = mkShape("SGNote_x", { SGNOTE: "1", SGNOTE_COLOR: "blue", ...extra }); s.textFrame.textRange.text = text; s.fillColor = "#BFE3FF"; return s; };

test("Done fill is a faint tint of the note's own color", () => {
  assert.strictEqual(api._doneFill("yellow"), "#EEE7C6");
  assert.notStrictEqual(api._doneFill("blue"), api._doneFill("pink"));
});

test("Toggle Done on a selected note: tinted, gray, ✓, tagged; others untouched", async () => {
  resetPrefs();
  const a = tagged("Fix chart"), b = tagged("Other");
  const deck = mkDeck([[a, b]]);
  deck.selectedShapes = [{ id: a.id }];
  await api.toggleDone();
  assert.strictEqual(a.textFrame.textRange.text, "✓ Fix chart");
  assert.strictEqual(a.fillColor, api._doneFill("blue"));
  assert.strictEqual(a.lineFormat.color, "#A6A6A6");
  assert.strictEqual(a.textFrame.textRange.font.color, "#6B6B6B");
  assert.strictEqual(a._tags.SGNOTE_DONE, "1");
  assert.strictEqual(b.textFrame.textRange.text, "Other");
  assert.strictEqual(api.ui.lastMsg, "Marked 1 note done");
});

test("Smart toggle: mixed selection marks all done (no double ✓); all done reopens all", async () => {
  resetPrefs();
  const a = tagged("A"), b = tagged("B");
  const deck = mkDeck([[a, b]]);
  deck.selectedShapes = [{ id: a.id }];
  await api.toggleDone();                    // A done
  deck.selectedShapes = [{ id: a.id }, { id: b.id }];
  await api.toggleDone();                    // mixed -> both done
  assert.deepStrictEqual([a.textFrame.textRange.text, b.textFrame.textRange.text], ["✓ A", "✓ B"]);
  assert.strictEqual(api.ui.lastMsg, "Marked 1 note done");
  await api.toggleDone();                    // all done -> reopen
  assert.deepStrictEqual([a.textFrame.textRange.text, b.textFrame.textRange.text], ["A", "B"]);
  assert.deepStrictEqual([a.fillColor, a.lineFormat.color, a.textFrame.textRange.font.color], ["#BFE3FF", "#5FA8E0", "#333333"]);
  assert.strictEqual(a._tags.SGNOTE_DONE, "0");
  assert.strictEqual(api.ui.lastMsg, "Reopened 2 notes");
});

test("Toggle Done with nothing selected applies to every note on the slide", async () => {
  resetPrefs();
  const a = tagged("A"), b = tagged("B"), c = tagged("C");
  const deck = mkDeck([[a, b], [c]]);
  await api.toggleDone();
  assert.deepStrictEqual([a._tags.SGNOTE_DONE, b._tags.SGNOTE_DONE, c._tags.SGNOTE_DONE], ["1", "1", undefined]);
  mkDeck([[]]);
  await api.toggleDone();
  assert.strictEqual(api.ui.lastMsg, "No notes on this slide");
});

test("Stamp text: Who, else prefix without colon, else date only", () => {
  const d = new Date(2026, 9, 23);
  assert.strictEqual(api._stampText({ who: "FL", prefix: "X:" }, d), "— Done FL · Oct 23");
  assert.strictEqual(api._stampText({ who: " ", prefix: "FL:" }, d), "— Done FL · Oct 23");
  assert.strictEqual(api._stampText({ who: "", prefix: "" }, d), "— Done Oct 23");
});

test("Stamp added as last line; reopen removes ✓ and stamp; re-mark adds a fresh stamp", async () => {
  resetPrefs({ stampOn: true, who: "FL" });
  const a = tagged("FL: Check Q3");
  const deck = mkDeck([[a]]);
  await api.toggleDone();
  const stamp = api._stampText(api.prefs());
  assert.strictEqual(a.textFrame.textRange.text, "✓ FL: Check Q3\n" + stamp);
  assert.strictEqual(a._tags.SGNOTE_STAMP, stamp);
  await api.toggleDone();
  assert.strictEqual(a.textFrame.textRange.text, "FL: Check Q3");
  assert.strictEqual(a._tags.SGNOTE_STAMP, undefined);
  await api.toggleDone();
  assert.strictEqual(a.textFrame.textRange.text, "✓ FL: Check Q3\n" + stamp);
});

test("Reopen leaves a hand-edited stamp line alone (still removes ✓); PowerPoint \\r breaks work", async () => {
  resetPrefs();
  const edited = tagged("✓ Fix\n— Done FL · Oct 2", { SGNOTE_DONE: "1", SGNOTE_STAMP: "— Done FL · Oct 23" });
  const crlf = tagged("✓ Fix\r— Done FL · Oct 23", { SGNOTE_DONE: "1", SGNOTE_STAMP: "— Done FL · Oct 23" });
  const deck = mkDeck([[edited], [crlf]]);
  await api.toggleDone();
  assert.strictEqual(edited.textFrame.textRange.text, "Fix\n— Done FL · Oct 2");
  deck.selectedSlide = 1;
  await api.toggleDone();
  assert.strictEqual(crlf.textFrame.textRange.text, "Fix");
});

test("v1.0 notes without a color tag reopen as yellow", async () => {
  resetPrefs();
  const old = mkShape("SGNote_x", { SGNOTE: "1" }); old.textFrame.textRange.text = "Old";
  const deck = mkDeck([[old]]);
  await api.toggleDone(); await api.toggleDone();
  assert.strictEqual(old.fillColor, "#FFE97A");
});

test("Delete Done: counts, cancel keeps, confirm removes only done notes", async () => {
  resetPrefs();
  const d1 = tagged("✓ a", { SGNOTE_DONE: "1" }), d2 = tagged("✓ b", { SGNOTE_DONE: "1" }), o = tagged("open"), other = mkShape("Title");
  const deck = mkDeck([[d1, o, other], [d2]]);
  await api.askDeleteDone();
  assert.strictEqual(api._pending(), true);
  api.cancelDeleteAll();
  await api.confirmPending();
  assert.strictEqual(d1.deleted, false);
  await api.askDeleteDone();
  await api.confirmPending();
  assert.deepStrictEqual([d1.deleted, d2.deleted, o.deleted, other.deleted], [true, true, false, false]);
  assert.strictEqual(api.ui.lastMsg, "Deleted 2 done notes");
  await api.askDeleteDone();
  assert.strictEqual(api.ui.lastMsg, "No done notes");
  assert.strictEqual(api._pending(), false);
});

test("Delete All still confirms and deletes everything after the refactor", async () => {
  const deck = mkDeck([[tagged("a"), tagged("✓ b", { SGNOTE_DONE: "1" })]]);
  await api.askDeleteAll();
  await api.confirmDeleteAll();
  assert.strictEqual(notesOn(deck, 0).length, 0);
});

test("Skip done notes: on skips done-only slides; all done says No open notes; off visits all", async () => {
  resetPrefs();
  const deck = mkDeck([[], [tagged("✓ x", { SGNOTE_DONE: "1" })], [tagged("open")]]);
  deck.selectedSlide = 0;
  await api.next(); assert.strictEqual(deck.selectedSlide, 2);
  api.savePrefs({ skipDone: false });
  deck.selectedSlide = 0;
  await api.next(); assert.strictEqual(deck.selectedSlide, 1);
  resetPrefs();
  const d2 = mkDeck([[tagged("✓ x", { SGNOTE_DONE: "1" })], []]);
  d2.selectedSlide = 1;
  await api.next();
  assert.strictEqual(d2.selectedSlide, 1);
  assert.strictEqual(api.ui.lastMsg, "No open notes");
});

test("Recolor all keeps done notes' done look, tinted with the new color", async () => {
  resetPrefs({ color: "green" });
  const d = tagged("✓ x", { SGNOTE_DONE: "1" }), o = tagged("y");
  mkDeck([[d, o]]);
  await api.recolorAll();
  assert.deepStrictEqual([d.fillColor, d.lineFormat.color], [api._doneFill("green"), "#A6A6A6"]);
  assert.strictEqual(o.fillColor, "#C8F0B8");
});

test("List model: hidden notes left out; counts; text without ✓ or stamp", async () => {
  const a = tagged("✓ Fix\n— Done FL · Oct 23", { SGNOTE_DONE: "1", SGNOTE_STAMP: "— Done FL · Oct 23" });
  const b = tagged("Line one\nline two"), h = tagged("hidden one"); h.visible = false;
  mkDeck([[b], [a, h]]);
  const s = await PowerPoint.run((ctx) => api.scan(ctx));
  const m = api._listModel(s.notes);
  assert.deepStrictEqual([m.total, m.done, m.open, m.hidden], [2, 1, 1, 1]);
  assert.strictEqual(api._listText(s.notes.find((n) => n.done)), "Fix");
  assert.strictEqual(api._listText(s.notes.find((n) => n.text.startsWith("Line"))), "Line one line two");
});

test("Settings: new prefs default and persist", () => {
  resetPrefs();
  assert.deepStrictEqual([api.prefs().stampOn, api.prefs().who, api.prefs().skipDone], [false, "", true]);
  api.savePrefs({ stampOn: true, who: "FL", skipDone: false });
  assert.deepStrictEqual(api._reloadPrefs(), { position: "tr", color: "yellow", prefix: "", stampOn: true, who: "FL", skipDone: false });
});
