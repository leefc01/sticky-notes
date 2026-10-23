/*
 * StatusGo Sticky Notes for PowerPoint
 * Plain JavaScript, no framework. Runs in a shared runtime: ribbon commands
 * and the task pane share this one page, so commands can update the pane.
 *
 * A note = a rectangle named "SGNote_<id>" with tag SGNOTE=1.
 * The name prefix is a cheap first filter; the tag is the real identifier.
 */
(function (root) {
  "use strict";

  var PREFIX = "SGNote_";
  var TAG = "SGNOTE";
  var NOTE_W = 144; // 2 in
  var NOTE_H = 108; // 1.5 in
  var OFFSET = 18; // 0.25 in cascade for each extra note on a slide
  var TEXT = "#333333";
  var PLACEHOLDER = "Note";
  var PREFIX_MAX = 20;

  var COLORS = {
    yellow: { name: "Yellow", fill: "#FFE97A", border: "#E0B000" },
    pink:   { name: "Pink",   fill: "#FFC2D6", border: "#E07A9E" },
    blue:   { name: "Blue",   fill: "#BFE3FF", border: "#5FA8E0" },
    green:  { name: "Green",  fill: "#C8F0B8", border: "#6FB85A" },
    orange: { name: "Orange", fill: "#FFD3A3", border: "#E0944A" },
    purple: { name: "Purple", fill: "#DCCBFF", border: "#9A7AE0" }
  };
  var POSITIONS = { tr: "Top right", tl: "Top left", tc: "Top center" };

  // ---------- preferences (per user, per device; defaults if storage is unavailable) ----------

  var PREFS_KEY = "sgnotes.prefs.v1";
  var DEFAULTS = { position: "tr", color: "yellow", prefix: "" };

  function loadPrefs() {
    var p = {};
    try { p = JSON.parse(root.localStorage.getItem(PREFS_KEY) || "{}") || {}; } catch (e) { p = {}; }
    return {
      position: POSITIONS[p.position] ? p.position : DEFAULTS.position,
      color: COLORS[p.color] ? p.color : DEFAULTS.color,
      prefix: typeof p.prefix === "string" ? p.prefix.slice(0, PREFIX_MAX) : DEFAULTS.prefix
    };
  }
  var prefs = loadPrefs();

  function savePrefs(changes) {
    Object.keys(changes).forEach(function (k) { prefs[k] = changes[k]; });
    prefs = (function (p) {
      return {
        position: POSITIONS[p.position] ? p.position : DEFAULTS.position,
        color: COLORS[p.color] ? p.color : DEFAULTS.color,
        prefix: String(p.prefix || "").slice(0, PREFIX_MAX)
      };
    })(prefs);
    try { root.localStorage.setItem(PREFS_KEY, JSON.stringify(prefs)); } catch (e) { /* not persisted */ }
    renderSettings();
    return prefs;
  }

  /** Where the Nth note (0-based) on a slide goes, flush to the top edge. */
  function notePosition(position, slideW, n) {
    var d = n * OFFSET;
    if (position === "tl") return { left: d, top: d };
    if (position === "tc") return { left: Math.max(0, (slideW - NOTE_W) / 2), top: d };
    return { left: Math.max(0, slideW - NOTE_W - d), top: d }; // tr (default)
  }

  /** Initial text and which part to select so typing replaces just the placeholder. */
  function initialText(prefix) {
    var p = String(prefix || "").trim();
    if (!p) return { text: PLACEHOLDER, start: 0, length: PLACEHOLDER.length };
    var head = p + " ";
    return { text: head + PLACEHOLDER, start: head.length, length: PLACEHOLDER.length };
  }

  // ---------- UI (no-ops when the pane DOM isn't there, e.g. in tests) ----------

  function el(id) {
    return typeof document !== "undefined" ? document.getElementById(id) : null;
  }
  function setText(id, text) {
    var e = el(id);
    if (e) e.textContent = text;
  }
  function setMsg(text) { setText("msg", text); ui.lastMsg = text; }
  function setStatus(text) { setText("status", text); ui.lastStatus = text; }
  var ui = { lastMsg: "", lastStatus: "" };

  function plural(n, word) { return n + " " + word + (n === 1 ? "" : "s"); }

  function hasApi(version) {
    return !!(root.Office && Office.context && Office.context.requirements &&
      Office.context.requirements.isSetSupported("PowerPointApi", version));
  }

  function showPane() {
    try {
      if (root.Office && Office.addin && Office.addin.showAsTaskpane) {
        return Office.addin.showAsTaskpane();
      }
    } catch (e) { /* pane is optional */ }
    return Promise.resolve();
  }

  // ---------- core ----------

  function run(fn) {
    return PowerPoint.run(fn).catch(function (e) {
      setMsg("Error: " + (e && e.message ? e.message : String(e)));
      if (root.console) console.error(e);
    });
  }

  /** One pass over the deck: slide ids, shape names, and tags only for candidates. 2–3 syncs total. */
  function scan(context) {
    var slides = context.presentation.slides;
    slides.load("items/id");
    var colls;
    return context.sync().then(function () {
      colls = slides.items.map(function (s) {
        var c = s.shapes;
        c.load("items/id,items/name");
        return c;
      });
      return context.sync();
    }).then(function () {
      var candidates = [];
      colls.forEach(function (c, i) {
        c.items.forEach(function (sh) {
          if (sh.name && sh.name.indexOf(PREFIX) === 0) {
            sh.tags.load("items/key,items/value");
            candidates.push({ slideIndex: i, slideId: slides.items[i].id, shape: sh });
          }
        });
      });
      return (candidates.length ? context.sync() : Promise.resolve()).then(function () {
        var notes = candidates.filter(function (n) {
          return n.shape.tags.items.some(function (t) {
            return String(t.key).toUpperCase() === TAG && t.value === "1";
          });
        });
        return { slides: slides.items, notes: notes };
      });
    });
  }

  function currentSlideIndex(context, slides) {
    var sel = context.presentation.getSelectedSlides();
    sel.load("items/id");
    return context.sync().then(function () {
      if (!sel.items.length) return 0;
      for (var i = 0; i < slides.length; i++) {
        if (slides[i].id === sel.items[0].id) return i;
      }
      return 0;
    });
  }

  function statusFor(notes) {
    if (!notes.length) return "No notes in this deck";
    var slideCount = new Set(notes.map(function (n) { return n.slideId; })).size;
    return plural(notes.length, "note") + " on " + plural(slideCount, "slide");
  }

  function slideWidth(context) {
    // pageSetup may not exist in every build; fall back to 16:9 (960 pt).
    try {
      var ps = context.presentation.pageSetup;
      if (ps) {
        ps.load("slideWidth");
        return context.sync().then(function () { return ps.slideWidth || 960; }, function () { return 960; });
      }
    } catch (e) { /* ignore */ }
    return Promise.resolve(960);
  }

  // ---------- commands ----------

  function addNote() {
    setMsg("");
    return run(function (context) {
      var scanned, idx;
      return scan(context).then(function (s) {
        scanned = s;
        if (!s.slides.length) { setMsg("This deck has no slides"); return null; }
        return currentSlideIndex(context, s.slides);
      }).then(function (i) {
        if (i === null) return;
        idx = i;
        return slideWidth(context).then(function (w) {
          var onSlide = scanned.notes.filter(function (n) { return n.slideIndex === idx; }).length;
          var pos = notePosition(prefs.position, w, onSlide);
          var color = COLORS[prefs.color];
          var init = initialText(prefs.prefix);
          var slide = scanned.slides[idx];
          var shape = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle, {
            left: pos.left, top: pos.top, width: NOTE_W, height: NOTE_H
          });
          shape.name = PREFIX + Date.now().toString(36) + Math.floor(Math.random() * 1296).toString(36);
          shape.fill.setSolidColor(color.fill);
          shape.lineFormat.color = color.border;
          shape.lineFormat.weight = 0.75;
          shape.textFrame.verticalAlignment = PowerPoint.TextVerticalAlignment.top;
          shape.textFrame.textRange.text = init.text;
          shape.textFrame.textRange.font.color = TEXT;
          shape.textFrame.textRange.font.size = 12;
          shape.textFrame.textRange.paragraphFormat.horizontalAlignment = PowerPoint.ParagraphHorizontalAlignment.left;
          shape.tags.add(TAG, "1");
          shape.tags.add("SGNOTE_CREATED", new Date().toISOString());
          shape.tags.add("SGNOTE_COLOR", prefs.color);
          if (hasApi("1.10")) {
            shape.altTextTitle = "Sticky note";
            shape.visible = true; // new notes are always visible, even after Hide all
          }
          shape.load("id");
          return context.sync().then(function () {
            scanned.notes.push({ slideIndex: idx, slideId: slide.id, shape: shape });
            setStatus(statusFor(scanned.notes));
            return selectPlaceholder(context, slide, shape, init);
          });
        });
      });
    });
  }

  /**
   * Put the cursor in the new note: select the placeholder word so typing replaces it.
   * Falls back to selecting the whole shape if text selection isn't available.
   * Never fails the Add note command.
   */
  function selectPlaceholder(context, slide, shape, init) {
    function selectShape() {
      try {
        slide.setSelectedShapes([shape.id]);
        return context.sync().then(function () { return "shape"; }, function () { return "none"; });
      } catch (e) { return Promise.resolve("none"); }
    }
    try {
      shape.textFrame.textRange.getSubstring(init.start, init.length).setSelected();
      return context.sync().then(function () { return "text"; }, selectShape);
    } catch (e) {
      return selectShape();
    }
  }

  /** Recolor every note in the deck to the current color preference. */
  function recolorAll() {
    setMsg("");
    var color = COLORS[prefs.color];
    return run(function (context) {
      return scan(context).then(function (s) {
        setStatus(statusFor(s.notes));
        if (!s.notes.length) { setMsg("No notes in this deck"); return 0; }
        s.notes.forEach(function (n) {
          n.shape.fill.setSolidColor(color.fill);
          n.shape.lineFormat.color = color.border;
          n.shape.tags.add("SGNOTE_COLOR", prefs.color);
        });
        return context.sync().then(function () {
          setMsg("Recolored " + plural(s.notes.length, "note") + " " + color.name.toLowerCase());
          return s.notes.length;
        });
      });
    });
  }

  function goTo(dir) {
    setMsg("");
    return run(function (context) {
      var scanned;
      return scan(context).then(function (s) {
        scanned = s;
        if (!s.notes.length) { setMsg("No notes in this deck"); setStatus(statusFor(s.notes)); return null; }
        return currentSlideIndex(context, s.slides);
      }).then(function (cur) {
        if (cur === null) return;
        var withNotes = Array.from(new Set(scanned.notes.map(function (n) { return n.slideIndex; })))
          .sort(function (a, b) { return a - b; });
        var target;
        if (dir > 0) {
          var after = withNotes.filter(function (i) { return i > cur; });
          target = after.length ? after[0] : withNotes[0]; // wrap to first
        } else {
          var before = withNotes.filter(function (i) { return i < cur; });
          target = before.length ? before[before.length - 1] : withNotes[withNotes.length - 1]; // wrap to last
        }
        context.presentation.setSelectedSlides([scanned.slides[target].id]);
        return context.sync().then(function () {
          setStatus(statusFor(scanned.notes));
          return target;
        });
      });
    });
  }

  function deleteNotes() {
    setMsg("");
    return run(function (context) {
      var scanned, sel;
      return scan(context).then(function (s) {
        scanned = s;
        sel = context.presentation.getSelectedShapes();
        sel.load("items/id");
        return context.sync();
      }).then(function () {
        var selected = scanned.notes.filter(function (n) {
          return sel.items.some(function (s) { return s.id === n.shape.id; });
        });
        if (selected.length) return selected;
        // Nothing (or no note) selected: every note on the current slide.
        return currentSlideIndex(context, scanned.slides).then(function (idx) {
          return scanned.notes.filter(function (n) { return n.slideIndex === idx; });
        });
      }).then(function (targets) {
        if (!targets.length) { setMsg("No notes on this slide"); return 0; }
        targets.forEach(function (n) { n.shape.delete(); });
        return context.sync().then(function () {
          var gone = new Set(targets.map(function (n) { return n.shape.id; }));
          var left = scanned.notes.filter(function (n) { return !gone.has(n.shape.id); });
          setMsg("Deleted " + plural(targets.length, "note"));
          setStatus(statusFor(left));
          return targets.length;
        });
      });
    });
  }

  var pendingDeleteAll = false;

  function showConfirm(show, text) {
    pendingDeleteAll = show;
    var box = el("confirm");
    if (box) box.hidden = !show;
    if (text) setText("confirm-text", text);
  }

  function askDeleteAll() {
    setMsg("");
    return run(function (context) {
      return scan(context).then(function (s) {
        setStatus(statusFor(s.notes));
        if (!s.notes.length) { showConfirm(false); setMsg("No notes in this deck"); return showPane(); }
        var slideCount = new Set(s.notes.map(function (n) { return n.slideId; })).size;
        showConfirm(true, "Delete " + plural(s.notes.length, "note") + " on " + plural(slideCount, "slide") + "?");
        return showPane();
      });
    });
  }

  function confirmDeleteAll() {
    if (!pendingDeleteAll) return Promise.resolve(0);
    showConfirm(false);
    return run(function (context) {
      return scan(context).then(function (s) { // rescan: the deck may have changed meanwhile
        s.notes.forEach(function (n) { n.shape.delete(); });
        return context.sync().then(function () {
          setMsg("Deleted " + plural(s.notes.length, "note"));
          setStatus(statusFor([]));
          return s.notes.length;
        });
      });
    });
  }

  function cancelDeleteAll() {
    showConfirm(false);
    setMsg("Nothing deleted");
  }

  function setAllVisible(visible) {
    setMsg("");
    if (!hasApi("1.10")) {
      setMsg("Update Office to use Hide all / Show all");
      return showPane();
    }
    return run(function (context) {
      return scan(context).then(function (s) {
        setStatus(statusFor(s.notes));
        if (!s.notes.length) { setMsg("No notes in this deck"); return 0; }
        s.notes.forEach(function (n) { n.shape.visible = visible; });
        return context.sync().then(function () {
          setMsg(visible ? "Notes shown" : "Notes hidden");
          return s.notes.length;
        });
      });
    });
  }

  function refresh() {
    return run(function (context) {
      return scan(context).then(function (s) { setStatus(statusFor(s.notes)); });
    });
  }

  // ---------- wiring ----------

  /** Wrap a command for the ribbon: always signal completion to Office. */
  function ribbon(fn) {
    return function (event) {
      Promise.resolve().then(fn).then(function () { event.completed(); }, function () { event.completed(); });
    };
  }

  // ---------- settings UI ----------

  function renderSettings() {
    if (typeof document === "undefined") return;
    var posBtns = document.querySelectorAll("[data-position]");
    Array.prototype.forEach.call(posBtns, function (b) {
      b.setAttribute("aria-pressed", String(b.getAttribute("data-position") === prefs.position));
    });
    var swatches = document.querySelectorAll("[data-color]");
    Array.prototype.forEach.call(swatches, function (b) {
      b.setAttribute("aria-pressed", String(b.getAttribute("data-color") === prefs.color));
    });
    var input = el("prefix");
    if (input && document.activeElement !== input) input.value = prefs.prefix;
    setText("recolor-label", "Apply " + COLORS[prefs.color].name.toLowerCase() + " to all notes");
  }

  function bindSettings() {
    if (typeof document === "undefined") return;
    Array.prototype.forEach.call(document.querySelectorAll("[data-position]"), function (b) {
      b.addEventListener("click", function () { savePrefs({ position: b.getAttribute("data-position") }); });
    });
    Array.prototype.forEach.call(document.querySelectorAll("[data-color]"), function (b) {
      var c = COLORS[b.getAttribute("data-color")];
      if (c) { b.style.background = c.fill; b.style.borderColor = c.border; b.title = c.name; }
      b.addEventListener("click", function () { savePrefs({ color: b.getAttribute("data-color") }); });
    });
    var input = el("prefix");
    if (input) {
      input.maxLength = PREFIX_MAX;
      input.addEventListener("input", function () { savePrefs({ prefix: input.value }); });
    }
    bindButton("btn-recolor", recolorAll);
    renderSettings();
  }

  var api = {
    addNote: addNote,
    recolorAll: recolorAll,
    prefs: function () { return prefs; },
    savePrefs: savePrefs,
    _reloadPrefs: function () { prefs = loadPrefs(); return prefs; },
    _notePosition: notePosition,
    _initialText: initialText,
    next: function () { return goTo(1); },
    previous: function () { return goTo(-1); },
    deleteNotes: deleteNotes,
    askDeleteAll: askDeleteAll,
    confirmDeleteAll: confirmDeleteAll,
    cancelDeleteAll: cancelDeleteAll,
    hideAll: function () { return setAllVisible(false); },
    showAll: function () { return setAllVisible(true); },
    refresh: refresh,
    scan: scan,
    ui: ui,
    _pending: function () { return pendingDeleteAll; }
  };

  function bindButton(id, fn) {
    var b = el(id);
    if (b) b.addEventListener("click", function () { fn(); });
  }

  if (root.Office && Office.onReady) {
    Office.onReady(function () {
      Office.actions.associate("addNote", ribbon(api.addNote));
      Office.actions.associate("nextNote", ribbon(api.next));
      Office.actions.associate("previousNote", ribbon(api.previous));
      Office.actions.associate("deleteNote", ribbon(api.deleteNotes));
      Office.actions.associate("deleteAllNotes", ribbon(api.askDeleteAll));
      Office.actions.associate("hideAllNotes", ribbon(api.hideAll));
      Office.actions.associate("showAllNotes", ribbon(api.showAll));

      bindButton("btn-add", api.addNote);
      bindButton("btn-prev", api.previous);
      bindButton("btn-next", api.next);
      bindButton("btn-del", api.deleteNotes);
      bindButton("btn-delall", api.askDeleteAll);
      bindButton("btn-hide", api.hideAll);
      bindButton("btn-show", api.showAll);
      bindButton("btn-confirm-yes", api.confirmDeleteAll);
      bindButton("btn-confirm-no", api.cancelDeleteAll);
      bindButton("btn-refresh", api.refresh);
      bindSettings();

      if (!hasApi("1.10")) {
        setMsg("Update Office to use Sticky Notes. Hide all / Show all need a newer PowerPoint.");
      }
      api.refresh();
    });
  }

  root.SGNotes = api;
  if (typeof module !== "undefined" && module.exports) module.exports = api;
})(typeof window !== "undefined" ? window : globalThis);
