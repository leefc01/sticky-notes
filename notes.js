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
  var MARGIN = 16;
  var OFFSET = 18; // 0.25 in cascade for each extra note on a slide
  var FILL = "#FFE97A";
  var BORDER = "#E0B000";
  var TEXT = "#333333";

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
          var shape = scanned.slides[idx].shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle, {
            left: Math.max(0, w - MARGIN - NOTE_W - onSlide * OFFSET),
            top: MARGIN + onSlide * OFFSET,
            width: NOTE_W,
            height: NOTE_H
          });
          shape.name = PREFIX + Date.now().toString(36) + Math.floor(Math.random() * 1296).toString(36);
          shape.fill.setSolidColor(FILL);
          shape.lineFormat.color = BORDER;
          shape.lineFormat.weight = 0.75;
          shape.textFrame.verticalAlignment = PowerPoint.TextVerticalAlignment.top;
          shape.textFrame.textRange.text = "Note";
          shape.textFrame.textRange.font.color = TEXT;
          shape.textFrame.textRange.font.size = 12;
          shape.textFrame.textRange.paragraphFormat.horizontalAlignment = PowerPoint.ParagraphHorizontalAlignment.left;
          shape.tags.add(TAG, "1");
          shape.tags.add("SGNOTE_CREATED", new Date().toISOString());
          if (hasApi("1.10")) {
            shape.altTextTitle = "Sticky note";
            shape.visible = true; // new notes are always visible, even after Hide all
          }
          return context.sync().then(function () {
            scanned.notes.push({ slideIndex: idx, slideId: scanned.slides[idx].id, shape: shape });
            setStatus(statusFor(scanned.notes));
          });
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

  var api = {
    addNote: addNote,
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

      if (!hasApi("1.10")) {
        setMsg("Update Office to use Sticky Notes. Hide all / Show all need a newer PowerPoint.");
      }
      api.refresh();
    });
  }

  root.SGNotes = api;
  if (typeof module !== "undefined" && module.exports) module.exports = api;
})(typeof window !== "undefined" ? window : globalThis);
