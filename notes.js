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

  var APP_VERSION = "1.3.1"; // product version major.minor.patch; manifest = APP_VERSION + ".0" (tools/build.py)
  var PREFIX = "SGNote_";
  var TAG = "SGNOTE";
  var NOTE_W = 144; // 2 in
  var NOTE_H = 108; // 1.5 in
  var OFFSET = 18; // 0.25 in cascade for each extra note on a slide
  var TEXT = "#333333";
  var PLACEHOLDER = "Note";
  var PREFIX_MAX = 20;

  // Done look (v1.3): faint tint of the note's own color, gray border and text, "✓ " prefix.
  var CHECK = "\u2713 ";
  var DONE_BASE = "#E6E6E6";
  var DONE_TINT = 0.3; // share of the note's own color in the done fill
  var DONE_BORDER = "#A6A6A6";
  var DONE_TEXT = "#6B6B6B";
  var MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

  var COLORS = {
    yellow: { name: "Yellow", fill: "#FFE97A", border: "#E0B000" },
    pink:   { name: "Pink",   fill: "#FFC2D6", border: "#E07A9E" },
    blue:   { name: "Blue",   fill: "#BFE3FF", border: "#5FA8E0" },
    green:  { name: "Green",  fill: "#C8F0B8", border: "#6FB85A" },
    orange: { name: "Orange", fill: "#FFD3A3", border: "#E0944A" },
    purple: { name: "Purple", fill: "#DCCBFF", border: "#9A7AE0" }
  };
  var POSITIONS = { tl: "Top Left", tc: "Top Center", tr: "Top Right" };

  // ---------- preferences (per user, per device; defaults if storage is unavailable) ----------

  var PREFS_KEY = "sgnotes.prefs.v1";
  var DEFAULTS = { position: "tr", color: "yellow", prefix: "", stampOn: false, who: "", skipDone: true };

  function cleanPrefs(p) {
    return {
      position: POSITIONS[p.position] ? p.position : DEFAULTS.position,
      color: COLORS[p.color] ? p.color : DEFAULTS.color,
      prefix: typeof p.prefix === "string" ? p.prefix.slice(0, PREFIX_MAX) : DEFAULTS.prefix,
      stampOn: typeof p.stampOn === "boolean" ? p.stampOn : DEFAULTS.stampOn,
      who: typeof p.who === "string" ? p.who.slice(0, PREFIX_MAX) : DEFAULTS.who,
      skipDone: typeof p.skipDone === "boolean" ? p.skipDone : DEFAULTS.skipDone
    };
  }

  function loadPrefs() {
    var p = {};
    try { p = JSON.parse(root.localStorage.getItem(PREFS_KEY) || "{}") || {}; } catch (e) { p = {}; }
    return cleanPrefs(p);
  }
  var prefs = loadPrefs();

  function savePrefs(changes) {
    var next = {};
    Object.keys(prefs).forEach(function (k) { next[k] = prefs[k]; });
    Object.keys(changes).forEach(function (k) { next[k] = changes[k]; });
    prefs = cleanPrefs(next);
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

  /**
   * First free slot (0, 1, 2, …) at a position: a slot is taken only if a note
   * actually sits there (within 2 pt). Notes elsewhere on the slide, or ones the
   * user dragged away, don't push a new note down.
   */
  function freeSlot(position, slideW, existing) {
    for (var n = 0; n < 40; n++) {
      var p = notePosition(position, slideW, n);
      var taken = existing.some(function (e) {
        return Math.abs(e.left - p.left) <= 2 && Math.abs(e.top - p.top) <= 2;
      });
      if (!taken) return n;
    }
    return 0;
  }

  /** Initial text and which part to select so typing replaces just the placeholder. */
  function initialText(prefix) {
    var p = String(prefix || "").trim();
    if (!p) return { text: PLACEHOLDER, start: 0, length: PLACEHOLDER.length };
    var head = p + " ";
    return { text: head + PLACEHOLDER, start: head.length, length: PLACEHOLDER.length };
  }

  /** Blend two #RRGGBB colors: share = weight of the first. */
  function blend(hexA, hexB, share) {
    function ch(h, i) { return parseInt(h.slice(1 + i * 2, 3 + i * 2), 16); }
    var out = "#";
    for (var i = 0; i < 3; i++) {
      var v = Math.round(ch(hexA, i) * share + ch(hexB, i) * (1 - share));
      out += ("0" + v.toString(16)).slice(-2).toUpperCase();
    }
    return out;
  }
  function doneFill(colorKey) { return blend((COLORS[colorKey] || COLORS.yellow).fill, DONE_BASE, DONE_TINT); }

  /** "— Done FL · Oct 23": Who, else the prefix without a trailing colon, else date only. */
  function stampText(p, date) {
    var d = date || new Date();
    var when = MONTHS[d.getMonth()] + " " + d.getDate();
    var who = String(p.who || "").trim() || String(p.prefix || "").trim().replace(/:+$/, "").trim();
    return "\u2014 Done " + (who ? who + " \u00b7 " : "") + when;
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
            sh.load(hasApi("1.10") ? "left,top,visible" : "left,top"); // placement + hidden state
            sh.textFrame.textRange.load("text"); // list text and done/stamp checks
            candidates.push({ slideIndex: i, slideId: slides.items[i].id, shape: sh });
          }
        });
      });
      return (candidates.length ? context.sync() : Promise.resolve()).then(function () {
        var notes = candidates.filter(function (n) {
          var tags = {};
          n.shape.tags.items.forEach(function (t) { tags[String(t.key).toUpperCase()] = t.value; });
          if (tags[TAG] !== "1") return false;
          n.done = tags.SGNOTE_DONE === "1";
          n.colorKey = COLORS[tags.SGNOTE_COLOR] ? tags.SGNOTE_COLOR : "yellow"; // v1.0 notes have no color tag
          n.stamp = tags.SGNOTE_STAMP || "";
          n.text = (n.shape.textFrame && n.shape.textFrame.textRange && n.shape.textFrame.textRange.text) || "";
          n.hidden = n.shape.visible === false;
          return true;
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
          var onSlide = scanned.notes
            .filter(function (n) { return n.slideIndex === idx; })
            .map(function (n) { return { left: n.shape.left, top: n.shape.top }; });
          var pos = notePosition(prefs.position, w, freeSlot(prefs.position, w, onSlide));
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
          // Done notes keep their done look, tinted with the new color.
          n.shape.fill.setSolidColor(n.done ? doneFill(prefs.color) : color.fill);
          n.shape.lineFormat.color = n.done ? DONE_BORDER : color.border;
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
        var pool = prefs.skipDone ? s.notes.filter(function (n) { return !n.done; }) : s.notes;
        if (!pool.length) { setMsg("No open notes"); return null; }
        scanned.pool = pool;
        return currentSlideIndex(context, s.slides);
      }).then(function (cur) {
        if (cur === null) return;
        var withNotes = Array.from(new Set(scanned.pool.map(function (n) { return n.slideIndex; })))
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

  var pendingAction = null; // "all" | "done" while a confirmation is showing

  function showConfirm(action, text) {
    pendingAction = action || null;
    var show = !!action;
    var box = el("confirm");
    if (box) box.hidden = !show;
    if (text) setText("confirm-text", text);
  }

  function askDeleteAll() {
    setMsg("");
    return run(function (context) {
      return scan(context).then(function (s) {
        setStatus(statusFor(s.notes));
        if (!s.notes.length) { showConfirm(null); setMsg("No notes in this deck"); return showPane(); }
        var slideCount = new Set(s.notes.map(function (n) { return n.slideId; })).size;
        showConfirm("all", "Delete " + plural(s.notes.length, "note") + " on " + plural(slideCount, "slide") + "?");
        return showPane();
      });
    });
  }

  function askDeleteDone() {
    setMsg("");
    return run(function (context) {
      return scan(context).then(function (s) {
        setStatus(statusFor(s.notes));
        var done = s.notes.filter(function (n) { return n.done; });
        if (!done.length) { showConfirm(null); setMsg("No done notes"); return showPane(); }
        var slideCount = new Set(done.map(function (n) { return n.slideId; })).size;
        showConfirm("done", "Delete " + plural(done.length, "done note") + " on " + plural(slideCount, "slide") + "?");
        return showPane();
      });
    });
  }

  /** Runs whichever delete is waiting for confirmation (Delete All or Delete Done). */
  function confirmPending() {
    var action = pendingAction;
    if (!action) return Promise.resolve(0);
    showConfirm(null);
    return run(function (context) {
      return scan(context).then(function (s) { // rescan: the deck may have changed meanwhile
        var targets = action === "done" ? s.notes.filter(function (n) { return n.done; }) : s.notes;
        targets.forEach(function (n) { n.shape.delete(); });
        return context.sync().then(function () {
          setMsg("Deleted " + plural(targets.length, action === "done" ? "done note" : "note"));
          setStatus(statusFor(s.notes.filter(function (n) { return targets.indexOf(n) < 0; })));
          afterChange();
          return targets.length;
        });
      });
    });
  }

  function cancelPending() {
    showConfirm(null);
    setMsg("Nothing deleted");
  }

  // ---------- done ----------

  var LINE_BREAKS = ["\r\n", "\r", "\n", "\u000b"];

  /** Queue the edits that mark one note done. */
  function applyDone(n, stamp) {
    var tr = n.shape.textFrame.textRange;
    var text = n.text;
    if (!text) {
      tr.text = CHECK + (stamp ? "\n" + stamp : "");
    } else {
      if (stamp) tr.getSubstring(text.length - 1, 1).text = text.slice(-1) + "\n" + stamp; // end first
      if (text.indexOf(CHECK) !== 0) tr.getSubstring(0, 1).text = CHECK + text.charAt(0);
    }
    tr.font.color = DONE_TEXT;
    n.shape.fill.setSolidColor(doneFill(n.colorKey));
    n.shape.lineFormat.color = DONE_BORDER;
    n.shape.tags.add("SGNOTE_DONE", "1");
    if (stamp) n.shape.tags.add("SGNOTE_STAMP", stamp);
    else if (n.stamp) n.shape.tags.delete("SGNOTE_STAMP");
  }

  /** Queue the edits that reopen one note: undo exactly what applyDone added, if still there. */
  function applyReopen(n) {
    var tr = n.shape.textFrame.textRange;
    var text = n.text;
    if (n.stamp) {
      for (var i = 0; i < LINE_BREAKS.length; i++) {
        var tail = LINE_BREAKS[i] + n.stamp;
        if (text.length >= tail.length && text.slice(-tail.length) === tail) {
          tr.getSubstring(text.length - tail.length, tail.length).text = "";
          text = text.slice(0, -tail.length);
          break;
        }
      }
      n.shape.tags.delete("SGNOTE_STAMP");
    }
    if (text.indexOf(CHECK) === 0) tr.getSubstring(0, CHECK.length).text = "";
    tr.font.color = TEXT;
    var c = COLORS[n.colorKey] || COLORS.yellow;
    n.shape.fill.setSolidColor(c.fill);
    n.shape.lineFormat.color = c.border;
    n.shape.tags.add("SGNOTE_DONE", "0");
  }

  /** Smart toggle: if any target is open, mark them all done; if all are done, reopen them all. */
  function toggleTargets(context, targets) {
    if (!targets.length) return Promise.resolve(null);
    var open = targets.filter(function (n) { return !n.done; });
    var stamp = prefs.stampOn ? stampText(prefs) : "";
    if (open.length) open.forEach(function (n) { applyDone(n, stamp); });
    else targets.forEach(applyReopen);
    return context.sync().then(function () {
      setMsg(open.length ? "Marked " + plural(open.length, "note") + " done" : "Reopened " + plural(targets.length, "note"));
      afterChange();
      return open.length ? "done" : "reopened";
    });
  }

  function toggleDone() {
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
          return sel.items.some(function (x) { return x.id === n.shape.id; });
        });
        if (selected.length) return selected;
        return currentSlideIndex(context, scanned.slides).then(function (idx) {
          return scanned.notes.filter(function (n) { return n.slideIndex === idx; });
        });
      }).then(function (targets) {
        if (!targets.length) { setMsg("No notes on this slide"); return null; }
        return toggleTargets(context, targets);
      });
    });
  }

  /** Toggle one note by shape id (the list checkboxes). */
  function toggleNote(shapeId) {
    setMsg("");
    return run(function (context) {
      return scan(context).then(function (s) {
        var n = s.notes.filter(function (x) { return x.shape.id === shapeId; });
        if (!n.length) { setMsg("That note no longer exists"); afterChange(); return null; }
        return toggleTargets(context, n);
      });
    });
  }

  /** Jump to a note's slide and select it (the list text). */
  function goToNote(shapeId) {
    return run(function (context) {
      return scan(context).then(function (s) {
        var n = s.notes.filter(function (x) { return x.shape.id === shapeId; })[0];
        if (!n) { setMsg("That note no longer exists"); afterChange(); return; }
        context.presentation.setSelectedSlides([n.slideId]);
        return context.sync().then(function () {
          try {
            s.slides[n.slideIndex].setSelectedShapes([shapeId]);
            return context.sync().catch(function () {});
          } catch (e) { /* selecting the shape is optional */ }
        });
      });
    });
  }

  // ---------- notes list + progress ----------

  /** Note text for the list: no ✓, no stamp line, one line. */
  function listText(n) {
    var text = n.text || "";
    if (n.done && n.stamp) {
      for (var i = 0; i < LINE_BREAKS.length; i++) {
        var tail = LINE_BREAKS[i] + n.stamp;
        if (text.slice(-tail.length) === tail) { text = text.slice(0, -tail.length); break; }
      }
    }
    if (text.indexOf(CHECK) === 0) text = text.slice(CHECK.length);
    text = text.replace(/[\r\n\u000b]+/g, " ").trim();
    return text || "(empty note)";
  }

  /** Counts and grouping the pane shows; hidden notes are left out of the list and progress. */
  function listModel(notes) {
    var shown = notes.filter(function (n) { return !n.hidden; });
    var done = shown.filter(function (n) { return n.done; });
    var open = shown.filter(function (n) { return !n.done; });
    function bySlide(a, b) { return a.slideIndex - b.slideIndex; }
    return {
      total: shown.length, done: done.length, open: open.length,
      hidden: notes.length - shown.length,
      openNotes: open.sort(bySlide), doneNotes: done.sort(bySlide)
    };
  }

  var lastSignature = "";

  function renderList(notes) {
    if (typeof document === "undefined") return;
    var m = listModel(notes);
    var sig = JSON.stringify([m.hidden, m.openNotes.concat(m.doneNotes).map(function (n) {
      return [n.shape.id, n.slideIndex, n.done, listText(n)];
    })]);
    if (sig === lastSignature) return;
    lastSignature = sig;

    setText("progress-label", m.total ? m.done + " of " + m.total + " done" : "No notes yet");
    setText("hidden-count", m.hidden ? "\u00b7 " + m.hidden + " hidden" : "");
    var fill = el("progress-fill");
    if (fill) fill.style.width = (m.total ? Math.round(100 * m.done / m.total) : 0) + "%";
    setText("list-summary", "Notes (" + m.open + " open \u00b7 " + m.done + " done)");
    setText("done-summary", "Done (" + m.done + ")");
    var doneBox = el("done-list");
    if (doneBox) doneBox.hidden = !m.done;
    fillItems(el("open-items"), m.openNotes, "No open notes");
    fillItems(el("done-items"), m.doneNotes, "");
  }

  function fillItems(container, notes, emptyText) {
    if (!container) return;
    while (container.firstChild) container.removeChild(container.firstChild);
    if (!notes.length && emptyText) {
      var e = document.createElement("p");
      e.className = "empty";
      e.textContent = emptyText;
      container.appendChild(e);
    }
    var lastSlide = -1;
    notes.forEach(function (n) {
      if (n.slideIndex !== lastSlide) {
        lastSlide = n.slideIndex;
        var h = document.createElement("p");
        h.className = "slide-head";
        h.textContent = "Slide " + (n.slideIndex + 1);
        container.appendChild(h);
      }
      var row = document.createElement("div");
      row.className = "row" + (n.done ? " done" : "");
      var cb = document.createElement("input");
      cb.type = "checkbox";
      cb.checked = n.done;
      cb.setAttribute("aria-label", (n.done ? "Reopen: " : "Mark done: ") + listText(n));
      cb.addEventListener("change", function () { toggleNote(n.shape.id); });
      var b = document.createElement("button");
      b.className = "note-text";
      b.textContent = listText(n);
      b.title = listText(n);
      b.addEventListener("click", function () { goToNote(n.shape.id); });
      row.appendChild(cb);
      row.appendChild(b);
      container.appendChild(row);
    });
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

  var refreshing = false;
  function refresh() {
    if (refreshing) return Promise.resolve();
    refreshing = true;
    return run(function (context) {
      return scan(context).then(function (s) { setStatus(statusFor(s.notes)); renderList(s.notes); });
    }).then(function () { refreshing = false; }, function () { refreshing = false; });
  }

  /** After any change, refresh the pane soon (no-op outside the pane, e.g. in tests). */
  var refreshTimer = null;
  function afterChange() {
    if (typeof document === "undefined") return;
    clearTimeout(refreshTimer);
    refreshTimer = setTimeout(refresh, 50);
  }

  /** Background refresh every 5 s while the pane is visible and nobody is typing in it. */
  function startPolling() {
    if (typeof document === "undefined") return;
    setInterval(function () {
      if (document.visibilityState && document.visibilityState !== "visible") return;
      var a = document.activeElement;
      if (a && (a.tagName === "INPUT" || a.tagName === "TEXTAREA")) return;
      refresh();
    }, 5000);
  }

  // ---------- wiring ----------

  /** Wrap a command for the ribbon: always signal completion to Office. */
  function ribbon(fn) {
    return function (event) {
      Promise.resolve().then(fn).then(function () { event.completed(); afterChange(); }, function () { event.completed(); });
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
    var who = el("who");
    if (who && document.activeElement !== who) who.value = prefs.who;
    if (who) who.disabled = !prefs.stampOn;
    var stampOn = el("stamp-on");
    if (stampOn) stampOn.checked = prefs.stampOn;
    var skip = el("skip-done");
    if (skip) skip.checked = prefs.skipDone;
    setText("stamp-preview", prefs.stampOn ? "Adds \u201c" + stampText(prefs) + "\u201d as the last line" : "Off");
    setText("recolor-label", "Apply " + COLORS[prefs.color].name + " to All Notes");
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
    var who = el("who");
    if (who) {
      who.maxLength = PREFIX_MAX;
      who.addEventListener("input", function () { savePrefs({ who: who.value }); });
    }
    var stampOn = el("stamp-on");
    if (stampOn) stampOn.addEventListener("change", function () { savePrefs({ stampOn: stampOn.checked }); });
    var skip = el("skip-done");
    if (skip) skip.addEventListener("change", function () { savePrefs({ skipDone: skip.checked }); });
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
    _freeSlot: freeSlot,
    version: APP_VERSION,
    _initialText: initialText,
    next: function () { return goTo(1); },
    previous: function () { return goTo(-1); },
    deleteNotes: deleteNotes,
    askDeleteAll: askDeleteAll,
    askDeleteDone: askDeleteDone,
    confirmDeleteAll: confirmPending, // confirms whichever delete is pending
    confirmPending: confirmPending,
    cancelDeleteAll: cancelPending,
    toggleDone: toggleDone,
    toggleNote: toggleNote,
    goToNote: goToNote,
    _stampText: stampText,
    _doneFill: doneFill,
    _listModel: listModel,
    _listText: listText,
    hideAll: function () { return setAllVisible(false); },
    showAll: function () { return setAllVisible(true); },
    refresh: refresh,
    scan: scan,
    ui: ui,
    _pending: function () { return !!pendingAction; }
  };

  function bindButton(id, fn) {
    var b = el(id);
    if (b) b.addEventListener("click", function () { Promise.resolve(fn()).then(afterChange); });
  }

  if (root.Office && Office.onReady) {
    Office.onReady(function () {
      Office.actions.associate("addNote", ribbon(api.addNote));
      Office.actions.associate("nextNote", ribbon(api.next));
      Office.actions.associate("previousNote", ribbon(api.previous));
      Office.actions.associate("deleteNote", ribbon(api.deleteNotes));
      Office.actions.associate("deleteAllNotes", ribbon(api.askDeleteAll));
      Office.actions.associate("toggleDone", ribbon(api.toggleDone));
      Office.actions.associate("deleteDoneNotes", ribbon(api.askDeleteDone));
      Office.actions.associate("hideAllNotes", ribbon(api.hideAll));
      Office.actions.associate("showAllNotes", ribbon(api.showAll));

      bindButton("btn-add", api.addNote);
      bindButton("btn-prev", api.previous);
      bindButton("btn-next", api.next);
      bindButton("btn-del", api.deleteNotes);
      bindButton("btn-delall", api.askDeleteAll);
      bindButton("btn-toggle-done", api.toggleDone);
      bindButton("btn-deldone", api.askDeleteDone);
      bindButton("btn-hide", api.hideAll);
      bindButton("btn-show", api.showAll);
      bindButton("btn-confirm-yes", api.confirmDeleteAll);
      bindButton("btn-confirm-no", api.cancelDeleteAll);
      bindButton("btn-refresh", api.refresh);
      bindSettings();
      setText("version", "Sticky Notes v" + APP_VERSION.replace(/\.0$/, "")); // 1.2.0 -> v1.2, 1.2.1 -> v1.2.1

      if (!hasApi("1.10")) {
        setMsg("Update Office to use Sticky Notes. Hide all / Show all need a newer PowerPoint.");
      }
      api.refresh();
      startPolling();
    });
  }

  root.SGNotes = api;
  if (typeof module !== "undefined" && module.exports) module.exports = api;
})(typeof window !== "undefined" ? window : globalThis);
