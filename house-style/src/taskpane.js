/*
 * Task pane controller. Talks to Outlook through a small "host" adapter so the
 * same UI also runs in a plain browser against a sample email (preview mode).
 */
(function () {
  "use strict";

  var STYLE = window.HOUSE_STYLE;
  var fmt = window.createFormatter(STYLE, window);
  var tools = window.createAttachmentTools(STYLE);
  var storage = null;
  try { storage = window.localStorage; } catch (e) { /* not available */ }
  // Formatting actions shared with the ribbon buttons (actions.js).
  var actions = window.createActions(STYLE, fmt, { win: window, storage: storage });
  var parser = new DOMParser();
  var host = null;
  var busy = false;
  var undoSnapshot = null;
  var statusTimer = null;

  function $(sel, root) { return (root || document).querySelector(sel); }
  function $all(sel, root) { return Array.prototype.slice.call((root || document).querySelectorAll(sel)); }
  function el(tag, cls, text) {
    var e = document.createElement(tag);
    if (cls) e.className = cls;
    if (text != null) e.textContent = text;
    return e;
  }

  function UserError(message) { this.userMessage = message; }


  // ── Outlook host (shared with the ribbon buttons: outlook-host.js) ──────
  function outlookHost() { return window.createOutlookHost(Office); }

  // ── Browser preview host ──────────────────────────────────────────────────
  function demoHost() {
    var editor = $("#demo-body");
    var subject = $("#demo-subject");
    var sample = window.DEMO_SAMPLE;
    var saved = null;
    var lastField = "body";
    var files = [];
    var seq = 0;
    var listeners = [];
    var BLOCK_HTML = /^(P|DIV|OL|UL|TABLE|H[1-6]|BLOCKQUOTE|HR)$/;

    function later(value) { return new Promise(function (r) { setTimeout(function () { r(value); }, 90); }); }
    function fileChanged() { renderChips(); listeners.forEach(function (fn) { fn(); }); }

    function renderChips() {
      var ul = $("#demo-attachments");
      ul.textContent = "";
      files.forEach(function (f) {
        var li = el("li", "file-chip");
        var parts = tools.splitName(f.name);
        var ext = f.type === "item" ? "MSG" : (parts.ext.replace(".", "").toUpperCase() || "FILE");
        var ft = el("span", "ft " + (/PDF/.test(ext) ? "pdf" : /DOC/.test(ext) ? "doc" : /MSG|EML/.test(ext) ? "msg" : ""), ext);
        li.appendChild(ft);
        li.appendChild(el("span", "nm", f.name));
        li.appendChild(el("span", "sz", tools.formatSize(f.size)));
        li.title = f.name;
        ul.appendChild(li);
      });
      $("#demo-attachments-row").hidden = !files.length;
    }

    function reset() {
      editor.innerHTML = sample.body;
      subject.value = sample.subject;
      saved = null;
      files = sample.attachments.map(function (a) {
        return { id: "att" + (++seq), name: a.name, size: a.size, type: a.type, content: btoa("demo:" + a.name) };
      });
      renderChips();
    }
    reset();
    $("#demo-reset").addEventListener("click", function () {
      reset();
      undoSnapshot = null;
      refreshMarkings();
      loadAttachments(true);
      showStatus("Sample email restored.", "info");
    });

    document.addEventListener("selectionchange", function () {
      var sel = window.getSelection();
      if (sel.rangeCount && editor.contains(sel.anchorNode)) {
        saved = sel.getRangeAt(0).cloneRange();
        lastField = "body";
      }
    });
    subject.addEventListener("focus", function () { lastField = "subject"; });
    editor.addEventListener("focus", function () { lastField = "body"; });

    function currentRange() {
      if (saved && editor.contains(saved.commonAncestorContainer)) return saved;
      var r = document.createRange();
      r.setStart(editor, 0);
      r.collapse(true);
      return r;
    }
    function lineOf(node) {
      while (node && node.parentNode !== editor) node = node.parentNode;
      return node;
    }
    function blank(frag) {
      return !(frag.querySelector && frag.querySelector("img")) && !String(frag.textContent || "").replace(/[\s ​]/g, "");
    }
    function isBlockHtml(html) {
      var body = parser.parseFromString(html, "text/html").body;
      return Array.prototype.some.call(body.children, function (c) { return BLOCK_HTML.test(c.tagName); });
    }
    // Like Outlook: the cursor ends up at the end of what was inserted, inside
    // its last paragraph, and the message keeps the keyboard focus.
    function placeCaretAtEnd(node) {
      if (!node || !node.parentNode) return;
      var target = node;
      while (target.lastChild) target = target.lastChild;
      var r = document.createRange();
      if (target.nodeType === 3) r.setStart(target, target.nodeValue.length);
      else r.setStartAfter(target);
      r.collapse(true);
      editor.focus();
      var sel = window.getSelection();
      sel.removeAllRanges();
      sel.addRange(r);
      saved = r.cloneRange();
    }

    function replace(html) {
      var r = currentRange();
      var startLine = lineOf(r.startContainer === editor ? editor.childNodes[r.startOffset] || editor.lastChild : r.startContainer);
      var endLine = lineOf(r.endContainer === editor ? editor.childNodes[Math.max(0, r.endOffset - 1)] || editor.lastChild : r.endContainer);

      if (!isBlockHtml(html) || !startLine || !endLine) {
        r.deleteContents();
        var frag = r.createContextualFragment(html);
        var last = frag.lastChild;
        r.insertNode(frag);
        placeCaretAtEnd(last);
        return;
      }
      // Block content: split the surrounding lines like a mail editor would.
      var before = document.createRange();
      before.setStartBefore(startLine);
      before.setEnd(r.startContainer, r.startOffset);
      var after = document.createRange();
      after.setStart(r.endContainer, r.endOffset);
      after.setEndAfter(endLine);
      var beforeFrag = before.cloneContents();
      var afterFrag = after.cloneContents();
      var anchor = endLine.nextSibling;
      var n = startLine;
      while (n) {
        var next = n.nextSibling;
        editor.removeChild(n);
        if (n === endLine) break;
        n = next;
      }
      var ctx = document.createRange();
      ctx.selectNodeContents(editor);
      var inserted = ctx.createContextualFragment(html);
      var lastInserted = inserted.lastChild;
      var out = document.createDocumentFragment();
      if (!blank(beforeFrag)) out.appendChild(beforeFrag);
      out.appendChild(inserted);
      if (!blank(afterFrag)) out.appendChild(afterFrag);
      editor.insertBefore(out, anchor);
      placeCaretAtEnd(lastInserted);
    }

    return {
      name: "demo",
      bodyType: function () { return Promise.resolve("html"); },
      getSelection: function () {
        if (lastField === "subject") return Promise.resolve({ html: "", source: "subject" });
        var div = document.createElement("div");
        div.appendChild(currentRange().cloneContents());
        return Promise.resolve({ html: div.innerHTML, source: "body" });
      },
      // Outlook can't put formatted text in the subject line, so it refuses.
      replaceSelection: function (html) {
        if (lastField === "subject") return Promise.reject(new Error("InvalidDataFormat"));
        replace(html);
        return Promise.resolve();
      },
      getBody: function () { return Promise.resolve(editor.innerHTML); },
      getBodyText: function () { return Promise.resolve(editor.innerText || editor.textContent); },
      setBody: function (html) {
        editor.innerHTML = parser.parseFromString(html, "text/html").body.innerHTML;
        saved = null;
        return Promise.resolve();
      },
      prependBody: function (html) {
        var ctx = document.createRange();
        ctx.selectNodeContents(editor);
        editor.insertBefore(ctx.createContextualFragment(html), editor.firstChild);
        return Promise.resolve();
      },
      getSubject: function () { return Promise.resolve(subject.value); },
      setSubject: function (s) { subject.value = s; return Promise.resolve(); },

      canEditAttachments: true,
      getAttachments: function () {
        return later(files.map(function (f) {
          return { id: f.id, name: f.name, size: f.size, type: f.type, isInline: false };
        }));
      },
      getAttachmentContent: function (id) {
        var f = files.filter(function (x) { return x.id === id; })[0];
        if (!f) return Promise.reject(new Error("That attachment is no longer on the email."));
        return later(f.type === "file" ? { format: "base64", content: f.content } : { format: "eml", content: "" });
      },
      addFileFromBase64: function (content, name) {
        var src = files.filter(function (x) { return x.content === content; })[0];
        var f = { id: "att" + (++seq), name: name, size: src ? src.size : 0, type: "file", content: content };
        files.push(f);
        fileChanged();
        return later(f.id);
      },
      removeAttachment: function (id) {
        files = files.filter(function (x) { return x.id !== id; });
        fileChanged();
        return later();
      },
      onAttachmentsChanged: function (fn) { listeners.push(fn); },
    };
  }

  // ── Status line ───────────────────────────────────────────────────────────
  function showStatus(message, kind, offerUndo, sticky) {
    var box = $("#status");
    clearTimeout(statusTimer);
    box.className = "status" + (kind === "error" ? " is-error" : kind === "info" ? " is-info" : "");
    $("#status-text").textContent = message;
    $("#status-undo").hidden = !offerUndo;
    box.hidden = false;
    if (kind !== "error" && !offerUndo && !sticky) {
      statusTimer = setTimeout(function () { box.hidden = true; }, 6000);
    }
  }

  function friendly(err) { return actions.describeError(err); }

  function run(button, task) {
    if (busy) return;
    busy = true;
    if (button) button.setAttribute("aria-busy", "true");
    Promise.resolve()
      .then(task)
      .then(function (result) {
        if (!result) return;
        if (typeof result === "string") showStatus(result, "ok");
        else showStatus(result.message, result.kind || "ok", result.undo);
      }, function (err) {
        if (!(err && err.userMessage) && window.console) console.error(err);
        showStatus(friendly(err), "error");
      })
      .then(function () {
        busy = false;
        if (button) button.removeAttribute("aria-busy");
      });
  }

  function requireHtml() {
    return host.bodyType().then(function (t) {
      var plain = String(t).toLowerCase() !== "html";
      $("#plain-warning").hidden = !plain;
      if (plain) throw new UserError("This message is in plain text. Switch it to HTML (Format Text › HTML) and try again.");
    });
  }

  // ── Format tab: text styles ───────────────────────────────────────────────
  var KIND_INFO = {
    head: { example: "", describe: "Heading" },
    body: { example: "", describe: "Body text" },
    alpha: { example: "(a)", describe: "Lettered list (a), (b), (c)" },
    roman: { example: "(i)", describe: "Numbered list (i), (ii), (iii)" },
    bullet: { example: "•", describe: "Bulleted list" },
    quote: { example: "", describe: "Indented quotation" },
  };

  function indentIcon(level) {
    var ns = "http://www.w3.org/2000/svg";
    var svg = document.createElementNS(ns, "svg");
    svg.setAttribute("width", "16");
    svg.setAttribute("height", "10");
    svg.setAttribute("viewBox", "0 0 16 10");
    svg.setAttribute("aria-hidden", "true");
    [0, 4, 8].forEach(function (y) {
      var r = document.createElementNS(ns, "rect");
      var x = level * 4;
      r.setAttribute("x", String(x));
      r.setAttribute("y", String(y));
      r.setAttribute("width", String(16 - x));
      r.setAttribute("height", "2");
      r.setAttribute("fill", "currentColor");
      svg.appendChild(r);
    });
    return svg;
  }

  function renderStyleGrid() {
    var grid = $("#style-grid");
    var levels = STYLE.layout.levels;
    for (var l = 0; l < levels; l++) {
      var head = el("div", "sg-col");
      head.appendChild(indentIcon(l));
      head.appendChild(document.createTextNode("Level " + (l + 1)));
      head.setAttribute("aria-hidden", "true");
      grid.appendChild(head);
    }
    STYLE.styles.forEach(function (row, r) {
      fmt.styles.filter(function (s) { return s.kind === row.kind; }).forEach(function (s) {
        var info = KIND_INFO[s.kind] || { example: "", describe: "" };
        var btn = el("button", "sg-btn k-" + s.kind + " lvl-" + s.level + (r === 0 ? " row-first" : ""));
        btn.type = "button";
        btn.dataset.style = s.id;
        var spec = el("span", "spec");
        if (info.example) spec.appendChild(el("span", "lbl", info.example));
        spec.appendChild(document.createTextNode(s.name));
        btn.appendChild(spec);
        btn.title = s.name + ": " + info.describe + ", level " + (s.level + 1);
        btn.setAttribute("aria-label", "Apply " + s.name + " (" + info.describe + ", level " + (s.level + 1) + ")");
        btn.addEventListener("click", function () { applyStyle(btn, s); });
        grid.appendChild(btn);
      });
    });
  }

  // With text selected: restyle it. With just a cursor: start a new paragraph
  // in the style at the cursor, ready to type (actions.js).
  function applyStyle(btn, s) {
    run(btn, function () { return actions.applyStyle(host, s.id); });
  }

  // ── Format tab: colours ───────────────────────────────────────────────────
  function swatch(hex, label, extraClass) {
    var b = el("button", "swatch" + (extraClass ? " " + extraClass : ""));
    b.type = "button";
    b.style.backgroundColor = hex;
    b.title = label + " (" + hex + ")";
    b.setAttribute("aria-label", label);
    b.addEventListener("click", function () { applyColour(b, hex, label); });
    return b;
  }

  function renderPalette() {
    var row = $("#theme-row");
    var grid = $("#shade-grid");
    fmt.theme.forEach(function (t) {
      row.appendChild(swatch(t.hex, t.name + ", " + t.slot));
    });
    for (var r = 0; r < 5; r++) {
      fmt.theme.forEach(function (t) {
        var s = t.shades[r];
        grid.appendChild(swatch(s.hex, t.name + ", " + t.slot + ", " + s.label, r === 0 ? "first" : ""));
      });
    }
    $("#colour-auto").addEventListener("click", function () { applyColour(this, null, "Automatic"); });
  }

  // With text selected: colour it. With just a cursor: type in that colour.
  function applyColour(btn, hex, label) {
    run(btn, function () { return actions.applyColour(host, hex, label); });
  }

  // ── Format tab: clean up ──────────────────────────────────────────────────
  function wireCleanUp() {
    $("#clear-format").addEventListener("click", function () {
      var btn = this;
      run(btn, function () {
        return actions.clearFormatting(host);
      });
    });

    $("#tidy").addEventListener("click", function () {
      var btn = this;
      run(btn, function () {
        var before;
        return requireHtml()
          .then(function () { return host.getBody(); })
          .then(function (html) {
            before = html;
            var res = fmt.tidyBody(html, { colours: $("#tidy-colours").checked });
            if (!res.changed) return { message: "This email already follows the house style.", kind: "info" };
            return host.setBody(res.html).then(function () {
              undoSnapshot = before;
              var n = res.changed + (res.changed === 1 ? " item" : " items");
              return {
                message: "Tidied " + n + "." + (res.quoteFound ? " The quoted thread was left as it was." : ""),
                undo: true,
              };
            });
          });
      });
    });

    $("#status-undo").addEventListener("click", function () {
      if (!undoSnapshot) return;
      var snapshot = undoSnapshot;
      run(this, function () {
        return host.setBody(snapshot).then(function () {
          undoSnapshot = null;
          refreshMarkings();
          return "Restored the email as it was.";
        });
      });
    });
  }

  // ── Insert tab ────────────────────────────────────────────────────────────
  function renderMarkings() {
    var list = $("#marking-list");
    STYLE.markings.forEach(function (m) {
      var chip = el("button", "chip", m.label);
      chip.type = "button";
      chip.dataset.marking = m.id;
      chip.setAttribute("aria-pressed", "false");
      chip.addEventListener("click", function () { toggleMarking(chip, m); });
      list.appendChild(chip);
    });
  }

  function refreshMarkings() {
    if (!host) return Promise.resolve();
    return host.getBodyText().then(function (text) {
      var present = fmt.detectMarkings(text);
      $all("[data-marking]").forEach(function (chip) {
        chip.setAttribute("aria-pressed", present.indexOf(chip.dataset.marking) >= 0 ? "true" : "false");
      });
    }, function () { /* body not readable yet; leave chips as they are */ });
  }

  function toggleMarking(chip, m) {
    var tagSubject = $("#mark-subject").checked;
    run(chip, function () {
      return actions.toggleMarking(host, m.id, tagSubject)
        .then(function (result) { return refreshMarkings().then(function () { return result; }); });
    });
  }

  function wireInsertTab() {
    $("#date-preview").textContent = fmt.formatDate(new Date());

    $("#insert-ref").addEventListener("click", function () {
      var btn = this;
      var refs = { ours: $("#ref-ours").value, yours: $("#ref-yours").value, matter: $("#ref-matter").value };
      run(btn, function () { return actions.insertReference(host, refs, $("#ref-subject").checked); });
    });

    $("#insert-date").addEventListener("click", function () {
      var btn = this;
      run(btn, function () { return actions.insertDate(host); });
    });

    $("#insert-disclaimer").addEventListener("click", function () {
      var btn = this;
      run(btn, function () { return actions.insertDisclaimer(host); });
    });
  }

  // ── Attachments tab ───────────────────────────────────────────────────────
  var att = { loaded: false, saving: false, original: [], items: [], fixed: [] };

  function attDirty() {
    return att.items.some(function (it, i) {
      return it.base.trim() + it.ext !== it.origName || !att.original[i] || att.original[i].id !== it.id;
    });
  }
  function attInvalid() {
    return att.items.some(function (it) { return tools.validate(it.base, it.ext); });
  }

  function loadAttachments(force) {
    if (!host.canEditAttachments) {
      $("#att-unsupported").hidden = false;
      $("#att-tools").hidden = true;
      $("#att-actions").hidden = true;
      $("#att-reload").hidden = true;
      return Promise.resolve();
    }
    if (!force && att.loaded && attDirty()) { $("#att-stale").hidden = false; return Promise.resolve(); }
    return host.getAttachments().then(function (list) {
      var files = list.filter(function (a) { return !a.isInline; });
      var editable = files.filter(function (a) { return a.type === "file"; });
      att.original = editable.map(function (a) { return { id: a.id, name: a.name }; });
      att.items = editable.map(function (a) {
        var parts = tools.splitName(a.name);
        return { id: a.id, base: parts.base, ext: parts.ext, origName: a.name, size: a.size };
      });
      att.fixed = files.filter(function (a) { return a.type !== "file"; });
      att.loaded = true;
      $("#att-stale").hidden = true;
      renderAttachments();
    }, function (err) {
      showStatus("Couldn’t read the attachments: " + friendly(err), "error");
    });
  }

  function gripIcon() {
    var ns = "http://www.w3.org/2000/svg";
    var svg = document.createElementNS(ns, "svg");
    svg.setAttribute("width", "8");
    svg.setAttribute("height", "14");
    svg.setAttribute("viewBox", "0 0 8 14");
    svg.setAttribute("aria-hidden", "true");
    [[1, 1], [6, 1], [1, 6], [6, 6], [1, 11], [6, 11]].forEach(function (p) {
      var c = document.createElementNS(ns, "circle");
      c.setAttribute("cx", String(p[0]));
      c.setAttribute("cy", String(p[1] + 0.5));
      c.setAttribute("r", "1.3");
      c.setAttribute("fill", "currentColor");
      svg.appendChild(c);
    });
    return svg;
  }

  function arrow(up) {
    var ns = "http://www.w3.org/2000/svg";
    var svg = document.createElementNS(ns, "svg");
    svg.setAttribute("width", "10");
    svg.setAttribute("height", "6");
    svg.setAttribute("viewBox", "0 0 10 6");
    svg.setAttribute("aria-hidden", "true");
    var p = document.createElementNS(ns, "path");
    p.setAttribute("d", up ? "M1 5l4-4 4 4" : "M1 1l4 4 4-4");
    p.setAttribute("fill", "none");
    p.setAttribute("stroke", "currentColor");
    p.setAttribute("stroke-width", "1.6");
    svg.appendChild(p);
    return svg;
  }

  function updateRowState(li, it, index) {
    var renamed = it.base.trim() + it.ext !== it.origName;
    var origIndex = -1;
    att.original.forEach(function (o, k) { if (o.id === it.id) origIndex = k; });
    var moved = origIndex !== index;
    var err = tools.validate(it.base, it.ext);
    $(".att-name", li).classList.toggle("is-invalid", !!err);
    $("input", li).setAttribute("aria-invalid", err ? "true" : "false");
    var errEl = $(".att-error", li);
    errEl.textContent = err || "";
    errEl.hidden = !err;
    var meta = $(".att-meta", li);
    meta.textContent = "";
    meta.appendChild(el("span", null, tools.typeLabel(it.ext) + " · " + tools.formatSize(it.size)));
    if (renamed || moved) meta.appendChild(el("span", "att-tag", renamed && moved ? "Renamed, moved" : renamed ? "Renamed" : "Moved"));
    if (renamed) meta.appendChild(el("span", "att-was", "Was: " + it.origName));
  }

  function updateAttActions() {
    var dirty = attDirty();
    $("#att-save").disabled = !dirty || attInvalid() || att.saving;
    $("#att-discard").disabled = !dirty || att.saving;
    $("#att-number").disabled = !att.items.length || att.saving;
    $("#att-unnumber").disabled = !att.items.some(function (it) { return tools.stripNumber(it.base) !== it.base; }) || att.saving;
    var total = att.items.length + att.fixed.length;
    var count = $("#att-count");
    count.textContent = String(total);
    count.hidden = !total;
  }

  function moveItem(from, to, focusSel) {
    if (to < 0 || to >= att.items.length || from === to) return;
    var it = att.items.splice(from, 1)[0];
    att.items.splice(to, 0, it);
    renderAttachments(it.id, focusSel);
  }

  var dragId = null;

  function renderAttachments(focusId, focusSel) {
    var list = $("#att-list");
    list.textContent = "";
    var n = att.items.length;
    att.items.forEach(function (it, i) {
      var li = el("li", "att-row");
      li.dataset.id = it.id;

      var grip = el("div", "att-grip");
      grip.draggable = true;
      grip.title = "Drag to reorder";
      grip.appendChild(el("span", "att-pos", String(i + 1)));
      grip.appendChild(gripIcon());
      li.appendChild(grip);

      var main = el("div", "att-main");
      var nameBox = el("div", "att-name");
      var input = el("input");
      input.type = "text";
      input.value = it.base;
      input.id = "att-name-" + it.id;
      input.spellcheck = false;
      input.setAttribute("aria-label", "Name of attachment " + (i + 1) + (it.ext ? " (" + it.ext + ")" : ""));
      input.addEventListener("input", function () {
        it.base = input.value;
        updateRowState(li, it, att.items.indexOf(it));
        updateAttActions();
      });
      input.addEventListener("keydown", function (e) {
        if (!e.altKey) return;
        var k = att.items.indexOf(it);
        if (e.key === "ArrowUp") { e.preventDefault(); moveItem(k, k - 1, "input"); }
        if (e.key === "ArrowDown") { e.preventDefault(); moveItem(k, k + 1, "input"); }
      });
      nameBox.appendChild(input);
      if (it.ext) nameBox.appendChild(el("span", "att-ext", it.ext));
      main.appendChild(nameBox);
      main.appendChild(el("div", "att-meta"));
      var errEl = el("div", "att-error");
      errEl.setAttribute("role", "alert");
      main.appendChild(errEl);
      li.appendChild(main);

      var move = el("div", "att-move");
      var up = el("button");
      up.type = "button";
      up.className = "up";
      up.disabled = i === 0;
      up.setAttribute("aria-label", "Move " + it.base + " up");
      up.appendChild(arrow(true));
      up.addEventListener("click", function () { var k = att.items.indexOf(it); moveItem(k, k - 1, ".up"); });
      var down = el("button");
      down.type = "button";
      down.className = "down";
      down.disabled = i === n - 1;
      down.setAttribute("aria-label", "Move " + it.base + " down");
      down.appendChild(arrow(false));
      down.addEventListener("click", function () { var k = att.items.indexOf(it); moveItem(k, k + 1, ".down"); });
      move.appendChild(up);
      move.appendChild(down);
      li.appendChild(move);

      // Drag and drop (mouse). The arrow buttons and Alt+↑/↓ do the same.
      grip.addEventListener("dragstart", function (e) {
        dragId = it.id;
        li.classList.add("is-dragging");
        e.dataTransfer.effectAllowed = "move";
        try { e.dataTransfer.setData("text/plain", it.id); e.dataTransfer.setDragImage(li, 12, 12); } catch (ignore) { /* older engines */ }
      });
      grip.addEventListener("dragend", function () {
        dragId = null;
        $all(".att-row", list).forEach(function (r) { r.classList.remove("is-dragging", "drop-before", "drop-after"); });
      });
      li.addEventListener("dragover", function (e) {
        if (!dragId) return;
        e.preventDefault();
        e.dataTransfer.dropEffect = "move";
        var rect = li.getBoundingClientRect();
        var after = e.clientY > rect.top + rect.height / 2;
        $all(".att-row", list).forEach(function (r) { r.classList.remove("drop-before", "drop-after"); });
        if (dragId !== it.id) li.classList.add(after ? "drop-after" : "drop-before");
      });
      li.addEventListener("drop", function (e) {
        if (!dragId) return;
        e.preventDefault();
        var rect = li.getBoundingClientRect();
        var after = e.clientY > rect.top + rect.height / 2;
        var from = -1;
        att.items.forEach(function (x, k) { if (x.id === dragId) from = k; });
        var target = att.items.indexOf(it) + (after ? 1 : 0);
        if (from < target) target--;
        dragId = null;
        if (from === target) renderAttachments();
        else moveItem(from, target, null);
      });

      list.appendChild(li);
      updateRowState(li, it, i);
    });

    var fixedList = $("#att-fixed-list");
    fixedList.textContent = "";
    att.fixed.forEach(function (a) {
      var li = el("li");
      li.appendChild(el("span", "why", a.type === "item" ? "Attached email:" : "Linked file:"));
      li.appendChild(el("span", "nm", a.name));
      li.title = a.name;
      fixedList.appendChild(li);
    });
    $("#att-fixed").hidden = !att.fixed.length;
    $("#att-empty").hidden = !!(att.items.length || att.fixed.length);
    $("#att-tools").hidden = !att.items.length;
    $("#att-actions").hidden = !att.items.length;
    updateAttActions();

    if (focusId) {
      var row = $('.att-row[data-id="' + focusId + '"]', list);
      var target = row && focusSel ? $(focusSel, row) : null;
      if (target && target.disabled) target = $(focusSel === ".up" ? ".down" : ".up", row);
      if (target) target.focus();
    }
  }

  function wireAttachmentsTab() {
    var sample = tools.numberName("Board resolution", 1, 9);
    $("#att-number-help").textContent = "Adds numbers to the start of each name in the order shown, for example “" + sample + "”.";

    $("#att-number").addEventListener("click", function () {
      var n = att.items.length;
      att.items.forEach(function (it, i) { it.base = tools.numberName(it.base, i + 1, n); });
      renderAttachments();
      showStatus("Numbered " + n + (n === 1 ? " file" : " files") + ". Save to apply.", "info");
    });
    $("#att-unnumber").addEventListener("click", function () {
      att.items.forEach(function (it) { it.base = tools.stripNumber(it.base); });
      renderAttachments();
      showStatus("Numbers removed. Save to apply.", "info");
    });
    $("#att-discard").addEventListener("click", function () {
      var byId = {};
      att.items.forEach(function (it) { byId[it.id] = it; });
      att.items = att.original.map(function (o) {
        var it = byId[o.id];
        var parts = tools.splitName(o.name);
        it.base = parts.base;
        it.ext = parts.ext;
        return it;
      });
      renderAttachments();
      showStatus("Changes discarded.", "info");
    });
    $("#att-reload").addEventListener("click", function () { loadAttachments(true); });
    $("#att-stale-reload").addEventListener("click", function () { loadAttachments(true); });

    $("#att-save").addEventListener("click", function () {
      var btn = this;
      run(btn, function () {
        if (attInvalid()) throw new UserError("Fix the highlighted names first.");
        var target = att.items.map(function (it) { return { id: it.id, name: it.base.trim() + it.ext }; });
        att.saving = true;
        updateAttActions();
        return tools.saveChanges(host, att.original, target, function (k, total) {
          showStatus("Saving attachments… " + k + " of " + total, "info", false, true);
        }).then(function (res) {
          att.saving = false;
          return loadAttachments(true).then(function () {
            if (res.leftovers.length) {
              return {
                kind: "error",
                message: "Saved, but Outlook kept the original of " + res.leftovers.map(function (n) { return "“" + n + "”"; }).join(", ") +
                  ". Remove the extra copy from the email by hand.",
              };
            }
            var parts = [];
            if (res.renamed) parts.push(res.renamed + (res.renamed === 1 ? " file renamed" : " files renamed"));
            if (res.moved) parts.push("order updated");
            return "Saved: " + (parts.join(", ") || "attachments updated") + ".";
          });
        }, function (err) {
          att.saving = false;
          return loadAttachments(true).then(function () {
            var why = (err && (err.userMessage || err.message)) || "Outlook refused the change.";
            throw new UserError("Couldn’t save: " + why + " Your attachments are as they were.");
          });
        });
      });
    });

    host.onAttachmentsChanged(function () {
      if (att.saving) return;
      loadAttachments(false);
    });
  }

  // ── Tabs ──────────────────────────────────────────────────────────────────
  var selectTab = function () {};

  function wireTabs() {
    var tabs = $all('[role="tab"]');
    function select(tab, focus) {
      tabs.forEach(function (t) {
        var on = t === tab;
        t.setAttribute("aria-selected", on ? "true" : "false");
        t.tabIndex = on ? 0 : -1;
        $("#" + t.getAttribute("aria-controls")).hidden = !on;
      });
      if (focus) tab.focus();
      if (tab.id === "tab-insert") refreshMarkings();
      // Refresh on opening, unless the user has unsaved edits there.
      if (tab.id === "tab-attachments" && !att.saving && !attDirty()) loadAttachments(true);
    }
    selectTab = function (name) {
      var tab = $("#tab-" + name);
      if (tab) select(tab, false);
    };
    tabs.forEach(function (tab, i) {
      tab.addEventListener("click", function () { select(tab, false); });
      tab.addEventListener("keydown", function (e) {
        var next = null;
        if (e.key === "ArrowRight") next = tabs[(i + 1) % tabs.length];
        if (e.key === "ArrowLeft") next = tabs[(i - 1 + tabs.length) % tabs.length];
        if (e.key === "Home") next = tabs[0];
        if (e.key === "End") next = tabs[tabs.length - 1];
        if (next) { e.preventDefault(); select(next, true); }
      });
    });
  }

  // ── Preview only: the ribbon the add-in adds to Outlook ───────────────────
  // Built from the same definition as the manifest (ribbon.js) and running
  // the same actions as the real ribbon buttons (commands.js).
  function iconSrc(name) {
    var inlined = window.RIBBON_ICONS && window.RIBBON_ICONS[name];
    if (inlined) return inlined;
    return name === "icon" ? "../assets/icon-32.png" : "../assets/ribbon/" + name + "-32.png";
  }

  function showNotice(message, kind) {
    var bar = $("#demo-notice");
    $("#demo-notice-text").textContent = message;
    bar.className = "compose-notice" + (kind === "error" ? " is-error" : "");
    bar.hidden = false;
  }

  function renderMockRibbon() {
    var ribbon = window.createRibbon(STYLE, fmt);
    var body = $("#demo-ribbon-body");
    var openMenu = null;
    var NOTES = {
      house: "House Style tab: shown in classic Outlook for Windows. Outlook on Mac, on the web and new Outlook don’t show add-in tabs; there, the same tools are in the group on the Message tab.",
      message: "Message tab group: shown in every Outlook, including Mac. In new Outlook and on the web it may sit under Apps until pinned.",
    };
    $("#demo-notice-close").addEventListener("click", function () { $("#demo-notice").hidden = true; });

    function closeMenu() {
      if (!openMenu) return;
      openMenu.menu.remove();
      openMenu.button.setAttribute("aria-expanded", "false");
      openMenu = null;
    }
    // Like Outlook: ribbon buttons act on the email and leave a message in the bar.
    function activate(item) {
      closeMenu();
      if (item.pane) {
        selectTab(item.pane);
        $("#pane").scrollIntoView({ block: "nearest", behavior: "smooth" });
        return;
      }
      Promise.resolve()
        .then(function () { return item.run(actions, host, { tagSubject: true }); })
        .then(function (res) {
          if (res && item.fn.indexOf("mark_") === 0) refreshMarkings();
          if (res && (res.notify || res.kind === "error")) showNotice(res.message, res.kind);
          else $("#demo-notice").hidden = true;
        }, function (err) { showNotice(actions.describeError(err), "error"); });
    }
    function keepCursor(e) { e.preventDefault(); } // clicking the ribbon doesn't move the cursor

    function controlButton(c) {
      var b = el("button", "rb-control");
      b.type = "button";
      b.title = c.tip;
      b.dataset.id = c.id;
      var img = el("img");
      img.src = iconSrc(c.icon);
      img.alt = "";
      b.appendChild(img);
      b.appendChild(el("span", "rb-label", c.label));
      b.addEventListener("mousedown", keepCursor);
      if (c.type !== "menu") {
        b.addEventListener("click", function () { activate(c); });
        return b;
      }
      b.appendChild(el("span", "caret", "▼"));
      b.setAttribute("aria-haspopup", "menu");
      b.setAttribute("aria-expanded", "false");
      b.addEventListener("click", function () {
        var same = openMenu && openMenu.button === b;
        closeMenu();
        if (same) return;
        var menu = el("div", "rb-menu");
        menu.setAttribute("role", "menu");
        menu.setAttribute("aria-label", c.label);
        c.items.forEach(function (it) {
          var mi = el("button", "rb-item");
          mi.type = "button";
          mi.setAttribute("role", "menuitem");
          mi.title = it.tip;
          var ic = el("img");
          ic.src = iconSrc(it.icon);
          ic.alt = "";
          mi.appendChild(ic);
          mi.appendChild(el("span", null, it.label));
          mi.dataset.fn = it.fn || "";
          mi.dataset.id = it.id;
          mi.addEventListener("mousedown", keepCursor);
          mi.addEventListener("click", function () { activate(it); });
          menu.appendChild(mi);
        });
        document.body.appendChild(menu);
        var r = b.getBoundingClientRect();
        menu.style.left = Math.max(8, Math.min(r.left, window.innerWidth - menu.offsetWidth - 8)) + "px";
        menu.style.top = r.bottom + 2 + "px";
        b.setAttribute("aria-expanded", "true");
        openMenu = { button: b, menu: menu };
        menu.addEventListener("keydown", function (e) {
          if (e.key === "Escape") { closeMenu(); b.focus(); }
        });
      });
      return b;
    }

    function groupEl(label, controls, id) {
      var g = el("div", "ribbon-group");
      if (id) g.dataset.group = id;
      var row = el("div", "ribbon-controls");
      controls.forEach(function (c) { row.appendChild(controlButton(c)); });
      g.appendChild(row);
      g.appendChild(el("div", "ribbon-group-label", label));
      return g;
    }

    function show(which) {
      closeMenu();
      body.textContent = "";
      if (which === "house") {
        ribbon.tab.groups.forEach(function (g) { body.appendChild(groupEl(g.label, g.controls, g.id)); });
      } else {
        body.appendChild(groupEl(ribbon.groupLabel, ribbon.controls, ribbon.groupId));
      }
      $all("[data-rtab]").forEach(function (t) { t.setAttribute("aria-pressed", t.dataset.rtab === which ? "true" : "false"); });
      $("#demo-ribbon-note").textContent = NOTES[which];
    }
    $all("[data-rtab]").forEach(function (t) {
      t.addEventListener("mousedown", keepCursor);
      t.addEventListener("click", function () { show(t.dataset.rtab); });
    });
    show("house");

    document.addEventListener("mousedown", function (e) {
      if (openMenu && !openMenu.menu.contains(e.target) && !openMenu.button.contains(e.target)) closeMenu();
    });
    window.addEventListener("resize", closeMenu);
    window.addEventListener("scroll", closeMenu, true);
  }

  // ── Start ─────────────────────────────────────────────────────────────────
  function start(h) {
    host = h;
    $("#needs-script").hidden = true;
    var demo = h.name === "demo";
    document.body.classList.toggle("demo-mode", demo);
    $("#demo").hidden = !demo;
    $("#firm-name").textContent = STYLE.firm.name;
    $("#monogram").textContent = STYLE.firm.monogram || STYLE.firm.name.charAt(0);
    $("#pane-foot").textContent = STYLE.firm.name + " house style";

    if (demo) renderMockRibbon();
    renderStyleGrid();
    renderPalette();
    renderMarkings();
    wireTabs();
    wireCleanUp();
    wireInsertTab();
    wireAttachmentsTab();

    h.bodyType().then(function (t) {
      $("#plain-warning").hidden = String(t).toLowerCase() === "html";
    }, function () {});
    refreshMarkings();
    // Ribbon menu items open the pane on a particular tab: taskpane.html?tab=attachments
    var wanted = (/[?&]tab=([a-z]+)/.exec(window.location.search) || [])[1];
    if (wanted && wanted !== "format") selectTab(wanted);
    loadAttachments(false);
    window.addEventListener("focus", refreshMarkings);
  }

  function boot() {
    var started = false;
    function go(h) { if (!started) { started = true; start(h); } }
    if (window.Office && typeof Office.onReady === "function") {
      Office.onReady(function (info) {
        var inOutlook = info && info.host === Office.HostType.Outlook &&
          Office.context && Office.context.mailbox && Office.context.mailbox.item;
        go(inOutlook ? outlookHost() : demoHost());
      });
    } else {
      go(demoHost());
    }
  }

  boot();
})();
