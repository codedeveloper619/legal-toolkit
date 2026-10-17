/*
 * Legal Toolkit: side panel (Attachments, Clauses, Matter, Check, Settings).
 */
(function () {
  "use strict";

  var VERSION = "1.0.0";
  var CONFIG = LT.config;
  var $ = function (sel) { return document.querySelector(sel); };

  function el(tag, attrs, children) {
    var node = document.createElement(tag);
    Object.keys(attrs || {}).forEach(function (k) {
      var v = attrs[k];
      if (v === undefined || v === null || v === false) return;
      if (k === "text") node.textContent = v;
      else if (k === "html") node.innerHTML = v;
      else if (k.slice(0, 2) === "on") node.addEventListener(k.slice(2), v);
      else if (k === "className") node.className = v;
      else node.setAttribute(k, v === true ? "" : v);
    });
    (children || []).forEach(function (c) {
      if (c === null || c === undefined) return;
      node.appendChild(typeof c === "string" ? document.createTextNode(c) : c);
    });
    return node;
  }

  var toastTimer;
  function toast(message, kind) {
    var t = $("#toast");
    t.textContent = message;
    t.className = "toast" + (kind ? " " + kind : "");
    t.hidden = false;
    clearTimeout(toastTimer);
    toastTimer = setTimeout(function () { t.hidden = true; }, kind === "error" ? 7000 : 4000);
  }

  function errorText(err) {
    return err && err.message ? err.message : String(err);
  }

  /* ------------------------------------------------------------------ */
  /* Navigation                                                          */
  /* ------------------------------------------------------------------ */

  var views = ["attachments", "clauses", "matter", "check", "settings"];
  var opened = {};

  function show(view) {
    if (views.indexOf(view) === -1) view = "attachments";
    views.forEach(function (v) { $("#view-" + v).hidden = v !== view; });
    document.querySelectorAll(".tabs button").forEach(function (b) {
      b.setAttribute("aria-selected", b.getAttribute("data-view") === view ? "true" : "false");
    });
    var first = !opened[view];
    opened[view] = true;
    if (view === "attachments" && first) Attachments.load();
    if (view === "clauses" && first) Clauses.render();
    if (view === "matter") Matter.render();
    if (view === "check") Check.run();
    if (view === "settings") Settings.render();
  }

  /* ------------------------------------------------------------------ */
  /* Attachments                                                         */
  /* ------------------------------------------------------------------ */

  var Attachments = (function () {
    var state = { original: [], rows: [], others: [], findings: {}, busy: false };
    var list = $("#att-list");

    function supported() {
      return LT.supports("1.8") && typeof LT.item().getAttachmentsAsync === "function";
    }

    async function fetchFileAttachments() {
      var all = await LT.call(LT.item(), "getAttachmentsAsync");
      var types = Office.MailboxEnums.AttachmentType;
      var files = [], others = [];
      all.forEach(function (a) {
        if (a.isInline) return;
        if (a.attachmentType === types.File) files.push(a);
        else others.push(a);
      });
      return { files: files, others: others };
    }

    async function load() {
      if (!supported()) {
        var box = $("#att-unsupported");
        box.hidden = false;
        box.textContent = "This version of Outlook can't edit attachments. Update Outlook (Help > Check for Updates) to use this feature.";
        $("#att-actions").hidden = true;
        return;
      }
      try {
        var res = await fetchFileAttachments();
        state.original = res.files.map(function (a) { return { id: a.id, name: a.name, size: a.size }; });
        state.others = res.others;
        reset();
      } catch (err) {
        toast("Couldn't read the attachments: " + errorText(err), "error");
      }
    }

    function reset() {
      state.rows = state.original.map(function (a, i) {
        var parts = LT.splitExt(a.name);
        return { id: a.id, origIndex: i, originalName: a.name, size: a.size, base: parts.base, ext: parts.ext };
      });
      render();
    }

    function finalName(row) {
      return LT.safeFileName(row.base) + row.ext;
    }

    function isDirty() {
      if (state.rows.length !== state.original.length) return true;
      return state.rows.some(function (r, i) {
        return r.id !== state.original[i].id || finalName(r) !== state.original[i].name;
      });
    }

    function updateButtons() {
      var dirty = isDirty();
      $("#att-save").disabled = !dirty || state.busy;
      $("#att-revert").disabled = !dirty || state.busy;
      $("#att-count").textContent = state.rows.length ? String(state.rows.length) : "";
      document.querySelectorAll(".att").forEach(function (li, i) {
        var r = state.rows[i];
        if (!r) return;
        var changed = finalName(r) !== r.originalName || r.origIndex !== i;
        li.classList.toggle("changed", changed);
        li.querySelector(".num").textContent = String(i + 1);
        var was = li.querySelector(".was");
        was.textContent = finalName(r) !== r.originalName ? "was " + r.originalName : "";
        li.querySelector(".undo-one").hidden = finalName(r) === r.originalName;
        li.querySelector(".up").disabled = i === 0 || state.busy;
        li.querySelector(".down").disabled = i === state.rows.length - 1 || state.busy;
      });
    }

    var arrowUp = '<svg viewBox="0 0 10 10" aria-hidden="true"><path d="M1 7l4-4 4 4" fill="none" stroke="currentColor" stroke-width="1.6"/></svg>';
    var arrowDown = '<svg viewBox="0 0 10 10" aria-hidden="true"><path d="M1 3l4 4 4-4" fill="none" stroke="currentColor" stroke-width="1.6"/></svg>';

    function render() {
      list.innerHTML = "";
      $("#att-empty").hidden = state.rows.length > 0;
      state.rows.forEach(function (row, i) {
        var kind = LT.fileKind(row.originalName);
        var badgeText = (row.ext || "").replace(".", "").toUpperCase().slice(0, 4) || "FILE";
        var input = el("input", { type: "text", value: row.base, "aria-label": "File name for attachment " + (i + 1), spellcheck: "false" });
        input.addEventListener("input", function () {
          row.base = input.value;
          updateButtons();
        });
        var findings = state.findings[row.id];
        var li = el("li", { className: "att", "data-index": String(i) }, [
          el("span", { className: "grip", title: "Drag to reorder", "aria-hidden": "true", text: "⋮⋮" }),
          el("span", { className: "num", text: String(i + 1) }),
          el("div", { className: "att-main" }, [
            el("div", { className: "name-edit" }, [input, el("span", { className: "ext", text: row.ext })]),
            el("div", { className: "meta" }, [
              el("span", { className: "badge " + kind, text: badgeText }),
              el("span", { text: LT.formatSize(row.size) }),
              el("span", { className: "was" }),
              el("button", {
                className: "link undo-one", type: "button", text: "undo",
                onclick: function () { var p = LT.splitExt(row.originalName); row.base = p.base; row.ext = p.ext; render(); }
              })
            ]),
            findings ? el("div", { className: "findings " + (findings.clean ? "ok" : "bad"), text: findings.text }) : null
          ]),
          el("div", { className: "moves" }, [
            el("button", { className: "up", type: "button", title: "Move up", "aria-label": "Move up", html: arrowUp, onclick: function () { move(i, i - 1); } }),
            el("button", { className: "down", type: "button", title: "Move down", "aria-label": "Move down", html: arrowDown, onclick: function () { move(i, i + 1); } })
          ])
        ]);
        wireDrag(li, i);
        list.appendChild(li);
      });
      var others = state.others.length;
      $("#att-others").hidden = !others;
      if (others) {
        $("#att-others").textContent = others + " attached email" + (others === 1 ? "" : "s") +
          " or cloud link" + (others === 1 ? "" : "s") + " can't be renamed or moved and will stay where " + (others === 1 ? "it is." : "they are.");
      }
      updateButtons();
    }

    function move(from, to) {
      if (to < 0 || to >= state.rows.length || from === to || state.busy) return;
      var row = state.rows.splice(from, 1)[0];
      state.rows.splice(to, 0, row);
      render();
      var btn = list.children[to] && list.children[to].querySelector(to > from ? ".down" : ".up");
      if (btn && !btn.disabled) btn.focus();
    }

    var dragFrom = null;
    function wireDrag(li, index) {
      var grip = li.querySelector(".grip");
      grip.addEventListener("mousedown", function () { li.draggable = true; });
      grip.addEventListener("mouseup", function () { li.draggable = false; });
      li.addEventListener("dragstart", function (e) {
        dragFrom = index;
        li.classList.add("dragging");
        e.dataTransfer.effectAllowed = "move";
        try { e.dataTransfer.setData("text/plain", String(index)); } catch (x) {}
      });
      li.addEventListener("dragend", function () {
        li.draggable = false;
        li.classList.remove("dragging");
        clearDropMarks();
        dragFrom = null;
      });
      li.addEventListener("dragover", function (e) {
        if (dragFrom === null) return;
        e.preventDefault();
        var rect = li.getBoundingClientRect();
        var after = e.clientY > rect.top + rect.height / 2;
        clearDropMarks();
        li.classList.add(after ? "drop-after" : "drop-before");
      });
      li.addEventListener("drop", function (e) {
        if (dragFrom === null) return;
        e.preventDefault();
        var rect = li.getBoundingClientRect();
        var after = e.clientY > rect.top + rect.height / 2;
        var to = index + (after ? 1 : 0);
        if (dragFrom < to) to -= 1;
        var from = dragFrom;
        dragFrom = null;
        clearDropMarks();
        move(from, to);
      });
    }
    function clearDropMarks() {
      list.querySelectorAll(".drop-before,.drop-after").forEach(function (n) { n.classList.remove("drop-before", "drop-after"); });
    }

    /* ---- naming tools ---- */

    var LABELED = /^(?:Tab|Exhibit|Schedule|Annex|Appendix)\s+[A-Za-z0-9]{1,4}(?:\s+-\s+|\s*_\s*|\s+)/i;
    var NUMBERED = /^\d{2,3}(?:\s+-\s+|\s*_\s*|\s+)/;
    var DATED = /^(?:\d{4}-\d{2}-\d{2}|\d{8})(?:\s+-\s+|\s*_\s*|\s+)/;

    function stripApplied(base, prefix) {
      var changed = true;
      while (changed) {
        changed = false;
        var before = base;
        if (prefix && base.toLowerCase().indexOf(prefix.toLowerCase()) === 0) {
          base = base.slice(prefix.length).replace(/^(?:\s+-\s+|\s*_\s*|\s+)/, "");
        }
        base = base.replace(DATED, "").replace(LABELED, "").replace(NUMBERED, "");
        if (base !== before) changed = true;
      }
      return base;
    }

    function tidy(base) {
      return base
        .replace(/^(copy of\s+)+/i, "")
        .replace(/\s*(\(\d+\)|-\s*copy(\s*\d+)?|\s+copy(\s*\d+)?)$/i, "")
        .replace(/_+/g, " ")
        .replace(/\s{2,}/g, " ")
        .trim();
    }

    function toLetters(n) {
      return LT.letterFor(n - 1).toUpperCase();
    }

    function lettersToNumber(s) {
      var n = 0;
      s.toUpperCase().split("").forEach(function (ch) { n = n * 26 + (ch.charCodeAt(0) - 64); });
      return n;
    }

    function label(scheme, n) {
      switch (scheme) {
        case "num": return (n < 10 ? "0" : "") + n;
        case "tab": return "Tab " + n;
        case "exhibitA": return "Exhibit " + toLetters(n);
        case "exhibit1": return "Exhibit " + n;
        case "schedule": return "Schedule " + n;
        case "annex": return "Annex " + toLetters(n);
        case "appendix": return "Appendix " + n;
        default: return "";
      }
    }

    function today() {
      var d = new Date();
      var p = function (x) { return (x < 10 ? "0" : "") + x; };
      return d.getFullYear() + "-" + p(d.getMonth() + 1) + "-" + p(d.getDate());
    }

    function applyNaming() {
      var scheme = $("#nm-scheme").value;
      var startRaw = $("#nm-start").value.trim() || "1";
      var start = /^[A-Za-z]{1,3}$/.test(startRaw) ? lettersToNumber(startRaw) : parseInt(startRaw, 10);
      if (!start || start < 1) start = 1;
      var prefix = LT.safeFileName($("#nm-prefix").value);
      var addDate = $("#nm-date").checked;
      var doTidy = $("#nm-tidy").checked;
      var find = $("#nm-find").value;
      var repl = $("#nm-replace").value;
      var sep = " - ";
      state.rows.forEach(function (row, i) {
        var base = stripApplied(row.base, prefix);
        if (doTidy) base = tidy(base);
        if (find) base = base.split(find).join(repl);
        var parts = [];
        if (prefix) parts.push(prefix);
        if (addDate) parts.push(today());
        if (scheme) parts.push(label(scheme, start + i));
        parts.push(base.trim() || "Attachment " + (i + 1));
        row.base = parts.join(sep);
      });
      render();
      toast("Names updated. Press Save to email to apply them.");
    }

    function sortAZ() {
      state.rows.sort(function (a, b) {
        return finalName(a).localeCompare(finalName(b), undefined, { numeric: true, sensitivity: "base" });
      });
      render();
    }

    /* ---- save ---- */

    function setBusy(busy, text, fraction) {
      state.busy = busy;
      var p = $("#att-progress");
      p.hidden = !busy;
      if (busy) {
        p.querySelector("span").textContent = text || "";
        p.querySelector("i").style.width = Math.round((fraction || 0) * 100) + "%";
      }
      list.querySelectorAll("input,button").forEach(function (n) { n.disabled = busy; });
      updateButtons();
    }

    async function getBase64(id, name) {
      var content = await LT.call(LT.item(), "getAttachmentContentAsync", id);
      if (content.format !== Office.MailboxEnums.AttachmentContentFormat.Base64) {
        throw new Error("“" + name + "” can't be re-attached by Outlook (it is a " + content.format + " attachment).");
      }
      return content.content;
    }

    async function save() {
      var names = state.rows.map(finalName);
      for (var i = 0; i < state.rows.length; i++) {
        if (!LT.safeFileName(state.rows[i].base)) {
          toast("Attachment " + (i + 1) + " needs a name.", "error");
          return;
        }
      }
      var lower = names.map(function (n) { return n.toLowerCase(); });
      var dup = lower.find(function (n, i) { return lower.indexOf(n) !== i; });
      if (dup) {
        toast("Two attachments would both be called “" + names[lower.indexOf(dup)] + "”. Give them different names.", "error");
        return;
      }
      // Attachments before the first change stay as they are; the rest are re-attached in the new order.
      var k = 0;
      while (k < state.rows.length && state.rows[k].id === state.original[k].id && names[k] === state.original[k].name) k++;
      if (k === state.rows.length) {
        toast("Nothing to save.");
        return;
      }
      var todo = state.rows.slice(k).map(function (row, j) { return { row: row, name: names[k + j] }; });
      var item = LT.item();
      try {
        setBusy(true, "Reading attachments…", 0.05);
        try {
          for (var a = 0; a < todo.length; a++) todo[a].content = await getBase64(todo[a].row.id, todo[a].row.originalName);
        } catch (firstErr) {
          // Some Outlook versions can only read attachments of a saved draft: save once and retry.
          await LT.call(item, "saveAsync");
          var fresh = (await fetchFileAttachments()).files;
          if (fresh.length !== state.original.length) throw firstErr;
          state.rows.forEach(function (r) { r.id = fresh[r.origIndex].id; });
          state.original.forEach(function (o, idx) { o.id = fresh[idx].id; });
          for (var b = 0; b < todo.length; b++) todo[b].content = await getBase64(todo[b].row.id, todo[b].row.originalName);
        }

        var added = [];
        try {
          for (var c = 0; c < todo.length; c++) {
            setBusy(true, "Re-attaching " + (c + 1) + " of " + todo.length + "…", 0.1 + 0.7 * (c / todo.length));
            added.push(await LT.call(item, "addFileAttachmentFromBase64Async", todo[c].content, todo[c].name));
          }
        } catch (addErr) {
          for (var d = 0; d < added.length; d++) {
            try { await LT.call(item, "removeAttachmentAsync", added[d]); } catch (ignore) {}
          }
          throw new Error("Outlook couldn't re-attach the files, so nothing was changed. (" + errorText(addErr) + ")");
        }

        setBusy(true, "Removing old copies…", 0.9);
        var leftovers = [];
        for (var e = 0; e < todo.length; e++) {
          try {
            await LT.call(item, "removeAttachmentAsync", todo[e].row.id);
          } catch (rmErr) {
            leftovers.push(todo[e].row.originalName);
          }
        }
        setBusy(false);
        state.findings = {};
        await load();
        if (leftovers.length) {
          toast("Saved, but the old copy of " + leftovers.join(", ") + " is still attached. Remove it by hand.", "warn");
        } else {
          toast("Attachments updated.", "ok");
        }
      } catch (err) {
        setBusy(false);
        toast(errorText(err), "error");
      }
    }

    /* ---- attachment list & hidden-content scan ---- */

    async function insertList() {
      if (!state.rows.length) {
        toast("There are no attachments to list.", "warn");
        return;
      }
      var heading = $("#enc-label").value;
      var style = LT.getSettings().style;
      var lines = state.rows.map(function (r) { return [{ text: LT.safeFileName(r.base), f: {} }]; });
      var html = LT.buildBlock("body", [[{ text: heading + ":", f: { b: true } }]], style) + LT.buildBlock("numbers", lines, style);
      try {
        await LT.call(LT.item().body, "setSelectedDataAsync", html, { coercionType: Office.CoercionType.Html });
        toast(isDirty() ? "List inserted. Remember to press Save to email so the files match." : "List inserted.", isDirty() ? "warn" : "ok");
      } catch (err) {
        toast(errorText(err), "error");
      }
    }

    async function scan() {
      var targets = state.rows.filter(function (r) {
        var k = LT.fileKind(r.originalName);
        return k === "word" || k === "excel" || k === "powerpoint";
      });
      if (!targets.length) {
        toast("No Word, Excel or PowerPoint attachments to check.");
        return;
      }
      var btn = $("#att-scan");
      btn.disabled = true;
      btn.textContent = "Checking…";
      var problems = 0;
      for (var i = 0; i < targets.length; i++) {
        var r = targets[i];
        try {
          var result = await LT.inspectOfficeFile(r.originalName, await getBase64(r.id, r.originalName));
          state.findings[r.id] = {
            clean: result.clean,
            text: result.clean ? "No hidden content found." : "Found: " + result.findings.join(", ") + "."
          };
          if (!result.clean) problems++;
        } catch (err) {
          state.findings[r.id] = { clean: false, text: "Couldn't check this file: " + errorText(err) };
          problems++;
        }
      }
      btn.disabled = false;
      btn.textContent = "Check attachments";
      render();
      toast(problems ? problems + " file" + (problems === 1 ? " needs" : "s need") + " a look before sending." : "No hidden content found.", problems ? "warn" : "ok");
    }

    $("#att-refresh").addEventListener("click", function () { state.findings = {}; load(); });
    $("#att-revert").addEventListener("click", reset);
    $("#att-save").addEventListener("click", save);
    $("#nm-apply").addEventListener("click", applyNaming);
    $("#nm-sort").addEventListener("click", sortAZ);
    $("#enc-insert").addEventListener("click", insertList);
    $("#att-scan").addEventListener("click", scan);

    return { load: load, getBase64: getBase64, fetchFileAttachments: fetchFileAttachments };
  })();

  /* ------------------------------------------------------------------ */
  /* Clauses                                                             */
  /* ------------------------------------------------------------------ */

  var Clauses = (function () {
    var pending = null;

    function userClauses() {
      return LT.roamingGet(LT.KEYS.clauses, []) || [];
    }

    function all() {
      var mine = userClauses().map(function (c) { return Object.assign({ group: "My text", mine: true }, c); });
      return mine.concat(CONFIG.clauses || []);
    }

    function fieldsOf(text) {
      var seen = {}, out = [];
      String(text).replace(/\{\{\s*([^}]+?)\s*\}\}/g, function (_, name) {
        if (!seen[name.toLowerCase()]) {
          seen[name.toLowerCase()] = true;
          out.push(name);
        }
        return _;
      });
      return out;
    }

    function render() {
      var q = LT.normalizeText($("#cl-search").value);
      var wrap = $("#cl-list");
      wrap.innerHTML = "";
      var groups = {};
      var order = [];
      all().forEach(function (c) {
        if (q && LT.normalizeText(c.title + " " + c.text).indexOf(q) === -1) return;
        if (!groups[c.group]) {
          groups[c.group] = [];
          order.push(c.group);
        }
        groups[c.group].push(c);
      });
      if (!order.length) {
        wrap.appendChild(el("p", { className: "empty", text: "No standard text matches “" + $("#cl-search").value + "”." }));
        return;
      }
      order.forEach(function (g) {
        var box = el("div", { className: "cl-group" }, [el("h3", { text: g })]);
        groups[g].forEach(function (c) {
          var preview = el("div", { className: "p", text: c.text.replace(/\{\{\s*([^}]+?)\s*\}\}/g, "[$1]") });
          box.appendChild(
            el("div", { className: "clause" }, [
              el("div", { className: "t", text: c.title }),
              el("div", { className: "btns" }, [
                c.mine ? el("button", { className: "link danger", type: "button", text: "Delete", onclick: function () { remove(c); } }) : null,
                el("button", { className: "link", type: "button", text: "Insert", onclick: function () { start(c); } })
              ]),
              preview
            ])
          );
          preview.addEventListener("click", function () { preview.classList.toggle("expanded"); });
        });
        wrap.appendChild(box);
      });
    }

    function start(clause) {
      var fields = fieldsOf(clause.text);
      if (!fields.length) {
        insert(clause.text);
        return;
      }
      pending = clause;
      $("#cl-fill-title").textContent = clause.title;
      var box = $("#cl-fields");
      box.innerHTML = "";
      fields.forEach(function (f, i) {
        box.appendChild(el("label", {}, [f, el("input", { type: "text", "data-field": f, autocomplete: "off" })]));
      });
      $("#cl-fill").hidden = false;
      $("#cl-fill").scrollIntoView({ block: "nearest" });
      var firstInput = box.querySelector("input");
      if (firstInput) firstInput.focus();
    }

    function fillAndInsert() {
      if (!pending) return;
      var values = {};
      $("#cl-fields").querySelectorAll("input").forEach(function (i) {
        values[i.getAttribute("data-field").toLowerCase()] = i.value.trim();
      });
      var text = pending.text.replace(/\{\{\s*([^}]+?)\s*\}\}/g, function (_, name) {
        return values[name.toLowerCase()] || "[insert " + name + "]";
      });
      pending = null;
      $("#cl-fill").hidden = true;
      insert(text);
    }

    async function insert(text) {
      try {
        await LT.call(LT.item().body, "setSelectedDataAsync", LT.textToBodyHtml(text), { coercionType: Office.CoercionType.Html });
        toast("Inserted.", "ok");
      } catch (err) {
        toast(errorText(err), "error");
      }
    }

    async function add() {
      var title = $("#cl-new-title").value.trim();
      var text = $("#cl-new-text").value.trim();
      if (!title || !text) {
        toast("Add a title and the text.", "warn");
        return;
      }
      var list = userClauses().concat([{ id: String(Date.now()), title: title, text: text }]);
      if (JSON.stringify(list).length > 24000) {
        toast("Your saved text is full. Delete something first.", "error");
        return;
      }
      try {
        await LT.roamingSet(LT.KEYS.clauses, list);
        $("#cl-new-title").value = "";
        $("#cl-new-text").value = "";
        $("#cl-add").open = false;
        render();
        toast("Saved to My text.", "ok");
      } catch (err) {
        toast("Couldn't save: " + errorText(err), "error");
      }
    }

    async function remove(clause) {
      var list = userClauses().filter(function (c) { return c.id !== clause.id; });
      try {
        await LT.roamingSet(LT.KEYS.clauses, list);
        render();
        toast("Deleted.");
      } catch (err) {
        toast("Couldn't delete: " + errorText(err), "error");
      }
    }

    $("#cl-search").addEventListener("input", render);
    $("#cl-cancel").addEventListener("click", function () { pending = null; $("#cl-fill").hidden = true; });
    $("#cl-insert").addEventListener("click", fillAndInsert);
    $("#cl-fields").addEventListener("keydown", function (e) { if (e.key === "Enter") fillAndInsert(); });
    $("#cl-save").addEventListener("click", add);

    return { render: render };
  })();

  /* ------------------------------------------------------------------ */
  /* Matter                                                              */
  /* ------------------------------------------------------------------ */

  var Matter = (function () {
    function recent() {
      return LT.roamingGet(LT.KEYS.matters, []) || [];
    }

    function tagFor(number) {
      var fmt = LT.getSettings().matter.subjectFormat || "[{matter}]";
      return fmt.replace("{matter}", number);
    }

    async function remember(number, name) {
      var list = recent().filter(function (m) { return m.number.toLowerCase() !== number.toLowerCase(); });
      list.unshift({ number: number, name: name || "" });
      try { await LT.roamingSet(LT.KEYS.matters, list.slice(0, 10)); } catch (e) {}
      render();
    }

    function current() {
      var number = $("#mt-number").value.trim();
      var name = $("#mt-name").value.trim();
      if (!number) {
        toast("Type the matter number first.", "warn");
        $("#mt-number").focus();
        return null;
      }
      return { number: number, name: name };
    }

    async function addToSubject() {
      var m = current();
      if (!m) return;
      try {
        var item = LT.item();
        var subject = await LT.call(item.subject, "getAsync");
        var tagged = LT.addSubjectTag(subject, tagFor(m.number));
        if (tagged === subject) {
          toast("The subject already has this matter number.");
        } else {
          await LT.call(item.subject, "setAsync", tagged);
          toast("Added to subject.", "ok");
        }
        remember(m.number, m.name);
      } catch (err) {
        toast(errorText(err), "error");
      }
    }

    async function insertRef() {
      var m = current();
      if (!m) return;
      var runs = [{ text: "Our ref: ", f: { b: true } }, { text: m.number + (m.name ? " — " + m.name : ""), f: {} }];
      try {
        await LT.call(LT.item().body, "setSelectedDataAsync", LT.buildBlock("body", [runs]), { coercionType: Office.CoercionType.Html });
        toast("Reference inserted.", "ok");
        remember(m.number, m.name);
      } catch (err) {
        toast(errorText(err), "error");
      }
    }

    async function removeFromSubject() {
      try {
        var item = LT.item();
        var subject = await LT.call(item.subject, "getAsync");
        var cleaned = subject;
        var numbers = recent().map(function (m) { return m.number; });
        var typed = $("#mt-number").value.trim();
        if (typed) numbers.unshift(typed);
        numbers.forEach(function (n) { cleaned = LT.removeSubjectTag(cleaned, tagFor(n)); });
        if (cleaned === subject) {
          toast("No matter number from this panel found in the subject.");
          return;
        }
        await LT.call(item.subject, "setAsync", cleaned);
        toast("Removed from subject.", "ok");
      } catch (err) {
        toast(errorText(err), "error");
      }
    }

    function render() {
      var ul = $("#mt-recent");
      ul.innerHTML = "";
      var items = recent();
      $("#mt-recent-empty").hidden = items.length > 0;
      items.forEach(function (m) {
        ul.appendChild(
          el("li", {}, [
            el("button", {
              className: "pick", type: "button",
              onclick: function () { $("#mt-number").value = m.number; $("#mt-name").value = m.name || ""; }
            }, [el("b", { text: m.number }), m.name ? el("span", { className: "nm", text: " — " + m.name }) : null])
          ])
        );
      });
    }

    $("#mt-subject").addEventListener("click", addToSubject);
    $("#mt-ref").addEventListener("click", insertRef);
    $("#mt-remove").addEventListener("click", removeFromSubject);
    return { render: render };
  })();

  /* ------------------------------------------------------------------ */
  /* Pre-send check                                                      */
  /* ------------------------------------------------------------------ */

  var Check = (function () {
    var running = false;
    var ATTACH_WORDS = /\b(attach(ed|ment|ments|ing)?|enclos(ed|ure|ures|ing)|please find)\b/i;
    var PLACEHOLDERS = /\[\s*(?:●|•|\.{3}|…|_+|x+|tbc|tbd|insert[^\]]{0,60}|date|name|client|amount)\s*\]|\{\{[^}]{1,60}\}\}|\bX{3,}\b|\bTB[CD]\b|<<[^>]{1,40}>>/gi;

    function addr(list) {
      return (list || []).filter(function (r) { return r && r.emailAddress; });
    }

    function isInternal(domain, firm) {
      return firm.some(function (f) { return domain === f || domain.slice(-(f.length + 1)) === "." + f; });
    }

    async function run() {
      if (running) return;
      running = true;
      var summary = $("#ck-summary");
      var ul = $("#ck-results");
      summary.className = "summary";
      summary.innerHTML = "";
      summary.appendChild(el("div", { className: "spinner" }));
      summary.appendChild(el("div", { className: "big", text: "Checking…" }));
      ul.innerHTML = "";
      var results = [];
      function add(level, title, detail, action) {
        results.push({ level: level, title: title, detail: detail, action: action });
      }
      try {
        var item = LT.item();
        var settings = LT.getSettings();
        var got = await Promise.all([
          LT.call(item.to, "getAsync"),
          LT.call(item.cc, "getAsync"),
          LT.call(item.bcc, "getAsync"),
          LT.call(item.subject, "getAsync"),
          LT.call(item.body, "getAsync", Office.CoercionType.Text)
        ]);
        var to = addr(got[0]), cc = addr(got[1]), bcc = addr(got[2]);
        var subject = got[3] || "";
        var own = LT.ownText(got[4] || "");
        var everyone = to.concat(cc, bcc);
        var firm = settings.firmDomains;

        // Recipients
        var external = everyone.filter(function (r) { return !isInternal(LT.domainOf(r.emailAddress), firm); });
        var extDomains = [];
        external.forEach(function (r) {
          var d = LT.domainOf(r.emailAddress);
          if (extDomains.indexOf(d) === -1) extDomains.push(d);
        });
        if (!everyone.length) {
          add("fail", "No recipients yet", "Add at least one recipient.");
        } else if (!external.length) {
          add("ok", "All recipients are internal", firm.length ? "Everyone is at " + firm.join(", ") + "." : "");
        } else if (extDomains.length === 1) {
          add("info", "External: " + extDomains[0], external.length + " outside recipient" + (external.length === 1 ? "" : "s") + ". Confirm the address is right.");
        } else {
          add("warn", "Going to " + extDomains.length + " outside organizations", extDomains.join(", ") + ". Confirm each one should receive this email.");
        }
        if (!firm.length) {
          add("info", "Firm domain not set", "Add your firm's email domain in Settings so external recipients can be recognised.", { label: "Open Settings", run: function () { show("settings"); } });
        }
        var compareTo = firm.concat(extDomains);
        var pairs = {};
        extDomains.forEach(function (d) {
          compareTo.forEach(function (other) {
            if (other === d || Math.min(d.length, other.length) < 5) return;
            var key = [d, other].sort().join("|");
            if (pairs[key]) return;
            pairs[key] = true;
            if (LT.levenshtein(d, other) <= 2) {
              var known = firm.indexOf(other) !== -1 ? other : null;
              add("fail", "Look-alike email domains",
                known
                  ? "“" + d + "” is one or two letters away from your firm's domain “" + known + "”. Check for a typo or a spoofed address."
                  : "“" + d + "” and “" + other + "” are almost identical. One may be a typo or a spoofed address.");
            }
          });
        });
        everyone.forEach(function (r) {
          var shown = String(r.displayName || "");
          var m = /[\w.+-]+@[\w-]+(\.[\w-]+)+/.exec(shown);
          if (m && m[0].toLowerCase() !== r.emailAddress.toLowerCase()) {
            add("fail", "Name and address don't match", "“" + shown + "” actually sends to " + r.emailAddress + ".");
          }
        });
        var large = (CONFIG.check && CONFIG.check.largeDistribution) || 15;
        if (everyone.length >= large) add("warn", "Large distribution", everyone.length + " recipients. Check the list, or use BCC.");
        if (bcc.length) add("info", "BCC: " + bcc.length + " recipient" + (bcc.length === 1 ? "" : "s"), bcc.map(function (r) { return r.emailAddress; }).join(", "));

        // Subject and matter number
        if (!subject.trim()) add("fail", "Subject is empty", "Add a subject line.");
        var matterRe = null;
        try { matterRe = new RegExp(settings.matter.pattern || "$^"); } catch (e) {}
        var hasMatter = matterRe ? matterRe.test(subject) : false;
        if (hasMatter) add("ok", "Matter number in subject", "");
        else if (settings.matter.requireInSubject) add("warn", "No matter number in subject", "Your firm asks for a matter number in the subject.", { label: "Add matter number", run: function () { show("matter"); } });

        // Markings
        var markings = CONFIG.markings || {};
        var found = Object.keys(markings).filter(function (id) {
          var m = markings[id];
          return subject.toLowerCase().indexOf(m.subjectTag.toLowerCase()) !== -1 ||
            LT.hasBannerLine(own, m.banner);
        });
        if (found.length) {
          add("ok", "Marked: " + found.map(function (id) { return markings[id].label; }).join(", "), "");
          var privileged = found.indexOf("privileged") !== -1 || found.indexOf("workProduct") !== -1;
          if (privileged && extDomains.length > 1) {
            add("warn", "Privileged email to several outside organizations", "Sharing privileged advice with third parties can waive privilege. Confirm everyone on the email is within the privilege.");
          }
        } else {
          add("info", "No privilege or confidentiality marking", "Use the Mark menu on the ribbon if this email needs one.");
        }

        // Placeholders
        var holes = own.match(PLACEHOLDERS) || [];
        if (holes.length) {
          var uniq = holes.filter(function (h, i) { return holes.indexOf(h) === i; }).slice(0, 6);
          add("fail", "Unfinished blanks in the email", uniq.join("  "));
        }

        // Attachments
        if (LT.supports("1.8") && typeof item.getAttachmentsAsync === "function") {
          var files = (await Attachments.fetchFileAttachments()).files;
          // Ignore disclaimer wording such as "this email and any attachments are confidential".
          var prose = own.split(/\n/).filter(function (line) {
            return !/confidential|intended recipient|privileged|disclaimer|virus/i.test(line);
          }).join("\n").replace(/\b(any|its|all|and|or|with|including)\s+(of\s+)?(its\s+|their\s+)?attachments?\b/gi, "");
          var mentions = ATTACH_WORDS.test(prose);
          if (mentions && !files.length) {
            add("fail", "Mentions an attachment, but nothing is attached", "The email talks about an attachment or enclosure.");
          }
          if (files.length) {
            var total = files.reduce(function (s, f) { return s + (f.size || 0); }, 0);
            add("ok", files.length + " attachment" + (files.length === 1 ? "" : "s") + " (" + LT.formatSize(total) + ")", files.map(function (f) { return f.name; }).join(", "));
            var maxMb = (CONFIG.check && CONFIG.check.maxTotalAttachmentMB) || 20;
            if (total > maxMb * 1024 * 1024) add("warn", "Attachments are large", LT.formatSize(total) + " may be rejected by the recipient's mail server. Consider a secure file-sharing link.");
            var words = (CONFIG.check && CONFIG.check.attachmentWarnWords) || [];
            var flagged = files.filter(function (f) {
              var n = f.name.toLowerCase();
              return words.some(function (w) { return n.indexOf(w.toLowerCase()) !== -1; });
            });
            if (flagged.length) add("warn", "Check these file names", flagged.map(function (f) { return f.name; }).join(", "));
            for (var i = 0; i < files.length; i++) {
              var kind = LT.fileKind(files[i].name);
              if (kind !== "word" && kind !== "excel" && kind !== "powerpoint") continue;
              try {
                var res = await LT.inspectOfficeFile(files[i].name, await Attachments.getBase64(files[i].id, files[i].name));
                if (!res.clean) {
                  add("fail", "Hidden content in " + files[i].name, res.findings.join(", ") + ". Clean it up or send a PDF.");
                }
              } catch (err) {
                add("info", "Couldn't look inside " + files[i].name, errorText(err));
              }
            }
          }
        } else {
          add("info", "Attachments not checked", "Update Outlook to check attachments.");
        }
      } catch (err) {
        add("fail", "The check couldn't finish", errorText(err));
      }
      running = false;
      draw(results);
    }

    var ORDER = { fail: 0, warn: 1, info: 2, ok: 3 };
    var ICON = { fail: "!", warn: "!", info: "i", ok: "✓" };

    function draw(results) {
      results.sort(function (a, b) { return ORDER[a.level] - ORDER[b.level]; });
      var fails = results.filter(function (r) { return r.level === "fail"; }).length;
      var warns = results.filter(function (r) { return r.level === "warn"; }).length;
      var summary = $("#ck-summary");
      summary.innerHTML = "";
      var level = fails ? "fail" : warns ? "warn" : "ok";
      summary.className = "summary " + level;
      var headline = fails
        ? fails + " problem" + (fails === 1 ? "" : "s") + " to fix" + (warns ? ", " + warns + " to review" : "")
        : warns ? warns + " thing" + (warns === 1 ? "" : "s") + " to review" : "Ready to send";
      summary.appendChild(el("div", { className: "big", text: headline }));
      var ul = $("#ck-results");
      ul.innerHTML = "";
      results.forEach(function (r) {
        ul.appendChild(
          el("li", { className: r.level }, [
            el("span", { className: "ic", text: ICON[r.level], "aria-label": r.level }),
            el("div", {}, [
              el("div", { className: "ttl", text: r.title }),
              r.detail ? el("div", { className: "det", text: r.detail }) : null,
              r.action ? el("div", { className: "act" }, [el("button", { className: "link", type: "button", text: r.action.label, onclick: r.action.run })]) : null
            ])
          ])
        );
      });
    }

    $("#ck-run").addEventListener("click", run);
    return { run: run };
  })();

  /* ------------------------------------------------------------------ */
  /* Settings                                                            */
  /* ------------------------------------------------------------------ */

  var Settings = (function () {
    var returnTo = "attachments";

    function selectValue(select, value) {
      var found = Array.prototype.some.call(select.options, function (o) { return o.value === String(value); });
      if (!found && value) select.appendChild(el("option", { value: String(value), text: String(value).split(",")[0].replace(/'/g, "") }));
      select.value = String(value);
    }

    function render() {
      var s = LT.getSettings();
      $("#st-domains").value = s.firmDomains.join(", ");
      selectValue($("#st-font"), s.style.fontFamily);
      selectValue($("#st-size"), s.style.bodySizePt);
      $("#st-color").value = /^#[0-9a-f]{6}$/i.test(s.style.headingColor) ? s.style.headingColor : "#1f3864";
      $("#st-require-matter").checked = !!s.matter.requireInSubject;
      $("#st-version").textContent = "Legal Toolkit " + VERSION + " · " + (CONFIG.firmName || "");
      preview();
    }

    function formStyle() {
      return Object.assign({}, CONFIG.style, {
        fontFamily: $("#st-font").value,
        bodySizePt: parseFloat($("#st-size").value),
        headingColor: $("#st-color").value
      });
    }

    function preview() {
      var style = formStyle();
      $("#st-preview").innerHTML =
        LT.buildBlock("h1", [[{ text: "Heading", f: {} }]], style) +
        LT.buildBlock("body", [[{ text: "Body text looks like this.", f: {} }]], style) +
        LT.buildBlock("bullets", [[{ text: "A bullet point", f: {} }]], style);
    }

    async function save() {
      var domains = $("#st-domains").value.split(/[,;\s]+/).map(function (d) { return d.trim().toLowerCase().replace(/^@/, ""); }).filter(Boolean);
      // Only keep domains the user actually changed, so later firm-wide updates still reach them.
      var defaults = (CONFIG.firmDomains || []).length ? CONFIG.firmDomains : [LT.domainOf(LT.userEmail())];
      if (domains.join(",") === defaults.join(",").toLowerCase()) domains = [];
      var user = {
        firmDomains: domains,
        style: { fontFamily: $("#st-font").value, bodySizePt: parseFloat($("#st-size").value), headingColor: $("#st-color").value },
        matter: { requireInSubject: $("#st-require-matter").checked }
      };
      try {
        await LT.roamingSet(LT.KEYS.settings, user);
        toast("Settings saved.", "ok");
        show(returnTo);
      } catch (err) {
        toast("Couldn't save settings: " + errorText(err), "error");
      }
    }

    async function resetDefaults() {
      try {
        await LT.roamingSet(LT.KEYS.settings, null);
        render();
        toast("Back to firm defaults.", "ok");
      } catch (err) {
        toast(errorText(err), "error");
      }
    }

    ["#st-font", "#st-size", "#st-color"].forEach(function (id) { $(id).addEventListener("input", preview); });
    $("#st-save").addEventListener("click", save);
    $("#st-reset").addEventListener("click", resetDefaults);
    $("#st-back").addEventListener("click", function () { show(returnTo); });
    $("#open-settings").addEventListener("click", function () {
      var visible = views.filter(function (v) { return !$("#view-" + v).hidden; })[0];
      if (visible && visible !== "settings") returnTo = visible;
      show("settings");
    });

    return { render: render };
  })();

  /* ------------------------------------------------------------------ */
  /* Start                                                               */
  /* ------------------------------------------------------------------ */

  document.querySelectorAll(".tabs button").forEach(function (b) {
    b.addEventListener("click", function () { show(b.getAttribute("data-view")); });
  });

  Office.onReady(function (info) {
    var inOutlook = info && info.host === Office.HostType.Outlook && Office.context.mailbox && Office.context.mailbox.item;
    if (!inOutlook) {
      $("#not-in-outlook").hidden = false;
      document.querySelector(".tabs").hidden = true;
      return;
    }
    var params = new URLSearchParams(location.search);
    show(params.get("view") || window.LT_START_VIEW || "attachments");
  });
})();
