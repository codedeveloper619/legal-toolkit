/*
 * The formatting actions, shared by the task pane and the ribbon buttons so
 * they always behave the same. Each action takes a "host" (Outlook, or the
 * browser preview) and returns a promise of a result:
 *   { message: "…", kind: "ok" | "info" | "error", notify: true | false }
 * Problems the user can fix are thrown as errors with a `userMessage`.
 */
(function (root) {
  "use strict";

  var SUBJECT_MSG = "Outlook says your cursor is in the subject line. Click in the message body, then try again.";
  var PLAIN_MSG = "This message is in plain text. Switch it to HTML (Format Text › HTML) and try again.";
  var LISTS_KEY = "house-style:lists";
  var LISTS_EXPIRE_MS = 30 * 60 * 1000;

  function UserError(message) {
    this.name = "UserError";
    this.userMessage = message;
    this.message = message;
  }
  UserError.prototype = Object.create(Error.prototype);
  UserError.prototype.constructor = UserError;

  function createActions(style, fmt, opts) {
    opts = opts || {};
    var win = opts.win || root;
    var parser = new win.DOMParser();
    var storage = opts.storage || null;
    var now = opts.now || function () { return Date.now(); };
    var memory = null;

    function hasContent(html) {
      var body = parser.parseFromString(html || "", "text/html").body;
      return !!(body.querySelector("img") || String(body.textContent || "").replace(/[\s ​]/g, ""));
    }

    function isSubject(sel) { return String((sel && sel.source) || "").toLowerCase() === "subject"; }

    function requireHtml(host) {
      return Promise.resolve(host.bodyType()).then(function (t) {
        if (String(t).toLowerCase() !== "html") throw new UserError(PLAIN_MSG);
      });
    }

    // Some versions of Outlook misreport which field the cursor is in, so
    // nothing is refused on that basis. If the cursor really is in the
    // subject, Outlook rejects the HTML when it's inserted.
    function readSelection(host, needText, what) {
      return requireHtml(host).then(function () { return host.getSelection(); }).then(function (sel) {
        if (needText && !hasContent(sel.html)) {
          throw new UserError("Select the " + (what || "text you want to change") + " first.");
        }
        return sel;
      });
    }

    function insertAtCursor(host, sel, html) {
      return Promise.resolve(host.replaceSelection(html)).catch(function (err) {
        if (isSubject(sel)) throw new UserError(SUBJECT_MSG);
        throw err;
      });
    }

    // ── List numbering carried on between clicks ────────────────────────────
    // Kept in browser storage so the ribbon and the pane share it, and tied
    // to the email (conversation and subject) so it doesn't carry over into
    // a different message. It lapses after 30 minutes without list use.
    function contextOf(host) {
      var conv = host.contextKey ? host.contextKey() : "";
      return Promise.resolve(host.getSubject ? host.getSubject() : "").then(function (subject) {
        return String(conv || "") + "|" + String(subject || "");
      }, function () { return String(conv || ""); });
    }
    function loadCounts(context) {
      var data = memory;
      try {
        var raw = storage && storage.getItem(LISTS_KEY);
        if (raw) data = JSON.parse(raw);
      } catch (e) { /* storage unavailable */ }
      if (!data || data.context !== context || now() - data.at > LISTS_EXPIRE_MS) return {};
      return data.counts || {};
    }
    function saveCounts(context, counts) {
      memory = { context: context, at: now(), counts: counts };
      try { if (storage) storage.setItem(LISTS_KEY, JSON.stringify(memory)); } catch (e) { /* ignore */ }
    }
    function keyOf(s) { return s.kind + "-" + s.level; }
    // A list style continues its own list and ends deeper lists (and other
    // list kinds at its level). Head and Body end lists at their level and
    // below; a quotation only ends the deeper ones.
    function recordStyle(counts, s, n) {
      Object.keys(counts).forEach(function (k) {
        var kind = k.split("-")[0];
        var level = Number(k.split("-")[1]);
        var ends = s.label ? (level > s.level || (level === s.level && kind !== s.kind))
          : s.kind === "quote" ? level > s.level : level >= s.level;
        if (ends) delete counts[k];
      });
      if (s.label) counts[keyOf(s)] = n;
      return counts;
    }

    function findStyle(id) {
      for (var i = 0; i < fmt.styles.length; i++) if (fmt.styles[i].id === id) return fmt.styles[i];
      throw new Error("Unknown house style: " + id);
    }

    // ── Actions ─────────────────────────────────────────────────────────────
    // With text selected: restyle it. With just a cursor: start a new
    // paragraph in the style at the cursor, ready to type.
    function applyStyle(host, id) {
      var s = findStyle(id);
      var sel, context, counts;
      return readSelection(host, false)
        .then(function (x) { sel = x; return contextOf(host); })
        .then(function (c) {
          context = c;
          counts = loadCounts(context);
          if (!hasContent(sel.html)) {
            var n = s.label ? (counts[keyOf(s)] || 0) + 1 : 0;
            return insertAtCursor(host, sel, fmt.typingParagraph(s.id, n)).then(function () {
              saveCounts(context, recordStyle(counts, s, n));
              var label = fmt.labelFor(s.id, n);
              return { message: "Start typing" + (label ? " item " + label : "") + " in “" + s.name + "”.", kind: "ok", typing: true };
            });
          }
          var res = fmt.applyStyleInfo(s.id, sel.html);
          if (!res.html) throw new UserError("Select the text you want to change first.");
          return insertAtCursor(host, sel, res.html).then(function () {
            saveCounts(context, recordStyle(counts, s, res.count));
            return { message: "Applied “" + s.name + "”.", kind: "ok" };
          });
        });
    }

    // With text selected: colour it. With just a cursor: type in that colour.
    function applyColour(host, hex, label) {
      return readSelection(host, false).then(function (sel) {
        var typing = !hasContent(sel.html);
        var html = typing ? fmt.typingColour(hex) : fmt.applyColour(hex, sel.html);
        return insertAtCursor(host, sel, html).then(function () {
          var name = hex ? label : "Automatic (black)";
          return { message: typing ? "Start typing in " + name + "." : "Text coloured " + name + ".", kind: "ok" };
        });
      });
    }

    function clearFormatting(host) {
      return readSelection(host, true, "text you want to clear").then(function (sel) {
        return insertAtCursor(host, sel, fmt.clearFormatting(sel.html)).then(function () {
          return { message: "Formatting cleared to body text.", kind: "ok" };
        });
      });
    }

    // Adds the marking's banner (and subject tag), or removes them if the
    // email already has that marking.
    function toggleMarking(host, id, tagSubject) {
      var m = null;
      fmt.markings.forEach(function (x) { if (x.id === id) m = x; });
      if (!m) throw new Error("Unknown marking: " + id);
      return requireHtml(host)
        .then(function () { return host.getBodyText(); })
        .then(function (text) {
          var present = fmt.detectMarkings(text).indexOf(m.id) >= 0;
          if (!present) {
            return Promise.resolve(host.prependBody(fmt.markingBanner(m.id)))
              .then(function () {
                if (!tagSubject || !m.subjectTag) return;
                return Promise.resolve(host.getSubject()).then(function (s) { return host.setSubject(fmt.tagSubject(s, m.subjectTag)); });
              })
              .then(function () { return { message: "Marked “" + m.label + "”.", kind: "ok", notify: true }; });
          }
          var removed = 0;
          return Promise.resolve(host.getBody())
            .then(function (html) {
              var r = fmt.removeMarking(html, m.id);
              removed = r.removed;
              if (removed) return host.setBody(r.html);
            })
            .then(function () { return host.getSubject(); })
            .then(function (s) {
              var next = m.subjectTag ? fmt.untagSubject(s, m.subjectTag) : s;
              if (next !== s) return host.setSubject(next);
            })
            .then(function () {
              if (removed) return { message: "Removed “" + m.label + "”.", kind: "ok", notify: true };
              return {
                message: "Couldn’t find the “" + m.label + "” banner to remove. Delete it from the message by hand.",
                kind: "error", notify: true,
              };
            });
        });
    }

    function insertDate(host) {
      var date = fmt.formatDate(new Date());
      return readSelection(host, false).then(function (sel) {
        return insertAtCursor(host, sel, fmt.esc(date));
      }).then(function () { return { message: "Inserted " + date + ".", kind: "ok" }; });
    }

    function insertDisclaimer(host) {
      return readSelection(host, false).then(function (sel) {
        return insertAtCursor(host, sel, fmt.disclaimer());
      }).then(function () { return { message: "Firm disclaimer inserted.", kind: "ok" }; });
    }

    function insertReference(host, refs, tagMatter) {
      var html = fmt.referenceLine(refs);
      if (!html) return Promise.reject(new UserError("Enter at least one reference first."));
      return readSelection(host, false)
        .then(function (sel) { return insertAtCursor(host, sel, html); })
        .then(function () {
          var matter = String((refs && refs.matter) || "").trim();
          if (!matter || !tagMatter) return { message: "Reference line inserted.", kind: "ok" };
          return Promise.resolve(host.getSubject())
            .then(function (s) { return host.setSubject(fmt.tagSubject(s, fmt.matterSubjectTag(matter))); })
            .then(function () { return { message: "Reference line inserted and matter number added to the subject.", kind: "ok" }; });
        });
    }

    // A message for any error, whether the user can fix it or not.
    function describeError(err) {
      if (err && err.userMessage) return err.userMessage;
      var msg = (err && (err.message || err.name)) || "";
      if (/plain text|bodytype/i.test(msg)) return PLAIN_MSG;
      return "Outlook couldn’t complete that" + (msg ? ": " + msg : ".") + " Try again, or click in the message body first.";
    }

    return {
      applyStyle: applyStyle,
      applyColour: applyColour,
      clearFormatting: clearFormatting,
      toggleMarking: toggleMarking,
      insertDate: insertDate,
      insertDisclaimer: insertDisclaimer,
      insertReference: insertReference,
      describeError: describeError,
      hasContent: hasContent,
      requireHtml: requireHtml,
    };
  }

  createActions.UserError = UserError;
  createActions.SUBJECT_MSG = SUBJECT_MSG;
  createActions.PLAIN_MSG = PLAIN_MSG;

  if (typeof module !== "undefined" && module.exports) module.exports = createActions;
  else root.createActions = createActions;
})(typeof self !== "undefined" ? self : this);
