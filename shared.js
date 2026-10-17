/*
 * Legal Toolkit: code shared by the ribbon commands and the side panel.
 * Everything runs inside Outlook on the user's computer; nothing is sent
 * to any server.
 */
(function (global) {
  "use strict";

  var LT = (global.LT = global.LT || {});
  var CONFIG = global.LT_CONFIG || {};
  LT.config = CONFIG;

  /* ------------------------------------------------------------------ */
  /* Office helpers                                                      */
  /* ------------------------------------------------------------------ */

  // Calls an Office "...Async(args, callback)" method and returns a Promise.
  LT.call = function (target, method) {
    var args = Array.prototype.slice.call(arguments, 2);
    return new Promise(function (resolve, reject) {
      if (!target || typeof target[method] !== "function") {
        reject(new Error("This version of Outlook does not support " + method + "."));
        return;
      }
      args.push(function (result) {
        if (result.status === Office.AsyncResultStatus.Failed) {
          var err = result.error || {};
          var e = new Error(err.message || "Outlook reported an error.");
          e.code = err.code;
          reject(e);
        } else {
          resolve(result.value);
        }
      });
      target[method].apply(target, args);
    });
  };

  LT.item = function () {
    return Office.context.mailbox.item;
  };

  LT.supports = function (version) {
    try {
      return Office.context.requirements.isSetSupported("Mailbox", version);
    } catch (e) {
      return false;
    }
  };

  // Shows a message in the yellow/blue bar above the email.
  LT.notify = function (message, kind) {
    try {
      var item = LT.item();
      var types = Office.MailboxEnums.ItemNotificationMessageType;
      var details = { message: String(message).slice(0, 150) };
      if (kind === "error") {
        details.type = types.ErrorMessage;
      } else {
        details.type = types.InformationalMessage;
        details.icon = "Icon.16";
        details.persistent = false;
      }
      item.notificationMessages.replaceAsync("lt-status", details, function () {});
    } catch (e) {
      /* notifications are best-effort */
    }
  };

  LT.clearNotice = function () {
    try {
      LT.item().notificationMessages.removeAsync("lt-status", function () {});
    } catch (e) {}
  };

  /* ------------------------------------------------------------------ */
  /* Settings (firm defaults + per-user overrides saved in the mailbox)  */
  /* ------------------------------------------------------------------ */

  var KEYS = { settings: "lt.settings.v1", matters: "lt.matters.v1", clauses: "lt.clauses.v1" };
  LT.KEYS = KEYS;

  LT.roamingGet = function (key, fallback) {
    try {
      var v = Office.context.roamingSettings.get(key);
      return v === undefined || v === null ? fallback : v;
    } catch (e) {
      return fallback;
    }
  };

  LT.roamingSet = function (key, value) {
    var rs = Office.context.roamingSettings;
    if (value === undefined || value === null) rs.remove(key);
    else rs.set(key, value);
    return LT.call(rs, "saveAsync");
  };

  LT.userEmail = function () {
    try {
      return Office.context.mailbox.userProfile.emailAddress || "";
    } catch (e) {
      return "";
    }
  };

  LT.domainOf = function (address) {
    var at = String(address || "").lastIndexOf("@");
    return at === -1 ? "" : address.slice(at + 1).trim().toLowerCase();
  };

  LT.getSettings = function () {
    var user = LT.roamingGet(KEYS.settings, {}) || {};
    var style = Object.assign({}, CONFIG.style || {}, user.style || {});
    var matter = Object.assign({}, CONFIG.matter || {}, user.matter || {});
    var domains = user.firmDomains && user.firmDomains.length ? user.firmDomains : CONFIG.firmDomains || [];
    domains = domains.map(function (d) { return String(d).trim().toLowerCase().replace(/^@/, ""); }).filter(Boolean);
    if (!domains.length) {
      var own = LT.domainOf(LT.userEmail());
      if (own) domains = [own];
    }
    return { style: style, matter: matter, firmDomains: domains, user: user };
  };

  /* ------------------------------------------------------------------ */
  /* Text helpers                                                        */
  /* ------------------------------------------------------------------ */

  LT.escape = function (s) {
    return String(s == null ? "" : s)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  };

  LT.normalizeText = function (s) {
    return String(s || "")
      .replace(/[‒-―]/g, "-")
      .replace(/[\s ]+/g, " ")
      .trim()
      .toLowerCase();
  };

  // The part of a plain-text body the user wrote, i.e. everything above the
  // quoted thread of a reply or forward.
  LT.ownText = function (text) {
    text = String(text || "");
    var markers = [
      /^[ \t]*-{2,}[ \t]*Original Message[ \t]*-{2,}/im,
      /^[ \t]*From:[^\n]*\r?\n[ \t]*(Sent|Date):/im,
      /^[ \t]*On [^\n]{4,200} wrote:[ \t]*$/im,
      /^[ \t]*Begin forwarded message:/im,
      /^[ \t]*_{20,}[ \t]*$/m
    ];
    var cut = text.length;
    markers.forEach(function (re) {
      var m = re.exec(text);
      if (m && m.index < cut) cut = m.index;
    });
    return text.slice(0, cut);
  };

  // True when the banner text appears as a line of its own (so the word
  // "confidential" inside a disclaimer doesn't count as a CONFIDENTIAL marking).
  LT.hasBannerLine = function (text, banner) {
    var target = LT.normalizeText(banner);
    return String(text || "").split(/\r?\n/).some(function (line) {
      return LT.normalizeText(line) === target;
    });
  };

  LT.levenshtein = function (a, b) {
    if (a === b) return 0;
    var prev = [], cur = [], i, j;
    for (j = 0; j <= b.length; j++) prev[j] = j;
    for (i = 1; i <= a.length; i++) {
      cur = [i];
      for (j = 1; j <= b.length; j++) {
        cur[j] = Math.min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
      }
      prev = cur;
    }
    return prev[b.length];
  };

  LT.formatSize = function (bytes) {
    if (!bytes && bytes !== 0) return "";
    if (bytes < 1024) return bytes + " B";
    if (bytes < 1024 * 1024) return Math.round(bytes / 1024) + " KB";
    return (bytes / (1024 * 1024)).toFixed(1) + " MB";
  };

  LT.splitExt = function (name) {
    var m = /^(.*?)(\.[A-Za-z0-9]{1,8})?$/.exec(String(name || ""));
    var base = m[1], ext = m[2] || "";
    if (!base) return { base: name, ext: "" }; // e.g. ".hidden"
    return { base: base, ext: ext };
  };

  // Removes characters that are not allowed in file names.
  LT.safeFileName = function (name) {
    return String(name || "")
      .replace(/[\\/:*?"<>|\u0000-\u001f]/g, "-")
      .replace(/\s+/g, " ")
      .trim();
  };

  /* ------------------------------------------------------------------ */
  /* Formatting: turn the selected HTML into clean, firm-styled HTML     */
  /* ------------------------------------------------------------------ */

  var BLOCK_TAGS = /^(P|DIV|LI|UL|OL|H[1-6]|BLOCKQUOTE|TR|TABLE|TBODY|THEAD|TFOOT|SECTION|ARTICLE|HEADER|FOOTER|PRE|DT|DD|DL|HR|ADDRESS|CENTER)$/;
  var SKIP_TAGS = /^(STYLE|SCRIPT|HEAD|TITLE|META|LINK|XML|O:SMARTTAGTYPE)$/;

  // Breaks selected HTML into lines of formatted "runs". Only bold, italic,
  // underline, strike, superscript, subscript and links are kept; fonts,
  // sizes and colours (e.g. from text pasted out of Word) are dropped.
  LT.parseSelection = function (html) {
    var doc = new DOMParser().parseFromString("<!doctype html><html><body>" + html + "</body></html>", "text/html");
    var body = doc.body;
    var info = { hasImage: !!body.querySelector("img"), hasTable: !!body.querySelector("td,th") };
    var lines = [];
    var cur = [];
    function flush() {
      lines.push(cur);
      cur = [];
    }
    function styleOf(el, fmt) {
      var f = Object.assign({}, fmt);
      var tag = el.tagName.toUpperCase();
      if (tag === "B" || tag === "STRONG") f.b = true;
      if (tag === "I" || tag === "EM") f.i = true;
      if (tag === "U" || tag === "INS") f.u = true;
      if (tag === "S" || tag === "STRIKE" || tag === "DEL") f.s = true;
      if (tag === "SUP") f.sup = true;
      if (tag === "SUB") f.sub = true;
      if (tag === "A") {
        var href = el.getAttribute("href") || "";
        if (/^(https?:|mailto:|tel:)/i.test(href)) f.href = href;
      }
      var st = el.style;
      if (st) {
        var fw = (st.fontWeight || "").toLowerCase();
        if (fw === "bold" || fw === "bolder" || parseInt(fw, 10) >= 600) f.b = true;
        else if (fw === "normal" || fw === "400" || fw === "lighter") f.b = false;
        var fs = (st.fontStyle || "").toLowerCase();
        if (fs === "italic" || fs === "oblique") f.i = true;
        else if (fs === "normal") f.i = false;
        var td = ((st.textDecorationLine || "") + " " + (st.textDecoration || "")).toLowerCase();
        if (td.indexOf("underline") !== -1) f.u = true;
        if (td.indexOf("line-through") !== -1) f.s = true;
        var va = (st.verticalAlign || "").toLowerCase();
        if (va === "super") f.sup = true;
        if (va === "sub") f.sub = true;
      }
      return f;
    }
    function walk(node, fmt) {
      if (node.nodeType === 3) {
        var t = node.nodeValue.replace(/[\s ]+/g, " ");
        if (t) cur.push({ text: t, f: fmt });
        return;
      }
      if (node.nodeType !== 1) return;
      var tag = node.tagName.toUpperCase();
      if (SKIP_TAGS.test(tag)) return;
      if (tag === "BR") {
        flush();
        return;
      }
      var block = BLOCK_TAGS.test(tag);
      if (block && cur.length) flush();
      var f = styleOf(node, fmt);
      for (var c = node.firstChild; c; c = c.nextSibling) walk(c, f);
      if (block && cur.length) flush();
    }
    walk(body, {});
    if (cur.length) flush();
    // Tidy whitespace and drop empty lines.
    var out = [];
    lines.forEach(function (runs) {
      var tidy = [];
      runs.forEach(function (r) {
        var text = r.text;
        var prev = tidy[tidy.length - 1];
        if ((!prev || / $/.test(prev.text)) && text.charAt(0) === " ") text = text.slice(1);
        if (text) tidy.push({ text: text, f: r.f });
      });
      while (tidy.length && / $/.test(tidy[tidy.length - 1].text)) {
        var last = tidy[tidy.length - 1];
        last.text = last.text.replace(/ +$/, "");
        if (!last.text) tidy.pop();
      }
      if (tidy.length) out.push(tidy);
    });
    info.lines = out;
    return info;
  };

  LT.lineText = function (runs) {
    return runs.map(function (r) { return r.text; }).join("");
  };

  // Strips a typed list marker ("•", "-", "1.", "(a)", "(iv)") from the start of a line.
  var MARKER = /^(?:[•●◦▪■‣⁃·*–—-]|\(?\d{1,3}[.)]|\(?[A-Za-z][.)]|\([ivxlcIVXLC]{1,5}\)|[ivxlc]{1,4}[.)])\s+/;
  LT.stripMarker = function (runs) {
    var m = MARKER.exec(LT.lineText(runs));
    if (!m) return runs;
    var remove = m[0].length;
    var out = [];
    runs.forEach(function (r) {
      if (remove >= r.text.length) {
        remove -= r.text.length;
        return;
      }
      out.push({ text: r.text.slice(remove), f: r.f });
      remove = 0;
    });
    return out;
  };

  LT.renderRuns = function (runs, opts) {
    opts = opts || {};
    var merged = [];
    runs.forEach(function (r) {
      var f = opts.plain ? {} : r.f;
      var prev = merged[merged.length - 1];
      if (prev && JSON.stringify(prev.f) === JSON.stringify(f)) prev.text += r.text;
      else merged.push({ text: r.text, f: f });
    });
    return merged
      .map(function (r) {
        var h = LT.escape(r.text);
        var f = r.f;
        if (f.sub) h = "<sub>" + h + "</sub>";
        if (f.sup) h = "<sup>" + h + "</sup>";
        if (f.s) h = "<s>" + h + "</s>";
        if (f.u && !f.href) h = "<u>" + h + "</u>";
        if (f.i) h = "<i>" + h + "</i>";
        if (f.b) h = "<b>" + h + "</b>";
        if (f.href) h = '<a href="' + LT.escape(f.href) + '" style="color:#0563C1;text-decoration:underline;">' + h + "</a>";
        return h;
      })
      .join("");
  };

  function fontCss(style, sizePt, color) {
    var family = String(style.fontFamily || "Calibri, Arial, sans-serif").replace(/"/g, "'");
    return "font-family:" + family + ";font-size:" + sizePt + "pt;color:" + (color || style.textColor || "#000000") + ";";
  }

  // kind: h1 | h2 | body | bullets | numbers | letters | quote
  LT.buildBlock = function (kind, lines, style) {
    style = style || LT.getSettings().style;
    var body = fontCss(style, style.bodySizePt || 11);
    var parts = [];
    if (kind === "h1" || kind === "h2") {
      var size = kind === "h1" ? style.heading1SizePt || 14 : style.heading2SizePt || 12;
      var margin = kind === "h1" ? "14pt 0 6pt 0" : "10pt 0 4pt 0";
      lines.forEach(function (runs) {
        parts.push(
          '<p style="margin:' + margin + ";" + fontCss(style, size, style.headingColor) + 'font-weight:bold;">' +
            LT.renderRuns(runs, { plain: true }) + "</p>"
        );
      });
    } else if (kind === "body") {
      lines.forEach(function (runs) {
        parts.push('<p style="margin:0 0 10pt 0;' + body + '">' + LT.renderRuns(runs) + "</p>");
      });
    } else if (kind === "bullets" || kind === "numbers") {
      var tag = kind === "bullets" ? "ul" : "ol";
      var type = kind === "bullets" ? "list-style-type:disc;" : "list-style-type:decimal;";
      parts.push('<' + tag + ' style="margin:0 0 10pt 0;padding-left:28pt;' + type + body + '">');
      lines.forEach(function (runs) {
        parts.push('<li style="margin:0 0 4pt 0;' + body + '">' + LT.renderRuns(LT.stripMarker(runs)) + "</li>");
      });
      parts.push("</" + tag + ">");
    } else if (kind === "letters") {
      // A borderless two-column table keeps "(a)" hanging-indents aligned in every email program.
      parts.push('<table cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;margin:0 0 10pt 0;">');
      lines.forEach(function (runs, i) {
        parts.push(
          '<tr><td valign="top" style="width:24pt;padding:0 8pt 6pt 18pt;white-space:nowrap;' + body + '">(' + LT.letterFor(i) + ")</td>" +
            '<td valign="top" style="padding:0 0 6pt 0;' + body + '">' + LT.renderRuns(LT.stripMarker(runs)) + "</td></tr>"
        );
      });
      parts.push("</table>");
    } else if (kind === "quote") {
      lines.forEach(function (runs) {
        parts.push(
          '<p style="margin:0 36pt 8pt 24pt;padding-left:10pt;border-left:3px solid #A6A6A6;' +
            fontCss(style, style.bodySizePt || 11, "#404040") + '">' + LT.renderRuns(runs) + "</p>"
        );
      });
    }
    return parts.join("");
  };

  LT.letterFor = function (i) {
    var s = "";
    i += 1;
    while (i > 0) {
      var r = (i - 1) % 26;
      s = String.fromCharCode(97 + r) + s;
      i = Math.floor((i - 1) / 26);
    }
    return s;
  };

  // Plain text (e.g. a clause) to firm-styled body paragraphs.
  LT.textToBodyHtml = function (text, style) {
    var lines = String(text || "")
      .split(/\r?\n/)
      .map(function (l) { return l.trim(); })
      .filter(Boolean)
      .map(function (l) { return [{ text: l, f: {} }]; });
    return LT.buildBlock("body", lines, style);
  };

  LT.FORMAT_NAMES = {
    h1: "Heading",
    h2: "Subheading",
    body: "Body Text",
    bullets: "Bullet List",
    numbers: "Numbered List",
    letters: "(a) (b) (c) List",
    quote: "Quote Block"
  };

  // Formats whatever is selected in the email body.
  LT.formatSelection = async function (kind) {
    var item = LT.item();
    var type = await LT.call(item.body, "getTypeAsync");
    if (type !== Office.CoercionType.Html) {
      throw new Error("This email is in plain-text format. Switch it to HTML (Format menu > HTML) to use formatting.");
    }
    var sel = await LT.call(item, "getSelectedDataAsync", Office.CoercionType.Html);
    if (sel && sel.sourceProperty === "subject") throw new Error("Select text in the body of the email, not the subject line.");
    var parsed = LT.parseSelection((sel && sel.data) || "");
    if (!parsed.lines.length) {
      throw new Error("First select the text you want to format, then choose " + LT.FORMAT_NAMES[kind] + ".");
    }
    if (parsed.hasImage || parsed.hasTable) {
      throw new Error("Your selection includes a picture or table. Select only text to apply formatting.");
    }
    var html = LT.buildBlock(kind, parsed.lines);
    await LT.call(item.body, "setSelectedDataAsync", html, { coercionType: Office.CoercionType.Html });
  };

  /* ------------------------------------------------------------------ */
  /* Subject tags and markings                                           */
  /* ------------------------------------------------------------------ */

  var REPLY_PREFIX = /^((?:\s*(?:RE|FW|FWD|AW|WG|SV|VS|TR|RV|ANTW)\s*(?:\[\d+\])?\s*:)*)\s*/i;

  LT.addSubjectTag = function (subject, tag) {
    subject = String(subject || "");
    if (!tag || subject.toLowerCase().indexOf(tag.toLowerCase()) !== -1) return subject;
    var m = REPLY_PREFIX.exec(subject);
    var prefix = m ? m[1].trim() : "";
    var rest = subject.slice(m ? m[0].length : 0).trim();
    return (prefix ? prefix + " " : "") + tag + (rest ? " " + rest : "");
  };

  LT.removeSubjectTag = function (subject, tag) {
    subject = String(subject || "");
    if (!tag) return subject;
    var re = new RegExp("\\s*" + tag.replace(/[.*+?^${}()|[\]\\]/g, "\\$&") + "\\s*", "ig");
    return subject.replace(re, " ").replace(/\s{2,}/g, " ").trim();
  };

  LT.markingIds = function () {
    return Object.keys(CONFIG.markings || {});
  };

  LT.bannerHtml = function (marking) {
    var style = LT.getSettings().style;
    return (
      '<p style="margin:0 0 12pt 0;' + fontCss(style, Math.max(9, (style.bodySizePt || 11) - 1), CONFIG.markingColor || "#C00000") +
      'font-weight:bold;letter-spacing:0.3pt;">' + LT.escape(marking.banner) + "</p>"
    );
  };

  LT.applyMarking = async function (id) {
    var marking = (CONFIG.markings || {})[id];
    if (!marking) throw new Error("Unknown marking: " + id);
    var item = LT.item();
    var subject = await LT.call(item.subject, "getAsync");
    var tagged = LT.addSubjectTag(subject, marking.subjectTag);
    if (tagged !== subject) await LT.call(item.subject, "setAsync", tagged);

    var type = await LT.call(item.body, "getTypeAsync");
    var text = await LT.call(item.body, "getAsync", Office.CoercionType.Text);
    var already = LT.hasBannerLine(LT.ownText(text), marking.banner);
    if (!already) {
      if (type === Office.CoercionType.Html) {
        await LT.call(item.body, "prependAsync", LT.bannerHtml(marking), { coercionType: Office.CoercionType.Html });
      } else {
        await LT.call(item.body, "prependAsync", marking.banner + "\n\n", { coercionType: Office.CoercionType.Text });
      }
    }
    return { subjectChanged: tagged !== subject, bannerAdded: !already };
  };

  // Removes every marking banner from the user's part of the email and every marking tag from the subject.
  LT.removeMarkings = async function () {
    var item = LT.item();
    var markings = CONFIG.markings || {};
    var ids = Object.keys(markings);
    var subject = await LT.call(item.subject, "getAsync");
    var newSubject = subject;
    ids.forEach(function (id) { newSubject = LT.removeSubjectTag(newSubject, markings[id].subjectTag); });
    if (newSubject !== subject) await LT.call(item.subject, "setAsync", newSubject);

    var banners = ids.map(function (id) { return LT.normalizeText(markings[id].banner); });
    var type = await LT.call(item.body, "getTypeAsync");
    var removed = 0;
    if (type === Office.CoercionType.Html) {
      var html = await LT.call(item.body, "getAsync", Office.CoercionType.Html);
      var doc = new DOMParser().parseFromString(html, "text/html");
      var seen = {};
      // Only the first occurrence of each banner (ours, at the top); older ones in the quoted thread stay.
      var candidates = doc.body.querySelectorAll("p,div,span,b,strong,font,h1,h2,h3,h4,h5,h6,td");
      for (var i = 0; i < candidates.length; i++) {
        var el = candidates[i];
        if (!el.isConnected) continue;
        var t = LT.normalizeText(el.textContent);
        var idx = banners.indexOf(t);
        if (idx === -1 || seen[idx]) continue;
        // Remove the outermost element that contains exactly the banner text.
        var target = el;
        while (target.parentElement && target.parentElement !== doc.body &&
               LT.normalizeText(target.parentElement.textContent) === t) {
          target = target.parentElement;
        }
        target.parentNode.removeChild(target);
        seen[idx] = true;
        removed++;
      }
      if (removed) {
        await LT.call(item.body, "setAsync", "<!DOCTYPE html>" + doc.documentElement.outerHTML, { coercionType: Office.CoercionType.Html });
      }
    } else {
      var text = await LT.call(item.body, "getAsync", Office.CoercionType.Text);
      var own = LT.ownText(text);
      var rest = text.slice(own.length);
      var kept = own.split(/\r?\n/).filter(function (line) {
        var hit = banners.indexOf(LT.normalizeText(line)) !== -1;
        if (hit) removed++;
        return !hit;
      });
      if (removed) {
        await LT.call(item.body, "setAsync", kept.join("\n").replace(/^\s+/, "") + rest, { coercionType: Office.CoercionType.Text });
      }
    }
    return { removed: removed, subjectChanged: newSubject !== subject };
  };

  /* ------------------------------------------------------------------ */
  /* Office file inspection (tracked changes, comments, hidden content)  */
  /* ------------------------------------------------------------------ */

  LT.fileKind = function (name) {
    var ext = (LT.splitExt(name).ext || "").toLowerCase();
    if (/^\.(docx|docm|dotx|dotm)$/.test(ext)) return "word";
    if (/^\.(xlsx|xlsm|xltx|xltm)$/.test(ext)) return "excel";
    if (/^\.(pptx|pptm|potx|potm)$/.test(ext)) return "powerpoint";
    if (ext === ".pdf") return "pdf";
    if (/^\.(png|jpe?g|gif|bmp|tiff?|heic|webp)$/.test(ext)) return "image";
    if (/^\.(doc|xls|ppt|rtf)$/.test(ext)) return "legacy-office";
    if (/^\.(zip|7z|rar)$/.test(ext)) return "archive";
    if (/^\.(eml|msg)$/.test(ext)) return "email";
    return "file";
  };

  LT.base64ToBytes = function (b64) {
    var bin = atob(String(b64).replace(/\s+/g, ""));
    var bytes = new Uint8Array(bin.length);
    for (var i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
    return bytes;
  };

  function readZipDirectory(bytes) {
    var dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    var eocd = -1;
    for (var i = bytes.length - 22; i >= Math.max(0, bytes.length - 65557); i--) {
      if (dv.getUint32(i, true) === 0x06054b50) {
        eocd = i;
        break;
      }
    }
    if (eocd < 0) throw new Error("not a valid Office file");
    var count = dv.getUint16(eocd + 10, true);
    var p = dv.getUint32(eocd + 16, true);
    var entries = {};
    var decoder = new TextDecoder();
    for (var n = 0; n < count && p + 46 <= bytes.length; n++) {
      if (dv.getUint32(p, true) !== 0x02014b50) break;
      var nameLen = dv.getUint16(p + 28, true);
      var extraLen = dv.getUint16(p + 30, true);
      var commentLen = dv.getUint16(p + 32, true);
      var name = decoder.decode(bytes.subarray(p + 46, p + 46 + nameLen));
      entries[name] = {
        method: dv.getUint16(p + 10, true),
        compressedSize: dv.getUint32(p + 20, true),
        offset: dv.getUint32(p + 42, true)
      };
      p += 46 + nameLen + extraLen + commentLen;
    }
    return entries;
  }

  async function readZipEntry(bytes, entry) {
    var dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    var p = entry.offset;
    if (dv.getUint32(p, true) !== 0x04034b50) throw new Error("damaged Office file");
    var start = p + 30 + dv.getUint16(p + 26, true) + dv.getUint16(p + 28, true);
    var data = bytes.subarray(start, start + entry.compressedSize);
    if (entry.method === 0) return new TextDecoder().decode(data);
    if (entry.method !== 8) throw new Error("unsupported compression");
    if (typeof DecompressionStream === "undefined") throw new Error("this version of macOS can't inspect Office files");
    var stream = new Blob([data]).stream().pipeThrough(new DecompressionStream("deflate-raw"));
    return new TextDecoder().decode(await new Response(stream).arrayBuffer());
  }

  function count(re, s) {
    var m = String(s || "").match(re);
    return m ? m.length : 0;
  }

  // Returns { findings: [string], clean: bool } for .docx/.xlsx/.pptx content.
  LT.inspectOfficeFile = async function (name, base64) {
    var kind = LT.fileKind(name);
    var bytes = LT.base64ToBytes(base64);
    var entries = readZipDirectory(bytes);
    var names = Object.keys(entries);
    var findings = [];
    async function read(n) {
      return entries[n] ? readZipEntry(bytes, entries[n]) : "";
    }
    function plural(n, word) {
      return n + " " + word + (n === 1 ? "" : "s");
    }

    if (kind === "word") {
      var parts = names.filter(function (n) {
        return /^word\/(document|header\d*|footer\d*|footnotes|endnotes)\.xml$/.test(n);
      });
      var revisions = 0, formatting = 0, hidden = 0;
      for (var i = 0; i < parts.length; i++) {
        var xml = await read(parts[i]);
        revisions += count(/<w:(ins|del|moveFrom|moveTo)[\s>]/g, xml);
        formatting += count(/<w:(rPrChange|pPrChange|sectPrChange|tblPrChange|trPrChange|tcPrChange)[\s>]/g, xml);
        hidden += count(/<w:vanish(\s+w:val="(1|true|on)")?\s*\/>/g, xml);
      }
      var comments = count(/<w:comment[\s>]/g, await read("word/comments.xml"));
      if (revisions) findings.push(plural(revisions, "tracked change") + " not accepted");
      if (formatting) findings.push(plural(formatting, "tracked formatting change"));
      if (comments) findings.push(plural(comments, "comment"));
      if (hidden) findings.push("hidden text (" + plural(hidden, "passage") + ")");
      if (entries["word/vbaProject.bin"]) findings.push("contains macros");
    } else if (kind === "excel") {
      var xlComments = 0;
      for (var j = 0; j < names.length; j++) {
        if (/^xl\/comments\d*\.xml$/.test(names[j])) xlComments += count(/<comment[\s>]/g, await read(names[j]));
        if (/^xl\/threadedComments\/.+\.xml$/.test(names[j])) xlComments += count(/<threadedComment[\s>]/g, await read(names[j]));
      }
      var workbook = await read("xl/workbook.xml");
      var hiddenSheets = count(/<sheet\b[^>]*state="(hidden|veryHidden)"/g, workbook);
      if (xlComments) findings.push(plural(xlComments, "comment"));
      if (hiddenSheets) findings.push(plural(hiddenSheets, "hidden sheet"));
      if (entries["xl/vbaProject.bin"]) findings.push("contains macros");
    } else if (kind === "powerpoint") {
      var pptComments = 0, notes = 0;
      for (var k = 0; k < names.length; k++) {
        if (/^ppt\/comments\/.+\.xml$/.test(names[k])) pptComments += count(/<(p:cm|p188:cm)[\s>]/g, await read(names[k]));
        if (/^ppt\/notesSlides\/notesSlide\d+\.xml$/.test(names[k])) {
          var noteXml = await read(names[k]);
          // Count slides whose notes contain real text (not just the slide number placeholder).
          var texts = (noteXml.match(/<a:t>([^<]*)<\/a:t>/g) || []).map(function (t) { return t.replace(/<\/?a:t>/g, ""); });
          if (texts.join("").replace(/\d+/g, "").trim().length > 0) notes++;
        }
      }
      if (pptComments) findings.push(plural(pptComments, "comment"));
      if (notes) findings.push("speaker notes on " + plural(notes, "slide"));
      if (entries["ppt/vbaProject.bin"]) findings.push("contains macros");
    }

    var core = await read("docProps/core.xml");
    var author = (/<dc:creator>([^<]*)</.exec(core) || [])[1] || "";
    var lastBy = (/<cp:lastModifiedBy>([^<]*)</.exec(core) || [])[1] || "";
    return { findings: findings, clean: findings.length === 0, author: author, lastModifiedBy: lastBy };
  };
})(window);
