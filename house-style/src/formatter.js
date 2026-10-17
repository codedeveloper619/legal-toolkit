/*
 * Formatting engine. Pure HTML-in / HTML-out functions with no Office.js
 * dependency, so they can be unit-tested and reused in the browser preview.
 *
 *   var fmt = createFormatter(HOUSE_STYLE, window);
 *   fmt.applyStyle("alpha-1", selectedHtml)    → HTML to put back
 *   fmt.applyColour("#A41D30", selectedHtml)   → HTML to put back
 *   fmt.tidyBody(bodyHtml, { colours: true })  → { html, changed, quoteFound }
 */
(function (root) {
  "use strict";

  var BLOCK_TAGS = /^(P|DIV|H[1-6]|LI|BLOCKQUOTE|PRE|ADDRESS|CENTER|DD|DT|SECTION|ARTICLE|HEADER|FOOTER)$/;
  var KEEP_TAGS = /^(UL|OL|TABLE)$/;
  var LIST_TAGS = /^(UL|OL)$/;
  var DEFAULTED_TAGS = /^(P|DIV|LI|TD|TH|H[1-6]|BLOCKQUOTE|PRE|DD|DT|ADDRESS)$/;
  var SKIP_TAGS = /^(HEAD|STYLE|SCRIPT|META|TITLE|LINK|BASE|XML|TEMPLATE)$/;
  var TABLE_TAGS = /^(TABLE|THEAD|TBODY|TFOOT|TR|TD|TH|COLGROUP|COL)$/;
  var COLOUR_PROPS = /^(color|background|background-color|border(-(top|right|bottom|left))?(-color)?)$/;
  var KEPT_ATTRS = /^(href|src|alt|title|width|height|colspan|rowspan|target|style|type|start|align|valign|cellpadding|cellspacing|border|bgcolor|role)$/;
  var CHARACTER_PROPS = /^(font-weight|font-style|text-decoration|text-decoration-line|text-underline|vertical-align)$/;
  var NAMED_COLOURS = {
    black: "#000000", white: "#ffffff", red: "#ff0000", blue: "#0000ff",
    green: "#008000", navy: "#000080", gray: "#808080", grey: "#808080",
    yellow: "#ffff00", maroon: "#800000", purple: "#800080", silver: "#c0c0c0",
  };
  // Labels this add-in writes: (a) (aa) (iv) (1) and bullets.
  var STRICT_LABEL = /^[\s ]*(?:\((?:([a-z])\1{0,3}|[ivxlcdm]{1,7}|\d{1,3})\)|[•·▪◦‣])[\s \t]+/i;
  // Also typed or pasted labels such as 1.  a)  iv.  -  *  (used for list styles only).
  var LOOSE_LABEL = /^[\s ]*(?:\((?:([a-z])\1{0,3}|[ivxlcdm]{1,7}|\d{1,3})\)|[•·▪◦‣*\-–]|\d{1,3}[.)]|[a-z][.)]|[ivx]{1,5}[.)])[\s \t]+/i;
  var ZWSP = "​";
  var QUOTE_HEADER =/^\s*(From|Von|De|Da|Van|Från|Fra|Od|Lähettäjä|Sent|Expéditeur|Remitente|من)\s*:/i;
  var QUOTE_SELECTORS = [
    "#appendonsend",
    "[id^='divRplyFwdMsg']",
    "#mail-editor-reference-message-container",
    "#OLK_SRC_BODY_SECTION",
    "div.OutlookMessageHeader",
    ".gmail_quote",
    "blockquote[type='cite']",
  ];

  function createFormatter(style, win) {
    var parser = new win.DOMParser();
    var ELEMENT = 1, TEXT = 3, COMMENT = 8;
    var L = style.layout;

    // ── Small helpers ────────────────────────────────────────────────────────
    function esc(s) {
      return String(s == null ? "" : s)
        .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;").replace(/'/g, "&#39;");
    }

    function assign(target) {
      for (var i = 1; i < arguments.length; i++) {
        var src = arguments[i] || {};
        Object.keys(src).forEach(function (k) { target[k] = src[k]; });
      }
      return target;
    }

    function pt(n) { return n ? +Number(n).toFixed(2) + "pt" : "0"; }

    function resolve(prop, value) {
      value = String(value);
      if (prop === "font-family" && style.fonts[value]) return style.fonts[value];
      if (COLOUR_PROPS.test(prop)) {
        return value.split(/\s+/).map(function (t) {
          return style.colours[t] || t;
        }).join(" ");
      }
      return value;
    }

    // css({...}, {...}) → "prop:value;prop:value" with tokens resolved.
    function css() {
      var out = [];
      var seen = {};
      for (var i = 0; i < arguments.length; i++) {
        var obj = arguments[i] || {};
        Object.keys(obj).forEach(function (k) {
          var v = resolve(k, obj[k]);
          if (seen[k] != null) out[seen[k]] = k + ":" + v;
          else { seen[k] = out.length; out.push(k + ":" + v); }
        });
      }
      return out.join(";");
    }

    function parseStyle(str) {
      var decls = [];
      var buf = "", depth = 0, quote = null;
      str = str || "";
      for (var i = 0; i <= str.length; i++) {
        var c = str.charAt(i);
        if (i === str.length || (c === ";" && depth === 0 && !quote)) {
          var idx = buf.indexOf(":");
          if (idx > 0) {
            var prop = buf.slice(0, idx).trim().toLowerCase();
            var val = buf.slice(idx + 1).trim();
            if (prop && val) decls.push([prop, val]);
          }
          buf = "";
          continue;
        }
        if (quote) { if (c === quote) quote = null; }
        else if (c === "'" || c === '"') quote = c;
        else if (c === "(") depth++;
        else if (c === ")") depth = Math.max(0, depth - 1);
        buf += c;
      }
      return decls;
    }

    function getDecl(decls, prop) {
      for (var i = decls.length - 1; i >= 0; i--) if (decls[i][0] === prop) return decls[i][1];
      return null;
    }
    function setDecl(decls, prop, value) {
      var found = false;
      for (var i = 0; i < decls.length; i++) {
        if (decls[i][0] === prop) {
          if (!found) { decls[i][1] = value; found = true; } else { decls.splice(i, 1); i--; }
        }
      }
      if (!found) decls.push([prop, value]);
    }
    function removeDecl(decls, test) {
      for (var i = decls.length - 1; i >= 0; i--) if (test(decls[i][0], decls[i][1])) decls.splice(i, 1);
    }
    function writeStyle(el, decls) {
      var s = decls.map(function (d) { return d[0] + ":" + d[1]; }).join(";");
      if (s) el.setAttribute("style", s); else el.removeAttribute("style");
    }

    function toPt(v) {
      if (!v) return null;
      v = String(v).trim().toLowerCase().replace(/\s*!important$/, "");
      var m = /^(\d*\.?\d+)\s*(pt|px)?$/.exec(v);
      if (m) return m[2] === "px" ? parseFloat(m[1]) * 0.75 : parseFloat(m[1]);
      var kw = { "xx-small": 7, "x-small": 7.5, small: 10, medium: 12, large: 13.5, "x-large": 18, "xx-large": 24 };
      return kw[v] != null ? kw[v] : null;
    }

    function normColour(v) {
      if (!v) return null;
      v = String(v).trim().toLowerCase().replace(/\s*!important$/, "");
      if (/^#[0-9a-f]{3}$/.test(v)) return "#" + v[1] + v[1] + v[2] + v[2] + v[3] + v[3];
      if (/^#[0-9a-f]{6}$/.test(v)) return v;
      var m = /^rgba?\(\s*(\d+)[\s,]+(\d+)[\s,]+(\d+)/.exec(v);
      if (m) {
        return "#" + [m[1], m[2], m[3]].map(function (n) {
          return ("0" + Math.min(255, parseInt(n, 10)).toString(16)).slice(-2);
        }).join("");
      }
      return NAMED_COLOURS[v] || null;
    }

    function colourIn(value) {
      var tokens = String(value).match(/#[0-9a-f]{3,6}\b|rgba?\([^)]*\)|[a-z-]+/gi) || [];
      for (var i = 0; i < tokens.length; i++) {
        var c = normColour(tokens[i]);
        if (c) return c;
      }
      return null;
    }

    function firstFamily(f) {
      return String(f || "").split(",")[0].replace(/["']/g, "").trim().toLowerCase();
    }

    // ── Theme colours (Word's tint and shade rules) ──────────────────────────
    function hexToRgb(hex) {
      var h = normColour(hex).slice(1);
      return [0, 2, 4].map(function (i) { return parseInt(h.slice(i, i + 2), 16) / 255; });
    }
    function rgbToHex(rgb) {
      return "#" + rgb.map(function (v) {
        var n = Math.max(0, Math.min(255, Math.round(v * 255)));
        return ("0" + n.toString(16)).slice(-2);
      }).join("").toUpperCase();
    }
    function rgbToHls(rgb) {
      var r = rgb[0], g = rgb[1], b = rgb[2];
      var max = Math.max(r, g, b), min = Math.min(r, g, b), l = (max + min) / 2;
      if (max === min) return [0, l, 0];
      var range = max - min;
      var s = l <= 0.5 ? range / (max + min) : range / (2 - max - min);
      var rc = (max - r) / range, gc = (max - g) / range, bc = (max - b) / range;
      var h = r === max ? bc - gc : g === max ? 2 + rc - bc : 4 + gc - rc;
      h = ((h / 6) % 1 + 1) % 1;
      return [h, l, s];
    }
    function hlsToRgb(h, l, s) {
      if (s === 0) return [l, l, l];
      var m2 = l <= 0.5 ? l * (1 + s) : l + s - l * s;
      var m1 = 2 * l - m2;
      function v(hue) {
        hue = ((hue % 1) + 1) % 1;
        if (hue < 1 / 6) return m1 + (m2 - m1) * hue * 6;
        if (hue < 0.5) return m2;
        if (hue < 2 / 3) return m1 + (m2 - m1) * (2 / 3 - hue) * 6;
        return m1;
      }
      return [v(h + 1 / 3), v(h), v(h - 1 / 3)];
    }
    function shadesFor(hex) {
      var hls = rgbToHls(hexToRgb(hex));
      var l = hls[1];
      var steps;
      if (l >= 0.999) steps = [["D", 5], ["D", 15], ["D", 25], ["D", 35], ["D", 50]];
      else if (l <= 0.001) steps = [["L", 50], ["L", 35], ["L", 25], ["L", 15], ["L", 5]];
      else if (l < 0.2) steps = [["L", 90], ["L", 75], ["L", 50], ["L", 25], ["L", 10]];
      else if (l > 0.8) steps = [["D", 10], ["D", 25], ["D", 50], ["D", 75], ["D", 90]];
      else steps = [["L", 80], ["L", 60], ["L", 40], ["D", 25], ["D", 50]];
      return steps.map(function (st) {
        var f = st[1] / 100;
        var nl = st[0] === "L" ? l + (1 - l) * f : l * (1 - f);
        return { hex: rgbToHex(hlsToRgb(hls[0], nl, hls[2])), label: (st[0] === "L" ? "Lighter " : "Darker ") + st[1] + "%" };
      });
    }
    var themeGrid = (style.theme || []).map(function (t) {
      return { name: t.name, slot: t.slot, hex: rgbToHex(hexToRgb(t.hex)), shades: shadesFor(t.hex) };
    });

    // ── Paragraph styles (from the config's rows × levels) ──────────────────
    var styleList = [];
    style.styles.forEach(function (row) {
      row.names.slice(0, L.levels).forEach(function (name, level) {
        var cssObj, before = 0, after = style.text.spaceAfter, right = 0, hang = 0;
        var left = level * L.indentStep;
        if (row.kind === "head") {
          var h = style.head.levels[Math.min(level, style.head.levels.length - 1)];
          cssObj = assign({}, style.head.css, { "font-size": h.size });
          before = h.spaceBefore || 0;
          after = h.spaceAfter;
        } else {
          cssObj = assign({}, style.text.css);
          if (row.label) { left = (level + 1) * L.indentStep; hang = L.labelWidth; }
          if (row.kind === "quote") { left = (level + 1) * L.indentStep; right = L.quoteRightIndent; }
        }
        cssObj.margin = [pt(before), pt(right), pt(after), pt(left)].join(" ");
        if (hang) cssObj["text-indent"] = "-" + pt(hang);
        // Browsers collapse margins between paragraphs; Outlook for Windows
        // adds them. The padding gives browsers the same gap above a heading.
        if (before) cssObj["padding-top"] = pt(before);
        styleList.push({
          id: row.kind + "-" + (level + 1),
          name: name,
          kind: row.kind,
          level: level,
          label: row.label || null,
          css: cssObj,
        });
      });
    });

    function findStyle(id) {
      for (var i = 0; i < styleList.length; i++) if (styleList[i].id === id) return styleList[i];
      throw new Error("Unknown house style: " + id);
    }

    var bodyStyle = findStyle("body-1");
    var bodyFont = resolve("font-family", style.text.css["font-family"]);
    var bodySize = style.text.css["font-size"];
    var bodyColour = resolve("color", style.text.css.color);
    var linkColour = style.colours.link || bodyColour;

    // ── What counts as "house style" (derived from the config) ───────────────
    var cssBlocks = styleList.map(function (s) { return s.css; });
    cssBlocks.push(style.text.css, style.smallPrint.css, style.banner.css, style.banner.noteCss,
      style.reference.css, style.reference.labelCss);

    var allowedFamilies = {};
    Object.keys(style.fonts).forEach(function (k) { allowedFamilies[firstFamily(style.fonts[k])] = true; });

    // Each house size, and whether it only appears in bold (e.g. 12pt = Head).
    var allowedSizes = [];
    cssBlocks.forEach(function (block) {
      var p = toPt(block && block["font-size"]);
      if (p == null) return;
      var bold = /^(bold|[6-9]00)$/.test(String(block["font-weight"] || ""));
      var existing = allowedSizes.filter(function (a) { return a.pt === p; })[0];
      if (existing) existing.boldOnly = existing.boldOnly && bold;
      else allowedSizes.push({ pt: p, boldOnly: bold });
    });

    var palette = { "#ffffff": true };
    Object.keys(style.colours).forEach(function (k) {
      var c = normColour(style.colours[k]);
      if (c) palette[c] = true;
    });
    themeGrid.forEach(function (t) {
      palette[t.hex.toLowerCase()] = true;
      t.shades.forEach(function (s) { palette[s.hex.toLowerCase()] = true; });
    });

    function sizeAllowed(v, el) {
      var p = toPt(v);
      if (p == null) return false;
      return allowedSizes.some(function (a) {
        return Math.abs(a.pt - p) < 0.3 && (!a.boldOnly || isBold(el));
      });
    }

    // Bold set on the element or inherited (<b>, <strong>, font-weight).
    function isBold(el) {
      for (var n = el; n && n.nodeType === ELEMENT; n = n.parentNode) {
        var tag = tagOf(n);
        if (tag === "BODY") return false;
        var w = getDecl(parseStyle(n.getAttribute("style")), "font-weight");
        if (w != null) return /^(bold|bolder|[6-9]00)/i.test(w);
        if (tag === "B" || tag === "STRONG" || /^H[1-6]$/.test(tag)) return true;
      }
      return false;
    }
    function familyAllowed(v) { return !!allowedFamilies[firstFamily(v)]; }
    function colourAllowed(v) {
      var c = normColour(v);
      return !!(c && palette[c]);
    }
    function backgroundAllowed(v) {
      var t = String(v).trim().toLowerCase();
      if (/^(transparent|none|initial|inherit|unset)$/.test(t)) return true;
      var c = colourIn(t);
      return !c || !!palette[c];
    }

    // ── DOM helpers ─────────────────────────────────────────────────────────
    function parseFragment(html) {
      var doc = parser.parseFromString("<!DOCTYPE html><html><body>" + (html || "") + "</body></html>", "text/html");
      return doc.body;
    }

    function parseDocument(html) {
      var isDocument = /<html[\s>]/i.test(html || "");
      var doc = parser.parseFromString(isDocument ? html : "<!DOCTYPE html><html><body>" + (html || "") + "</body></html>", "text/html");
      return {
        doc: doc,
        serialize: function () {
          var out = isDocument ? doc.documentElement.outerHTML : doc.body.innerHTML;
          if (isDocument && /^\s*<!doctype/i.test(html)) out = "<!DOCTYPE html>" + out;
          return out;
        },
      };
    }

    function unwrap(el) {
      var parent = el.parentNode;
      if (!parent) return;
      while (el.firstChild) parent.insertBefore(el.firstChild, el);
      parent.removeChild(el);
    }

    function elementsIn(node) {
      return Array.prototype.slice.call(node.getElementsByTagName("*"));
    }

    function tagOf(el) { return el.tagName.toUpperCase(); }

    function hasBlockChild(el) {
      for (var c = el.firstChild; c; c = c.nextSibling) {
        if (c.nodeType === ELEMENT && (BLOCK_TAGS.test(tagOf(c)) || KEEP_TAGS.test(tagOf(c)))) return true;
      }
      return false;
    }

    function isBlank(node) {
      if (node.querySelector && node.querySelector("img")) return false;
      return !String(node.textContent || "").replace(/[\s ​]/g, "");
    }

    // Split a selection into paragraphs. Lists and tables are kept whole.
    function collectBlocks(node) {
      var blocks = [];
      var run = [];
      function flush() {
        if (run.length) blocks.push({ type: "inline", nodes: run });
        run = [];
      }
      Array.prototype.slice.call(node.childNodes).forEach(function (child) {
        if (child.nodeType === COMMENT) return;
        if (child.nodeType === ELEMENT) {
          var tag = tagOf(child);
          if (SKIP_TAGS.test(tag)) return;
          if (tag === "BR") { flush(); return; }
          if (KEEP_TAGS.test(tag)) { flush(); blocks.push({ type: "keep", node: child }); return; }
          if (BLOCK_TAGS.test(tag)) {
            flush();
            if (hasBlockChild(child)) blocks = blocks.concat(collectBlocks(child));
            else blocks.push({ type: "block", node: child });
            return;
          }
        }
        run.push(child);
      });
      flush();
      return blocks;
    }

    function containerOf(nodes, doc) {
      var div = doc.createElement("div");
      nodes.forEach(function (n) { div.appendChild(n.cloneNode(true)); });
      return div;
    }

    function blockContainer(block, doc) {
      return containerOf(block.type === "inline" ? block.nodes : Array.prototype.slice.call(block.node.childNodes), doc);
    }

    // List items (nested lists flattened), each as a container of inline content.
    function listItems(listEl, doc) {
      var out = [];
      Array.prototype.forEach.call(listEl.children, function (child) {
        var tag = tagOf(child);
        if (LIST_TAGS.test(tag)) { out = out.concat(listItems(child, doc)); return; }
        if (tag !== "LI") return;
        var own = Array.prototype.filter.call(child.childNodes, function (n) {
          return !(n.nodeType === ELEMENT && LIST_TAGS.test(tagOf(n)));
        });
        var c = containerOf(own, doc);
        flattenBlocks(c);
        out.push(c);
        Array.prototype.forEach.call(child.children, function (n) {
          if (LIST_TAGS.test(tagOf(n))) out = out.concat(listItems(n, doc));
        });
      });
      return out;
    }

    // Replace paragraph wrappers inside a container with line breaks.
    function flattenBlocks(container) {
      elementsIn(container).filter(function (e) { return BLOCK_TAGS.test(tagOf(e)); }).forEach(function (el) {
        if (el.previousSibling && !isBlank(el.previousSibling)) el.parentNode.insertBefore(container.ownerDocument.createElement("br"), el);
        unwrap(el);
      });
    }

    function hasBlocks(container) {
      for (var c = container.firstChild; c; c = c.nextSibling) {
        if (c.nodeType !== ELEMENT) continue;
        var tag = tagOf(c);
        if (tag === "BR" || BLOCK_TAGS.test(tag) || KEEP_TAGS.test(tag)) return true;
      }
      return false;
    }

    // Strip pasted fonts, sizes, colours, spacing and Word classes, keeping
    // bold, italics, underline, links, images and line breaks.
    function cleanInline(container) {
      elementsIn(container).forEach(function (el) {
        var tag = tagOf(el);
        if (tag === "O:P") { if (isBlank(el)) el.parentNode.removeChild(el); else unwrap(el); return; }
        if (SKIP_TAGS.test(tag)) { el.parentNode.removeChild(el); return; }
        if (tag === "FONT") { unwrap(el); return; }
        Array.prototype.slice.call(el.attributes).forEach(function (a) {
          if (!KEPT_ATTRS.test(a.name.toLowerCase())) el.removeAttribute(a.name);
        });
        var decls = parseStyle(el.getAttribute("style"));
        removeDecl(decls, function (p) { return !CHARACTER_PROPS.test(p); });
        if (tag === "A") {
          setDecl(decls, "color", linkColour);
          setDecl(decls, "text-decoration", "underline");
        }
        writeStyle(el, decls);
        if (tag === "SPAN" && !el.attributes.length) unwrap(el);
      });
      var walker = container.ownerDocument.createTreeWalker(container, 128 /* SHOW_COMMENT */, null, false);
      var comments = [];
      while (walker.nextNode()) comments.push(walker.currentNode);
      comments.forEach(function (c) { c.parentNode.removeChild(c); });
      // Drop the invisible placeholders left by "start typing" paragraphs.
      var texts = container.ownerDocument.createTreeWalker(container, 4 /* SHOW_TEXT */, null, false);
      while (texts.nextNode()) {
        if (texts.currentNode.nodeValue.indexOf(ZWSP) >= 0) texts.currentNode.nodeValue = texts.currentNode.nodeValue.split(ZWSP).join("");
      }
      return container;
    }

    // Remove the first `count` characters of text, then any wrappers left empty.
    function dropLeadingText(container, count) {
      var walker = container.ownerDocument.createTreeWalker(container, 4 /* SHOW_TEXT */, null, false);
      var left = count;
      while (left > 0 && walker.nextNode()) {
        var t = walker.currentNode;
        var v = t.nodeValue;
        if (v.length <= left) { left -= v.length; t.nodeValue = ""; }
        else { t.nodeValue = v.slice(left); left = 0; }
      }
      (function prune(node) {
        while (node.firstChild) {
          var c = node.firstChild;
          if (c.nodeType === TEXT && !c.nodeValue) { node.removeChild(c); continue; }
          if (c.nodeType === COMMENT) { node.removeChild(c); continue; }
          if (c.nodeType === ELEMENT && !/^(BR|IMG)$/.test(tagOf(c)) && isBlank(c)) { node.removeChild(c); continue; }
          if (c.nodeType === ELEMENT) prune(c);
          return;
        }
      })(container);
    }

    function stripLeadingLabel(container, loose) {
      var m = (loose ? LOOSE_LABEL : STRICT_LABEL).exec(container.textContent || "");
      if (!m) return false;
      dropLeadingText(container, m[0].length);
      return true;
    }

    function textWithBreaks(node) {
      var clone = node.cloneNode(true);
      elementsIn(clone).forEach(function (el) {
        if (tagOf(el) === "BR") el.parentNode.replaceChild(clone.ownerDocument.createTextNode("\n"), el);
      });
      return String(clone.textContent || "").replace(/ /g, " ").replace(/[ \t]+\n/g, "\n").trim();
    }

    function restyleTable(node) {
      var el = node.cloneNode(true);
      cleanInline(el);
      var props = parseStyle(css({
        "font-family": style.text.css["font-family"],
        "font-size": style.text.css["font-size"],
        color: style.text.css.color,
      }));
      elementsIn(el).filter(function (e) { return /^(TD|TH)$/.test(tagOf(e)); }).forEach(function (t) {
        var decls = parseStyle(t.getAttribute("style"));
        props.forEach(function (d) { setDecl(decls, d[0], d[1]); });
        writeStyle(t, decls);
      });
      return el.outerHTML;
    }

    // ── Numbering ───────────────────────────────────────────────────────────
    function alpha(n) {
      var letter = String.fromCharCode(97 + ((n - 1) % 26));
      return new Array(Math.floor((n - 1) / 26) + 2).join(letter);
    }
    function roman(n) {
      var map = [[1000, "m"], [900, "cm"], [500, "d"], [400, "cd"], [100, "c"], [90, "xc"],
        [50, "l"], [40, "xl"], [10, "x"], [9, "ix"], [5, "v"], [4, "iv"], [1, "i"]];
      var out = "";
      map.forEach(function (p) { while (n >= p[0]) { out += p[1]; n -= p[0]; } });
      return out;
    }
    function labelText(template, n) {
      return template.replace("{a}", alpha(n)).replace("{i}", roman(n)).replace("{1}", String(n));
    }
    // The label sits in the hanging indent. Browsers use the inline-block
    // width; Outlook for Windows (Word) ignores it and uses the real tab.
    function labelHtml(s, n) {
      return '<span style="display:inline-block;width:' + pt(L.labelWidth) + ';text-indent:0">' +
        esc(labelText(s.label, n)) + '<span style="mso-tab-count:1"> </span></span>';
    }

    // ── Public: paragraph styles ────────────────────────────────────────────
    // Each paragraph in the selection gets the style. Numbered styles label
    // the selected paragraphs (a), (b), (c) … in order, replacing old labels.
    function applyStyle(id, html) {
      return applyStyleInfo(id, html).html;
    }

    // Same as applyStyle, also reporting how many paragraphs were styled.
    function applyStyleInfo(id, html) {
      var s = findStyle(id);
      var root = parseFragment(html);
      var doc = root.ownerDocument;
      var parts = [];
      collectBlocks(root).forEach(function (block) {
        if (block.type === "keep") {
          if (tagOf(block.node) === "TABLE") { parts.push({ table: restyleTable(block.node) }); return; }
          listItems(block.node, doc).forEach(function (c) { parts.push({ para: c }); });
          return;
        }
        parts.push({ para: blockContainer(block, doc) });
      });
      var n = 0;
      var out = [];
      var styleAttr = esc(css(s.css));
      parts.forEach(function (p) {
        if (p.table) { out.push(p.table); return; }
        var c = p.para;
        stripLeadingLabel(c, !!s.label);
        cleanInline(c);
        if (isBlank(c)) return;
        n++;
        out.push('<p style="' + styleAttr + '">' + (s.label ? labelHtml(s, n) : "") + c.innerHTML.trim() + "</p>");
      });
      return { html: out.join(""), count: n };
    }

    // A new, empty paragraph in a style, inserted at the cursor so the user
    // can type straight into it. `n` is the list number to show, e.g. 2 → (b).
    // The invisible placeholder gives the cursor somewhere to sit and carries
    // the font, so typed text picks up the style.
    function typingParagraph(id, n) {
      var s = findStyle(id);
      var run = {};
      ["font-family", "font-size", "font-weight", "color"].forEach(function (k) {
        if (s.css[k] != null) run[k] = s.css[k];
      });
      return '<p style="' + esc(css(s.css)) + '">' + (s.label ? labelHtml(s, n || 1) : "") +
        '<span style="' + esc(css(run)) + '">' + ZWSP + "</span></p>";
    }

    // Colour for text the user is about to type.
    function typingColour(hex) {
      var colour = hex ? (normColour(hex) || bodyColour).toUpperCase() : bodyColour;
      return '<span style="color:' + colour + '">' + ZWSP + "</span>";
    }

    // Plain body text. Inline selections stay inline; paragraphs stay paragraphs.
    function clearFormatting(html) {
      var root = parseFragment(html);
      var doc = root.ownerDocument;
      var bodyCss = esc(css(bodyStyle.css));
      if (!hasBlocks(root)) {
        var c0 = containerOf(Array.prototype.slice.call(root.childNodes), doc);
        stripLeadingLabel(c0, false);
        var text = textWithBreaks(c0);
        if (!text) return "";
        var inline = css({ "font-family": bodyFont, "font-size": bodySize, color: bodyColour, "font-weight": "normal", "font-style": "normal" });
        return '<span style="' + esc(inline) + '">' + esc(text) + "</span>";
      }
      var out = [];
      function push(container) {
        stripLeadingLabel(container, false);
        var t = textWithBreaks(container);
        if (t) out.push('<p style="' + bodyCss + '">' + esc(t).replace(/\n/g, "<br>") + "</p>");
      }
      collectBlocks(root).forEach(function (block) {
        if (block.type === "keep" && LIST_TAGS.test(tagOf(block.node))) {
          listItems(block.node, doc).forEach(push);
          return;
        }
        push(block.type === "keep" ? containerOf([block.node], doc) : blockContainer(block, doc));
      });
      return out.join("");
    }

    // ── Public: text colour ─────────────────────────────────────────────────
    // Colours every piece of text in the selection. Pass null for Automatic.
    function applyColour(hex, html) {
      var colour = hex ? (normColour(hex) || bodyColour).toUpperCase() : bodyColour;
      var root = parseFragment(html);
      var doc = root.ownerDocument;
      // Inline colours are replaced; a paragraph's own colour can stay because
      // the new colour is set on the text inside it.
      elementsIn(root).forEach(function (el) {
        if (tagOf(el) === "FONT") el.removeAttribute("color");
        if (DEFAULTED_TAGS.test(tagOf(el))) return;
        var decls = parseStyle(el.getAttribute("style"));
        var n = decls.length;
        removeDecl(decls, function (p) { return p === "color"; });
        if (decls.length !== n) writeStyle(el, decls);
        if (/^(SPAN|FONT)$/.test(tagOf(el)) && !el.attributes.length) unwrap(el);
      });
      var walker = doc.createTreeWalker(root, 4, null, false);
      var texts = [];
      while (walker.nextNode()) {
        if (walker.currentNode.nodeValue.replace(/[\s ]/g, "")) texts.push(walker.currentNode);
      }
      texts.forEach(function (t) {
        var span = doc.createElement("span");
        span.setAttribute("style", "color:" + colour);
        t.parentNode.replaceChild(span, t);
        span.appendChild(t);
      });
      return root.innerHTML;
    }

    // ── Public: whole-message tidy ──────────────────────────────────────────
    function findQuoteBoundary(body) {
      var candidates = [];
      QUOTE_SELECTORS.forEach(function (sel) {
        var el = body.querySelector(sel);
        if (el) candidates.push(el);
      });
      elementsIn(body).forEach(function (el) {
        var tag = tagOf(el);
        if (tag !== "DIV" && tag !== "P" && tag !== "HR") return;
        var st = (el.getAttribute("style") || "").toLowerCase();
        if (tag === "HR" || /border-top\s*:[^;]*solid/.test(st)) {
          var probe = tag === "HR" ? el.nextElementSibling : el;
          if (probe && QUOTE_HEADER.test(probe.textContent || "")) candidates.push(el);
        }
      });
      if (!candidates.length) return null;
      candidates.sort(function (a, b) {
        if (a === b) return 0;
        return a.compareDocumentPosition(b) & 4 /* FOLLOWING */ ? -1 : 1;
      });
      return candidates[0];
    }

    function signatureRegions(body) {
      return elementsIn(body).filter(function (el) {
        return (el.getAttribute("id") || "").toLowerCase().indexOf("signature") >= 0;
      });
    }

    function tidyBody(html, opts) {
      opts = opts || {};
      var colours = opts.colours !== false;
      var parsed = parseDocument(html);
      var body = parsed.doc.body;
      var boundary = findQuoteBoundary(body);
      var sigs = signatureRegions(body);
      var changed = 0;

      function inScope(el) {
        if (boundary) {
          if (el === boundary || el.contains(boundary) || boundary.contains(el)) return false;
          if (boundary.compareDocumentPosition(el) & 4 /* el FOLLOWS boundary */) return false;
        }
        for (var i = 0; i < sigs.length; i++) if (sigs[i] === el || sigs[i].contains(el)) return false;
        return true;
      }

      elementsIn(body).forEach(function (el) {
        if (tagOf(el) === "FONT" && inScope(el)) { unwrap(el); changed++; }
      });

      elementsIn(body).forEach(function (el) {
        var tag = tagOf(el);
        if (SKIP_TAGS.test(tag) || tag === "O:P" || tag === "BR" || tag === "IMG" || tag === "HR") return;
        if (!inScope(el)) return;
        var before = el.getAttribute("style") || "";
        var decls = parseStyle(before);
        var isBlock = DEFAULTED_TAGS.test(tag);
        var isTablePart = TABLE_TAGS.test(tag);

        if (getDecl(decls, "font") != null) removeDecl(decls, function (p) { return p === "font"; });

        var fam = getDecl(decls, "font-family");
        if (fam != null && !familyAllowed(fam)) {
          if (isBlock) setDecl(decls, "font-family", bodyFont);
          else removeDecl(decls, function (p) { return p === "font-family"; });
        } else if (fam == null && isBlock) setDecl(decls, "font-family", bodyFont);

        var size = getDecl(decls, "font-size");
        if (size != null && !sizeAllowed(size, el)) {
          if (isBlock) setDecl(decls, "font-size", bodySize);
          else removeDecl(decls, function (p) { return p === "font-size"; });
        } else if (size == null && isBlock) setDecl(decls, "font-size", bodySize);

        if (colours) {
          var col = getDecl(decls, "color");
          if (tag === "A") {
            if (col == null || !colourAllowed(col)) setDecl(decls, "color", linkColour);
          } else if (col != null && !colourAllowed(col)) {
            if (isBlock) setDecl(decls, "color", bodyColour);
            else removeDecl(decls, function (p) { return p === "color"; });
          } else if (col == null && isBlock) setDecl(decls, "color", bodyColour);

          if (!isTablePart) {
            removeDecl(decls, function (p, v) {
              return (p === "background" || p === "background-color") && !backgroundAllowed(v);
            });
          }
        }

        writeStyle(el, decls);
        if ((el.getAttribute("style") || "") !== before) changed++;
      });

      return { html: parsed.serialize(), changed: changed, quoteFound: !!boundary };
    }

    // ── Public: legal markings ──────────────────────────────────────────────
    function findMarking(id) {
      for (var i = 0; i < style.markings.length; i++) if (style.markings[i].id === id) return style.markings[i];
      throw new Error("Unknown marking: " + id);
    }

    function markingBanner(id) {
      var m = findMarking(id);
      var bg = resolve("color", style.banner.css.background || "");
      var tdCss = css(style.banner.css, { "background-color": style.banner.css.background });
      var note = m.note ? '<br><span style="' + esc(css(style.banner.noteCss)) + '">' + esc(m.note) + "</span>" : "";
      return '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" ' +
        'style="border-collapse:collapse;width:100%;margin:0 0 12pt 0">' +
        '<tr><td bgcolor="' + esc(bg) + '" style="' + esc(tdCss) + '">' + esc(m.banner) + note +
        "</td></tr></table>";
    }

    // Longest banners are matched first and blanked out, so "WITHOUT PREJUDICE
    // SAVE AS TO COSTS" is not also reported as plain "WITHOUT PREJUDICE".
    function detectMarkings(text) {
      var t = String(text || "").replace(/\s+/g, " ").toUpperCase();
      var found = {};
      style.markings.slice().sort(function (a, b) { return b.banner.length - a.banner.length; })
        .forEach(function (m) {
          var b = m.banner.replace(/\s+/g, " ").toUpperCase();
          var i = t.indexOf(b);
          if (i >= 0) {
            found[m.id] = true;
            t = t.slice(0, i) + " " + t.slice(i + b.length);
          }
        });
      return style.markings.filter(function (m) { return found[m.id]; }).map(function (m) { return m.id; });
    }

    // Text of a banner cell up to the first line break (without its note).
    // Walks depth-first because Outlook for Windows may wrap the cell's
    // contents in its own <p> when the message is re-saved.
    function bannerHeadline(td) {
      var text = "";
      var stop = false;
      (function walk(node) {
        for (var n = node.firstChild; n && !stop; n = n.nextSibling) {
          if (n.nodeType === TEXT) text += n.nodeValue;
          else if (n.nodeType === ELEMENT) {
            if (tagOf(n) === "BR") { stop = true; return; }
            walk(n);
          }
        }
      })(td);
      return text.replace(/\s+/g, " ").trim().toUpperCase();
    }

    function removeMarking(html, id) {
      var m = findMarking(id);
      var parsed = parseDocument(html);
      var target = m.banner.replace(/\s+/g, " ").toUpperCase();
      var removed = 0;
      Array.prototype.slice.call(parsed.doc.body.getElementsByTagName("td")).forEach(function (td) {
        if (bannerHeadline(td) !== target) return;
        var table = td.closest ? td.closest("table") : null;
        if (table && table.parentNode && table.getElementsByTagName("td").length === 1) {
          table.parentNode.removeChild(table);
          removed++;
        }
      });
      return { html: parsed.serialize(), removed: removed };
    }

    function tagSubject(subject, tag) {
      subject = String(subject || "");
      var bracketed = "[" + tag + "]";
      if (subject.toUpperCase().indexOf(bracketed.toUpperCase()) >= 0) return subject;
      return (bracketed + " " + subject).trim();
    }

    function untagSubject(subject, tag) {
      var bracketed = "[" + tag + "]";
      var re = new RegExp("\\s*" + bracketed.replace(/[.*+?^${}()|[\]\\]/g, "\\$&") + "\\s*", "gi");
      return String(subject || "").replace(re, " ").replace(/\s{2,}/g, " ").trim();
    }

    function matterSubjectTag(matter) {
      return String(style.reference.subjectFormat || "[{matter}]").replace("{matter}", matter).replace(/^\[|\]$/g, "");
    }

    // ── Public: snippets ────────────────────────────────────────────────────
    function referenceLine(refs) {
      refs = refs || {};
      var labels = style.reference.labels;
      var sep = esc(style.reference.separator).replace(/ /g, "&nbsp;");
      var labelCss = esc(css(style.reference.labelCss));
      var parts = [];
      [["ours", refs.ours], ["yours", refs.yours], ["matter", refs.matter]].forEach(function (p) {
        var v = String(p[1] || "").trim();
        if (v) parts.push('<span style="' + labelCss + '">' + esc(labels[p[0]]) + ":</span> " + esc(v));
      });
      if (!parts.length) return "";
      return '<p style="' + esc(css(style.reference.css)) + '">' + parts.join(sep) + "</p>";
    }

    function formatDate(date) {
      return (date || new Date()).toLocaleDateString(style.firm.dateLocale || "en-GB", {
        day: "numeric", month: "long", year: "numeric",
      });
    }

    function disclaimer() {
      var text = String(style.disclaimer || "").replace(/\{firm\}/g, style.firm.name);
      return '<p style="' + esc(css(style.smallPrint.css, { margin: "12pt 0 0 0" })) + '">' + esc(text) + "</p>";
    }

    return {
      css: css,
      esc: esc,
      styles: styleList,
      theme: themeGrid,
      markings: style.markings,
      automaticColour: bodyColour,
      applyStyle: applyStyle,
      applyStyleInfo: applyStyleInfo,
      typingParagraph: typingParagraph,
      typingColour: typingColour,
      labelFor: function (id, n) { var s = findStyle(id); return s.label ? labelText(s.label, n) : ""; },
      applyColour: applyColour,
      clearFormatting: clearFormatting,
      tidyBody: tidyBody,
      markingBanner: markingBanner,
      detectMarkings: detectMarkings,
      removeMarking: removeMarking,
      tagSubject: tagSubject,
      untagSubject: untagSubject,
      matterSubjectTag: matterSubjectTag,
      referenceLine: referenceLine,
      formatDate: formatDate,
      disclaimer: disclaimer,
      // exposed for tests
      _internals: { toPt: toPt, normColour: normColour, alpha: alpha, roman: roman, findQuoteBoundary: findQuoteBoundary },
    };
  }

  if (typeof module !== "undefined" && module.exports) module.exports = createFormatter;
  else root.createFormatter = createFormatter;
})(typeof self !== "undefined" ? self : this);
