/*
 * The ribbon, in two forms:
 *  - a "House Style" tab with a button for each style (classic Outlook for
 *    Windows, the only Outlook that shows add-in tabs), and
 *  - one group on the Message tab with six controls (Outlook's limit for an
 *    add-in on a built-in tab), used by Outlook on Mac, on the web and new
 *    Outlook, which don't show add-in tabs.
 *
 * This one definition drives the manifest (scripts/build-manifest.js), the
 * ribbon commands (commands.js) and the mock ribbon in the browser preview.
 * Each item either runs an action straight away (`fn`) or opens the pane at
 * a tab (`pane`).
 */
(function (root) {
  "use strict";

  var KIND_TIPS = {
    head: "Heading",
    body: "Body text",
    quote: "Indented quotation",
    alpha: "Lettered list: (a), (b), (c)",
    roman: "Numbered list: (i), (ii), (iii)",
    bullet: "Bulleted list",
  };
  var STYLE_TIP = " Select text to restyle it, or click where you want to write and type.";

  function createRibbon(style, fmt) {
    function styleItem(s) {
      return {
        id: "HS.Style." + s.id,
        label: s.name,
        tip: KIND_TIPS[s.kind] + ", level " + (s.level + 1) + "." + STYLE_TIP,
        icon: "st-" + s.id,
        fn: "style_" + s.id.replace(/-/g, "_"),
        run: function (actions, host) { return actions.applyStyle(host, s.id); },
      };
    }
    function stylesOf(kinds) {
      return fmt.styles.filter(function (s) { return kinds.indexOf(s.kind) >= 0; }).map(styleItem);
    }

    var colourItems = [{
      id: "HS.Colour.auto",
      label: "Automatic",
      tip: "Black body text. Select text to change it, or choose and type.",
      icon: "co-auto",
      fn: "colour_auto",
      run: function (actions, host) { return actions.applyColour(host, null, "Automatic"); },
    }];
    fmt.theme.forEach(function (t, i) {
      var label = t.name + " (" + t.slot + ")";
      colourItems.push({
        id: "HS.Colour." + i,
        label: label,
        tip: t.hex + ". Select text to colour it, or choose and type.",
        icon: "co-" + t.hex.replace("#", "").toLowerCase(),
        fn: "colour_" + i,
        run: function (actions, host) { return actions.applyColour(host, t.hex, t.name + ", " + t.slot); },
      });
    });
    colourItems.push({
      id: "HS.Colour.more",
      label: "More shades…",
      tip: "Opens the House Style pane with all the theme shades.",
      icon: "co-more",
      pane: "format",
    });

    var markingItems = fmt.markings.map(function (m) {
      return {
        id: "HS.Mark." + m.id,
        label: m.label,
        tip: "Adds the “" + m.banner + "” banner" + (m.subjectTag ? " and tags the subject" : "") + ". Choose again to remove it.",
        icon: "mk",
        fn: "mark_" + m.id,
        run: function (actions, host, opts) { return actions.toggleMarking(host, m.id, !opts || opts.tagSubject !== false); },
      };
    });

    var toolItems = [
      {
        id: "HS.Tool.clear", label: "Clear formatting", icon: "tl-clear", fn: "tool_clear",
        tip: "Turns the selected text into plain body text.",
        run: function (actions, host) { return actions.clearFormatting(host); },
      },
      {
        id: "HS.Tool.date", label: "Today’s date", icon: "tl-date", fn: "tool_date",
        tip: "Inserts today’s date, e.g. " + fmt.formatDate(new Date(2026, 9, 17)) + ".",
        run: function (actions, host) { return actions.insertDate(host); },
      },
      {
        id: "HS.Tool.disclaimer", label: "Firm disclaimer", icon: "tl-disclaimer", fn: "tool_disclaimer",
        tip: "Inserts the firm disclaimer in small print.",
        run: function (actions, host) { return actions.insertDisclaimer(host); },
      },
      {
        id: "HS.Tool.reference", label: "Reference line…", icon: "tl-ref", pane: "insert",
        tip: "Opens the pane to add Our ref, Your ref and Matter.",
      },
      {
        id: "HS.Tool.attachments", label: "Rename or reorder attachments…", icon: "tl-attach", pane: "attachments",
        tip: "Opens the pane to rename attachments or change their order.",
      },
    ];

    var controls = [
      { id: "HS.Text", type: "menu", label: "Text", icon: "ct-text", tip: "Head, Body and Quote styles at three levels.", items: stylesOf(["head", "body", "quote"]) },
      { id: "HS.Lists", type: "menu", label: "Lists", icon: "ct-lists", tip: "(a), (i) and bullet lists at three levels. Choose again for the next item.", items: stylesOf(["alpha", "roman", "bullet"]) },
      { id: "HS.Colour", type: "menu", label: "Font Colour", icon: "ct-fontcolour", tip: "Colour the selected text, or the text you’re about to type, in the firm’s theme colours.", items: colourItems },
      { id: "HS.Marking", type: "menu", label: "Marking", icon: "ct-marking", tip: "Privileged, Without Prejudice and other legal markings.", items: markingItems },
      { id: "HS.Tools", type: "menu", label: "Tools", icon: "ct-tools", tip: "Clear formatting, date, disclaimer, references and attachments.", items: toolItems },
      { id: "HS.Open", type: "button", label: "House Style", icon: "icon", tip: "Opens the House Style pane with every tool, including Tidy and Attachments.", pane: "format", pinnable: true },
    ];

    // ── The House Style tab (classic Outlook for Windows only) ──────────────
    // Outlook allows one add-in tab of up to 10 groups with up to 6 controls
    // each. Outlook on Mac, on the web and new Outlook don't show add-in tabs,
    // so they use the Message-tab group above instead.
    function copy(item, id, extra) {
      var c = {};
      Object.keys(item).forEach(function (k) { c[k] = item[k]; });
      c.id = id;
      Object.keys(extra || {}).forEach(function (k) { c[k] = extra[k]; });
      return c;
    }
    function styleButtons(kind) {
      return stylesOf([kind]).map(function (it) {
        return copy(it, "HT." + it.id.replace("HS.Style.", ""), { type: "button" });
      });
    }
    function listMenu(kind, label) {
      var items = stylesOf([kind]);
      return {
        id: "HT.menu." + kind, type: "menu", label: label, icon: items[0].icon,
        tip: KIND_TIPS[kind] + ". Choose a level; choose it again for the next item.",
        items: items.map(function (it) { return copy(it, "HT.item." + it.id.replace("HS.Style.", "")); }),
      };
    }
    function toolButton(it, label) { return copy(it, it.id.replace("HS.", "HT."), { type: "button", label: label || it.label }); }

    var tab = {
      id: "HS.Tab",
      label: "House Style",
      groups: [
        { id: "HT.G.body", label: "Body", controls: styleButtons("body") },
        { id: "HT.G.head", label: "Head", controls: styleButtons("head") },
        { id: "HT.G.quote", label: "Quote", controls: styleButtons("quote") },
        { id: "HT.G.lists", label: "Lists", controls: [listMenu("alpha", "alpha (a)"), listMenu("roman", "roman (i)"), listMenu("bullet", "Bullets")] },
        {
          id: "HT.G.colour", label: "Font Colour",
          controls: [{
            id: "HT.Colour", type: "menu", label: "Font Colour", icon: "ct-fontcolour",
            tip: "Colour the selected text, or the text you’re about to type, in the firm’s theme colours.",
            items: colourItems.map(function (it) { return copy(it, it.id.replace("HS.", "HT.")); }),
          }],
        },
        {
          id: "HT.G.mark", label: "Marking",
          controls: markingItems.map(function (it) { return copy(it, it.id.replace("HS.", "HT."), { type: "button" }); }),
        },
        {
          id: "HT.G.insert", label: "Insert",
          controls: [toolButton(toolItems[1], "Date"), toolButton(toolItems[2], "Disclaimer"), toolButton(toolItems[3], "Reference")],
        },
        {
          id: "HT.G.tools", label: "Tools",
          controls: [
            toolButton(toolItems[0]),
            { id: "HT.Tool.tidy", type: "button", label: "Tidy Email", icon: "tl-tidy", pane: "format",
              tip: "Opens the House Style pane to tidy fonts, sizes and colours across the whole email (with Undo)." },
            toolButton(toolItems[4], "Attachments"),
            { id: "HT.Open", type: "button", label: "House Style Pane", icon: "icon", pane: "format", pinnable: true,
              tip: "Opens the House Style pane with every tool. You can pin it open." },
          ],
        },
      ],
    };

    function all() {
      var out = [];
      function add(c) { out.push(c); (c.items || []).forEach(function (i) { out.push(i); }); }
      controls.forEach(add);
      tab.groups.forEach(function (g) { g.controls.forEach(add); });
      return out;
    }

    return {
      groupId: "HS.Group",
      groupLabel: style.firm.name,
      controls: controls,
      tab: tab,
      all: all,
      // One entry per function, even when several buttons share it.
      functions: function () {
        var seen = {};
        return all().filter(function (x) { return x.fn && !seen[x.fn] && (seen[x.fn] = true); });
      },
      icons: function () {
        var seen = {};
        return all().map(function (x) { return x.icon; }).filter(function (n) { return !seen[n] && (seen[n] = true); });
      },
    };
  }

  if (typeof module !== "undefined" && module.exports) module.exports = createRibbon;
  else root.createRibbon = createRibbon;
})(typeof self !== "undefined" ? self : this);
