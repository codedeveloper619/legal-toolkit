/*
 * The ribbon: one group on the Message tab with six controls (Outlook's
 * limit for an add-in on a built-in tab, and the only option in Outlook on
 * Mac, which can't show an add-in's own tab).
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
      { id: "HS.Colour", type: "menu", label: "Colour", icon: "ct-colour", tip: "The firm’s theme colours.", items: colourItems },
      { id: "HS.Marking", type: "menu", label: "Marking", icon: "ct-marking", tip: "Privileged, Without Prejudice and other legal markings.", items: markingItems },
      { id: "HS.Tools", type: "menu", label: "Tools", icon: "ct-tools", tip: "Clear formatting, date, disclaimer, references and attachments.", items: toolItems },
      { id: "HS.Open", type: "button", label: "House Style", icon: "icon", tip: "Opens the House Style pane with every tool, including Tidy and Attachments.", pane: "format", pinnable: true },
    ];

    function all() {
      var out = [];
      controls.forEach(function (c) { out.push(c); (c.items || []).forEach(function (i) { out.push(i); }); });
      return out;
    }

    return {
      groupId: "HS.Group",
      groupLabel: style.firm.name,
      controls: controls,
      all: all,
      functions: function () { return all().filter(function (x) { return x.fn; }); },
      icons: function () {
        var seen = {};
        return all().map(function (x) { return x.icon; }).filter(function (n) { return !seen[n] && (seen[n] = true); });
      },
    };
  }

  if (typeof module !== "undefined" && module.exports) module.exports = createRibbon;
  else root.createRibbon = createRibbon;
})(typeof self !== "undefined" ? self : this);
