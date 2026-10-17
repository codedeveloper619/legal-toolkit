/*
 * ─────────────────────────────────────────────────────────────────────────────
 *  HOUSE STYLE CONFIGURATION
 *  This is the only file most firms need to edit.
 *
 *  Everything the add-in inserts or tidies is driven from here: the firm
 *  name, font, theme colours, text styles, legal markings, the disclaimer and
 *  attachment numbering. "Tidy whole email" treats every font, size and colour
 *  used below (including all theme colour shades) as house-approved and snaps
 *  anything else back to body text.
 *
 *  Measurements were taken from the firm's style sheet: Arial throughout,
 *  0.5" (36pt) indent per level, labels (a) (i) • hanging in the indent.
 * ─────────────────────────────────────────────────────────────────────────────
 */
(function (root) {
  "use strict";

  var HOUSE_STYLE = {
    firm: {
      name: "Ali Sharif Zu'bi",
      monogram: "ASZ",
      // Used for "Today's date". en-GB → 17 October 2026, en-US → October 17, 2026
      dateLocale: "en-GB",
    },

    fonts: {
      body: "Arial, Helvetica, sans-serif",
      heading: "Arial, Helvetica, sans-serif",
    },

    // Theme colours in the same order as Word's colour picker. The five shades
    // under each colour are calculated the same way Word calculates them, so
    // only these ten values need to change if the theme changes.
    theme: [
      { name: "White", slot: "Background 1", hex: "#FFFFFF" },
      { name: "Black", slot: "Text 1", hex: "#1E1E1E" },
      { name: "Off-white", slot: "Background 2", hex: "#F3F3F1" },
      { name: "Taupe", slot: "Text 2", hex: "#8C8179" },
      { name: "Burgundy", slot: "Accent 1", hex: "#A41D30" },
      { name: "Slate grey", slot: "Accent 2", hex: "#8896A0" },
      { name: "Pale sage", slot: "Accent 3", hex: "#BAC5C6" },
      { name: "Taupe", slot: "Accent 4", hex: "#8C8179" },
      { name: "Dark burgundy", slot: "Accent 5", hex: "#63111D" },
      { name: "Charcoal blue", slot: "Accent 6", hex: "#566470" },
    ],

    // Named colours used by the styles below.
    colours: {
      text: "#000000", // body text ("Automatic")
      heading: "#A41D30", // Burgundy, Accent 1
      accent: "#A41D30", // legal markings
      accentTint: "#F3F3F1", // Off-white, Background 2: banner background
      muted: "#566470", // Charcoal blue, Accent 6: small print, reference line
      link: "#0563C1", // Office's standard hyperlink blue
    },

    // All indents are in points. 36pt = 0.5 inch = 1.27 cm.
    layout: {
      levels: 3,
      indentStep: 36, // each level moves the text in by this much
      labelWidth: 36, // space from (a) / (i) / • to the start of the text
      quoteRightIndent: 36, // quotations are also indented from the right
    },

    // Body text, used by Body, Alpha, Roman, Bullet and Quote styles.
    text: {
      css: {
        "font-family": "body",
        "font-size": "11pt",
        color: "text",
        "line-height": "115%",
      },
      spaceAfter: 10,
    },

    // Head styles. One entry per level: Head, Head 1, Head 2.
    head: {
      css: {
        "font-family": "heading",
        "font-weight": "bold",
        color: "heading",
        "line-height": "115%",
      },
      levels: [
        { size: "12pt", spaceBefore: 6, spaceAfter: 10 },
        { size: "12pt", spaceBefore: 6, spaceAfter: 10 },
        { size: "11pt", spaceBefore: 8, spaceAfter: 5 },
      ],
    },

    // The style set shown in the pane: one row per kind, one name per level.
    // For numbered kinds, {a} is a letter (a, b … z, aa, bb), {i} a roman
    // numeral and {1} a number.
    styles: [
      { kind: "head", names: ["Head", "Head 1", "Head 2"] },
      { kind: "body", names: ["Body", "Body 1", "Body 2"] },
      { kind: "alpha", names: ["Alpha 1", "Alpha 2", "Alpha 3"], label: "({a})" },
      { kind: "roman", names: ["Roman 1", "Roman 2", "Roman 3"], label: "({i})" },
      { kind: "bullet", names: ["Bullet 1", "Bullet 2", "Bullet 3"], label: "•" },
      { kind: "quote", names: ["Quote", "Quote 1", "Quote 2"] },
    ],

    smallPrint: {
      css: {
        "font-family": "body",
        "font-size": "8pt",
        color: "muted",
        "line-height": "115%",
      },
    },

    // Legal markings: a banner at the top of the message plus an optional
    // subject-line tag. Edit the wording to match the firm's policy.
    markings: [
      {
        id: "privileged",
        label: "Privileged & Confidential",
        banner: "PRIVILEGED & CONFIDENTIAL",
        note: "Legally privileged. If you have received this in error, please tell the sender and delete it.",
        subjectTag: "PRIVILEGED & CONFIDENTIAL",
      },
      {
        id: "wp",
        label: "Without Prejudice",
        banner: "WITHOUT PREJUDICE",
        note: "",
        subjectTag: "WITHOUT PREJUDICE",
      },
      {
        id: "wpsatc",
        label: "Without Prejudice Save as to Costs",
        banner: "WITHOUT PREJUDICE SAVE AS TO COSTS",
        note: "",
        subjectTag: "WP SATC",
      },
      {
        id: "stc",
        label: "Subject to Contract",
        banner: "SUBJECT TO CONTRACT",
        note: "Nothing in this email creates a binding agreement.",
        subjectTag: "SUBJECT TO CONTRACT",
      },
      {
        id: "draft",
        label: "Draft for discussion",
        banner: "DRAFT — FOR DISCUSSION PURPOSES ONLY",
        note: "",
        subjectTag: "DRAFT",
      },
    ],

    banner: {
      css: {
        "font-family": "heading",
        "font-size": "9pt",
        "font-weight": "bold",
        "letter-spacing": "0.6pt",
        color: "accent",
        background: "accentTint",
        "border-left": "4pt solid accent",
        padding: "6pt 10pt",
      },
      noteCss: {
        "font-family": "body",
        "font-size": "8pt",
        "font-weight": "normal",
        "letter-spacing": "0",
        color: "muted",
      },
    },

    reference: {
      labels: { ours: "Our ref", yours: "Your ref", matter: "Matter" },
      separator: "  |  ",
      labelCss: { "font-weight": "bold", color: "text" },
      css: {
        "font-family": "body",
        "font-size": "9pt",
        color: "muted",
        margin: "0 0 10pt 0",
      },
      // e.g. "[12345-0001] Original subject"
      subjectFormat: "[{matter}]",
    },

    // Check this wording with the firm before rollout.
    disclaimer:
      "This email and any attachments are confidential and may be legally privileged. If you are not the intended recipient, please notify the sender immediately and delete it; you must not copy, use or disclose its contents. {firm} will never change its bank details by email. Always confirm by telephone before transferring funds.",

    attachments: {
      // Used by "Number in this order". {nn} = 01, 02 … ; {n} = 1, 2 …
      numberFormat: "{nn} - {name}",
    },
  };

  if (typeof module !== "undefined" && module.exports) module.exports = HOUSE_STYLE;
  else root.HOUSE_STYLE = HOUSE_STYLE;
})(typeof self !== "undefined" ? self : this);
