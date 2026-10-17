/*
 * Attachment renaming and reordering.
 *
 * Outlook's add-in API has no "rename" or "move" call for attachments, so a
 * change is made by attaching a copy under the new name (in the new order)
 * and then removing the original. Copies are added before anything is
 * removed, so if a step fails nothing is lost: the copies made so far are
 * taken off again and the email is left as it was.
 *
 * Only file attachments can be copied this way. Attached emails and linked
 * (cloud) files are left where they are.
 */
(function (root) {
  "use strict";

  var INVALID_CHARS = /[\\/:*?"<>|\u0000-\u001f]/;
  // A number the user or this tool put at the start of a name: "01 - ", "3. ", "12_".
  var NUMBER_PREFIX = /^\s*\d{1,3}\s*[-–—._)]\s*/;
  var MAX_NAME = 255;
  var TYPES = {
    pdf: "PDF", doc: "Word", docx: "Word", rtf: "Word", xls: "Excel", xlsx: "Excel", xlsm: "Excel", csv: "CSV",
    ppt: "PowerPoint", pptx: "PowerPoint", msg: "Email", eml: "Email", txt: "Text", zip: "ZIP",
    jpg: "Image", jpeg: "Image", png: "Image", gif: "Image", tif: "Image", tiff: "Image", heic: "Image",
  };

  function createAttachmentTools(style) {
    var numberFormat = (style && style.attachments && style.attachments.numberFormat) || "{nn} - {name}";

    function splitName(name) {
      name = String(name || "");
      var m = /^(.*[^.\s])(\.[A-Za-z0-9]{1,8})$/.exec(name);
      return m ? { base: m[1], ext: m[2] } : { base: name, ext: "" };
    }

    // Returns a message explaining what's wrong, or null if the name is fine.
    function validate(base, ext) {
      var b = String(base || "").trim();
      if (!b) return "Enter a name.";
      if (INVALID_CHARS.test(b)) return "Names can’t contain \\ / : * ? \" < > |";
      if (/\.$/.test(b)) return "Names can’t end with a full stop.";
      if ((b + (ext || "")).length > MAX_NAME) return "Names can be up to " + MAX_NAME + " characters.";
      return null;
    }

    function stripNumber(base) {
      return String(base || "").replace(NUMBER_PREFIX, "");
    }

    function numberName(base, n, total) {
      var width = Math.max(2, String(total || n).length);
      var nn = (new Array(width + 1).join("0") + n).slice(-width);
      return numberFormat.replace("{nn}", nn).replace("{n}", String(n)).replace("{name}", stripNumber(base).trim());
    }

    function formatSize(bytes) {
      bytes = Number(bytes) || 0;
      if (bytes < 1024) return bytes + " bytes";
      if (bytes < 1024 * 1024) return Math.max(1, Math.round(bytes / 1024)) + " KB";
      return (bytes / (1024 * 1024)).toFixed(1).replace(/\.0$/, "") + " MB";
    }

    function typeLabel(ext) {
      var e = String(ext || "").replace(/^\./, "").toLowerCase();
      return TYPES[e] || (e ? e.toUpperCase() : "File");
    }

    // Which attachments must be re-attached to reach the target order and
    // names. Everything from the first difference onward has to be, because
    // Outlook always adds new attachments at the end.
    function planSave(original, target) {
      var i = 0;
      while (i < target.length && i < original.length &&
        target[i].id === original[i].id && target[i].name === original[i].name) i++;
      var byId = {};
      original.forEach(function (o) { byId[o.id] = o; });
      var redo = target.slice(i);
      return {
        redo: redo,
        renamed: target.filter(function (t) { return byId[t.id] && byId[t.id].name !== t.name; }).length,
        moved: target.some(function (t, k) { return !original[k] || original[k].id !== t.id; }),
      };
    }

    // host: { getAttachmentContent(id) → {format, content}, addFileFromBase64(content, name) → newId,
    //         removeAttachment(id) }
    async function saveChanges(host, original, target, onProgress) {
      var plan = planSave(original, target);
      var total = plan.redo.length;
      if (!total) return { changed: 0, renamed: 0, moved: false, leftovers: [] };
      var added = [];
      try {
        for (var k = 0; k < total; k++) {
          var item = plan.redo[k];
          if (onProgress) onProgress(k + 1, total);
          var content = await host.getAttachmentContent(item.id);
          if (!content || String(content.format).toLowerCase() !== "base64") {
            throw new Error("“" + item.name + "” can’t be copied, so it can’t be renamed or moved.");
          }
          added.push(await host.addFileFromBase64(content.content, item.name));
        }
      } catch (err) {
        for (var a = added.length - 1; a >= 0; a--) {
          try { await host.removeAttachment(added[a]); } catch (ignore) { /* best effort */ }
        }
        throw err;
      }
      // Every copy is in place; now take the originals off.
      var leftovers = [];
      for (var r = 0; r < total; r++) {
        try { await host.removeAttachment(plan.redo[r].id); }
        catch (e) { leftovers.push(plan.redo[r].name); }
      }
      return { changed: total, renamed: plan.renamed, moved: plan.moved, leftovers: leftovers };
    }

    return {
      splitName: splitName,
      validate: validate,
      stripNumber: stripNumber,
      numberName: numberName,
      formatSize: formatSize,
      typeLabel: typeLabel,
      planSave: planSave,
      saveChanges: saveChanges,
    };
  }

  if (typeof module !== "undefined" && module.exports) module.exports = createAttachmentTools;
  else root.createAttachmentTools = createAttachmentTools;
})(typeof self !== "undefined" ? self : this);
