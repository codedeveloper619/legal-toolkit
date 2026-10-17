/*
 * The connection to Outlook (Office.js), shared by the task pane and the
 * ribbon buttons. Turns Outlook's callback API into promises.
 */
(function (root) {
  "use strict";

  function HostError(message) {
    this.name = "HostError";
    this.userMessage = message;
    this.message = message;
  }
  HostError.prototype = Object.create(Error.prototype);
  HostError.prototype.constructor = HostError;

  function createOutlookHost(Office) {
    var mailbox = Office.context.mailbox;
    var item = mailbox.item;
    var HTML = { coercionType: Office.CoercionType.Html };

    function call(target, method, args) {
      return new Promise(function (resolve, reject) {
        if (!target || typeof target[method] !== "function") {
          reject(new HostError("This version of Outlook doesn’t support that action."));
          return;
        }
        target[method].apply(target, (args || []).concat(function (r) {
          if (r.status === Office.AsyncResultStatus.Succeeded) resolve(r.value);
          else reject(r.error || new Error("Outlook could not complete the action."));
        }));
      });
    }
    function supports(v) {
      try { return Office.context.requirements.isSetSupported("Mailbox", v); } catch (e) { return false; }
    }

    return {
      name: "outlook",
      bodyType: function () { return call(item.body, "getTypeAsync"); },
      getSelection: function () {
        return call(item, "getSelectedDataAsync", [Office.CoercionType.Html]).then(function (v) {
          return { html: (v && v.data) || "", source: (v && v.sourceProperty) || "body" };
        });
      },
      replaceSelection: function (html) { return call(item, "setSelectedDataAsync", [html, HTML]); },
      getBody: function () { return call(item.body, "getAsync", [Office.CoercionType.Html]); },
      getBodyText: function () { return call(item.body, "getAsync", [Office.CoercionType.Text]); },
      setBody: function (html) { return call(item.body, "setAsync", [html, HTML]); },
      prependBody: function (html) { return call(item.body, "prependAsync", [html, HTML]); },
      getSubject: function () { return call(item.subject, "getAsync"); },
      setSubject: function (s) { return call(item.subject, "setAsync", [s]); },
      // Identifies the email being written (replies and forwards), for list numbering.
      contextKey: function () { try { return item.conversationId || ""; } catch (e) { return ""; } },

      // A short message in the bar above the email (used by ribbon buttons).
      notify: function (message, kind) {
        var nm = item.notificationMessages;
        if (!nm) return Promise.resolve();
        var text = String(message || "").slice(0, 150);
        var details = kind === "error"
          ? { type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage, message: text }
          : { type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage, message: text, icon: "Icon.16", persistent: false };
        return call(nm, "replaceAsync", ["houseStyle", details]).catch(function () { /* best effort */ });
      },
      clearNotice: function () {
        var nm = item.notificationMessages;
        if (!nm) return Promise.resolve();
        return call(nm, "removeAsync", ["houseStyle"]).catch(function () { /* none to remove */ });
      },

      // Attachments (Mailbox 1.8+)
      canEditAttachments: supports("1.8"),
      getAttachments: function () {
        return call(item, "getAttachmentsAsync").then(function (list) {
          return (list || []).map(function (a) {
            return { id: a.id, name: a.name, size: a.size, type: String(a.attachmentType || "file").toLowerCase(), isInline: !!a.isInline };
          });
        });
      },
      getAttachmentContent: function (id) {
        return call(item, "getAttachmentContentAsync", [id]).then(function (v) {
          return { format: v.format, content: v.content };
        });
      },
      addFileFromBase64: function (base64, name) {
        return call(item, "addFileAttachmentFromBase64Async", [base64, name, { isInline: false }]);
      },
      removeAttachment: function (id) { return call(item, "removeAttachmentAsync", [id]); },
      onAttachmentsChanged: function (fn) {
        if (!supports("1.8")) return;
        try { item.addHandlerAsync(Office.EventType.AttachmentsChanged, fn); } catch (e) { /* not available */ }
      },
    };
  }

  if (typeof module !== "undefined" && module.exports) module.exports = createOutlookHost;
  else root.createOutlookHost = createOutlookHost;
})(typeof self !== "undefined" ? self : this);
