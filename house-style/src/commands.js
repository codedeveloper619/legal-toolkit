/*
 * Ribbon buttons that act straight away (Text, Lists, Colour, Marking and
 * Tools menus). Outlook loads this page invisibly and calls the functions
 * named in manifest.xml (<FunctionName>), which come from ribbon.js.
 *
 * Success is silent, so typing can carry on. Problems, and markings being
 * added or removed, are shown in the bar above the email.
 */
(function () {
  "use strict";

  var STYLE = window.HOUSE_STYLE;
  var fmt = window.createFormatter(STYLE, window);
  var storage = null;
  try { storage = window.localStorage; } catch (e) { /* not available */ }
  var actions = window.createActions(STYLE, fmt, { win: window, storage: storage });
  var ribbon = window.createRibbon(STYLE, fmt);
  var host = null;

  function getHost() {
    if (!host) host = window.createOutlookHost(Office);
    return host;
  }

  function finish(event) {
    try { if (event && typeof event.completed === "function") event.completed(); } catch (e) { /* ignore */ }
  }

  function run(event, item) {
    var h = null;
    Promise.resolve()
      .then(function () {
        h = getHost();
        return item.run(actions, h, { tagSubject: true });
      })
      .then(function (res) {
        if (res && (res.notify || res.kind === "error")) return h.notify(res.message, res.kind);
        return h.clearNotice();
      }, function (err) {
        if (!(err && err.userMessage) && window.console) console.error(err);
        try { return (h || getHost()).notify(actions.describeError(err), "error"); } catch (e) { return null; }
      })
      .then(function () { finish(event); }, function () { finish(event); });
  }

  var registered = ribbon.functions().map(function (item) {
    var fn = function (event) { run(event, item); };
    window[item.fn] = fn; // older Outlook versions look functions up by name
    return [item.fn, fn];
  });

  Office.onReady(function () {
    if (Office.actions && typeof Office.actions.associate === "function") {
      registered.forEach(function (r) {
        try { Office.actions.associate(r[0], r[1]); } catch (e) { /* already registered */ }
      });
    }
  });
})();
