/*
 * Legal Toolkit: ribbon buttons that act immediately (Format and Mark menus).
 * The function names must match <FunctionName> entries in manifest.xml.
 */
(function (global) {
  "use strict";

  Office.onReady(function () {});

  function run(event, work, doneMessage) {
    Promise.resolve()
      .then(work)
      .then(function (msg) {
        var text = typeof msg === "string" ? msg : doneMessage;
        if (text) LT.notify(text);
        else LT.clearNotice();
      })
      .catch(function (err) {
        LT.notify(err && err.message ? err.message : String(err), "error");
      })
      .then(function () {
        if (event && typeof event.completed === "function") event.completed();
      });
  }

  var commands = {};

  // Format menu
  Object.keys(LT.FORMAT_NAMES).forEach(function (kind) {
    var name = "format_" + kind;
    commands[name] = function (event) {
      run(event, function () { return LT.formatSelection(kind); }, null);
    };
  });

  // Mark menu
  Object.keys(LT.config.markings || {}).forEach(function (id) {
    commands["mark_" + id] = function (event) {
      run(event, function () {
        return LT.applyMarking(id).then(function (r) {
          var label = LT.config.markings[id].label;
          if (!r.subjectChanged && !r.bannerAdded) return label + " marking is already on this email.";
          return "Marked " + label + ".";
        });
      });
    };
  });

  commands.mark_remove = function (event) {
    run(event, function () {
      return LT.removeMarkings().then(function (r) {
        if (!r.removed && !r.subjectChanged) return "No markings found on this email.";
        return "Markings removed.";
      });
    });
  };

  Object.keys(commands).forEach(function (name) {
    global[name] = commands[name];
    if (global.Office && Office.actions && typeof Office.actions.associate === "function") {
      try { Office.actions.associate(name, commands[name]); } catch (e) {}
    }
  });
  global.LT_COMMANDS = Object.keys(commands);
})(window);
