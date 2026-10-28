/*
 * Email Composer: DIAGNOSTIC function command.
 *
 * Same job as the normal commands.js, but reports each step in the bar at the
 * top of the draft ("Email Composer step N: ...") and always tells Outlook it
 * has finished, so the button cannot spin forever. The last step shown is
 * where it stopped. Do not leave this version live.
 */
/* global Office, MacrotecFont */
(function () {
  "use strict";

  var NOTICE_KEY = "emailComposerDiag";
  var loadedAt = Date.now();
  var officeReady = false;

  function log(msg) {
    try { console.log("[Email Composer] " + msg); } catch (e) { /* ignore */ }
  }

  function show(item, isError, message) {
    log(message);
    try {
      item.notificationMessages.replaceAsync(NOTICE_KEY, isError
        ? { type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage, message: message }
        : { type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
            message: message, icon: "Icon.16x16", persistent: true },
        function (r) { if (r.status !== "succeeded") log("notification failed: " + JSON.stringify(r.error)); });
    } catch (e) {
      log("notification threw: " + e.message);
    }
  }

  function parse(html) {
    return new DOMParser().parseFromString(html, "text/html");
  }

  function applyHouseFont(event) {
    var finished = false;
    var step = "0 started";
    var done = function () {
      if (finished) return;
      finished = true;
      try { if (event && typeof event.completed === "function") event.completed(); }
      catch (e) { log("completed() threw: " + e.message); }
    };
    log("clicked; officeReady=" + officeReady + ", loaded " + (Date.now() - loadedAt) + " ms ago");

    var item;
    try {
      item = Office.context.mailbox.item;
    } catch (e) {
      log("step 1 failed, no mailbox item: " + e.message);
      done();
      return;
    }

    // Safety net: never leave the button spinning.
    setTimeout(function () {
      if (!finished) {
        show(item, true, "Email Composer: stopped at step " + step + " (timed out after 10 s).");
        done();
      }
    }, 10000);

    try {
      step = "2 reading body";
      show(item, false, "Email Composer step 2: reading the body (Office ready: " + officeReady + ").");
      item.body.getAsync(Office.CoercionType.Html, function (read) {
        try {
          if (read.status !== Office.AsyncResultStatus.Succeeded) {
            show(item, true, "Email Composer step 2 failed: " + JSON.stringify(read.error));
            done();
            return;
          }
          step = "3 formatting";
          var result = MacrotecFont.applyHouseFont(read.value, parse);
          step = "4 writing body";
          show(item, false, "Email Composer step 4: read " + read.value.length +
            " chars, changed " + result.changed + ", writing back.");
          item.body.setAsync(result.html, { coercionType: Office.CoercionType.Html }, function (write) {
            if (write.status !== Office.AsyncResultStatus.Succeeded) {
              show(item, true, "Email Composer step 4 failed: " + JSON.stringify(write.error));
            } else {
              show(item, false, "Email Composer step 5: done. Century Gothic 10 applied.");
            }
            done();
          });
        } catch (e) {
          show(item, true, "Email Composer failed at step " + step + ": " + e.message);
          done();
        }
      });
    } catch (e) {
      show(item, true, "Email Composer failed at step " + step + ": " + e.message);
      done();
    }
  }

  window.addEventListener("error", function (e) { log("page error: " + e.message); });
  window.applyHouseFont = applyHouseFont;
  Office.onReady(function (info) {
    officeReady = true;
    log("Office ready: host=" + (info && info.host) + ", platform=" + (info && info.platform));
    if (Office.actions && typeof Office.actions.associate === "function") {
      Office.actions.associate("applyHouseFont", applyHouseFont);
    }
  });
  log("commands.js loaded");
})();
