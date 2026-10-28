/*
 * Email Composer: Outlook function command.
 *
 * Runs when the "Email Composer" button is clicked in a compose window.
 * Reads the draft body, applies Century Gothic 10pt to the new text via
 * MacrotecFont.applyHouseFont, and writes it back. Touches only the item
 * being composed (ReadWriteItem permission). Makes no network requests.
 */
/* global Office, MacrotecFont */
(function () {
  "use strict";

  var NOTICE_KEY = "macrotecFont";

  function notify(item, type, message) {
    // Shows a short line at the top of the draft; failure here is harmless.
    try {
      item.notificationMessages.replaceAsync(NOTICE_KEY, {
        type: type,
        message: message,
        icon: "Icon.16x16",
        persistent: false,
      });
    } catch (e) {
      /* notifications are cosmetic */
    }
  }

  function parse(html) {
    return new DOMParser().parseFromString(html, "text/html");
  }

  function applyHouseFont(event) {
    var item = Office.context.mailbox.item;
    var done = function () {
      if (event && typeof event.completed === "function") event.completed();
    };

    item.body.getAsync(Office.CoercionType.Html, function (read) {
      if (read.status !== Office.AsyncResultStatus.Succeeded) {
        notify(item, Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
          "Email Composer: could not read the email body.");
        done();
        return;
      }

      var result;
      try {
        result = MacrotecFont.applyHouseFont(read.value, parse);
      } catch (e) {
        notify(item, Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
          "Email Composer: formatting failed (" + e.message + ").");
        done();
        return;
      }

      item.body.setAsync(result.html, { coercionType: Office.CoercionType.Html }, function (write) {
        if (write.status !== Office.AsyncResultStatus.Succeeded) {
          notify(item, Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
            "Email Composer: could not update the email body.");
        } else {
          notify(item, Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
            "Century Gothic 10 applied.");
        }
        done();
      });
    });
  }

  // Older clients look the function up by its global name; newer ones use
  // Office.actions.associate. Support both.
  window.applyHouseFont = applyHouseFont;
  Office.onReady(function () {
    if (Office.actions && typeof Office.actions.associate === "function") {
      Office.actions.associate("applyHouseFont", applyHouseFont);
    }
  });
})();
