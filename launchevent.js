/*
 * File on Send — Smart Alert handler.
 * Runs when you click Send. If a filing folder has been chosen (the message carries a
 * "File: <folder path>" category), the send goes ahead. Otherwise Outlook shows a prompt
 * with "Choose folder" (opens the picker) and "Send Anyway" (keeps it in Sent Items).
 */
var FILE_PREFIX = "File: ";

function onMessageSendHandler(event) {
  try {
    Office.context.mailbox.item.categories.getAsync(function (result) {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        // Never block mail because of an add-in problem.
        event.completed({ allowEvent: true });
        return;
      }
      var cats = result.value || [];
      var chosen = cats.some(function (c) {
        return c.displayName && c.displayName.indexOf(FILE_PREFIX) === 0;
      });
      if (chosen) {
        event.completed({ allowEvent: true });
        return;
      }
      var options = {
        allowEvent: false,
        errorMessage:
          "Where should this message be filed? Click Choose folder to pick one, " +
          "or Send Anyway to keep it in Sent Items."
      };
      // Mailbox 1.14+: relabel the "Don't Send" button and make it open the folder picker.
      if (Office.context.requirements.isSetSupported("Mailbox", "1.14")) {
        options.cancelLabel = "Choose folder";
        options.commandId = "msgComposeOpenPaneButton";
      }
      event.completed(options);
    });
  } catch (e) {
    event.completed({ allowEvent: true });
  }
}

Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
