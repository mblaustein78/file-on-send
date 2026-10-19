/*
 * File on Send — Smart Alert handler.
 * Runs when you click Send. If a filing folder has been chosen (the message carries the
 * add-in's "fileTo" custom property), the send goes ahead. Otherwise Outlook shows a prompt
 * with "Choose folder" (opens the picker) and "Send Anyway" (keeps it in Sent Items).
 */
function onMessageSendHandler(event) {
  var done = false;
  function finish(options) {
    if (done) return;
    done = true;
    event.completed(options);
  }
  // Never hold up mail: if Outlook is slow to answer, let the send go through.
  setTimeout(function () { finish({ allowEvent: true }); }, 4000);

  try {
    Office.context.mailbox.item.loadCustomPropertiesAsync(function (result) {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        finish({ allowEvent: true });
        return;
      }
      if (result.value.get("fileTo")) {
        finish({ allowEvent: true });
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
      finish(options);
    });
  } catch (e) {
    finish({ allowEvent: true });
  }
}

Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
