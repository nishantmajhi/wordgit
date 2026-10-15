/* global Office */

/**
 * The manifest declares a FunctionFile, which Word loads in a hidden frame so
 * that ribbon buttons can run code without opening the task pane. WordGit has
 * no such buttons: every command is in the task pane. This file exists so the
 * declared page resolves, and does nothing else.
 *
 * It previously held the Yeoman generator's sample command, which was the
 * Outlook one -- it built an `Office.NotificationMessageDetails` and called
 * `Office.context.mailbox.item.notificationMessages.replaceAsync`. There is no
 * mailbox in Word, so that function would have thrown on the first line that
 * touched it had anything ever invoked it.
 */

Office.onReady(() => {
  // Nothing to register.
});
