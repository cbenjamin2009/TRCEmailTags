/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office */

Office.onReady(() => {
  // If needed, Office.js is ready to be called.
});

function tagInfoOnly() {
  tagEmail("FYI");
}

function tagActionRequired() {
  tagEmail("[ACTION]");
}

function tagResponseRequested() {
  tagEmail("[Response Required]");
}

function tagUrgent() {
  tagEmail("[URGENT]");
  Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message &&
    Office.context.mailbox.item.setAsync({ importance: "high" });
}

function tagEmail(prefix) {
  const message = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: `Prefix: ${prefix}`,
    icon: "Icon.80x80",
    persistent: false,
  };

  Office.context.mailbox.item.notificationMessages.replaceAsync("action", message);

   Office.context.mailbox.item.subject.getAsync((prefix) => {
    const currentSubject = prefix.value;
    const newSubject = `${prefix} ${currentSubject}`;
    Office.context.mailbox.item.subject.setAsync(newSubject);
   // Office.context.mailbox.item.subject.replaceAsync(newSubject);
  });
}

/**
 * Shows a notification when the add-in command is executed.
 * @param event {Office.AddinCommands.Event}
 */
function action(event) {
  Office.context.mailbox.item.subject.setAsync("NEW SUBJECT");

  const message = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: "Rush ACTION EVENT Email Tags Applied.",
    icon: "Icon.80x80",
    persistent: false,
  };
  Office.context.mailbox.item.notificationMessages.replaceAsync("action", message);
  // Be sure to indicate when the add-in command function is complete
  event.completed();
}
// Register the function with Office.
Office.actions.associate("action", action);