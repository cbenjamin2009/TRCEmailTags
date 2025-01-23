/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office */

Office.onReady(() => {
  // If needed, Office.js is ready to be called.
});
/**
 * Shows a notification when the add-in command is executed.
 * @param event {Office.AddinCommands.Event}
 */
function tagInfoOnly(event) {
  tagEmail("[FYI]");
  event.completed();
}

function tagActionRequired(event) {
  tagEmail("[ACTION]");
  event.completed();
}

function tagResponseRequested(event) {
  tagEmail("[Response Required]");
  event.completed();
}

function tagUrgent(event) {
  tagEmail("[Response Required]");
    Office.context.mailbox.item = Office.MailboxEnums.Importance.High;
    event.completed(); // Ensure this is called regardless of success or failure
}

function tagEmail(prefix) {
  const message = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: `Prefix: ${prefix}`,
    icon: "Icon.80x80",
    persistent: false,
  };

  Office.context.mailbox.item.subject.getAsync((result) => {
    // We must first check that the currentSubject doesn't already contain a prefix, if it does, we need to exlude it and only apply the new prefix
    const prefixes = ["[FYI]", "[ACTION]", "[Response Required]", "[URGENT]"];
    let currentSubject = result.value;
    prefixes.forEach((prefix) => {
      if (currentSubject.startsWith(prefix)) {
      currentSubject = currentSubject.replace(prefix, "").trim();
      }
    });
    const newSubject = `${prefix} ${currentSubject}`;
    Office.context.mailbox.item.subject.setAsync(newSubject);
  });

}


function action(event) {
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
Office.actions.associate("tagInfoOnly", tagInfoOnly);
Office.actions.associate("tagActionRequired", tagActionRequired);
Office.actions.associate("tagResponseRequested", tagResponseRequested);
Office.actions.associate("tagUrgent", tagUrgent);
