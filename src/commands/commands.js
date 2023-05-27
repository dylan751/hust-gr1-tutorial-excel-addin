/* eslint-disable no-undef */
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global global, Office, self, window */

Office.onReady(() => {
  // If needed, Office.js is ready to be called
});

/**
 * Shows a notification when the add-in command is executed.
 * @param event {Office.AddinCommands.Event}
 */
function action(event) {
  const message = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: "Performed action.",
    icon: "Icon.80x80",
    persistent: true,
  };

  // Show a notification message
  Office.context.mailbox.item.notificationMessages.replaceAsync("action", message);

  // Be sure to indicate when the add-in command function is complete
  event.completed();
}

/**
 *
 * Note that we specify an args parameter to the function and the very last line of the function
 * calls args.completed. This is a requirement for all add-in commands of type ExecuteFunction.
 * It signals the Office client application that the function has finished and the UI can become responsive again.
 */
async function toggleProtection(args) {
  try {
    await Excel.run(async (context) => {
      // 1. Queue commands to reverse the protection status of the current worksheet.

      // This code uses the worksheet object's protection property in a standard toggle pattern
      const sheet = context.workbook.worksheets.getActiveWorksheet();

      // 2. Queue command to load the sheet's "protection.protected" property from
      // the document and re-synchronize the document and task pane.
      if (sheet.protection.protected) {
        sheet.protection.unprotect();
      } else {
        sheet.protection.protect();
      }

      await context.sync();
    });
  } catch (error) {
    // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
    console.error(error);
  }

  args.completed();
}

// Register the function
Office.actions.associate("toggleProtection", toggleProtection);

function getGlobal() {
  return typeof self !== "undefined"
    ? self
    : typeof window !== "undefined"
    ? window
    : typeof global !== "undefined"
    ? global
    : undefined;
}

const g = getGlobal();

// The add-in command functions need to be available in global scope
g.action = action;
