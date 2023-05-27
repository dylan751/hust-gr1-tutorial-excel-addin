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

      /**
       * In each function that you've created in this tutorial until now, you queued commands to write to the
       * Office document. Each function ended with a call to the context.sync() method, which sends the
       * queued commands to the document to be executed. However, the code you added in the last step
       * calls the sheet.protection.protected property. This is a significant difference from the earlier
       * functions you wrote, because the sheet object is only a proxy object that exists in your task pane's
       * script. The proxy object doesn't know the actual protection state of the document, so its
       * protection.protected property can't have a real value. To avoid an exception error, you must first
       * fetch the protection status from the document and use it set the value of
       * sheet.protection.protected. This fetching process has three steps:
       *    1. Queue a command to load (that is; fetch) the properties that your code needs to read.
       *    2. Call the context object's sync method to send the queued command to the document for
       *       execution and return the requested information.
       *    3. Because the sync method is asynchronous, ensure that it has completed before your code calls
       *       the properties that were fetched.
       */

      // 2. Queue command to load the sheet's "protection.protected" property from
      // the document and re-synchronize the document and task pane.

      /**
       * Every Excel object has a load method. You specify the properties of the object that you
       * want to read in the parameter as a string of comma-delimited names. In this case, the
       * property you need to read is a subproperty of the protection property. You reference the
       * subproperty almost exactly as you would anywhere else in your code, with the exception
       * that you use a forward slash ('/') character instead of a "." character.
       */

      /**
       * To ensure that the toggle logic, which reads sheet.protection.protected, doesn't run
       * until after the sync is complete and the sheet.protection.protected has been assigned
       * the correct value that is fetched from the document, it must come after the await
       * operator ensures sync has completed.
       */
      sheet.load("protection/protected");
      await context.sync();

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
