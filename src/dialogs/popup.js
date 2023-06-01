/* eslint-disable no-undef */

/**
 * Every page that calls APIs in the Office.js library
 * must first ensure that the library is fully initialized.
 * The best way to do that is to call the Office.onReady() function.
 * The call of Office.onReady() must run before any calls to Office.js;
 * hence the assignment is in a script file that is loaded by the page,
 * as it is in this case.
 */
Office.onReady((info) => {
  // 1. Assign handler to the OK button.
  document.getElementById("ok-button").onclick = () => tryCatch(sendStringToParentPage);
});

// 2. Create the OK button handler.
/**
 * The messageParent method passes its parameter to the parent page,
 * in this case, the page in the task pane. The parameter must be a string,
 * which includes anything that can be serialized as a string, such as XML or JSON,
 * or any type that can be cast to a string. This also adds the same tryCatch method
 * used in taskpane.js for error handling.
 */
function sendStringToParentPage() {
  const userName = document.getElementById("name-box").value;
  Office.context.ui.messageParent(userName);
}

/** Default helper for invoking an action and handling errors. */
async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
    console.error(error);
  }
}
