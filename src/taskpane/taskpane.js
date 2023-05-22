/* eslint-disable office-addins/load-object-before-read */
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

/**
 * Your Excel.js business logic will be added to the function that is passed to Excel.run. This logic does not execute immediately. Instead, it is added to a queue of pending commands.
 * The context.sync method sends all queued commands to Excel for execution.
 * The tryCatch function will be used by all the functions interacting with the workbook from the task pane. Catching Office JavaScript errors in this fashion is a convenient way to generically handle any uncaught errors.
 */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    // Assign event handlers and other initialization logic.
    document.getElementById("create-table").onclick = () => tryCatch(createTable);
    document.getElementById("filter-table").onclick = () => tryCatch(filterTable);
    document.getElementById("sort-table").onclick = () => tryCatch(sortTable);
  }
});

async function createTable() {
  await Excel.run(async (context) => {
    // 1. Queue table creation logic here.

    /**
     * The code creates a table by using the add method of a worksheet's table collection,
     * which always exists even if it is empty. This is the standard way that Excel.js objects are created.
     * There are no class constructor APIs, and you never use a new operator to create an Excel object.
     * Instead, you add to a parent collection object.
     */

    /**
     * The first parameter of the add method is the range of only the top row of the table,
     * not the entire range the table will ultimately use. This is because when the add-in
     * populates the data rows (in the next step), it will add new rows to the table instead of
     * writing values to the cells of existing rows. This is a common pattern, because the number
     * of rows a table will have is often unknown when the table is created.
     */
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    const expensesTable = currentWorksheet.tables.add("A1:D1", true /*hasHeaders*/);
    expensesTable.name = "ExpensesTable";

    // 2. Queue commands to populate the table with data.
    /**
     * The cell values of a range are set with an array of arrays.
     * New rows are created in a table by calling the add method of the table's row collection.
     * You can add multiple rows in a single call of add by including multiple cell value arrays
     * in the parent array that is passed as the second parameter.
     */
    expensesTable.getHeaderRowRange().values = [["Date", "Merchant", "Category", "Amount"]];
    expensesTable.rows.add(null /*add at the end*/, [
      ["1/1/2017", "The Phone Company", "Communications", "120"],
      ["1/2/2017", "Northwind Electric Cars", "Transportation", "142.33"],
      ["1/5/2017", "Best For You Organics Company", "Groceries", "27.9"],
      ["1/10/2017", "Coho Vineyard", "Restaurant", "33"],
      ["1/11/2017", "Bellows College", "Education", "350.1"],
      ["1/15/2017", "Trey Research", "Other", "135"],
      ["1/15/2017", "Best For You Organics Company", "Groceries", "97.88"],
    ]);

    // 3. Queue commands to format the table.
    /**
     * The code gets a reference to the Amount column by passing its zero-based index to the getItemAt method of the table's column collection.
     * The code then formats the range of the Amount column as Euros to the second decimal.
     * Finally, it ensures that the width of the columns and height of the rows is big enough to fit the longest (or tallest) data item. Notice that the code must get Range objects to format. TableColumn and TableRow objects do not have format properties.
     */
    expensesTable.columns.getItemAt(3).getRange().numberFormat = [["\u20AC#,##0.00"]];
    expensesTable.getRange().format.autofitColumns();
    expensesTable.getRange().format.autofitRows();

    await context.sync();
  });
}

async function filterTable() {
  await Excel.run(async (context) => {
    // 1. Queue commands to filter out all expense categories except Groceries and Education.
    /**
     * The code first gets a reference to the column that needs filtering by passing the column name
     *  to the getItem method, instead of passing its index to the getItemAt method as the createTable
     *  method does. Since users can move table columns, the column at a given index might change after
     * the table is created. Hence, it is safer to use the column name to get a reference to the column.
     *  We used getItemAt safely in the preceding tutorial, because we used it in the very same method
     *  that creates the table, so there is no chance that a user has moved the column.
     */

    /**
     * The applyValuesFilter method is one of several filtering methods on the Filter object.
     */
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    const expensesTable = currentWorksheet.tables.getItem("ExpensesTable");
    const categoryFilter = expensesTable.columns.getItem("Category").filter;
    categoryFilter.applyValuesFilter(["Education", "Groceries"]);

    await context.sync();
  });
}

async function sortTable() {
  await Excel.run(async (context) => {
    // 1. Queue commands to sort the table by Merchant name.
    /**
     * The code creates an array of `SortField` objects, which has just one member since the add-in only sorts on the Merchant column.
     * The `key` property of a `SortField` object is the zero-based index of the column used for sorting. The rows of the table are sorted based on the values in the referenced column.
     * The `sort` member of a `Table` is a `TableSort` object, not a method. The `SortFields` are passed to the `TableSort` object's `apply` method.
     */
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    const expensesTable = currentWorksheet.tables.getItem("ExpensesTable");
    const sortFields = [
      {
        key: 1, // Merchant column
        ascending: false,
      },
    ];

    expensesTable.sort.apply(sortFields);

    await context.sync();
  });
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
