﻿'use strict';

(function () {


    Office.onReady((info) => {
        if (info.host === Office.HostType.Excel) {
            // Assign event handlers and other initialization logic.
            document.getElementById("set-color").onclick = (() => tryCatch(setColor));
            document.getElementById("create-table").onclick = (() => tryCatch(createTable));
            document.getElementById("filter-table").onclick = (() => tryCatch(filterTable));
            document.getElementById("sort-table").onclick = (() => tryCatch(sortTable));
            document.getElementById("create-chart").onclick = (() => tryCatch(createChart));
            document.getElementById("freeze-header").onclick = (() => tryCatch(freezeHeader));
            document.getElementById("open-dialog").onclick = (() => tryCatch(openDialog));

        }
    });

    /**
     * This function set color for selected range.
     */
    async function setColor() {
        await Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange();
            range.format.fill.color = 'green';

            await context.sync();
        }).catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));

            }
        });
    }


    /**
     * This function creates a table with some sample data and formats the range to fit it.
     */
    async function createTable() {
        await Excel.run(async (context) => {
            const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
            const expensesTable = currentWorksheet.tables.add("A1:D1", true /*hasHeaders*/);
            expensesTable.name = "ExpensesTable";

            expensesTable.getHeaderRowRange().values =
                [["Date", "Merchant", "Category", "Amount"]];

            expensesTable.rows.add(null /*add at the end*/, [
                ["1/1/2023", "The Phone Company", "Communications", "120"],
                ["1/2/2023", "Northwind Electric Cars", "Transportation", "142.33"],
                ["1/5/2023", "Best For You Organics Company", "Groceries", "27.9"],
                ["1/10/2023", "Coho Vineyard", "Restaurant", "33"],
                ["1/11/2023", "Bellows College", "Education", "350.1"],
                ["1/15/2023", "Trey Research", "Other", "135"],
                ["1/15/2023", "Best For You Organics Company", "Groceries", "97.88"]
            ]);


            // Learn more about the Excel number format syntax in this article:
            // https://support.microsoft.com/office/5026bbd6-04bc-48cd-bf33-80f18b4eae68
            expensesTable.columns.getItemAt(3).getRange().numberFormat = [['\u20AC#,##0.00']];

            expensesTable.getRange().format.autofitColumns();
            expensesTable.getRange().format.autofitRows();

            await context.sync();
        });
    }


    /**
     * This function filters the "ExpensesTable" to only show rows
     * with categories of "Education" and "Groceries".
     */
    async function filterTable() {
        await Excel.run(async (context) => {
            const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();

            const expensesTable = currentWorksheet.tables.getItem('ExpensesTable');
            const categoryFilter = expensesTable.columns.getItem('Category').filter;
            categoryFilter.applyValuesFilter(['Education', 'Groceries']);

            await context.sync();
        });
    }


    /**
     * This function sorts the "ExpensesTable" based on values in the second column.
     */
    async function sortTable() {
        await Excel.run(async (context) => {
            const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
            const expensesTable = currentWorksheet.tables.getItem('ExpensesTable');
            const sortFields = [
                {
                    key: 1,            // Merchant column
                    ascending: false,
                }
            ];

            expensesTable.sort.apply(sortFields);
            await context.sync();
        });
    }


    /**
     * This function creates a clustered column chart based on the "ExpensesTable".
     */
    async function createChart() {
        await Excel.run(async (context) => {

            const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
            const expensesTable = currentWorksheet.tables.getItem('ExpensesTable');
            const dataRange = expensesTable.getDataBodyRange();

            const chart = currentWorksheet.charts.add('ColumnClustered', dataRange, 'Auto');

            chart.setPosition("A15", "F30");
            chart.title.text = "Expenses";
            chart.legend.position = "Right";
            chart.legend.format.fill.setSolidColor("white");
            chart.dataLabels.format.font.size = 15;
            chart.dataLabels.format.font.color = "black";
            chart.series.getItemAt(0).name = 'Value in \u20AC';

            await context.sync();
        });
    }


    /**
     * This function freezes the top row of the active Excel worksheet.
     */
    async function freezeHeader() {
        await Excel.run(async (context) => {

            const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
            currentWorksheet.freezePanes.freezeRows(1);

            await context.sync();
        });
    }


    let dialog = null;

    /**
     * This function opens a dialog that uses popup.html.
     */
    function openDialog() {
        Office.context.ui.displayDialogAsync(
            'https://localhost:3000/popup.html',
            { height: 45, width: 55 },

            function (result) {
                dialog = result.value;
                dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
            }
        );
    }


    /**
     * This function writes the string provided by the dialog to the "user-name" element in the task pane.
     * @param arg The value returned from the dialog.
     */
    function processMessage(arg) {
        document.getElementById("user-name").innerHTML = arg.message;
        dialog.close();
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


})();