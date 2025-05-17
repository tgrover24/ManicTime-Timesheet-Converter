# ManicTime Timesheet Converter - Excel Office Script

This Excel Office Script automates the process of converting a ManicTime data export into a structured timesheet format suitable for import into other systems or for reporting.

## Features

-   **Unpivots ManicTime Data**: Transforms the ManicTime export (where dates are columns) into a flat list of time entries.
-   **Extracts Project Numbers**: Derives project numbers from "Tag 1" in the ManicTime data.
-   **Handles Special Cases**: Includes specific logic for "Office" tags.
-   **Calculates Fiscal Periods**: Determines fiscal month, start/end dates, and fiscal year for each entry.
-   **Creates Formatted Output Table**: Generates a new sheet (or clears an existing one) named after the first month of data (e.g., "July 2024").
-   **Populates Table with Data**: Fills the table with processed time entries and static employee information.
-   **Applies Formulas**:
    -   Looks up Project Descriptions, Task Descriptions, and Job Code Descriptions from a "LOOKUPS" sheet.
    -   Calculates Job Codes based on project numbers.
    -   Generates daily hour totals and cumulative monthly totals.
    -   Adds a helper column for conditional formatting.
-   **Formatting**:
    -   Applies date and number formatting to relevant columns.
    -   Autofits columns for readability.
    -   Applies conditional formatting to rows based on an alternating day pattern (using helper column `Column1`) for visual grouping.
    -   Highlights weekend dates (Saturday, Sunday) in the "Date" column with a distinct background color.
-   **Robust Error Handling & Logging**: Includes console logs for script progress, warnings, and errors.

## Prerequisites

1.  **ManicTime Export Sheet**: The script must be run when the active Excel sheet is the raw data export from ManicTime.
    -   The ManicTime export should have "Tag 1" in the first column, "Notes" in the second, and then date columns.
    -   Dates in the header of the ManicTime export should be recognizable as Excel serial date numbers or parsable date strings.
2.  **"LOOKUPS" Sheet**: A sheet named "LOOKUPS" must exist in the workbook. This sheet must contain the following tables:
    -   `ProjectLookup`: Needs columns named "Number" (for project numbers) and "Description" (for project descriptions).
    -   `TaskCodes`: Structure expected by XLOOKUP formulas (details in script, involves dynamic lookup based on project number).
    -   `JobCodes`: Needs columns "Job Code" and "Description".

## How to Use

1.  **Prepare Your Excel File**:
    *   Ensure you have your ManicTime data exported into an Excel sheet.
    *   Verify that a "LOOKUPS" sheet is present and correctly populated with `ProjectLookup`, `TaskCodes`, and `JobCodes` tables.
2.  **Open the Script**:
    *   In Excel, go to the "Automate" tab.
    *   Open the "timesheetConverter.ts" script (or whatever you have named it) in the Code Editor.
3.  **Run the Script**:
    *   Make sure your ManicTime export sheet is the active sheet.
    *   Click "Run" in the Code Editor.
4.  **Review Output**:
    *   A new sheet will be created (or an existing one cleared) named after the month and year of the first data entry (e.g., "July 2024").
    *   This sheet will contain the formatted timesheet table.
    *   Check the "Script output" pane in the Code Editor for any logs, warnings, or errors.

## Script Configuration

The script contains some configuration constants at the beginning that you might need to adjust:

-   `lookupsSheetName`: Name of the sheet containing lookup tables (default: "LOOKUPS").
-   `projectLookupTableName`: Name of the project lookup table (default: "ProjectLookup").
-   `taskCodesTableName`: Name of the task codes table (default: "TaskCodes").
-   `jobCodesTableName`: Name of the job codes table (default: "JobCodes").
-   `outputTableStartCell`: Cell where the output table headers will begin (default: "B4").

Static values like Employee Number (27) and Employee Name ("Grover, Tarun") are hardcoded. Modify these directly in the script if needed.

## Troubleshooting

-   **"No active sheet found"**: Make sure you run the script from the ManicTime export sheet.
-   **"No data found on the ManicTime sheet"**: Ensure the ManicTime sheet is not empty.
-   **"No valid date columns found"**: Check that your ManicTime export has date headers that are recognizable by Excel.
-   **"Lookup sheet 'LOOKUPS' not found"**: Ensure the "LOOKUPS" sheet exists and is named correctly.
-   **"One or more lookup tables ... not found"**: Verify the names and existence of the required tables on the "LOOKUPS" sheet.
-   **"Range setValues: The number of rows or columns..."**: This usually indicates an issue with data preparation. The script now has detailed logging just before this step. Check the script output console in Excel for "CRITICAL" messages that will specify the mismatch.

## For Developers (Office Scripts)

-   The script is written in TypeScript.
-   It uses the ExcelScript API for interacting with Excel workbooks.
-   Refer to the [Office Scripts documentation](https://learn.microsoft.com/en-us/office/dev/scripts/) for more information on developing Office Scripts. 