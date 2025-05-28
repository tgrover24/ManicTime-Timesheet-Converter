# ManicTime Timesheet Converter - Excel Office Script

This Excel Office Script automates the process of converting a ManicTime data export into a structured timesheet format suitable for import into other systems or for reporting.

## Features

-   **Unpivots ManicTime Data**: Transforms the ManicTime export (where dates are columns) into a flat list of time entries.
-   **Extracts Project Numbers**: Derives project numbers from "Tag 1" in the ManicTime data with special handling for "Office" tags (assigns project 992024).
-   **Tag-Based Lookups**: Utilizes Tag 2 for task lookups and Tag 3 for job code lookups with project-specific logic.
-   **Calculates Fiscal Periods**: Determines fiscal month, start/end dates, and fiscal year for each entry (May=1 through April=12).
-   **Creates Formatted Output Table**: Generates a new sheet (or clears an existing one) named after the first month of data (e.g., "July 2024").
-   **Populates Table with Data**: Fills the table with processed time entries and static employee information.
-   **Advanced Formulas**:
    -   XLOOKUP formulas for Project Descriptions, Task Descriptions, and Job Code Descriptions from the "LOOKUPS" sheet.
    -   Dynamic task lookups based on project numbers with special handling for project 992024.
    -   Job code assignment logic (ADM for 992024, ENC as default, Tag 3 lookup for others).
    -   Daily hour totals and cumulative monthly totals calculations.
    -   Helper column for conditional formatting (alternating day pattern).
-   **Data Validation**: Dynamic dropdown lists for Task column based on selected project, with comprehensive error messaging.
-   **Professional Formatting**:
    -   Date and number formatting for relevant columns.
    -   Column widths copied from reference "April 2025" sheet for consistency.
    -   Conditional formatting for alternating day patterns using helper column.
    -   Weekend highlighting (Saturday, Sunday) in the "Date" column with grey background.
    -   Color-coded headers (green for lookup-dependent fields, white for others).
    -   Hidden utility columns (Column1, Tag 2, Tag 3) for cleaner appearance.
    -   **Output sheet positioning**: Always moved to immediately precede the LOOKUPS sheet.
-   **Robust Error Handling & Logging**: Comprehensive console logs for script progress, validation, warnings, and errors.

## Prerequisites

1.  **ManicTime Export Sheet**: The script must be run when the active Excel sheet is the raw data export from ManicTime.
    -   **Expected Column Structure**:
        -   Column A: "Tag 1" (project identification)
        -   Column B: "Tag 2" (task identification) 
        -   Column C: "Tag 3" (job code identification)
        -   Column D: "Notes" (comments/descriptions)
        -   Columns E onward: Date columns with time values
    -   Dates in the header should be Excel serial date numbers or parsable date strings.
    -   Time values should be numeric (hours as decimal values).

2.  **"LOOKUPS" Sheet**: A sheet named "LOOKUPS" must exist in the workbook with the following tables:
    -   **`ProjectLookup`**: Columns "Number" (project numbers) and "Description" (project descriptions).
    -   **`TaskCodes`**: Dynamic structure with project-specific task codes and descriptions (e.g., "992024 Codes", "992024 Desc" columns).
    -   **`JobCodes`**: Columns "Job Code" and "Description" for job code lookups.
    -   **`Tag2Lookup`**: Columns "Tag 2" and "Task" for general Tag 2 to task mapping.
    -   **`Tag3Lookup`**: Columns "Tag 3" and "Job Code" for Tag 3 to job code mapping.

3.  **Reference Sheet**: An "April 2025" sheet for copying column widths and A1 cell formatting (optional but recommended).

## How to Use

1.  **Prepare Your Excel File**:
    *   Export your ManicTime data into an Excel sheet with the expected column structure.
    *   Ensure the "LOOKUPS" sheet exists and is populated with all required lookup tables.
    *   Verify that an "April 2025" reference sheet exists for formatting consistency (optional).

2.  **Open the Script**:
    *   In Excel, go to the "Automate" tab.
    *   Open the "timesheetConverter.ts" script in the Code Editor.

3.  **Run the Script**:
    *   Make sure your ManicTime export sheet is the active sheet.
    *   Click "Run" in the Code Editor.

4.  **Review Output**:
    *   A new sheet will be created (or cleared) named after the month/year of the first data entry.
    *   The sheet contains a formatted timesheet table with data validation and conditional formatting.
    *   Use the Task dropdown in each row to select appropriate tasks for each project.
    *   Check the "Script output" pane for logs, warnings, or errors.

## Output Table Structure

The generated timesheet contains the following columns:

| Column | Description | Source/Calculation |
|--------|-------------|-------------------|
| Employee Number | Static employee ID | Hardcoded (27) |
| Employee Name | Static employee name | Hardcoded ("GROVER, TARUN") |
| Date | Work date | From ManicTime date columns |
| Project | Project number | Extracted from Tag 1 |
| Project Description | Project name | XLOOKUP from ProjectLookup |
| Task | Task code | Data validation dropdown, Tag 2 lookup |
| Task Description | Task description | XLOOKUP based on project and task |
| Job Code | Job classification | Logic-based assignment or Tag 3 lookup |
| Job Code Description | Job description | XLOOKUP from JobCodes |
| Hours | Time worked | From ManicTime data |
| Comment | Work notes | From ManicTime Notes column |
| Period | Fiscal month (1-12) | Calculated (May=1, Apr=12) |
| Start Date | Fiscal month start | Calculated |
| End Date | Fiscal month end | Calculated |
| Fiscal Year | Fiscal year | Calculated |
| Total Days | Cumulative hours | Formula-calculated |
| Transaction | Reserved field | Empty |
| Accum Days | Reserved field | Empty |
| Hour Total | Daily total hours | Formula-calculated |
| Column1* | Formatting helper | Hidden, alternating TRUE/FALSE |
| Tag 2* | Original Tag 2 value | Hidden reference |
| Tag 3* | Original Tag 3 value | Hidden reference |

*Hidden columns for internal use

## Script Configuration

Configuration constants at the beginning of the script:

-   `lookupsSheetName`: Name of lookup sheet (default: "LOOKUPS")
-   `projectLookupTableName`: Project lookup table name (default: "ProjectLookup")
-   `taskCodesTableName`: Task codes table name (default: "TaskCodes")
-   `jobCodesTableName`: Job codes table name (default: "JobCodes")
-   `outputTableStartCell`: Output table start position (default: "B4")

**Employee Information**: Update the hardcoded values in the script:
-   Employee Number: Line ~264 (`rowValues.push(27);`)
-   Employee Name: Line ~266 (`rowValues.push("GROVER, TARUN");`)

## Troubleshooting

### Common Issues

-   **"No active sheet found"**: Run the script from the ManicTime export sheet.
-   **"No data found"**: Ensure the ManicTime sheet contains data.
-   **"No valid date columns found"**: Check that date headers are recognizable by Excel.
-   **"Lookup sheet 'LOOKUPS' not found"**: Verify the LOOKUPS sheet exists.
-   **"One or more lookup tables not found"**: Check table names and existence on LOOKUPS sheet.
-   **"CRITICAL ROW/COLUMN MISMATCH"**: Data structure issue - check script console for details.

### Data Validation Issues

-   **Empty task dropdowns**: Ensure TaskCodes table has columns matching "[ProjectNumber] Codes" format.
-   **"Invalid Task" errors**: Select tasks only from the dropdown list for each project.

### Formatting Issues

-   **Incorrect column widths**: Ensure "April 2025" reference sheet exists.
-   **Missing conditional formatting**: Check that data contains multiple days for pattern recognition.

## For Developers

-   **Language**: TypeScript with ExcelScript API
-   **Key Technologies**: XLOOKUP formulas, conditional formatting, data validation, table management
-   **Error Handling**: Comprehensive try-catch blocks with detailed logging
-   **Performance**: Bulk data operations with `setValues()` for efficiency

Refer to the [Office Scripts documentation](https://learn.microsoft.com/en-us/office/dev/scripts/) for API details. 