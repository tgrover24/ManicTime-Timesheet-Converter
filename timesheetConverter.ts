// Office Script for transforming ManicTime data export into a timesheet format.
async function main(workbook: ExcelScript.Workbook) {
    console.log("Starting timesheet conversion script...");
  
    // --- Configuration Constants ---
    const lookupsSheetName = "LOOKUPS";
    const projectLookupTableName = "ProjectLookup";
    const taskCodesTableName = "TaskCodes";
    const jobCodesTableName = "JobCodes";
    const outputTableStartCell = "B4"; // Cell where the output table headers will begin
  
    // Get the active sheet which should be the ManicTime export
    const manicTimeSheet = workbook.getActiveWorksheet();
    if (!manicTimeSheet) {
      console.log("No active sheet found. Please run the script from the ManicTime export sheet.");
      return;
    }
  
    // --- Phase 1: Unpivot ManicTime Data ---
    // This section reads the ManicTime sheet, which has dates across columns,
    // and transforms it into a flat list where each row represents a single time entry for a specific tag and date.
    const unpivotedData: {
      date: Date;         // The date of the time entry
      projectNumber: string; // Project number extracted from Tag 1
      notes: string;        // Notes from the ManicTime export
      hours: number;        // Duration of the time entry
      originalTag: string;  // The original "Tag 1" value
      tag2: string;         // Tag 2 value for Task lookup
      tag3: string;         // Tag 3 value for Job Code lookup
    }[] = [];
  
    const manicTimeRange = manicTimeSheet.getUsedRange();
    if (!manicTimeRange) {
      console.log("No data found on the ManicTime sheet.");
      return;
    }
    const manicTimeValues = manicTimeRange.getValues();
  
    if (manicTimeValues.length === 0) {
      console.log("manicTimeValues is empty (no rows found, even headers). Aborting.");
      return;
    }
  
    const headerRowValues = manicTimeValues[0];
    const dateColumns: { index: number; date: Date }[] = []; // Stores column index and JS Date for each date column
    let firstMonthDate: Date | null = null; // Used for naming the output sheet
  
    // Parse date columns from the header row (starting at index 4, 5th column)
    for (let j = 4; j < headerRowValues.length; j++) {
      const headerCellRawValue = headerRowValues[j];
      if (typeof headerCellRawValue === 'string' && headerCellRawValue.toLowerCase() === "total") {
        break; // Stop if a "total" column header is found, common in ManicTime exports
      }
  
      let jsDate: Date | null = null;
      if (typeof headerCellRawValue === 'number') {
        // Convert Excel serial date number to JavaScript Date object
        jsDate = new Date(Math.round((headerCellRawValue - 25569) * 86400 * 1000));
      } else if (typeof headerCellRawValue === 'string') {
        // Fallback for string dates, though serial numbers are expected
        const parsedDate = new Date(headerCellRawValue);
        if (!isNaN(parsedDate.getTime())) {
          jsDate = parsedDate;
        }
      }
  
      if (jsDate && !isNaN(jsDate.getTime())) { // Ensure jsDate is a valid Date object
        dateColumns.push({ index: j, date: jsDate });
        if (!firstMonthDate) {
          firstMonthDate = jsDate;
        }
      } else {
        if (headerCellRawValue !== null && headerCellRawValue !== undefined && headerCellRawValue.toString().trim() !== "") {
          let attemptedParseForLog: Date | null = null;
          if (typeof headerCellRawValue === 'string') attemptedParseForLog = new Date(headerCellRawValue);
          else if (typeof headerCellRawValue === 'number') attemptedParseForLog = new Date(Math.round((headerCellRawValue - 25569) * 86400 * 1000));
  
          if (attemptedParseForLog && isNaN(attemptedParseForLog.getTime())) {
            console.log(`Warning: Header cell value '${headerCellRawValue}' at column index ${j} could not be parsed into a valid date. Skipping this column.`);
          } else if (!jsDate) {
            console.log(`Warning: Could not interpret header cell value '${headerCellRawValue}' at column index ${j} as a date. Skipping this column.`);
          }
        }
      }
    }
  
    if (dateColumns.length === 0) {
      console.log("No valid date columns found in ManicTime sheet header. Ensure dates are in the header or are recognizable.");
      return;
    }
  
    // Iterate through each data row of the ManicTime sheet (skip header row, index 0)
    for (let i = 1; i < manicTimeValues.length; i++) {
      const currentRow = manicTimeValues[i] as (string | number | boolean)[];
      const tag1 = currentRow[0]?.toString().trim() || ""; // "Tag 1" column
  
      // Stop processing if a "total" row or an empty tag in the first column is encountered
      if (tag1.toLowerCase() === "total" || tag1 === "") {
        break;
      }
  
      const tag2 = currentRow[1]?.toString().trim() || ""; // "Tag 2" column
      const tag3 = currentRow[2]?.toString().trim() || ""; // "Tag 3" column
      const notes = currentRow[3]?.toString().trim() || ""; // "Notes" column (moved from index 1)
      let projectNumber: string;
  
      // Determine project number based on "Tag 1"
      if (tag1.toLowerCase() === "office") {
        projectNumber = "992024"; // Special case for "Office" tag
      } else {
        const tagLength = tag1.length;
        if (tagLength >= 6) {
          projectNumber = tag1.substring(tagLength - 6); // Extract last 6 characters
          if (!/^\d{6}$/.test(projectNumber)) {
            console.log(`Warning: Extracted project number '${projectNumber}' from tag '${tag1}' is not 6 digits. Using as is.`);
          }
        } else {
          projectNumber = tag1; // Use the whole tag if shorter than 6 characters
          console.log(`Warning: Tag '${tag1}' is shorter than 6 characters. Using full tag as project number: '${projectNumber}'.`);
        }
      }
  
      // Iterate through the identified date columns for the current row
      for (const dc of dateColumns) {
        const hoursValue = currentRow[dc.index]; // Get hours value for the current tag and date
        let hours = 0;
        if (typeof hoursValue === 'number') {
          hours = hoursValue;
        } else if (typeof hoursValue === 'string') {
          hours = parseFloat(hoursValue);
          if (isNaN(hours)) hours = 0; // Treat non-numeric strings as 0 hours
        }
  
        // Only add entries with more than 0 hours
        if (hours > 0) {
          unpivotedData.push({
            date: dc.date,
            projectNumber: projectNumber,
            notes: notes,
            hours: hours,
            originalTag: tag1,
            tag2: tag2,
            tag3: tag3
          });
        }
      }
    }
  
    // Sort unpivoted data by date, then by original tag for consistent output order
    if (unpivotedData.length > 0) {
      unpivotedData.sort((a, b) => {
        const dateComparison = a.date.getTime() - b.date.getTime();
        if (dateComparison !== 0) return dateComparison;
        return a.originalTag.localeCompare(b.originalTag);
      });
      firstMonthDate = unpivotedData[0].date; // Update firstMonthDate based on actual earliest data point
    }
  
    if (!firstMonthDate) {
      console.log("Cannot determine a valid first month date from the data. Aborting script.");
      return;
    }
  
    // --- Phase 2: Prepare Output Sheet and Data --- 
    const monthNames = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
    const outputSheetName = `${monthNames[firstMonthDate.getUTCMonth()]} ${firstMonthDate.getUTCFullYear()}`;
  
    let outputSheet = workbook.getWorksheet(outputSheetName);
    if (outputSheet) {
      console.log(`Output sheet '${outputSheetName}' already exists. Clearing it.`);
      // Get all tables on the sheet
      const existingTables = outputSheet.getTables();
      // Iterate and delete each table
      existingTables.forEach(table => {
        let tableNameForLogging = "UNKNOWN_TABLE"; // Default in case getName fails or table is already invalid
        try {
          tableNameForLogging = table.getName(); // Get name FIRST
          console.log(`Attempting to delete existing table '${tableNameForLogging}' from sheet '${outputSheetName}'.`);
          table.delete(); // THEN delete
          console.log(`Successfully deleted table '${tableNameForLogging}'.`);
        } catch (e) {
          // This will log the error with the table name if getName succeeded, or UNKNOWN_TABLE if getName failed.
          console.log(`Error during processing/deletion of table '${tableNameForLogging}': ${e.message}`);
        }
      });
      outputSheet.getRange().clear(ExcelScript.ClearApplyTo.all);
    } else {
      outputSheet = workbook.addWorksheet(outputSheetName);
      console.log(`Created new output sheet: ${outputSheetName}`);
    }
    outputSheet.activate();
  
    const lookupsSheet = workbook.getWorksheet(lookupsSheetName);
    if (!lookupsSheet) {
      console.log(`Lookup sheet '${lookupsSheetName}' not found. Cannot proceed with lookups.`);
      return;
    }
  
    // Move the output sheet to immediately before the LOOKUPS sheet
    if (outputSheet) {
      try {
        let lookupsSheetPosition = lookupsSheet.getPosition();
        let outputSheetPosition = outputSheet.getPosition();
  
        if (outputSheetPosition > lookupsSheetPosition) {
          outputSheet.setPosition(lookupsSheetPosition);
        } else {
            outputSheet.setPosition(lookupsSheetPosition - 1);
        }
        
        console.log(`Moved output sheet '${outputSheetName}' before LOOKUPS sheet.`);
      } catch (e) {
        console.log(`Error moving output sheet before LOOKUPS: ${e.message}`);
      }
    }
  
    // --- Copy A1 format and column width from 'April 2025' sheet ---
    const sourceSheetNameForA1 = "April 2025";
    const sourceSheetForA1 = workbook.getWorksheet(sourceSheetNameForA1);
    if (sourceSheetForA1) {
      try {
        // Copy format
        const sourceA1 = sourceSheetForA1.getRange("A1");
        const targetA1 = outputSheet.getRange("A1");
        targetA1.copyFrom(sourceA1, ExcelScript.RangeCopyType.formats);
        // Copy column width
        const sourceAColWidth = sourceSheetForA1.getCell(0, 0).getFormat().getColumnWidth();
        outputSheet.getCell(0, 0).getFormat().setColumnWidth(sourceAColWidth);
        console.log(`Copied A1 format and column width from '${sourceSheetNameForA1}' to output sheet.`);
      } catch (e) {
        console.log(`Error copying A1 format/width from '${sourceSheetNameForA1}': ${e.message}`);
      }
    } else {
      console.log(`Source sheet '${sourceSheetNameForA1}' not found for copying A1 format/width.`);
    }
  
    outputSheet.getRange("A1").setValue(`${monthNames[firstMonthDate.getUTCMonth()]} ${firstMonthDate.getUTCFullYear()}`);
  
    // Define headers for the output table
    const headers = [
      "Employee Number", "Employee Name", "Date", "Project", "Project Description",
      "Task", "Task Description", "Job Code", "Job Code Description", "Hours",
      "Comment", "Period", "Start Date", "End Date", "Fiscal Year",
      "Total Days", "Transaction", "Accum Days", "Hour Total", "Column1",
      "Tag 2", "Tag 3"
    ];
  
    const outputTableStartRange = outputSheet.getRange(outputTableStartCell);
    const headerRange = outputTableStartRange.getResizedRange(0, headers.length - 1);
    headerRange.setValues([headers]);
    headerRange.getFormat().getFill().setColor("LightGray");
    headerRange.getFormat().getFont().setBold(true);
  
    // Get lookup tables from the "LOOKUPS" sheet
    const projectLookupTable = lookupsSheet.getTable(projectLookupTableName);
    const taskCodesTable = lookupsSheet.getTable(taskCodesTableName);
    const jobCodesTable = lookupsSheet.getTable(jobCodesTableName);
  
    if (!projectLookupTable || !taskCodesTable || !jobCodesTable) {
      console.log(`One or more lookup tables ('${projectLookupTableName}', '${taskCodesTableName}', '${jobCodesTableName}') not found on '${lookupsSheetName}'. Cannot proceed.`);
      return;
    }
  
    const actualProjectLookupTableName = projectLookupTable.getName();
    const actualTaskCodesTableName = taskCodesTable.getName();
    const actualJobCodesTableName = jobCodesTable.getName();
  
    // Prepare the data array (dataToWrite) for the output table rows
    const dataToWrite: (string | number | boolean)[][] = [];
  
    // Iterate through the unpivoted data to construct each row for the output table
    for (const [entryIndex, entry] of Array.from(unpivotedData.entries())) {
      const rowValues: (string | number | boolean)[] = [];
      const currentEntryDate = entry.date;
  
      // Column 1: Employee Number (Static)
      rowValues.push(27);
      // Column 2: Employee Name (Static)
      rowValues.push("GROVER, TARUN");
      // Column 3: Date (Convert JS Date to Excel serial number)
      rowValues.push((currentEntryDate.getTime() / 86400000) + 25569);
      // Column 4: Project (from unpivoted data)
      rowValues.push(entry.projectNumber || "");
      // Column 5: Project Description (Placeholder - formula applied later)
      rowValues.push("PENDING_FORMULA");
      // Column 6: Task (Blank initially - data validation and user input expected)
      rowValues.push("");
      // Column 7: Task Description (Placeholder - formula applied later)
      rowValues.push("PENDING_FORMULA");
      // Column 8: Job Code (Placeholder - formula applied later)
      rowValues.push("PENDING_FORMULA");
      // Column 9: Job Code Description (Placeholder - formula applied later)
      rowValues.push("PENDING_FORMULA");
      // Column 10: Hours (from unpivoted data)
      rowValues.push(entry.hours || 0);
      // Column 11: Comment (from unpivoted data notes)
      rowValues.push(entry.notes || "");
  
      // Calculate Fiscal Period, Start Date, End Date, Fiscal Year
      const entryUTCMonth = currentEntryDate.getUTCMonth(); // 0-11 (Jan-Dec)
      const entryUTCFullYear = currentEntryDate.getUTCFullYear();
  
      // Fiscal month: May=1, ..., April=12
      const fiscalMonth = (entryUTCMonth - 4 + 12) % 12 + 1;
      // Year for fiscal period calculation (year in which the fiscal period STARTS)
      const fiscalYearForPeriodCalc = entryUTCMonth >= 4 ? entryUTCFullYear : entryUTCFullYear - 1;
      // JS month (0-indexed) corresponding to the calculated fiscalMonth
      const jsFiscalMonthForCalc = (fiscalMonth - 1 + 4) % 12;
  
      const jsFiscalMonthStartDate = new Date(Date.UTC(fiscalYearForPeriodCalc, jsFiscalMonthForCalc, 1));
      const fiscalMonthStartDateSerial = (jsFiscalMonthStartDate.getTime() / 86400000) + 25569;
  
      const jsFiscalMonthEndDate = new Date(Date.UTC(fiscalYearForPeriodCalc, jsFiscalMonthForCalc + 1, 0)); // Day 0 of next month is last day of current
      const fiscalMonthEndDateSerial = (jsFiscalMonthEndDate.getTime() / 86400000) + 25569;
  
      // Fiscal Year for display (ending year of the fiscal period, e.g., May 2024 - April 2025 is FY 2025)
      const fiscalYearDisplaySerial = entryUTCMonth >= 4 ? entryUTCFullYear + 1 : entryUTCFullYear;
  
      // Column 12: Period
      rowValues.push(fiscalMonth);
      // Column 13: Start Date (of fiscal month)
      rowValues.push(fiscalMonthStartDateSerial);
      // Column 14: End Date (of fiscal month)
      rowValues.push(fiscalMonthEndDateSerial);
      // Column 15: Fiscal Year
      rowValues.push(fiscalYearDisplaySerial);
  
      // Column 16: Total Days (Placeholder - formula applied later)
      rowValues.push("");
      // Column 17: Transaction (Empty)
      rowValues.push("");
      // Column 18: Accum Days (Empty)
      rowValues.push("");
      // Column 19: Hour Total (Placeholder - formula applied later)
      rowValues.push("");
      // Column 20: Column1 (for conditional formatting - Placeholder, formula applied later)
      rowValues.push(false);
  
      // Column 21: Tag 2 (from unpivoted data)
      rowValues.push(entry.tag2 || "");
  
      // Column 22: Tag 3 (from unpivoted data)
      rowValues.push(entry.tag3 || "");
  
      dataToWrite.push(rowValues);
    }
  
    let numDataRows = 0; // Variable to store the actual number of data rows to write
  
    if (dataToWrite.length > 0) {
      numDataRows = dataToWrite.length;
      const dataBodyRangeStart = outputTableStartRange.getOffsetRange(1, 0);
      const dataBodyRange = dataBodyRangeStart.getResizedRange(numDataRows - 1, headers.length - 1);
  
      // --- Enhanced Debugging and Validation before setValues ---
      console.log("--- VALIDATION BEFORE SETVALUES ---");
      const expectedRows = dataToWrite.length;
      const actualRangeRows = dataBodyRange.getRowCount();
      console.log(`dataToWrite - Number of rows (expected for range): ${expectedRows}`);
      console.log(`dataBodyRange - Actual reported rows: ${actualRangeRows}`);
  
      if (expectedRows !== actualRangeRows) {
        console.error(`CRITICAL ROW MISMATCH: dataToWrite has ${expectedRows} rows, but dataBodyRange is dimensioned for ${actualRangeRows} rows.`);
      }
  
      const expectedCols = headers.length;
      const actualRangeCols = dataBodyRange.getColumnCount();
      console.log(`dataToWrite - Expected columns per row (expected for range): ${expectedCols}`);
      console.log(`dataBodyRange - Actual reported columns: ${actualRangeCols}`);
  
      if (expectedCols !== actualRangeCols) {
        console.error(`CRITICAL COLUMN MISMATCH: dataToWrite expects ${expectedCols} columns, but dataBodyRange is dimensioned for ${actualRangeCols} columns.`);
      }
  
      let columnCountConsistent = true;
      for (let k = 0; k < dataToWrite.length; k++) {
        if (!dataToWrite[k]) {
          console.error(`CRITICAL DATA ERROR: dataToWrite[${k}] is null or undefined!`);
          columnCountConsistent = false; // This is a fatal data issue
          break;
        }
        if (dataToWrite[k].length !== expectedCols) {
          console.error(`CRITICAL COLUMN COUNT ERROR: dataToWrite[${k}] has ${dataToWrite[k].length} columns, but ${expectedCols} were expected. Row content: ${JSON.stringify(dataToWrite[k])}`);
          columnCountConsistent = false; // Log all inconsistent rows
        }
      }
  
      if (expectedRows === actualRangeRows && expectedCols === actualRangeCols && columnCountConsistent) {
        console.log("Data dimensions appear consistent with range dimensions.");
      } else {
        console.error("Dimension or data inconsistency detected. Review logs above. setValues is likely to fail.");
      }
      console.log("Attempting dataBodyRange.setValues(dataToWrite)...");
      // --- End Enhanced Debugging ---
  
      // Write the prepared data to the sheet in bulk
      dataBodyRange.setValues(dataToWrite);
      console.log("dataBodyRange.setValues(dataToWrite) call completed successfully.");
  
      dataToWrite.length = 0; // Clear array as it's no longer needed
  
      // Apply date and number formatting to relevant columns
      const dateColumnNamesForFormatting = ["Date", "Start Date", "End Date"];
      for (const colName of dateColumnNamesForFormatting) {
        const colIndex = headers.indexOf(colName);
        if (colIndex !== -1 && numDataRows > 0) { // Ensure colIndex is valid and there's data
          const columnRange = dataBodyRange.getColumn(colIndex); // dataBodyRange is correctly sized
          columnRange.setNumberFormatLocal("yyyy-mm-dd;@");
        }
      }
      const hoursColIndex = headers.indexOf("Hours");
      if (hoursColIndex !== -1 && numDataRows > 0) {
        dataBodyRange.getColumn(hoursColIndex).setNumberFormatLocal("0.00");
      }
    }
  
    // Create an Excel Table from the populated data
    const outputTableNameToCreate = `TimesheetData_${String(firstMonthDate.getUTCMonth() + 1).padStart(2, '0')}_${firstMonthDate.getUTCFullYear()}`;
    let existingTableOnSheet = outputSheet.getTable(outputTableNameToCreate);
    if (existingTableOnSheet) {
      existingTableOnSheet.delete();
    }
  
    // Define the full range for the table, including headers and data rows (if any)
    const dataBodyForTableDefinition = numDataRows > 0
      ? outputTableStartRange.getOffsetRange(1, 0).getResizedRange(numDataRows - 1, headers.length - 1)
      : headerRange;
  
    const fullRangeForTable = headerRange.getBoundingRect(dataBodyForTableDefinition);
  
    let timesheetTable: ExcelScript.Table | null = null;
    if (fullRangeForTable.getRowCount() > 0 && fullRangeForTable.getColumnCount() > 0) {
      timesheetTable = outputSheet.addTable(fullRangeForTable, true /*hasHeaders*/);
      if (timesheetTable) {
        timesheetTable.setName(outputTableNameToCreate);
        timesheetTable.setShowTotals(false);
        try {
          timesheetTable.resize(fullRangeForTable); // Attempt to ensure table size is exact
        } catch (resizeError) {
          console.log(`Error during table resize: ${resizeError.message}. Current range: ${timesheetTable.getRange().getAddress()}`);
        }
        let verifiedTable = outputSheet.getTable(outputTableNameToCreate);
        if (verifiedTable) {
          timesheetTable = verifiedTable;
        } else {
          console.log(`Could not re-fetch table by name '${outputTableNameToCreate}' for verification.`);
          timesheetTable = null;
        }
      } else {
        timesheetTable = null;
      }
    } else {
      console.log("Cannot create table, fullRangeForTable is invalid or empty.");
      timesheetTable = null;
    }
  
    // --- Apply Formulas and Post-Processing to the Table ---
    // Apply formulas only if the table was created and has data rows
    if (timesheetTable && timesheetTable.getRowCount() > 0) {
      const tableActualDataStartCell = outputTableStartRange.getOffsetRange(1, 0);
  
      // Formula for Project Description
      const projectDescCol = timesheetTable.getColumnByName("Project Description");
      if (projectDescCol && numDataRows > 0) {
        const columnTargetRange = tableActualDataStartCell.getOffsetRange(0, projectDescCol.getIndex()).getResizedRange(numDataRows - 1, 0);
        const safeProjectLookupTableName = actualProjectLookupTableName.includes(" ") || /[^a-zA-Z0-9_]/.test(actualProjectLookupTableName) ? `'${actualProjectLookupTableName.replace(/'/g, "''")}'` : actualProjectLookupTableName;
        const formulaString = `=IFERROR(XLOOKUP([@Project],${safeProjectLookupTableName}[Number],${safeProjectLookupTableName}[Description],"Project Not Found"),"Project Not Found")`;
        columnTargetRange.setFormula(formulaString);
      }
  
      // Formula for Task Description
      const taskDescCol = timesheetTable.getColumnByName("Task Description");
      if (taskDescCol && numDataRows > 0) {
        const columnTargetRange = tableActualDataStartCell.getOffsetRange(0, taskDescCol.getIndex()).getResizedRange(numDataRows - 1, 0);
        if (taskCodesTable) {
          const taskDescFormula = `=IFERROR(XLOOKUP([@Task],INDIRECT("${actualTaskCodesTableName}["&[@Project]&" Codes]"),INDIRECT("${actualTaskCodesTableName}["&[@Project]&" Desc]"),"FIXXX",0),"")`;
          columnTargetRange.setFormula(taskDescFormula);
        } else {
          console.log("TaskCodes table not found, setting error message for Task Description column.");
          columnTargetRange.setValue("ERR: TaskCodes Table Missing");
        }
      }
  
      // Formula for Task column (using Tag 2 lookup with special handling for project 992024)
      const taskColumn = timesheetTable.getColumnByName("Task");
      if (taskColumn && numDataRows > 0) {
        const columnTargetRange = tableActualDataStartCell.getOffsetRange(0, taskColumn.getIndex()).getResizedRange(numDataRows - 1, 0);
        const taskFormula = `=IF([@Project]=992024,IF([@[Tag 2]]<>"",XLOOKUP([@[Tag 2]],${actualTaskCodesTableName}[992024 Desc],${actualTaskCodesTableName}[992024 Codes],""),""),IF([@[Tag 2]]<>"",XLOOKUP([@[Tag 2]],Tag2Lookup[Tag 2],Tag2Lookup[Task],""),""))`;
        columnTargetRange.setFormula(taskFormula);
      }
  
      // Formula for Job Code
      const jobCodeCol = timesheetTable.getColumnByName("Job Code");
      if (jobCodeCol && numDataRows > 0) {
        const columnTargetRange = tableActualDataStartCell.getOffsetRange(0, jobCodeCol.getIndex()).getResizedRange(numDataRows - 1, 0);
        const formulaString = `=IF([@Project]=992024,"ADM",IF([@[Tag 3]]<>"",XLOOKUP([@[Tag 3]],Tag3Lookup[Tag 3],Tag3Lookup[Job Code],"ENC"),"ENC"))`;
        columnTargetRange.setFormula(formulaString);
      }
  
      // Formula for Job Code Description
      const jobCodeDescCol = timesheetTable.getColumnByName("Job Code Description");
      if (jobCodeDescCol && numDataRows > 0) {
        const columnTargetRange = tableActualDataStartCell.getOffsetRange(0, jobCodeDescCol.getIndex()).getResizedRange(numDataRows - 1, 0);
        const safeJobCodesTableName = actualJobCodesTableName.includes(" ") || /[^a-zA-Z0-9_]/.test(actualJobCodesTableName) ? `'${actualJobCodesTableName.replace(/'/g, "''")}'` : actualJobCodesTableName;
        const formulaString = `=IFERROR(XLOOKUP([@[Job Code]],${safeJobCodesTableName}[Job Code],${safeJobCodesTableName}[Description],"Job Code Not Found"),"Job Code Not Found")`;
        columnTargetRange.setFormula(formulaString);
      }
  
      // Formulas for Column1, Hour Total, Total Days
      const dateCol = timesheetTable.getColumnByName("Date");
      const column1Col = timesheetTable.getColumnByName("Column1");
      const hourTotalCol = timesheetTable.getColumnByName("Hour Total");
      const totalDaysCol = timesheetTable.getColumnByName("Total Days");
      const hoursCol = timesheetTable.getColumnByName("Hours");
  
      // Formula for Column1 (alternating TRUE/FALSE for daily conditional formatting)
      if (dateCol && column1Col && numDataRows > 0) {
        const firstCellColumn1 = tableActualDataStartCell.getCell(0, column1Col.getIndex());
        firstCellColumn1.setFormula("=TRUE");
        if (numDataRows > 1) {
          const restOfColumn1StartCell = tableActualDataStartCell.getCell(1, column1Col.getIndex());
          const restOfColumn1TargetRange = restOfColumn1StartCell.getResizedRange(numDataRows - 2, 0);
          const column1Formula = `=IF([@Date]<>"", IF([@Date]=INDIRECT(ADDRESS(ROW()-1,COLUMN([@Date]))), INDIRECT(ADDRESS(ROW()-1,COLUMN([@Column1]))), NOT(INDIRECT(ADDRESS(ROW()-1,COLUMN([@Column1]))))), "")`;
          restOfColumn1TargetRange.setFormula(column1Formula);
        }
      } else {
        console.log("Could not find Date or Column1 column, or no data rows, for Column1 formulas.");
      }
  
      // Formula for Hour Total (sum of hours for the current day)
      if (dateCol && hourTotalCol && hoursCol && numDataRows > 0) {
        const hourTotalColumnTargetRange = tableActualDataStartCell.getOffsetRange(0, hourTotalCol.getIndex()).getResizedRange(numDataRows - 1, 0);
        const hourTotalFormula = `=IF([@Date]<>"", IF(OR([@Date]<>INDIRECT(ADDRESS(ROW()+1,COLUMN([@Date]))),INDIRECT(ADDRESS(ROW()+1,COLUMN([@Date])))=""),SUMIFS([Hours],[Date],[@Date]),""), "")`;
        hourTotalColumnTargetRange.setFormula(hourTotalFormula);
      } else {
        console.log("Could not find Date, Hour Total, or Hours column, or no data rows, for Hour Total formula.");
      }
  
      // Formula for Total Days (cumulative sum of hours for the month, up to the current day)
      if (dateCol && totalDaysCol && hoursCol && numDataRows > 0) {
        const totalDaysColumnTargetRange = tableActualDataStartCell.getOffsetRange(0, totalDaysCol.getIndex()).getResizedRange(numDataRows - 1, 0);
        const totalDaysFormula = `=IF([@Date]<>"", IF(OR([@Date]<>INDIRECT(ADDRESS(ROW()+1,COLUMN([@Date]))),INDIRECT(ADDRESS(ROW()+1,COLUMN([@Date])))=""),SUM(INDEX([Hours],1):[@Hours]),""), "")`;
        totalDaysColumnTargetRange.setFormula(totalDaysFormula);
      } else {
        console.log("Could not find Date, Total Days, or Hours column, or no data rows, for Total Days formula.");
      }
  
      // --- Add Data Validation to Task Column (This block should be within the main if condition too) ---
      const taskCol = timesheetTable.getColumnByName("Task");
      const projectColForValidation = timesheetTable.getColumnByName("Project"); // Renamed to avoid conflict
      if (taskCol && projectColForValidation && numDataRows > 0) {
        const taskColumnDataRange = tableActualDataStartCell.getOffsetRange(0, taskCol.getIndex()).getResizedRange(numDataRows - 1, 0);
        const dataValidationFormula = "=OFFSET(LOOKUPS!$N$2,1,MATCH(E5&\" Codes\",LOOKUPS!$N$2:$EF$2,0)-1, COUNTA(OFFSET(LOOKUPS!$N$2,1,MATCH(E5&\" Codes\",LOOKUPS!$N$2:$EF$2,0)-1,33)), 1)";
        const dv = taskColumnDataRange.getDataValidation();
        dv.setRule({ list: { source: dataValidationFormula, inCellDropDown: true } });
        dv.setPrompt({ showPrompt: false, message: "", title: "" });
        dv.setErrorAlert({ showAlert: true, style: ExcelScript.DataValidationAlertStyle.stop, message: "Please select a valid task from the list for the selected project. If the list is empty, the project may not have associated tasks defined in the LOOKUPS sheet.", title: "Invalid Task" });
        dv.setIgnoreBlanks(true);
        console.log(`Applied dynamic list data validation to 'Task' column (range: ${taskColumnDataRange.getAddress()}) using formula: ${dataValidationFormula}`);
      } else {
        console.log("Could not find 'Task' or 'Project' column, or no data rows, for applying data validation.");
      }
  
      // --- Apply Conditional Formatting and Specific Column Formatting ---
      console.log("Applying conditional formatting and Hour Total column color...");
      // Conditional formatting for columns up to and including "Accum Days" (index 17) based on [Column1]
      const accumDaysColIndex = headers.indexOf("Accum Days");
      const conditionalFormattingRange = tableActualDataStartCell.getResizedRange(numDataRows - 1, accumDaysColIndex); // 0 to accumDaysColIndex (inclusive)
  
      // Conditional formatting for Date column (weekends) - MOVED UP
      const dateColForWeekendFormatting = timesheetTable.getColumnByName("Date");
      if (dateColForWeekendFormatting && numDataRows > 0) {
        const dateColumnDataRange = tableActualDataStartCell.getOffsetRange(0, dateColForWeekendFormatting.getIndex()).getResizedRange(numDataRows - 1, 0);
        const cfWeekend = dateColumnDataRange.addConditionalFormat(ExcelScript.ConditionalFormatType.custom).getCustom();
        // WEEKDAY(cell, 2) returns 6 for Saturday and 7 for Sunday.
        const firstCellAddressInDateColumn = tableActualDataStartCell.getCell(0, dateColForWeekendFormatting.getIndex()).getAddress().split("!")[1];
        cfWeekend.getRule().setFormula(`=WEEKDAY(${firstCellAddressInDateColumn}, 2) > 5`);
        cfWeekend.getFormat().getFill().setColor("#AEAAAA"); // Grey color for weekends
        cfWeekend.getFormat().getFont().setColor("black");
        console.log(`Applied conditional formatting for weekends to 'Date' column (range: ${dateColumnDataRange.getAddress()}).`);
      } else {
        console.log("Could not find 'Date' column or no data rows for weekend conditional formatting.");
      }
  
      const cfTrue = conditionalFormattingRange.addConditionalFormat(ExcelScript.ConditionalFormatType.custom).getCustom();
      cfTrue.getRule().setFormula(`=$U5 = TRUE`);
      cfTrue.getFormat().getFill().setColor("#F8CBAD");
      cfTrue.getFormat().getFont().setColor("black");
  
      const cfFalse = conditionalFormattingRange.addConditionalFormat(ExcelScript.ConditionalFormatType.custom).getCustom();
      cfFalse.getRule().setFormula(`=$U5 = FALSE`);
      cfFalse.getFormat().getFill().setColor("#B4C6E7");
      cfFalse.getFormat().getFont().setColor("black");
  
      const hourTotalColIndex = headers.indexOf("Hour Total");
      if (hourTotalColIndex !== -1) {
        const hourTotalColDataRange = tableActualDataStartCell.getOffsetRange(0, hourTotalColIndex).getResizedRange(numDataRows - 1, 0);
        hourTotalColDataRange.getFormat().getFill().setColor("#5B9BD5");
        hourTotalColDataRange.getFormat().getFont().setColor("black");
      }
  
      // --- Hide specific columns by setting width to 0 ---
      console.log("Hiding Column1, Tag 2, and Tag 3 columns...");
      const columnsToHide = ["Column1", "Tag 2", "Tag 3"];
      for (const columnName of columnsToHide) {
        const columnIndex = headers.indexOf(columnName);
        if (columnIndex !== -1) {
          // Get the column letter (B=1, C=2, etc.) - table starts at column B
          const columnLetter = String.fromCharCode(66 + columnIndex); // 66 is ASCII for 'B'
          const columnRange = outputSheet.getRange(`${columnLetter}:${columnLetter}`);
          // Hide column by setting width to 0
          columnRange.getFormat().setColumnWidth(0);
          console.log(`Hidden column: ${columnName} (${columnLetter})`);
        } else {
          console.log(`Warning: Column '${columnName}' not found for hiding.`);
        }
      }
  
      // --- Apply Table Styling and Formatting ---
      console.log("Applying table style and basic formatting...");
      timesheetTable.setPredefinedTableStyle("TableStyleLight1");
      timesheetTable.setShowFilterButton(false);
      timesheetTable.setShowBandedRows(false);
      timesheetTable.setShowBandedColumns(false);
  
      // --- Set Column Widths (copied from "April 2025" sheet) ---
      console.log("Attempting to set column widths based on 'April 2025' sheet...");
      const sourceSheetNameForWidths = "April 2025";
      const sourceWidthSheet = workbook.getWorksheet(sourceSheetNameForWidths);
  
      if (sourceWidthSheet) {
        // Columns B to U correspond to 0-indexed 1 to 20, now extended to V and W for Tag 2 and Tag 3
        // Office Scripts getColumn is 0-indexed for number, or use string like "B"
        for (let i = 0; i < 22; i++) { // Loop for 22 columns (B to W)
          const targetColumnIndex = 1 + i; // Sheet column index (1 for B, 2 for C, ..., 22 for W)
          try {
            // Use getCell(0, columnIndex) to get a cell in the column, then get/set its column width
            const sourceColumnWidth = sourceWidthSheet.getCell(0, targetColumnIndex).getFormat().getColumnWidth();
            outputSheet.getCell(0, targetColumnIndex).getFormat().setColumnWidth(sourceColumnWidth);
          } catch (e) {
            console.log(`Error setting width for column index ${targetColumnIndex}: ${e.message}`);
            // It's possible a column doesn't exist or another error occurs
          }
        }
        console.log("Finished applying column widths from 'April 2025' sheet.");
      } else {
        console.log(`Warning: Source sheet '${sourceSheetNameForWidths}' not found. Column widths not set from source.`);
      }
  
      // --- Set Specific Header Background Colors ---
      console.log("Setting specific header background colors...");
      const headerRowRange = timesheetTable.getHeaderRowRange();
      const greenHeaders = [
        "Project", "Project Description", "Task Description", "Job Code Description", "Period"
      ];
      const greenColor = "#92D050";
      const whiteColor = "#FFFFFF";
  
      for (let colIndex = 0; colIndex < headers.length; colIndex++) {
        const headerName = headers[colIndex];
        const cellToFormat = headerRowRange.getCell(0, colIndex);
        if (greenHeaders.includes(headerName)) {
          cellToFormat.getFormat().getFill().setColor(greenColor);
        } else {
          cellToFormat.getFormat().getFill().setColor(whiteColor);
        }
      }
    }
  
    console.log("Timesheet conversion script finished.");
  }
  
  /**
   * Sanitizes a string to be used as part of an Excel Named Item name.
   * Named Items cannot start with numbers, be 'C', 'R', 'c', 'r', or contain many special characters.
   * @param namePart The string to sanitize.
   * @returns A sanitized string suitable for use in a Named Item name.
   */
  function sanitizeForNamedRange(namePart: string): string {
    const nameAsString = String(namePart);
    let sanitized = nameAsString.replace(/[^a-zA-Z0-9_.]/g, '_');
    if (/^[0-9]/.test(sanitized) || /^[CcRr]$/.test(sanitized) || sanitized.length === 0) {
      sanitized = '_' + sanitized;
    }
    return sanitized.substring(0, 250); // Max length for named items is 255, but 250 provides a buffer
  } 