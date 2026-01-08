/**
 * Monthly Attendance Processor
 * Processes monthly attendance data from Daxko Ops exports for any program
 * Automatically creates monthly tabs, archives data, and calculates daily averages
 * Generic design allows use across multiple programs without code changes
 */

// ===========================
// MENU SETUP
// ===========================

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Attendance Tools')
    .addItem('Process Attendance', 'processAttendance')
    .addToUi();
}

/**
 * Set up the Attendance Report sheet with program name input and formatting
 * Run this once to initialize the sheet
 */
function setupAttendanceReportSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Attendance Report');

  if (!sheet) {
    sheet = ss.insertSheet('Attendance Report', 0);
  }

  // Clear existing content
  sheet.clear();
  sheet.clearFormats();

  // Merge A2:D2
  sheet.getRange('A2:D2').merge();
  const headerCell = sheet.getRange('A2');
  headerCell.setValue('Enter the name of the program for which you are uploading attendance.');
  headerCell.setFontWeight('bold');
  headerCell.setFontSize(10);
  headerCell.setFontFamily('Verdana');
  headerCell.setVerticalAlignment('middle');
  headerCell.setHorizontalAlignment('left');

  // Add "Program Name:" label in A4
  const labelCell = sheet.getRange('A4');
  labelCell.setValue('Program Name:');
  labelCell.setFontWeight('bold');
  labelCell.setFontSize(9);
  labelCell.setFontFamily('Verdana');
  labelCell.setVerticalAlignment('middle');
  labelCell.setHorizontalAlignment('left');

  // Add input cell in B4
  const inputCell = sheet.getRange('B4');
  inputCell.setBackground('#FFFACD'); // Soft yellow
  inputCell.setBorder(null, null, true, null, null, null, '#666666', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  inputCell.setFontFamily('Verdana');
  inputCell.setFontSize(9);
  inputCell.setHorizontalAlignment('left');
  inputCell.setVerticalAlignment('middle');

  // Thin light border bottom of A8:D8
  sheet.getRange('A8:D8').setBorder(null, null, true, null, null, null, '#CCCCCC', SpreadsheetApp.BorderStyle.SOLID);

  // Thin light border on right of D1:D8
  sheet.getRange('D1:D8').setBorder(null, null, null, true, null, null, '#CCCCCC', SpreadsheetApp.BorderStyle.SOLID);

  // Move "Paste here" to A10
  const pasteCell = sheet.getRange('A10');
  pasteCell.setValue('Paste here');
  pasteCell.setBackground('#FFFACD'); // Soft yellow
  pasteCell.setFontFamily('Verdana');
  pasteCell.setFontSize(9);

  // Hide gridlines
  sheet.setHiddenGridlines(true);

  // Set default font for sheet
  const fullRange = sheet.getRange('A1:Z100');
  fullRange.setFontFamily('Verdana');
  fullRange.setFontSize(9);
  fullRange.setFontColor('#333333');

  // Auto-resize columns
  sheet.autoResizeColumns(1, 4);

  SpreadsheetApp.getActiveSpreadsheet().toast('Attendance Report sheet has been set up!', 'Setup Complete', 3);
}

// ===========================
// MAIN PROCESSING FUNCTION
// ===========================

function processAttendance() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sheet = ss.getSheetByName('Attendance Report');

  if (!sheet) {
    ui.alert('Error: "Attendance Report" sheet not found.');
    return;
  }

  // Read program name from B4
  const programNameRaw = sheet.getRange('B4').getValue();
  const programName = String(programNameRaw).trim();

  // Validation
  if (!programName) {
    ui.alert('No Program Name', 'Please enter a program name in cell B4', ui.ButtonSet.OK);
    return;
  }

  if (programName.length > 25) {
    ui.alert('Program Name Too Long', 'Program name must be 25 characters or less.\n\nCurrent length: ' + programName.length, ui.ButtonSet.OK);
    return;
  }

  // Call the processing function
  handleReportTypeSelection(programName);
}

/**
 * Handle the program name and process the attendance data
 * Called by processAttendance after validating the program name input
 */
function handleReportTypeSelection(reportType) {
  if (!reportType) {
    // User cancelled
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // Show processing status
  ss.toast(`Processing ${reportType} attendance report...`, 'Processing', -1);

  // Get the raw data sheet
  let rawSheet = ss.getSheetByName('Attendance Report');
  if (!rawSheet) {
    ui.alert('Error: "Attendance Report" sheet not found.');
    return;
  }

  // Check if there's data pasted
  const lastRow = rawSheet.getLastRow();
  const lastCol = rawSheet.getLastColumn();

  if (lastRow < 2 || lastCol < 4) {
    ui.alert('No data found. Please paste attendance data in cell A5 or below.');
    return;
  }

  // Get all data
  const allData = rawSheet.getRange(1, 1, lastRow, lastCol).getValues();

  // Find where the actual data starts (could be row 1 or row 5)
  const dataStartRow = findDataStartRow(allData);
  if (dataStartRow === -1) {
    ui.alert('Error: Could not find attendance data. Make sure the first row contains dates.');
    return;
  }

  // Extract just the data portion (from dataStartRow onwards)
  const actualData = allData.slice(dataStartRow);

  // Detect month from first row of actual data
  const month = detectMonth(actualData[0]);
  if (!month) {
    ui.alert('Error: Could not detect month from data. Make sure the first row of data contains dates.');
    return;
  }

  // Process the attendance data
  const processedData = processAttendanceData(actualData);

  // Calculate daily averages
  const dailyAverages = calculateDailyAverages(actualData, month, reportType);

  // Update daily averages sheet FIRST (so it positions correctly after Attendance Report)
  updateDailyAveragesSheet(ss, month, reportType, dailyAverages);

  // Create or update monthly summary sheet (positions after Daily Averages)
  const monthlySheet = createOrUpdateMonthlySheet(ss, month, reportType, processedData);

  // Create or update archive sheet
  createOrUpdateArchiveSheet(ss, month, reportType, actualData);

  // Clear the Attendance Report sheet
  clearAttendanceReportSheet(rawSheet);

  // Hide the processing toast
  ss.toast('', '', 1);

  // Show completion message
  const tabName = `${month}-${reportType}`;
  ui.alert(
    'Success!',
    `Attendance data for ${month} (${reportType}) has been processed.\n\n` +
    `• Monthly summary: "${tabName}" tab\n` +
    `• Raw data archived: "${tabName} Archive" tab\n` +
    `• Daily averages updated\n` +
    `• Attendance Report cleared and ready for next report`,
    ui.ButtonSet.OK
  );
}

// ===========================
// HELPER FUNCTIONS
// ===========================

/**
 * Find the row where actual attendance data starts
 * Looks for a row that has dates in the columns (could be row 0 or row 4)
 */
function findDataStartRow(allData) {
  for (let r = 0; r < Math.min(10, allData.length); r++) {
    const row = allData[r];

    // Check if this row has date-like values in the later columns
    // Look for Date objects or date-formatted strings
    for (let c = 3; c < Math.min(20, row.length); c++) {
      const cellValue = row[c];

      // Check if it's a Date object
      if (cellValue instanceof Date && !isNaN(cellValue)) {
        return r;
      }

      // Check if it's a date-formatted string
      const cellString = String(cellValue);
      if (cellString.match(/\d{1,2}-[A-Za-z]{3}/) || cellString.match(/\d{1,2}\/\d{1,2}\/\d{2,4}/)) {
        return r;
      }
    }
  }
  return -1; // Not found
}

/**
 * Detect month from first row of data
 * Handles both Date objects and string dates like "01-Dec", "15-Jan", etc.
 */
function detectMonth(firstRow) {
  const monthNames = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
                       'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

  for (let i = 0; i < firstRow.length; i++) {
    const cellValue = firstRow[i];

    // Check if it's a Date object
    if (cellValue instanceof Date && !isNaN(cellValue)) {
      return monthNames[cellValue.getMonth()];
    }

    // Check if it's a string with date pattern like "01-Dec" or "15-Jan"
    const cellString = String(cellValue);
    const match = cellString.match(/\d{1,2}-([A-Za-z]{3})/);
    if (match) {
      return match[1]; // Return just the month abbreviation
    }

    // Check if it's a formatted date string like "10/1/2024" or "Oct 1"
    if (cellString.match(/\d{1,2}\/\d{1,2}\/\d{2,4}/)) {
      const date = new Date(cellString);
      if (!isNaN(date)) {
        return monthNames[date.getMonth()];
      }
    }
  }
  return null;
}

/**
 * Process attendance data and calculate days attended per student
 */
function processAttendanceData(allData) {
  const results = [];

  // Start from row 3 (index 2) - skip row 1 (dates) and row 2 (AM/PM/Full headers)
  for (let r = 2; r < allData.length; r++) {
    const row = allData[r];

    // Skip empty rows
    if (!row[0] && !row[1] && !row[2]) continue;

    const firstName = row[0] || '';
    const lastName = row[1] || '';
    const location = row[2] || '';

    // Skip if no name
    if (!firstName && !lastName) continue;

    // Count attended days
    let daysAttended = 0;

    // Starting from column D (index 3), we have AM, PM, Full pattern
    // Column D = AM, E = PM, F = Full, G = AM, H = PM, I = Full, etc.
    for (let c = 3; c < row.length; c += 3) {
      const amValue = Number(row[c]) || 0;
      const pmValue = Number(row[c + 1]) || 0;
      const fullValue = Number(row[c + 2]) || 0;

      // Count this date as attended if:
      // - Full > 0 (full day attendance)
      // - OR both AM > 0 AND PM > 0 (counts as 1 day, not 2)
      // - OR just AM > 0 (half day)
      // - OR just PM > 0 (half day)

      if (fullValue > 0) {
        daysAttended++;
      } else if (amValue > 0 && pmValue > 0) {
        daysAttended++; // Both AM and PM at same site = 1 day
      } else if (amValue > 0 || pmValue > 0) {
        daysAttended++; // Just AM or just PM = 1 day
      }
    }

    // Add to results: Last Name, First Name | Attended Days | Location
    results.push({
      lastName: lastName,
      firstName: firstName,
      fullName: `${lastName}, ${firstName}`,
      daysAttended: daysAttended,
      location: location
    });
  }

  return results;
}

/**
 * Calculate daily average attendance by site
 */
function calculateDailyAverages(allData, month, reportType) {
  // Dictionary to hold Site -> (Date -> Count)
  const siteData = {};

  // Get all dates from row 1 (index 0)
  // Dates appear in columns every 3 positions starting from column D (index 3)
  // The pattern is: AM (index 3), PM (index 4), Full (index 5), then repeats
  // Check all columns in each group to find which one has the date
  const dates = [];
  for (let c = 3; c < allData[0].length; c += 3) {
    // Check all three columns in this group for a date
    let dateValue = null;
    for (let offset = 0; offset < 3; offset++) {
      const testValue = allData[0][c + offset];
      if (testValue instanceof Date || (typeof testValue === 'string' && testValue.match(/\d{1,2}-[A-Za-z]{3}/))) {
        dateValue = testValue;
        break;
      }
    }

    if (dateValue) {
      dates.push({col: c, date: String(dateValue)});
    }
  }

  // Process each student row (start from row 3, index 2)
  for (let r = 2; r < allData.length; r++) {
    const row = allData[r];
    const location = row[2] || '';

    if (!location) continue;

    // Initialize site in dictionary
    if (!siteData[location]) {
      siteData[location] = {};
    }

    // Check PM attendance for each date
    for (let i = 0; i < dates.length; i++) {
      const dateInfo = dates[i];
      const pmValue = Number(row[dateInfo.col + 1]) || 0;
      const fullValue = Number(row[dateInfo.col + 2]) || 0;

      // Count PM or Full attendance
      if (pmValue > 0 || fullValue > 0) {
        if (!siteData[location][dateInfo.date]) {
          siteData[location][dateInfo.date] = 0;
        }
        siteData[location][dateInfo.date]++;
      }
    }
  }

  // Calculate averages and peak attendance
  const averages = [];
  for (const site in siteData) {
    const dateCounts = siteData[site];
    const dateKeys = Object.keys(dateCounts);

    if (dateKeys.length === 0) continue;

    const sum = dateKeys.reduce((total, date) => total + dateCounts[date], 0);
    const average = Math.ceil(sum / dateKeys.length); // Round up

    // Find peak attendance (highest single day)
    const peak = Math.max(...Object.values(dateCounts));

    averages.push({
      site: site,
      month: month,
      reportType: reportType,
      average: average,
      peak: peak
    });
  }

  return averages;
}

/**
 * Create or update monthly summary sheet
 */
function createOrUpdateMonthlySheet(ss, month, reportType, processedData) {
  const tabName = `${month}-${reportType}`;
  let sheet = ss.getSheetByName(tabName);
  const dailyAvgSheet = ss.getSheetByName('Daily Averages');

  if (!sheet) {
    // Create new sheet - position it after Daily Averages
    if (dailyAvgSheet) {
      // Insert after Daily Averages
      sheet = ss.insertSheet(tabName, dailyAvgSheet.getIndex() + 1);
    } else {
      // Daily Averages doesn't exist yet, just create it
      sheet = ss.insertSheet(tabName);
    }
  } else {
    // Clear existing data
    sheet.clear();

    // Move sheet to position right after Daily Averages (so newest reprocessed reports stay closest)
    if (dailyAvgSheet) {
      sheet.activate();
      ss.moveActiveSheet(dailyAvgSheet.getIndex() + 1);
    }
  }

  // Set up the top section (rows 1-4)
  sheet.getRange('A1').setValue('Month:');
  sheet.getRange('B1').setValue(`${month} (${reportType})`);

  sheet.getRange('A2').setValue('Year:');
  const yearCell = sheet.getRange('B2');
  yearCell.setBackground('#FFFACD'); // Soft yellow
  yearCell.setHorizontalAlignment('left');
  yearCell.setBorder(true, true, true, true, false, false, '#CCCCCC', SpreadsheetApp.BorderStyle.DOTTED);

  sheet.getRange('A3').setValue('Eligible Days:');
  const eligibleCell = sheet.getRange('B3');
  eligibleCell.setBackground('#FFFACD'); // Soft yellow
  eligibleCell.setHorizontalAlignment('left');
  eligibleCell.setBorder(true, true, true, true, false, false, '#CCCCCC', SpreadsheetApp.BorderStyle.DOTTED);

  // Bold the labels
  sheet.getRange('A1:A3').setFontWeight('bold');

  // Headers in row 5
  const headers = ['Last Name, First Name', 'Attended Days', 'Location', 'Attend %'];
  sheet.getRange(5, 1, 1, headers.length).setValues([headers]);

  // Add data starting at row 6
  const dataRows = [];
  for (let i = 0; i < processedData.length; i++) {
    const student = processedData[i];

    // Formula for attendance %: = B6 / $B$3 (adjusted for new row positions)
    const attendFormula = `=IF($B$3>0, B${6+i}/$B$3, 0)`;

    dataRows.push([
      student.fullName,
      student.daysAttended,
      student.location,
      attendFormula
    ]);
  }

  if (dataRows.length > 0) {
    sheet.getRange(6, 1, dataRows.length, 4).setValues(dataRows);
  }

  // Apply formatting
  applyMonthlySheetFormatting(sheet, dataRows.length);

  return sheet;
}

/**
 * Create or update archive sheet (hidden)
 */
function createOrUpdateArchiveSheet(ss, month, reportType, rawData) {
  const archiveName = `${month}-${reportType} Archive`;
  let sheet = ss.getSheetByName(archiveName);

  if (!sheet) {
    sheet = ss.insertSheet(archiveName);
  } else {
    sheet.clear();
  }

  // Paste raw data
  if (rawData.length > 0) {
    sheet.getRange(1, 1, rawData.length, rawData[0].length).setValues(rawData);
  }

  // Hide the sheet
  sheet.hideSheet();

  return sheet;
}

/**
 * Update daily averages sheet
 */
function updateDailyAveragesSheet(ss, month, reportType, averages) {
  let sheet = ss.getSheetByName('Daily Averages');

  if (!sheet) {
    // Create new sheet - position it to the right of Attendance Report
    const attendanceReportSheet = ss.getSheetByName('Attendance Report');
    if (attendanceReportSheet) {
      // Insert after Attendance Report
      sheet = ss.insertSheet('Daily Averages', attendanceReportSheet.getIndex() + 1);
    } else {
      // Just create it at the end
      sheet = ss.insertSheet('Daily Averages');
    }

    // Add headers in row 4 (now with 5 columns including Program)
    sheet.getRange(4, 1, 1, 5).setValues([['Site', 'Month', 'Program', 'Average Attendance', 'Peak Attendance']]);
    applyDailyAveragesFormatting(sheet, 0);
  }

  // Find the last row with data
  let lastRow = sheet.getLastRow();
  if (lastRow < 4) lastRow = 4;

  // Check if this month-report combo already exists and remove old data
  const numCols = Math.max(sheet.getLastColumn(), 5);
  const existingData = sheet.getRange(5, 1, Math.max(1, lastRow - 4), numCols).getValues();

  // Filter out rows that match this month AND report type
  const filteredData = existingData.filter(row => {
    const rowMonth = row[1];
    const rowReportType = row[2];
    return !(rowMonth === month && rowReportType === reportType);
  });

  // Clear data area and rewrite
  if (lastRow > 4) {
    sheet.getRange(5, 1, lastRow - 4, numCols).clear();
  }

  // Combine filtered data with new averages
  const allData = filteredData.filter(row => row[0]); // Remove empty rows
  for (const avg of averages) {
    allData.push([avg.site, avg.month, avg.reportType, avg.average, avg.peak]);
  }

  // Write all data
  if (allData.length > 0) {
    sheet.getRange(5, 1, allData.length, 5).setValues(allData);
  }

  // Apply formatting
  applyDailyAveragesFormatting(sheet, allData.length);

  return sheet;
}

/**
 * Clear the Attendance Report sheet
 */
function clearAttendanceReportSheet(sheet) {
  // Only clear data from row 10 onwards (preserve header and program name input in rows 1-9)
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow >= 10) {
    sheet.getRange(10, 1, lastRow - 9, lastCol).clear();
  }

  // Re-add "Paste here" cell in A10
  const pasteCell = sheet.getRange('A10');
  pasteCell.setValue('Paste here');
  pasteCell.setBackground('#FFFACD'); // Soft yellow
  pasteCell.setFontFamily('Verdana');
  pasteCell.setFontSize(9);

  // Clear the program name input cell
  sheet.getRange('B4').clear();
  sheet.getRange('B4').setBackground('#FFFACD'); // Restore soft yellow background
  sheet.getRange('B4').setBorder(null, null, true, null, null, null, '#666666', SpreadsheetApp.BorderStyle.SOLID_MEDIUM); // Restore border
}

// ===========================
// FORMATTING FUNCTIONS
// ===========================

/**
 * Apply formatting to monthly sheet
 */
function applyMonthlySheetFormatting(sheet, dataRowCount) {
  // Remove gridlines
  sheet.setHiddenGridlines(true);

  // Set font for entire sheet
  const fullRange = sheet.getRange('A1:Z1000');
  fullRange.setFontFamily('Verdana');
  fullRange.setFontSize(9);
  fullRange.setFontColor('#333333');

  // Bold headers in row 5
  const headerRange = sheet.getRange(5, 1, 1, 4);
  headerRange.setFontWeight('bold');

  // Add thicker border under header row
  headerRange.setBorder(null, null, true, null, null, null, '#666666', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // Format attendance % column as percentage
  if (dataRowCount > 0) {
    sheet.getRange(6, 4, dataRowCount, 1).setNumberFormat('0%');
  }

  // Center the attended days data (column B, rows 6+)
  if (dataRowCount > 0) {
    sheet.getRange(6, 2, dataRowCount, 1).setHorizontalAlignment('center');
  }

  // Add borders around data area (A5 to D[last row])
  if (dataRowCount > 0) {
    const dataRange = sheet.getRange(5, 1, dataRowCount + 1, 4);

    // Outer border - soft black
    dataRange.setBorder(true, true, true, true, false, false, '#666666', SpreadsheetApp.BorderStyle.SOLID);

    // Inner borders - soft gray dotted
    dataRange.setBorder(null, null, null, null, true, true, '#CCCCCC', SpreadsheetApp.BorderStyle.DOTTED);

    // Re-apply thicker bottom border on header row (overwrites the dotted one)
    headerRange.setBorder(null, null, true, null, null, null, '#666666', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  }

  // Add conditional formatting for attendance % < 70%
  if (dataRowCount > 0) {
    const attendPercentRange = sheet.getRange(6, 4, dataRowCount, 1);

    const rule = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(0.70)
      .setBackground('#FFE5E5') // Light red
      .setRanges([attendPercentRange])
      .build();

    const rules = sheet.getConditionalFormatRules();
    rules.push(rule);
    sheet.setConditionalFormatRules(rules);
  }

  // Enable filter on header row - check if filter already exists first
  if (dataRowCount > 0) {
    const existingFilter = sheet.getFilter();
    if (existingFilter) {
      existingFilter.remove();
    }
    sheet.getRange(5, 1, dataRowCount + 1, 4).createFilter();
  }

  // Set specific column widths to ensure content is fully visible
  sheet.setColumnWidth(1, 200); // Last Name, First Name
  sheet.setColumnWidth(2, 120); // Attended Days
  sheet.setColumnWidth(3, 180); // Location
  sheet.setColumnWidth(4, 100); // Attend %
}

/**
 * Apply formatting to daily averages sheet
 */
function applyDailyAveragesFormatting(sheet, dataRowCount) {
  // Remove gridlines
  sheet.setHiddenGridlines(true);

  // Set font for entire sheet
  const fullRange = sheet.getRange('A1:Z1000');
  fullRange.setFontFamily('Verdana');
  fullRange.setFontSize(9);
  fullRange.setFontColor('#333333');

  // Bold headers in row 4 (now 5 columns)
  const headerRange = sheet.getRange(4, 1, 1, 5);
  headerRange.setFontWeight('bold');

  // Add thicker border under header row
  headerRange.setBorder(null, null, true, null, null, null, '#666666', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // Add borders around data area (now 5 columns)
  if (dataRowCount > 0) {
    const dataRange = sheet.getRange(4, 1, dataRowCount + 1, 5);

    // Outer border - soft black
    dataRange.setBorder(true, true, true, true, false, false, '#666666', SpreadsheetApp.BorderStyle.SOLID);

    // Inner borders - soft gray dotted
    dataRange.setBorder(null, null, null, null, true, true, '#CCCCCC', SpreadsheetApp.BorderStyle.DOTTED);

    // Re-apply thicker bottom border on header row
    headerRange.setBorder(null, null, true, null, null, null, '#666666', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  }

  // Enable filter on header row - check if filter already exists first
  if (dataRowCount > 0) {
    const existingFilter = sheet.getFilter();
    if (existingFilter) {
      existingFilter.remove();
    }
    sheet.getRange(4, 1, dataRowCount + 1, 5).createFilter();
  }

  // Set specific column widths to ensure headers are fully visible
  sheet.setColumnWidth(1, 180); // Site
  sheet.setColumnWidth(2, 80);  // Month
  sheet.setColumnWidth(3, 200); // Program (max 25 chars, so needs room)
  sheet.setColumnWidth(4, 150); // Average Attendance
  sheet.setColumnWidth(5, 150); // Peak Attendance
}
