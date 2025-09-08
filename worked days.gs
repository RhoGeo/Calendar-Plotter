function plotClockoutsAsSeparateCalendars() {
  const timeOutReportSheetName = "Time-out Report";
  const payslipSheetName = "Payslip";
  const expensesSheetName = "Riders";
  const calendarStartCell = "A20";

  // Fixed summary output rows
  const summaryStartRow = 15;
  const summaryStartCol = 1; // Column A
  const currencyFormat = "₱#,##0";

  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const timeOutSheet = spreadsheet.getSheetByName(timeOutReportSheetName);
    const payslipSheet = spreadsheet.getSheetByName(payslipSheetName);
    const expensesSheet = spreadsheet.getSheetByName(expensesSheetName);

    if (!timeOutSheet || !payslipSheet || !expensesSheet) {
      SpreadsheetApp.getUi().alert("Error: One or more required sheets were not found.");
      return;
    }

    const startDate = payslipSheet.getRange("B5").getValue();
    const endDate = payslipSheet.getRange("B6").getValue();

    if (!startDate || !endDate || !(startDate instanceof Date) || !(endDate instanceof Date)) {
      SpreadsheetApp.getUi().alert("Error: Invalid date range in the Payslip sheet.");
      return;
    }

    endDate.setHours(23, 59, 59, 999);

    const timeOutDataRange = timeOutSheet.getRange(2, 1, timeOutSheet.getLastRow() - 1, timeOutSheet.getLastColumn());
    const allTimeOutData = timeOutDataRange.getValues();

    const accounts = ["OCW Etcobanez", "OCW Mendoza", "OCW Seco"];
    const timezone = spreadsheet.getSpreadsheetTimeZone();
    let currentPlotRow = expensesSheet.getRange(calendarStartCell).getRow();
    const calendarStartCol = expensesSheet.getRange(calendarStartCell).getColumn();

    // Clear previous calendar content
    expensesSheet.getRange(currentPlotRow, calendarStartCol, expensesSheet.getMaxRows() - currentPlotRow, expensesSheet.getMaxColumns() - calendarStartCol).clearContent().clearFormat();

    // Clear the summary area: A15:C18
    expensesSheet.getRange(summaryStartRow, summaryStartCol, accounts.length + 1, 3).clearContent().clearFormat();

    let totalAmount = 0;

    accounts.forEach((accountName, index) => {
      const clockOutsByDate = {};

      allTimeOutData.forEach(row => {
        const timestamp = new Date(row[0]);
        const riderName = row[1];
        const account = row[2];

        if (/pp/i.test(riderName)) return;

        if (timestamp >= startDate && timestamp <= endDate && account === accountName) {
          const dateKey = Utilities.formatDate(timestamp, timezone, "M/d/yyyy");
          if (!clockOutsByDate[dateKey]) {
            clockOutsByDate[dateKey] = [];
          }
          clockOutsByDate[dateKey].push(riderName);
        }
      });

      const highlightCount = Object.keys(clockOutsByDate).length;
      const amount = highlightCount * 2300;
      totalAmount += amount;

      // --- Write summary to A15:C17 ---
      const summaryRow = summaryStartRow + index;
      expensesSheet.getRange(summaryRow, 1).setValue(`${accountName} Clock-outs`);
      expensesSheet.getRange(summaryRow, 2).setValue(highlightCount);
      expensesSheet.getRange(summaryRow, 3).setValue(amount).setNumberFormat(currencyFormat);

      // --- Write the calendar ---
      expensesSheet.getRange(currentPlotRow, calendarStartCol)
        .setValue(`${accountName} Clock-outs`)
        .setFontWeight("bold")
        .setFontSize(12);

      expensesSheet.getRange(currentPlotRow, calendarStartCol + 1)
        .setValue(highlightCount)
        .setFontWeight("bold")
        .setFontSize(12);

      expensesSheet.getRange(currentPlotRow, calendarStartCol + 2)
        .setValue(amount)
        .setNumberFormat(currencyFormat)
        .setFontWeight("bold")
        .setFontSize(12);

      currentPlotRow++;

      const startCalendarDate = new Date(startDate);
      startCalendarDate.setDate(startDate.getDate() - startCalendarDate.getDay()); // Sunday start

      const calendarHeader = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
      expensesSheet.getRange(currentPlotRow, calendarStartCol, 1, 7).setValues([calendarHeader]).setFontWeight("bold");
      currentPlotRow++;

      let currentDate = new Date(startCalendarDate);

      while (currentDate <= endDate || currentDate.getDay() !== 0) {
        const weekData = [];
        const rowBackgrounds = [];

        for (let i = 0; i < 7; i++) {
          const dateKey = Utilities.formatDate(currentDate, timezone, "M/d/yyyy");
          const dayOfMonth = currentDate.getDate();

          let cellText = dayOfMonth;
          let cellColor = "#ffffff";

          if (currentDate >= startDate && currentDate <= endDate) {
            if (clockOutsByDate[dateKey]) {
              cellText = `${dayOfMonth}\n${clockOutsByDate[dateKey].join(', ')}`;
              cellColor = "#c9daf8";
            }
          }

          weekData.push(cellText);
          rowBackgrounds.push(cellColor);
          currentDate.setDate(currentDate.getDate() + 1);
        }

        const rowRange = expensesSheet.getRange(currentPlotRow, calendarStartCol, 1, 7);
        rowRange.setValues([weekData]);
        rowRange.setBackgrounds([rowBackgrounds]);
        rowRange.setVerticalAlignment("top").setWrap(true);
        expensesSheet.setRowHeight(currentPlotRow, 100);
        currentPlotRow++;
      }

      currentPlotRow += 2;
    });

    // Total amount in C18
    expensesSheet.getRange(summaryStartRow + accounts.length, 3)
      .setValue(totalAmount)
      .setNumberFormat(currencyFormat)
      .setFontWeight("bold")
      .setFontSize(12)
      .setBackground("#d9ead3");

    // Set calendar column widths
    for (let i = 0; i < 7; i++) {
      expensesSheet.setColumnWidth(calendarStartCol + i, 120);
    }

    expensesSheet.setRowHeight(expensesSheet.getRange(calendarStartCell).getRow() + 1, 30);

    SpreadsheetApp.getUi().alert("Clock-out calendars and summaries generated successfully.");
  } catch (e) {
    SpreadsheetApp.getUi().alert("An error occurred: " + e.message);
  }
}
