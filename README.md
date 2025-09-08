# Clock‑Out Calendars & Summary (Google Apps Script)

Generate **per‑account monthly calendars** that highlight rider clock‑outs and a **cost summary** for a chosen date range — all inside your Google Sheet.

This README documents the `plotClockoutsAsSeparateCalendars()` function you shared and how to use/customize it in production.

---

## ✨ What it does

Given three sheets — **Time-out Report**, **Payslip**, and **Riders** — the script:

1. Reads the **start** and **end** dates from `Payslip!B5:B6`.
2. Scans **Time-out Report** rows for clock‑outs within that date range.
3. Groups clock‑outs **per account** and **per day** (excludes rider names that contain `pp`, case‑insensitive).
4. For each configured account:
   - Writes a **summary row** showing highlighted days and **amount = highlighted days × 2300**.
   - Renders a **calendar** (Sunday‑start weeks) where days that have clock‑outs are **light‑blue** and list the rider names for that day.
5. Writes the grand **Total** amount at the end of the summary, formats currency as `₱#,##0`, and sizes calendar columns and rows for readability.
6. Clears all prior output before writing new results to avoid duplication/overlap.

---

## 📄 Sheet & Data Model

### Required sheets
- **`Time-out Report`** — source of clock‑out records
- **`Payslip`** — supplies the date range
- **`Riders`** — the output canvas for the calendars + summary

### Expected columns in **Time-out Report**
The function expects (at minimum) these first three columns:

| Col | Field        | Type                     | Notes                           |
|-----|--------------|--------------------------|----------------------------------|
| A   | Timestamp    | Date/Time (Google Sheets)| When the rider clocked out      |
| B   | Rider Name   | Text                     | Rider identity (filters out `pp`)|
| C   | Account      | Text                     | Must match configured accounts   |

> Rows where the Rider Name contains `pp` (e.g., `PP Test`) are **ignored** via `/pp/i` filter.

### Required cells in **Payslip**
- `B5` → **Start Date** (date value)
- `B6` → **End Date** (date value; the script treats this as inclusive by setting `23:59:59.999`)

### Output location in **Riders**
- **Summary block**: `A15:C18` (cleared and rewritten)
- **Calendars start**: from **`A20`** downward (fully cleared to the sheet’s lower/right edge before rendering)

> The script uses your spreadsheet’s **Time zone** (`File → Settings → Locale & Time zone`) to bucket days.

---

## 🧠 How it works (logic overview)

1. **Validation**
   - Ensures the three sheets exist; alerts if not.
   - Validates `Payslip!B5` and `B6` are Dates; alerts if invalid.

2. **Load & filter data**
   - Pulls all rows from `Time-out Report` (data range from row 2 across all columns).
   - Applies date range filter (`>= startDate && <= endDate`), account match, and excludes riders with `/pp/i`.

3. **Aggregate per account / per day**
   - Keyed by day (`M/d/yyyy` in the sheet’s time zone).
   - For each day, collects a list of **rider names** that clocked out for that account.

4. **Summaries**
   - `highlightCount = number of unique days with at least one clock‑out`.
   - `amount = highlightCount × 2300` (configurable).
   - Grand `totalAmount` accumulated across all accounts.

5. **Calendar rendering**
   - Starts from the **Sunday** of the week containing the start date.
   - Iterates week by week until it passes the end date and ends on a Saturday.
   - Each day’s cell shows the **day number** and (if any) the **comma‑separated rider list**.
   - Highlight background: `#c9daf8` for days within the date range that have clock‑outs.

6. **Styling & layout**
   - Weekday header row (`Sunday ... Saturday`) in **bold**.
   - Calendar cells `wrap` text; row height set to `100` for visibility.
   - Calendar columns set to width `120`.
   - Total cell has background `#d9ead3` and bold currency formatting `₱#,##0`.

---

## 🔧 Configuration (edit these constants)

Inside the function:

```js
const timeOutReportSheetName = "Time-out Report";
const payslipSheetName = "Payslip";
const expensesSheetName = "Riders";
const calendarStartCell = "A20";

const summaryStartRow = 15;   // Summary starts at A15
const summaryStartCol = 1;    // Column A
const currencyFormat = "₱#,##0";

const accounts = ["OCW Etcobanez", "OCW Mendoza", "OCW Seco"]; // Edit to your accounts

// Rate per highlighted day
const DAILY_RATE = 2300; // <— replace 2300 everywhere if you prefer, or compute dynamically
```

**Common customizations**
- **Accounts**: Change the `accounts` array to your exact account labels.
- **Rate**: Replace `2300` with your current rate (or compute per account).
- **Currency**: Set `currencyFormat` to your desired locale pattern.
- **Output area**: Move `calendarStartCell` if you want calendars elsewhere.
- **Summary area**: Adjust `summaryStartRow`/`summaryStartCol` if needed.
- **Filter**: Remove or change the `/pp/i` rider filter if it’s not relevant.

> Note: In your original code, the multiplier `2300` is inline. You may factor it into a `DAILY_RATE` constant, as shown above, and apply it when computing `amount`.

---

## ✅ Usage

1. **Open your Google Sheet** that contains the three tabs above.
2. Go to **Extensions → Apps Script** and add the function below (or paste your existing one).
3. Make sure your **spreadsheet time zone** is correct (File → Settings).
4. Set `Payslip!B5` and `Payslip!B6` to your intended start/end dates.
5. **Run** `plotClockoutsAsSeparateCalendars()` from Apps Script (or attach it to a button).
6. Grant authorization on first run.

### Optional: add a custom menu
If you prefer a one‑click menu inside the Sheet, add this alongside your function:

```js
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Clock-outs")
    .addItem("Generate Calendars & Summary", "plotClockoutsAsSeparateCalendars")
    .addToUi();
}
```

---

## 🔐 Permissions

The script uses:
- `SpreadsheetApp` (read/write your spreadsheet, format cells, display UI alerts)
- `Utilities` (for date formatting)

You’ll be prompted to allow these the first time you run the script.

---

## 🧪 Testing checklist

- [ ] `Time-out Report` has at least three columns: **Timestamp**, **Rider Name**, **Account**.
- [ ] `Payslip!B5` and `B6` are valid date values (not plain text).
- [ ] `Riders` has space starting at `A20` (or update `calendarStartCell`).
- [ ] The **accounts** in your data exactly match entries in the `accounts` array.
- [ ] Time zone in **File → Settings** matches your intended locale.

---

## 🧰 Full function (reference)

> This is the function exactly as documented. You can paste it into Apps Script and tweak the constants to your needs.

```js
function plotClockoutsAsSeparateCalendars() {
  const timeOutReportSheetName = "Time-out Report";
  const payslipSheetName = "Payslip";
  const expensesSheetName = "Riders";
  const calendarStartCell = "A20";

  const summaryStartRow = 15;
  const summaryStartCol = 1;
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

    expensesSheet.getRange(currentPlotRow, calendarStartCol, expensesSheet.getMaxRows() - currentPlotRow, expensesSheet.getMaxColumns() - calendarStartCol).clearContent().clearFormat();
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

      const summaryRow = summaryStartRow + index;
      expensesSheet.getRange(summaryRow, 1).setValue(`${accountName} Clock-outs`);
      expensesSheet.getRange(summaryRow, 2).setValue(highlightCount);
      expensesSheet.getRange(summaryRow, 3).setValue(amount).setNumberFormat(currencyFormat);

      expensesSheet.getRange(currentPlotRow, calendarStartCol).setValue(`${accountName} Clock-outs`).setFontWeight("bold").setFontSize(12);
      expensesSheet.getRange(currentPlotRow, calendarStartCol + 1).setValue(highlightCount).setFontWeight("bold").setFontSize(12);
      expensesSheet.getRange(currentPlotRow, calendarStartCol + 2).setValue(amount).setNumberFormat(currencyFormat).setFontWeight("bold").setFontSize(12);
      currentPlotRow++;

      const startCalendarDate = new Date(startDate);
      startCalendarDate.setDate(startDate.getDate() - startCalendarDate.getDay());

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

    expensesSheet.getRange(summaryStartRow + accounts.length, 3).setValue(totalAmount).setNumberFormat(currencyFormat).setFontWeight("bold").setFontSize(12).setBackground("#d9ead3");

    for (let i = 0; i < 7; i++) {
      expensesSheet.setColumnWidth(calendarStartCol + i, 120);
    }

    expensesSheet.setRowHeight(expensesSheet.getRange(calendarStartCell).getRow() + 1, 30);

    SpreadsheetApp.getUi().alert("Clock-out calendars and summaries generated successfully.");
  } catch (e) {
    SpreadsheetApp.getUi().alert("An error occurred: " + e.message);
  }
}
```

---

## 🧯 Troubleshooting

- **“One or more required sheets were not found.”**  
  Ensure tabs are named exactly: `Time-out Report`, `Payslip`, `Riders` (or update the constants).

- **“Invalid date range in the Payslip sheet.”**  
  `Payslip!B5` and `B6` must be genuine Date values (not text). Use **Data → Data validation** to enforce Date.

- **No highlights appear**  
  - Confirm account names in your data match the `accounts` array exactly (case/spacing).
  - Confirm `Time-out Report!A` contains real date/times.
  - Check time zone under **File → Settings** (date bucketing depends on it).

- **Output overwrites other content**  
  Shift `calendarStartCell` lower/right (e.g., `A30`) or move other content away. The script clears from that cell to the sheet’s bottom/right edge.

- **You don’t want to exclude `pp` riders**  
  Remove `if (/pp/i.test(riderName)) return;` from the loop.

---

## ⏱️ Performance notes

- The script reads **all rows** in `Time-out Report` from row 2 down. If your sheet is huge, consider limiting the read range to a known data region or archiving old rows.
- Most time is spent formatting output; avoid extremely large date ranges if you don’t need them.

---

## 🔄 Versioning & Changelog

- **v1.0.0** — Initial public README describing the provided implementation.

---

## 📜 License

MIT. You’re free to use, modify, and distribute. Consider adding your copyright line.

```text
MIT License

Copyright (c) 2025 <Your Name>

Permission is hereby granted, free of charge, to any person obtaining a copy...
```

---

## 🙌 Credits

Authored from your provided Apps Script. Documentation by ChatGPT.
