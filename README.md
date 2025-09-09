
# Google Sheets Clock‑outs Toolkit (Apps Script)

Utilities for delivery/rider operations in Google Sheets:
- **Name formatter** for `Time-out Report!B2:B` (proper case with smart rules).
- **`SpellPeso` custom function** to convert numbers into Philippine Peso words.
- **Clock‑out calendars & summaries** that render weekly calendars with POD totals and a per‑day amount computation.

> Works as a single Apps Script file bound to your Google Sheet. No add-ons required.

---

## Contents

- [`formatNamesAndCleanSpaces`](#formatnamesandcleanspaces)
- [`SpellPeso`](#spellpeso)
- [`plotClockoutsAsSeparateCalendars`](#plotclockoutsasseparatecalendars)
- [Sheet setup](#sheet-setup)
- [Install](#install)
- [Usage](#usage)
- [How calculations work](#how-calculations-work)
- [Customize](#customize)
- [Troubleshooting](#troubleshooting)
- [FAQ](#faq)
- [License](#license)

---

## `formatNamesAndCleanSpaces`

Cleans and proper‑cases rider names in **`Time-out Report!B2:B`**.

**What it does**

- Replaces non‑breaking spaces and collapses multiple spaces.
- Proper‑cases tokens including special cases:
  - `O'Connor` → `O'Connor`
  - Hyphenated names like `ana-marie` → `Ana-Marie`
  - `Mc`/`Mac` patterns like `mcdonald` → `McDonald`, `macarthur` → `MacArthur`
- Keeps suffixes **uppercase** (with optional trailing dot kept): `JR`, `SR`, `II`, `III`, `IV`, `VI`, `VII`, `VIII`, `IX`, `X`.

**Run it:**  
`Extensions → Apps Script → Run → formatNamesAndCleanSpaces`  
or add a button drawing and assign this function.

---

## `SpellPeso`

Custom function you can use in cells to spell out Philippine Peso amounts in words.

**Example**

```gs
=SpellPeso(1234.56)
```

**Output**

```
One Thousand Two Hundred Thirty Four Pesos and Fifty Six Centavos Only
```

- `0` → `Zero Pesos Only`
- Correctly pluralizes `Peso`/`Pesos` and `Centavo`/`Centavos`.

---

## `plotClockoutsAsSeparateCalendars`

Generates **weekly calendars** and a **summary** under the `Riders` sheet from a date range in `Payslip`.

A custom menu is added on open:
```
Clock-outs » Plot calendars
```

### Inputs

- **Time-out Report** (source)
  - Column **A**: Date/Time (Clock‑out timestamp)
  - Column **B**: Rider Name
  - Column **C**: Account (e.g., `OCW Etcobanez`, `OCW Mendoza`, `OCW Seco`)
  - Column **D**: POD (number). If multiple rows per rider/day exist, the **max per rider/day** is used.
- **Payslip** (controls range)
  - **B5**: Start date
  - **B6**: End date
- **Riders** (output)
  - Calendars and a compact summary are rendered here starting **A20** and **row 15** for the summary block.

### Behavior

- Groups by **Account → Date → Rider**.
- For each **date**, collects riders that clocked out and sums their **per‑rider max POD** that day.
- Colors a date cell:
  - **Red** `#f4cccc` if **exactly 1 or 3 riders** worked (under/over staffing signal).
  - **Blue** `#c9daf8` otherwise.
- Produces a Summary table (Days, Total POD, Amount) and a grand total at the bottom.

---

## Sheet setup

Create or verify these sheets:

- `Time-out Report`
- `Payslip` with:
  - `B5` = start date
  - `B6` = end date
- `Riders` as the output canvas

> Spreadsheet time zone should match your operations (e.g., **File → Settings → Time zone**).

Expected minimal columns in **Time-out Report**:

| Col | Meaning          | Notes                          |
|-----|------------------|--------------------------------|
| A   | Timestamp        | Date/Time value                |
| B   | Rider Name       | Will be cleaned/proper‑cased   |
| C   | Account          | Must match an entry in `accounts` |
| D   | POD              | Number; per rider/day **max** is used |

---

## Install

1. Open your Google Sheet.
2. **Extensions → Apps Script**.
3. Replace the editor content with the file from this repo (see below to download).
4. Click **Save**.
5. Back in the sheet, reload to see **Clock-outs** menu.
6. On first run, accept any permission prompts (spreadsheet access only).

---

## Usage

### 1) Format names
- Run `formatNamesAndCleanSpaces` to normalize `Time-out Report!B2:B`.

### 2) Peso in words
- In any cell: `=SpellPeso(1234.56)`

### 3) Plot calendars & summary
- Set **Payslip!B5** and **Payslip!B6** to your date range.
- Click **Clock-outs → Plot calendars**.
- Calendars appear in **Riders** starting at **A20** with a weekly header; summary appears near **row 15**.

---

## How calculations work

- **Accounts processed:**  
  Defined inside the script:
  ```js
  const accounts = ["OCW Etcobanez", "OCW Mendoza", "OCW Seco"];
  ```

- **Per day per account:**
  - Identify all riders who clocked out on that day.
  - For each rider, take the **maximum** POD logged that day.
  - **Date POD** = sum of these rider maxima.

- **Color rule:**  
  - Exactly **1** or **3** riders → red `#f4cccc` (under/over staffing).
  - Otherwise → blue `#c9daf8`.

- **Summary math:**  
  - **Days** = number of dates that had at least one rider for that account.
  - **Total POD** = sum of **Date POD** across those days.
  - **Amount** = `Days × 2300` (configurable).

- **Grand totals** across all accounts are appended at the end of the summary.

---

## Customize

Open `plotClockoutsAsSeparateCalendars` and tweak:

```js
const timeOutReportSheetName = "Time-out Report";
const payslipSheetName       = "Payslip";
const expensesSheetName      = "Riders";
const calendarStartCell      = "A20";

const summaryStartRow = 15;
const summaryStartCol = 1;
const currencyFormat  = "₱#,##0";

const accounts = ["OCW Etcobanez", "OCW Mendoza", "OCW Seco"];
```

Other options inside the function:
- **Daily rate** per account/day:
  ```js
  const amount = highlightCount * 2300;
  ```
- **Colors**:
  ```js
  // red when exactly 1 or 3 riders, else blue
  const red  = "#f4cccc";
  const blue = "#c9daf8";
  ```
- **Skip pattern** (rows where rider name matches `/pp/i` are ignored). Adjust or remove as needed.

---

## Troubleshooting

- **“Sheet not found”**: Ensure the sheet names match the constants in the script.
- **“Invalid date range (B5–B6)”**: Make sure `Payslip!B5` and `B6` contain valid dates.
- **No data rows**: Check that `Time-out Report` has rows starting at **row 2** and that column order is correct.
- **Wrong totals**: Confirm POD values are numeric and that duplicates per rider/day are intentional (the **max** per rider/day is used).
- **Timezone drift**: Set the spreadsheet time zone (`File → Settings`) to your local time (e.g., Asia/Manila).

---

## FAQ

**Q: How do I add another account?**  
Edit the `accounts` array and rerun.

**Q: Can I change the red condition (1 or 3 riders)?**  
Yes. Change the logic where `numRiders` is evaluated.

**Q: Why is one rider/day with multiple rows not double-counted?**  
The script **uses the max POD per rider/day** to avoid overcounting duplicate log entries.

**Q: Can I place the calendars somewhere else?**  
Yes. Change `calendarStartCell` (e.g., `"B5"`).

**Q: Does `SpellPeso` handle very large numbers?**  
It supports up to the `Trillion` magnitude in this version.

---

## License

MIT © 2025
