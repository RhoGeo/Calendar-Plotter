function formatNamesAndCleanSpaces() {
  const sheetName = "Time-out Report";
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    safeNotify(ss, 'Name Correction', `Sheet "${sheetName}" not found.`);
    return;
  }
  const lastRow = Math.max(2, sheet.getLastRow());
  const height = lastRow - 1;
  if (height <= 0) {
    safeNotify(ss, 'Name Correction', 'No data rows to process.');
    return;
  }
  const range = sheet.getRange(2, 2, height, 1);
  const values = range.getValues();
  let changed = 0;
  const formattedValues = values.map(([v]) => {
    if (typeof v !== 'string' || v.trim() === '') return [v];
    const before = v;
    let s = String(v).replace(/\u00A0/g, ' ').trim().replace(/\s+/g, ' ');
    s = properCaseFullName(s);
    if (s !== before) changed++;
    return [s];
  });
  range.setValues(formattedValues);
  safeNotify(ss, 'Name Correction', `Formatted ${changed}/${height} name(s) in "${sheetName}".`);
}
function properCaseFullName(name) {
  return name
    .split(' ')
    .map(token => properCaseToken(token))
    .join(' ');
}
function properCaseToken(token) {
  if (!token) return token;
  const upperKeep = new Set(['II','III','IV','VI','VII','VIII','IX','X','JR','SR']);
  const cleaned = token.replace(/\.+$/,'');
  const suffixDot = token.endsWith('.') ? '.' : '';
  if (upperKeep.has(cleaned.toUpperCase())) {
    return cleaned.toUpperCase() + suffixDot;
  }
  const splitRe = /([\'-])/;
  const parts = token.split(splitRe).map(part => {
    if (part === "'" || part === "-") return part;
    let p = part.toLowerCase();
    if (p.startsWith('mc') && p.length >= 3) {
      return 'Mc' + capFirst(p.slice(2));
    }
    if (p.startsWith('mac') && p.length >= 4) {
      return 'Mac' + capFirst(p.slice(3));
    }
    return capFirst(p);
  });
  return parts
    .join('')
    .replace(/^o'([a-z])/, (_, c) => "O'" + c.toUpperCase());
}
function capFirst(s) {
  return s ? s.charAt(0).toUpperCase() + s.slice(1) : s;
}
function safeNotify(ss, title, message, seconds) {
  try {
    ss.toast(String(message), String(title || ''), seconds || 5);
  } catch (e) {
    Logger.log(`[${title}] ${message}`);
  }
}
function SpellPeso(input) {
  const number = parseFloat(input);
  if (isNaN(number)) {
    return "Not a valid number";
  }
  if (number === 0) {
    return "Zero Pesos Only";
  }
  const a = ['', 'One', 'Two', 'Three', 'Four', 'Five', 'Six', 'Seven', 'Eight', 'Nine', 'Ten', 'Eleven', 'Twelve', 'Thirteen', 'Fourteen', 'Fifteen', 'Sixteen', 'Seventeen', 'Eighteen', 'Nineteen'];
  const b = ['', '', 'Twenty', 'Thirty', 'Forty', 'Fifty', 'Sixty', 'Seventy', 'Eighty', 'Ninety'];
  const magnitudes = ['', 'Thousand', 'Million', 'Billion', 'Trillion'];
  function convertChunk(num) {
    let str = '';
    if (num >= 100) {
      str += a[Math.floor(num / 100)] + ' Hundred ';
      num %= 100;
    }
    if (num > 0) {
      if (num < 20) {
        str += a[num];
      } else {
        str += b[Math.floor(num / 10)];
        if (num % 10 > 0) {
          str += ' ' + a[num % 10];
        }
      }
    }
    return str.trim();
  }
  function numberToWords(num) {
    if (num === 0) return '';
    let words = '';
    let i = 0;
    do {
      const chunk = num % 1000;
      if (chunk !== 0) {
        const chunkWords = convertChunk(chunk);
        words = chunkWords + ' ' + magnitudes[i] + ' ' + words;
      }
      num = Math.floor(num / 1000);
      i++;
    } while (num > 0);
    return words.trim();
  }
  const pesos = Math.floor(number);
  const centavos = Math.round((number - pesos) * 100);
  let pesoWords = numberToWords(pesos);
  let centavoWords = numberToWords(centavos);
  let result = '';
  if (pesoWords) {
    result += pesoWords + (pesos === 1 ? ' Peso' : ' Pesos');
  }
  if (centavoWords) {
    if (result) {
      result += ' and ';
    }
    result += centavoWords + (centavos === 1 ? ' Centavo' : ' Centavos');
  }
  return (result + ' Only').trim();
}
function onOpen(e) {
  try {
    SpreadsheetApp.getUi()
      .createMenu("Clock-outs")
      .addItem("Plot calendars", "plotClockoutsAsSeparateCalendars")
      .addToUi();
  } catch (err) {}
}
function notify(ss, title, message, seconds) {
  try {
    ss.toast(String(message), String(title || ""), seconds || 5);
  } catch (e) {
    Logger.log(`[${title || "Clock-outs"}] ${message}`);
  }
}
function plotClockoutsAsSeparateCalendars() {
  const timeOutReportSheetName = "Time-out Report";
  const payslipSheetName = "Payslip";
  const expensesSheetName = "Riders";
  const calendarStartCell = "A20";
  const summaryStartRow = 15;
  const summaryStartCol = 1;
  const currencyFormat = "₱#,##0";
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const timeOutSheet = ss.getSheetByName(timeOutReportSheetName);
  const payslipSheet = ss.getSheetByName(payslipSheetName);
  const expensesSheet = ss.getSheetByName(expensesSheetName);
  if (!timeOutSheet || !payslipSheet || !expensesSheet) {
    notify(ss, "Clock-outs", "Error: Missing required sheet(s).", 8);
    return;
  }
  const startDate = payslipSheet.getRange("B5").getValue();
  const endDate = payslipSheet.getRange("B6").getValue();
  if (!startDate || !endDate || !(startDate instanceof Date) || !(endDate instanceof Date)) {
    notify(ss, "Clock-outs", "Error: Invalid date range in Payslip! (B5–B6)", 8);
    return;
  }
  notify(ss, "Clock-outs", "Generating calendars…", 5);
  const endDateClamped = new Date(endDate);
  endDateClamped.setHours(23, 59, 59, 999);
  const timeOutDataRange = timeOutSheet.getRange(
    2,
    1,
    Math.max(0, timeOutSheet.getLastRow() - 1),
    timeOutSheet.getLastColumn()
  );
  const allTimeOutData = timeOutDataRange.getValues();
  const accounts = ["OCW Etcobanez", "OCW Mendoza", "OCW Seco"];
  const tz = ss.getSpreadsheetTimeZone();
  let currentPlotRow = expensesSheet.getRange(calendarStartCell).getRow();
  const calendarStartCol = expensesSheet.getRange(calendarStartCell).getColumn();
  expensesSheet
    .getRange(
      currentPlotRow,
      calendarStartCol,
      expensesSheet.getMaxRows() - currentPlotRow + 1,
      expensesSheet.getMaxColumns() - calendarStartCol + 1
    )
    .clearContent()
    .clearFormat();
  expensesSheet
    .getRange(summaryStartRow, summaryStartCol, accounts.length + 1, 4)
    .clearContent()
    .clearFormat();
  let grandTotalAmount = 0;
  let grandTotalDays = 0;
  let grandTotalPOD = 0;
  accounts.forEach((accountName, idx) => {
    const daily = new Map();
    for (const row of allTimeOutData) {
      const ts = row[0] instanceof Date ? row[0] : new Date(row[0]);
      const riderName = String(row[1] || "").trim();
      const account = String(row[2] || "").trim();
      const podRaw = row[3];
      const pod = typeof podRaw === "number" ? podRaw : Number(podRaw) || 0;
      if (!ts || isNaN(ts.getTime())) continue;
      if (/pp/i.test(riderName)) continue;
      if (account !== accountName) continue;
      if (ts < startDate || ts > endDateClamped) continue;
      const dateKey = Utilities.formatDate(ts, tz, "M/d/yyyy");
      if (!daily.has(dateKey)) {
        daily.set(dateKey, { riders: new Set(), podByRider: new Map() });
      }
      const bucket = daily.get(dateKey);
      bucket.riders.add(riderName);
      const prev = bucket.podByRider.get(riderName) || 0;
      bucket.podByRider.set(riderName, Math.max(prev, pod));
    }
    const highlightCount = daily.size;
    const totalPODForAccount = Array.from(daily.values()).reduce((acc, b) => {
      let sum = 0;
      b.podByRider.forEach(v => (sum += (typeof v === "number" ? v : Number(v) || 0)));
      return acc + sum;
    }, 0);
    const amount = highlightCount * 2300;
    grandTotalAmount += amount;
    grandTotalDays += highlightCount;
    grandTotalPOD += totalPODForAccount;
    const summaryRow = summaryStartRow + idx;
    expensesSheet.getRange(summaryRow, 1).setValue(accountName + " Clock-outs");
    expensesSheet.getRange(summaryRow, 2).setValue(highlightCount);
    expensesSheet.getRange(summaryRow, 3).setValue(totalPODForAccount);
    expensesSheet.getRange(summaryRow, 4).setValue(amount).setNumberFormat(currencyFormat);
    expensesSheet.getRange(currentPlotRow, calendarStartCol).setValue(accountName + " Clock-outs").setFontWeight("bold").setFontSize(12);
    expensesSheet.getRange(currentPlotRow, calendarStartCol + 1).setValue(highlightCount).setFontWeight("bold").setFontSize(12);
    expensesSheet.getRange(currentPlotRow, calendarStartCol + 2).setValue(totalPODForAccount).setFontWeight("bold").setFontSize(12);
    expensesSheet.getRange(currentPlotRow, calendarStartCol + 3).setValue(amount).setNumberFormat(currencyFormat).setFontWeight("bold").setFontSize(12);
    currentPlotRow++;
    const startCalendarDate = new Date(startDate);
    startCalendarDate.setHours(0, 0, 0, 0);
    startCalendarDate.setDate(startCalendarDate.getDate() - startCalendarDate.getDay());
    const header = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
    expensesSheet.getRange(currentPlotRow, calendarStartCol, 1, 7).setValues([header]).setFontWeight("bold");
    currentPlotRow++;
    let curr = new Date(startCalendarDate);
    while (curr <= endDateClamped || curr.getDay() !== 0) {
      const weekTexts = [];
      const weekBgs = [];
      for (let i = 0; i < 7; i++) {
        const dateKey = Utilities.formatDate(curr, tz, "M/d/yyyy");
        const dayNum = curr.getDate();
        let text = String(dayNum);
        let bg = "#ffffff";
        if (curr >= startDate && curr <= endDateClamped && daily.has(dateKey)) {
          const info = daily.get(dateKey);
          const riders = Array.from(info.riders);
          let datePOD = 0;
          info.podByRider.forEach(v => (datePOD += (typeof v === "number" ? v : Number(v) || 0)));
          const ridersLine = riders.join(", ");
          text = `${dayNum}\n${ridersLine}\n${datePOD} POD`;
          const numRiders = riders.length;
          const isUnderOrOver = (numRiders === 1 || numRiders === 3);
          bg = isUnderOrOver ? "#f4cccc" : "#c9daf8";
        }
        weekTexts.push(text);
        weekBgs.push(bg);
        curr.setDate(curr.getDate() + 1);
      }
      const row = expensesSheet.getRange(currentPlotRow, calendarStartCol, 1, 7);
      row.setValues([weekTexts]);
      row.setBackgrounds([weekBgs]);
      row.setVerticalAlignment("top").setWrap(true);
      expensesSheet.setRowHeight(currentPlotRow, 100);
      currentPlotRow++;
    }
    currentPlotRow += 2;
  });
  const totalsRow = summaryStartRow + accounts.length;
  expensesSheet.getRange(totalsRow, 1).setValue("TOTALS").setFontWeight("bold");
  expensesSheet.getRange(totalsRow, 2).setValue(grandTotalDays).setFontWeight("bold");
  expensesSheet.getRange(totalsRow, 3).setValue(grandTotalPOD).setFontWeight("bold");
  expensesSheet.getRange(totalsRow, 4).setValue(grandTotalAmount).setNumberFormat(currencyFormat).setFontWeight("bold").setBackground("#d9ead3");
  for (let i = 0; i < 7; i++) {
    expensesSheet.setColumnWidth(calendarStartCol + i, 120);
  }
  expensesSheet.setRowHeight(expensesSheet.getRange(calendarStartCell).getRow() + 1, 30);
  notify(ss, "Clock-outs", "Clock-out calendars with combined POD and summaries generated successfully.", 5);
}
