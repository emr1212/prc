const SS_ID = '1MS5-MmDgreH4ouLdo3K_9NbPcP-V8M5vqufDghov1x4';

function doGet(e) {
  try {
    const ss    = SpreadsheetApp.openById(SS_ID);
    const tz    = Session.getScriptTimeZone();
    const now   = new Date();

    const shiftSheet = ss.getSheetByName('each shift cal');

    // ── shift start: L kolonunun en son dolu hucresi ──────────
    const lBottom = shiftSheet.getRange(shiftSheet.getMaxRows(), 12)
                              .getNextDataCell(SpreadsheetApp.Direction.UP);
    const shiftStartTime = lBottom.getValue() instanceof Date ? lBottom.getValue() : null;

    // ── since last job: B kolonunun en son dolu hucresi ───────
    const bBottom = shiftSheet.getRange(shiftSheet.getMaxRows(), 2)
                              .getNextDataCell(SpreadsheetApp.Direction.UP);
    const maxBDateVal = bBottom.getValue();
    let sinceLastJobCalc = '';
    if (maxBDateVal instanceof Date) {
      const sm = (now - maxBDateVal) / 60000;
      sinceLastJobCalc = Math.floor(sm/60) + ' hr ' + Math.floor(sm%60) + ' min';
    }

    // ── form responses: jobs after shift start ────────────────
    const fs     = ss.getSheetByName('Form responses 1');
    const fLast  = fs.getLastRow();
    const fStart = Math.max(2, fLast - 149);
    const fData  = fs.getRange(fStart, 1, fLast - fStart + 1, 4).getValues();
    const jobs = [];
    fData.forEach(function(row) {
      if (!row[0]) return;
      const tsDate = row[0] instanceof Date ? row[0] : new Date(row[0]);
      if (shiftStartTime && tsDate < shiftStartTime) return;
      const amt = parseFloat(String(row[1]).replace(/[^0-9.]/g, '')) || 0;
      if (!amt) return;
      jobs.push({
        time: Utilities.formatDate(tsDate, tz, 'HH:mm'),
        amt:  amt,
        type: String(row[2] || ''),
        info: String(row[3] || '').replace(/'/g, '').replace(/"/g, '')
      });
    });

    // ── Sheet2 batch read ─────────────────────────────────────
    const s2     = ss.getSheetByName('Sheet2');
    const s2data = s2.getRange('A1:H43').getValues();
    const todayEarnings = parseFloat(String(s2data[5][2]).replace(/[^0-9.-]/g, '')) || 0; // C6
    const b5val         = String(s2.getRange('B5').getValue());
    const e6val         = parseFloat(String(s2data[5][4]).replace(/[^0-9.-]/g, '')) || 0; // E6
    const d6val         = String(s2data[5][3]); // D6
    const weeklyTotal   = parseFloat(String(s2data[42][0]).replace(/[^0-9.-]/g, '')) || 0; // A43
    const e4note        = String(s2.getRange('E4').getValue());
    const e43val        = parseFloat(String(s2.getRange('E43').getValue()).replace(/[^0-9.-]/g, '')) || 0;

    // ── tax data ──────────────────────────────────────────────
    const taxIncl = s2.getRange('A129:B144').getValues();
    const taxExcl = s2.getRange('C129:D144').getValues();
    const taxData = {
      incl: {
        gross:     parseFloat(taxIncl[2][1]) || 0,
        loss:      parseFloat(taxIncl[1][1]) || 0,
        expenses:  parseFloat(taxIncl[3][1]) || 0,
        netProfit: parseFloat(taxIncl[4][1]) || 0,
        monthly:   parseFloat(taxIncl[8][1]) || 0,
        weekly:    parseFloat(taxIncl[9][1]) || 0,
        tax:       parseFloat(taxIncl[13][1]) || 0,
        receipts:  parseFloat(taxIncl[14][1]) || 0
      },
      excl: {
        gross:     parseFloat(taxExcl[2][1]) || 0,
        loss:      parseFloat(taxExcl[1][1]) || 0,
        expenses:  parseFloat(taxExcl[3][1]) || 0,
        netProfit: parseFloat(taxExcl[4][1]) || 0,
        monthly:   parseFloat(taxExcl[8][1]) || 0,
        weekly:    parseFloat(taxExcl[9][1]) || 0,
        tax:       parseFloat(taxExcl[13][1]) || 0,
        receipts:  parseFloat(taxExcl[14][1]) || 0
      }
    };

    // ── weekly cal ────────────────────────────────────────────
    const weeklyCalData = ss.getSheetByName('Weekly cal').getRange('Y1:AB18').getValues();
    const weeklyRows = [];
    for (let i = 1; i < weeklyCalData.length; i++) {
      const row = weeklyCalData[i];
      if (!row[0]) continue;
      weeklyRows.push({ label: String(row[0]), hours: parseFloat(row[1])||0, total: parseFloat(row[2])||0, hourly: parseFloat(row[3])||0 });
    }

    // ── report N1:N7 weekly total ─────────────────────────────
    const reportN = ss.getSheetByName('report').getRange('N1:N7').getValues();
    let weeklyTotalNew = 0;
    reportN.forEach(function(row) { weeklyTotalNew += parseFloat(String(row[0]).replace(/[^0-9.-]/g, '')) || 0; });

    // ── daily target remaining ────────────────────────────────
    const todayStr = Utilities.formatDate(now, tz, 'dd/MM/yyyy');
    const formEF   = fLast > 1 ? fs.getRange(2, 5, fLast - 1, 2).getValues() : [];
    let todaySumF  = 0;
    formEF.forEach(function(row) {
      const dv = row[0] instanceof Date ? Utilities.formatDate(row[0], tz, 'dd/MM/yyyy') : String(row[0]).substring(0,10);
      if (dv === todayStr) todaySumF += parseFloat(String(row[1]).replace(/[^0-9.-]/g, '')) || 0;
    });
    const dailyTargetRemaining = Math.round((e43val - todaySumF) * 100) / 100;

    // ── working time ──────────────────────────────────────────
    let b6val = '';
    if (shiftStartTime instanceof Date) {
      const totalMins = (now - shiftStartTime) / 60000;
      if (totalMins / 60 > 20) {
        b6val = 'shift finished';
      } else {
        b6val = Math.floor(totalMins/60) + ' hr ' + Math.floor(totalMins%60) + ' min';
      }
    }

    // ── day/night analyse I2 ──────────────────────────────────
    const i2val = parseFloat(String(ss.getSheetByName('day/night analyse').getRange('I2').getValue()).replace(/[^0-9.-]/g, '')) || 0;
    let f6val = '';
    if (i2val !== 0 && e6val !== 0) {
      const diff = (e6val - i2val) / i2val;
      f6val = diff > 0 ? '\u2B06\uFE0F +' + Math.abs(diff*100).toFixed(2) + '%' : '\u2B07\uFE0F -' + Math.abs(diff*100).toFixed(2) + '%';
    }

    // ── AT/AU columns (first 500 rows) ────────────────────────
    const atauData = shiftSheet.getRange(1, 46, 500, 2).getValues();

    // target date: last year same weekday
    const targetDate = (function() {
      const lastYear = new Date(now.getFullYear()-1, now.getMonth(), now.getDate());
      const todayDow = now.getDay() === 0 ? 7 : now.getDay();
      const lyDow    = lastYear.getDay() === 0 ? 7 : lastYear.getDay();
      lastYear.setDate(lastYear.getDate() - lyDow + todayDow);
      lastYear.setHours(0,0,0,0);
      return lastYear;
    })();
    const targetStr2 = Utilities.formatDate(targetDate, tz, 'dd/MM/yyyy');

    // last year same weekday perf
    let g8val = '';
    for (let i = 0; i < atauData.length; i++) {
      if (String(atauData[i][0]).trim() === targetStr2) { g8val = String(atauData[i][1]); break; }
    }

    // last year today perf
    const lyTodayDate = new Date(now); lyTodayDate.setDate(lyTodayDate.getDate() - 365);
    const lyTodayStr  = Utilities.formatDate(lyTodayDate, tz, 'dd/MM/yyyy');
    let lyTodayPerf   = lyTodayStr + ' Off day';
    for (let i = 0; i < atauData.length; i++) {
      if (String(atauData[i][0]).trim() === lyTodayStr) { lyTodayPerf = String(atauData[i][1]); break; }
    }

    // BC8 — last year same day star portion
    const bc8raw   = String(shiftSheet.getRange('BC8').getValue());
    const starMatch = bc8raw.match(/[★☆].*/);
    const ax2val   = starMatch ? starMatch[0].trim() : bc8raw;

    // ── same day last year shift note ─────────────────────────
    const lmData = shiftSheet.getRange(1, 12, 500, 6).getValues();
    const days2  = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];
    const mths   = ['01','02','03','04','05','06','07','08','09','10','11','12'];
    const fmtD   = function(d) {
      return String(d.getDate()).padStart(2,'0') + '/' + mths[d.getMonth()] + '/' + d.getFullYear()
        + ' ' + String(d.getHours()).padStart(2,'0') + ':' + String(d.getMinutes()).padStart(2,'0');
    };
    let sameYearNote = 'No shift last year';
    for (let i = 0; i < lmData.length; i++) {
      const lVal = lmData[i][0];
      if (lVal instanceof Date) {
        const lDate = new Date(lVal); lDate.setHours(0,0,0,0);
        if (lDate.getTime() === targetDate.getTime()) {
          const lEnd = lmData[i][1];
          if (lEnd instanceof Date) {
            const dm = (lEnd - lVal) / 60000;
            sameYearNote = 'Start: ' + days2[lVal.getDay()] + ' ' + fmtD(lVal)
              + ' - Finished: ' + fmtD(lEnd)
              + ' - Hours Worked: ' + Math.floor(dm/60) + 'h ' + Math.round(dm%60) + 'm'
              + ' - Total: \u00a3' + parseFloat(lmData[i][5]).toFixed(2)
              + ' - Total Jobs: ' + lmData[i][4];
          }
          break;
        }
      }
    }

    // ── monthly chart data ────────────────────────────────────
    const reportST   = ss.getSheetByName('report').getRange('S1:T31').getValues();
    const monthlyData = [];
    reportST.forEach(function(row) {
      if (!row[0]) return;
      const d = row[0] instanceof Date ? row[0] : new Date(row[0]);
      monthlyData.push({ day: d.getDate(), date: Utilities.formatDate(d, tz, 'dd/MM'), total: Math.round((parseFloat(row[1])||0)*100)/100 });
    });

    // ── weekly chart data ─────────────────────────────────────
    const reportMP = ss.getSheetByName('report').getRange('M1:P7').getValues();
    const weekDays = ['Mon','Tue','Wed','Thu','Fri','Sat','Sun'];
    const chartData = [];
    reportMP.forEach(function(row, i) {
      if (!row[0]) return;
      const d = row[0] instanceof Date ? row[0] : new Date(row[0]);
      chartData.push({ day: weekDays[i], date: Utilities.formatDate(d, tz, 'dd/MM'), total: parseFloat(row[1])||0, jobs: parseInt(row[3])||0 });
    });

    // ── monthly totals from Form responses 1 ────────────────────
    const allFormFull = fs.getRange(2, 1, fLast - 1, 2).getValues(); // A=timestamp, B=amount
    const monthTotals = {};
    const monthNames  = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    allFormFull.forEach(function(row) {
      if (!row[0] || !row[1]) return;
      const d   = row[0] instanceof Date ? row[0] : new Date(row[0]);
      if (isNaN(d.getTime())) return;
      const amt = parseFloat(String(row[1]).replace(/[^0-9.]/g, '')) || 0;
      if (!amt) return;
      const key = d.getFullYear() + '-' + String(d.getMonth()+1).padStart(2,'0');
      const lbl = monthNames[d.getMonth()] + ' ' + d.getFullYear();
      if (!monthTotals[key]) monthTotals[key] = { label: lbl, total: 0 };
      monthTotals[key].total += amt;
    });
    // Sort by key desc, take last 12 months with previous year comparison
    const sortedKeys = Object.keys(monthTotals).sort().reverse();
    const monthlyTotals = sortedKeys.slice(0, 12).map(function(k) {
      const parts = k.split('-');
      const prevKey = (parseInt(parts[0])-1) + '-' + parts[1];
      const prevTotal = monthTotals[prevKey] ? Math.round(monthTotals[prevKey].total * 100) / 100 : null;
      const total = Math.round(monthTotals[k].total * 100) / 100;
      const diff = prevTotal !== null ? Math.round((total - prevTotal) * 100) / 100 : null;
      const pct  = prevTotal !== null && prevTotal > 0 ? Math.round((total - prevTotal) / prevTotal * 1000) / 10 : null;
      return { label: monthTotals[k].label, total: total, prevTotal: prevTotal, diff: diff, pct: pct };
    });

    // ── airport pickups: this week and this month ───────────────
    const allFormData = fs.getRange(2, 4, fLast - 1, 2).getValues(); // D and E cols
    const nowDate = new Date(now);
    // Week start (Monday)
    const weekDay = nowDate.getDay() === 0 ? 6 : nowDate.getDay() - 1;
    const weekStart = new Date(nowDate); weekStart.setDate(nowDate.getDate() - weekDay); weekStart.setHours(0,0,0,0);
    // Month start
    const monthStart = new Date(nowDate.getFullYear(), nowDate.getMonth(), 1);
    let airportWeek = 0, airportMonth = 0;
    allFormData.forEach(function(row) {
      if (!row[0] || String(row[0]).toLowerCase().indexOf('airport') === -1) return;
      const d = row[1] instanceof Date ? row[1] : new Date(row[1]);
      if (isNaN(d.getTime())) return;
      if (d >= monthStart) airportMonth++;
      if (d >= weekStart)  airportWeek++;
    });
    const airportData = { week: airportWeek, month: airportMonth };

    // ── recent expenses ───────────────────────────────────────
    const expSS    = SpreadsheetApp.openById('1hRLCy3DtULdHOPFJSuXYJoD2ZnZNCuDddJyxUR69xkI');
    const expSheet = expSS.getSheetByName('Expenses form');
    const expLast  = expSheet.getLastRow();
    const expStart = Math.max(2, expLast - 4);
    const expData  = expSheet.getRange(expStart, 1, expLast - expStart + 1, 7).getValues();
    const expenses = [];
    expData.forEach(function(row) {
      if (!row[0] || !row[2]) return;
      const d = row[0] instanceof Date ? row[0] : new Date(row[0]);
      expenses.push({ date: Utilities.formatDate(d, tz, 'dd/MM/yyyy'), category: String(row[1]||''), name: String(row[2]||''), amount: parseFloat(row[4])||parseFloat(row[6])||0, payment: String(row[5]||'') });
    });
    expenses.sort(function(a,b){ return b.date > a.date ? 1 : -1; });
    const recentExpenses = expenses.slice(0, 5);

    // ── payload ───────────────────────────────────────────────
    const payload = JSON.stringify({
      jobs, weeklyTotal: weeklyTotalNew || weeklyTotal,
      todayTotal: todayEarnings, workingTime: b6val,
      hourlyRate: e6val, hourlyDiff: f6val, perf: d6val,
      lastYearPerf: lyTodayPerf, lastYearSameDay: g8val,
      sinceLastJob: sinceLastJobCalc || b5val, rightTitle: ax2val,
      lastYearNote: e4note, lastYearNoteDate: lyTodayStr,
      sameYearNote, dailyTargetRemaining, weeklyRows, chartData,
      recentExpenses, monthlyData, taxData, airportData, monthlyTotals
    });

    const cb = e && e.parameter && e.parameter.cb;
    if (cb) return ContentService.createTextOutput(cb + '(' + payload + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
    return ContentService.createTextOutput(payload).setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({ error: err.toString() })).setMimeType(ContentService.MimeType.JSON);
  }
}

function test() {
  const r = doGet({ parameter: {} });
  Logger.log(r.getContent());
}

function pushToFirebase() {
  try {
    const result = doGet({ parameter: {} });
    const data   = JSON.parse(result.getContent());
    const url    = 'https://firestore.googleapis.com/v1/projects/emr-dashboard-6f69d/databases/(default)/documents/dashboard/live';
    const token  = ScriptApp.getOAuthToken();
    const fields = {
      updatedAt:            { stringValue: new Date().toISOString() },
      weeklyTotal:          { doubleValue: data.weeklyTotal || 0 },
      todayTotal:           { doubleValue: data.todayTotal || 0 },
      workingTime:          { stringValue: data.workingTime || '' },
      hourlyRate:           { doubleValue: data.hourlyRate || 0 },
      hourlyDiff:           { stringValue: data.hourlyDiff || '' },
      perf:                 { stringValue: data.perf || '' },
      lastYearPerf:         { stringValue: data.lastYearPerf || '' },
      lastYearSameDay:      { stringValue: data.lastYearSameDay || '' },
      sinceLastJob:         { stringValue: data.sinceLastJob || '' },
      rightTitle:           { stringValue: data.rightTitle || '' },
      lastYearNote:         { stringValue: data.lastYearNote || '' },
      lastYearNoteDate:     { stringValue: data.lastYearNoteDate || '' },
      sameYearNote:         { stringValue: data.sameYearNote || '' },
      dailyTargetRemaining: { doubleValue: data.dailyTargetRemaining || 0 },
      jobs:                 { stringValue: JSON.stringify(data.jobs || []) },
      weeklyRows:           { stringValue: JSON.stringify(data.weeklyRows || []) },
      chartData:            { stringValue: JSON.stringify(data.chartData || []) },
      monthlyData:          { stringValue: JSON.stringify(data.monthlyData || []) },
      taxData:              { stringValue: JSON.stringify(data.taxData || {}) },
      recentExpenses:       { stringValue: JSON.stringify(data.recentExpenses || []) },
      airportData:          { stringValue: JSON.stringify(data.airportData || {}) },
      monthlyTotals:        { stringValue: JSON.stringify(data.monthlyTotals || []) }
    };
    const response = UrlFetchApp.fetch(url, {
      method: 'PATCH', contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + token },
      payload: JSON.stringify({ fields: fields }),
      muteHttpExceptions: true
    });
    Logger.log('Status: ' + response.getResponseCode());
    Logger.log('Firebase updated: ' + new Date().toISOString());
  } catch(err) {
    Logger.log('Firebase error: ' + err.toString());
  }
}

function testFirebase() { pushToFirebase(); }

function syncNewJobsToFirestore() {
  const ss        = SpreadsheetApp.openById(SS_ID);
  const formSheet = ss.getSheetByName('Form responses 1');
  const propStore = PropertiesService.getScriptProperties();
  const lastSynced = parseInt(propStore.getProperty('lastSyncedRow') || '1');
  const lastRow    = formSheet.getLastRow();
  if (lastRow <= lastSynced) return;
  const newRows = formSheet.getRange(lastSynced + 1, 1, lastRow - lastSynced, 4).getValues();
  const url   = 'https://firestore.googleapis.com/v1/projects/emr-dashboard-6f69d/databases/(default)/documents/jobs';
  const token = ScriptApp.getOAuthToken();
  newRows.forEach(function(row) {
    if (!row[0]) return;
    const ts  = row[0] instanceof Date ? Utilities.formatDate(row[0], Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm') : String(row[0]);
    const amt = parseFloat(String(row[1]).replace(/[^0-9.]/g, '')) || 0;
    if (!amt) return;
    UrlFetchApp.fetch(url, {
      method: 'POST', contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + token },
      payload: JSON.stringify({ fields: { timestamp: {stringValue:ts}, amount: {doubleValue:amt}, type: {stringValue:String(row[2]||'')}, info: {stringValue:String(row[3]||'')}, createdAt: {integerValue:String(Date.now())} } }),
      muteHttpExceptions: true
    });
  });
  propStore.setProperty('lastSyncedRow', String(lastRow));
}

function initSyncFromNow() {
  const ss = SpreadsheetApp.openById(SS_ID);
  const lastRow = ss.getSheetByName('Form responses 1').getLastRow();
  PropertiesService.getScriptProperties().setProperty('lastSyncedRow', String(lastRow));
  Logger.log('Initialized at row: ' + lastRow);
}
