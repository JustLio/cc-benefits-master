// Branch of the cc-benefits-master repo that cards.json is pulled from.
// Keep this on 'main' for production. Point it at a feature branch ONLY for
// temporary testing of unmerged card-data changes, then switch it back —
// the live app always reads whatever branch is named here.
var CARD_DATA_BRANCH = 'main';
var CARD_DATA_URL = 'https://raw.githubusercontent.com/JustLio/cc-benefits-master/refs/heads/' + CARD_DATA_BRANCH + '/cards.json';

// Bump this when the Dashboard / Annual Fee Analyzer static layout (column
// widths, header row heights, fonts) changes, so refresh re-applies it once
// instead of on every single edit. See refreshDashboard / refreshAnnualFeeAnalyzer.
var LAYOUT_VERSION = '1';

// Opens the HTML dialog as a modal popup
function showEntryForm() {
  var html = HtmlService.createHtmlOutputFromFile('Form')
    .setWidth(460)
    .setHeight(540)
    .setTitle('New Credit Card');
  SpreadsheetApp.getUi().showModalDialog(html, 'New Credit Card');
}

function showUpdatePointsForm() {
  var html = HtmlService.createHtmlOutputFromFile('UpdatePoints')
    .setWidth(400)
    .setHeight(300)
    .setTitle('Update Points');
  SpreadsheetApp.getUi().showModalDialog(html, 'Update Points');
}

function getCardsForUpdate() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Credit Card Raw Data');
  if (!sheet || sheet.getLastRow() < 2) return JSON.stringify([]);
  var rows  = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
  var cards = [];
  rows.forEach(function(r, idx) {
    if (!r[0]) return;
    cards.push({ bank: r[0], cardName: r[1], points: parseFloat(r[2]) || 0, rowIndex: idx + 2 });
  });
  return JSON.stringify(cards);
}

function updateCardPoints(dataString) {
  var data  = JSON.parse(dataString);
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Credit Card Raw Data');
  if (!sheet) return;
  sheet.getRange(data.rowIndex, 3).setValue(parseFloat(data.points) || 0);
  refreshDashboard();
}
function newCard(dataString) {
  var data  = JSON.parse(dataString);
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Credit Card Raw Data');

  // Fetch full JSON once — provides card data, network, annualFee, and CPP
  var json      = fetchAllCardData();
  var cardData  = json ? (json[data.bank + '|' + data.cardName] || null) : null;
  var network   = (cardData && cardData.network)          ? cardData.network   : '';
  var annualFee = (cardData && cardData.annualFee != null) ? cardData.annualFee : 0;
  var cpp       = 0;
  if (cardData && cardData.rewardsProgram && json && json.programs) {
    cpp = json.programs[cardData.rewardsProgram] || 0;
  }

  // Use user-reviewed benefits if they edited them in the preview step
  if (data.benefits && Array.isArray(data.benefits) && data.benefits.length > 0 && cardData) {
    cardData = Object.assign({}, cardData, { benefits: data.benefits });
  }

  // Build per-benefit override map for this card (written to Raw Data col I)
  var cardOverrides = {};
  if (cardData && cardData.benefits) {
    cardData.benefits.forEach(function(b) {
      if (b.valueOverrides) cardOverrides[b.name] = b.valueOverrides;
    });
  }

  // Save to Raw Data sheet
  sheet.appendRow([
    data.bank,
    data.cardName,
    parseFloat(data.points) || 0,
    cpp,
    annualFee,
    data.renewalDate,
    data.renewalReminder ? 'Yes' : 'No',
    network,
    Object.keys(cardOverrides).length ? JSON.stringify(cardOverrides) : ''
  ]);

  // Write to Credits Tracker
  setupBenefitsTracker();
  if (cardData) {
    addCardBenefits(data.bank, data.cardName, cardData);
  } else {
    Logger.log('No data found in JSON for: ' + data.bank + '|' + data.cardName);
  }

  refreshDashboard();
  refreshAnnualFeeAnalyzer();
}
function showDashboard() {
  var html = HtmlService.createHtmlOutputFromFile('Dashboard')
    .setWidth(1540)
    .setHeight(900)
    .setTitle('Card Dashboard');
  SpreadsheetApp.getUi().showModalDialog(html, 'Card Dashboard');
}

function getNextReset_(frequency) {
  var today = new Date(); today.setHours(0,0,0,0);
  var y = today.getFullYear(), m = today.getMonth();
  if (frequency === 'Monthly')     return new Date(y, m + 1, 0);
  if (frequency === 'Semi-Annual') return m <= 5 ? new Date(y, 6, 0) : new Date(y, 12, 0);
  if (frequency === 'Quarterly') {
    var ends = [2, 5, 8, 11];
    for (var i = 0; i < ends.length; i++) if (m <= ends[i]) return new Date(y, ends[i] + 1, 0);
  }
  return null;
}

// Returns the effective per-period value. overrideObj is a {month:value} map from Credit Card Raw Data
// (e.g. {"12":20} for a December-specific amount), or null if no override exists.
function resolvePerPeriodValue_(valStr, overrideObj) {
  var base = null;
  if (typeof valStr === 'number' && valStr > 0) base = valStr;
  else if (typeof valStr === 'string' && valStr.charAt(0) === '$') base = parseFloat(valStr.substring(1)) || null;
  if (base === null) return null;
  if (overrideObj) {
    var month = (new Date().getMonth() + 1).toString();
    if (overrideObj[month] !== undefined) return overrideObj[month];
  }
  return base;
}

function getDashboardData() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Credit Card Raw Data');
  var rows  = sheet.getDataRange().getValues();
  var cards = [];

  for (var i = 1; i < rows.length; i++) {
    var row = rows[i];
    if (!row[0]) continue;
    var renewalDate = '';
    if (row[5]) {
      var d = new Date(row[5]);
      renewalDate = Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    }
    cards.push({
      bank:            row[0],
      cardName:        row[1],
      points:          parseFloat(row[2]) || 0,
      cpp:             parseFloat(row[3]) || 0,
      annualFee:       parseFloat(row[4]) || 0,
      renewalDate:     renewalDate,
      renewalReminder: row[6] === 'Yes',
      network:         row[7] || ''
    });
  }

  // Build used-amounts map from Raw Data col J: { "bank|card": { "benefit name": 15 } }
  var gdUsed = {};
  var rawSheet2 = ss.getSheetByName('Credit Card Raw Data');
  if (rawSheet2 && rawSheet2.getLastRow() > 1) {
    rawSheet2.getRange(2, 1, rawSheet2.getLastRow() - 1, 10).getValues().forEach(function(r) {
      if (!r[0] || !r[9]) return;
      try { gdUsed[r[0] + '|' + r[1]] = JSON.parse(r[9]); } catch(e) {}
    });
  }

  var pinnedBenefits   = [];
  var recurringCredits = [];
  var rcFreqs  = { 'Monthly': true, 'Semi-Annual': true, 'Quarterly': true };
  var rcToday  = new Date(); rcToday.setHours(0,0,0,0);
  var btSheet = ss.getSheetByName('Credits Tracker');
  if (btSheet && btSheet.getLastRow() > 2) {
    var btRows = btSheet.getRange(3, 2, btSheet.getLastRow() - 2, 8).getValues(); // B-I
    btRows.forEach(function(r) {
      var bank = r[0], card = r[1];
      if (!bank || !card || !r[2]) return;
      if (r[5] === 'Earn Rate') return;
      var pinned = r[7] === true; // r[7] = PIN (col I)
      if (!pinned) return;
      var rd       = getNextReset_(r[4]);
      var daysLeft = rd ? Math.round((rd - rcToday) / 86400000) : null;
      var rdStr    = rd ? Utilities.formatDate(rd, Session.getScriptTimeZone(), 'MMM d') : null;
      var valNum   = null;
      if (typeof r[3] === 'string' && r[3].charAt(0) === '$') valNum = parseFloat(r[3].substring(1)) || 0;
      var notesUsed = (gdUsed[bank + '|' + card] || {})[r[2]] || 0;
      var remaining = valNum !== null ? Math.max(0, Math.round((valNum - notesUsed) * 100) / 100) : null;
      if (rcFreqs[r[4]]) {
        recurringCredits.push({
          bank: bank, cardName: card, name: r[2], value: r[3], valueNum: valNum,
          remaining: remaining,
          frequency: r[4], category: r[5], resetDate: rdStr, daysLeft: daysLeft
        });
      } else {
        pinnedBenefits.push({ bank: bank, cardName: card, name: r[2], value: r[3], frequency: r[4], category: r[5] });
      }
    });
  }

  return JSON.stringify({ cards: cards, pinnedBenefits: pinnedBenefits, recurringCredits: recurringCredits });
}
// No menu — just navigate to Dashboard and refresh on open
function onOpen() {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var dash = ss.getSheetByName('Dashboard');
  if (dash) {
    ss.setActiveSheet(dash);
    syncBenefitsTracker();
    // NOTE: JSON reconciliation (syncBenefitsFromJson) needs UrlFetchApp, which a
    // simple onOpen trigger is NOT authorized to call — it would silently fail here.
    // It runs from the installable trigger instead (onOpenSync). Run installTriggers()
    // once from the Apps Script editor to enable on-open card-data sync.
    refreshDashboard();
  }
  if (!ss.getSheetByName('Annual Fee Analyzer')) {
    setupAnnualFeeAnalyzer();
  } else {
    refreshAnnualFeeAnalyzer();
  }

  // Enforce tab order: Benefits Tracker = 2nd, Annual Fee Analyzer = last
  var btSheet  = ss.getSheetByName('Credits Tracker');
  var afaSheet = ss.getSheetByName('Annual Fee Analyzer');
  var total    = ss.getSheets().length;
  if (afaSheet) { ss.setActiveSheet(afaSheet); ss.moveActiveSheet(total); }
  if (btSheet)  { ss.setActiveSheet(btSheet);  ss.moveActiveSheet(2); }
  if (dash) ss.setActiveSheet(dash);
}

// Installable onOpen trigger target — runs WITH authorization, so (unlike the
// simple onOpen above) it can call UrlFetchApp to reconcile the Credits Tracker
// against the latest cards.json. Enable by running installTriggers() once.
function onOpenSync() {
  syncBenefitsFromJson();
  refreshDashboard();
  refreshAnnualFeeAnalyzer();
}

// Run ONCE from the Apps Script editor to enable on-open card-data sync.
// Creates (and de-dupes) an installable onOpen trigger pointing at onOpenSync.
function installTriggers() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'onOpenSync') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('onOpenSync')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onOpen()
    .create();
  SpreadsheetApp.getUi().alert('✅ Auto-sync enabled. Card data will refresh from GitHub each time you open the sheet.');
}

// Fires on cell edits — refreshes Dashboard on any Raw Data change,
// syncs Benefits Tracker only when bank/card columns are touched,
// live-updates dashboard + AFA when PIN column (col 9) is toggled
function onEdit(e) {
  var sheetName = e.range.getSheet().getName();
  if (sheetName === 'Credit Card Raw Data') {
    refreshDashboard();
    if (e.range.getColumn() <= 2) syncBenefitsTracker();
  }
  if (sheetName === 'Credits Tracker' && e.range.getColumn() === 9) {
    refreshDashboard();
    refreshAnnualFeeAnalyzer();
  }
  // Dashboard "USED" checkbox (col F / 6) toggled — persist usage so it survives refresh
  if (sheetName === 'Dashboard' && e.range.getColumn() === 6) {
    markCreditUsedFromDashboard_(e.range);
  }
}

// Persists a Dashboard "USED" checkbox toggle into Credit Card Raw Data col J so
// used / remaining amounts survive the next dashboard rebuild. Col J stores a
// per-card { "credit name": usedAmount } map. The checkbox is all-or-nothing:
// checked marks the full per-period value as used, unchecked clears it.
function markCreditUsedFromDashboard_(cell) {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var dash = cell.getSheet();
  var row  = cell.getRow();

  // Act only on genuine credit rows: a boolean checkbox, a named credit in col B,
  // and a numeric per-period value in col C.
  var isChecked = cell.getValue();
  if (isChecked !== true && isChecked !== false) return;
  var creditName = (dash.getRange(row, 2).getValue() || '').toString().trim();
  if (!creditName || creditName.indexOf('💳') === 0) return;
  var perPeriod = parseFloat(dash.getRange(row, 3).getValue());
  if (isNaN(perPeriod)) return;

  // Walk up to the nearest card header row ("💳  BANK  —  CARD") to identify the card.
  var bank = null, card = null;
  for (var r = row - 1; r >= 7; r--) {
    var v = (dash.getRange(r, 2).getValue() || '').toString();
    if (v.indexOf('💳  ') === 0 && v.indexOf('  —  ') !== -1) {
      var inner = v.substring(4);
      var sep   = inner.indexOf('  —  ');
      bank = inner.substring(0, sep).trim();
      card = inner.substring(sep + 5).trim();
      break;
    }
  }
  if (!bank || !card) return;

  // Update the used-amounts JSON in Raw Data col J for this card.
  var raw = ss.getSheetByName('Credit Card Raw Data');
  if (!raw || raw.getLastRow() < 2) return;
  var keys = raw.getRange(2, 1, raw.getLastRow() - 1, 2).getValues();
  for (var i = 0; i < keys.length; i++) {
    if (keys[i][0] === bank && keys[i][1] === card) {
      var cellJ    = raw.getRange(i + 2, 10);
      var used     = {};
      var existing = cellJ.getValue();
      if (existing) { try { used = JSON.parse(existing); } catch (e2) { used = {}; } }
      if (isChecked) used[creditName] = perPeriod;
      else           delete used[creditName];
      cellJ.setValue(Object.keys(used).length ? JSON.stringify(used) : '');
      break;
    }
  }

  refreshDashboard();
}

// Removes Benefits Tracker sections for cards no longer in Raw Data
function syncBenefitsTracker() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var raw   = ss.getSheetByName('Credit Card Raw Data');
  var sheet = ss.getSheetByName('Credits Tracker');
  // Migration: rename legacy "Benefits Tracker" tab on first open
  if (!sheet) {
    var legacy = ss.getSheetByName('Benefits Tracker');
    if (legacy) { legacy.setName('Credits Tracker'); sheet = legacy; }
  }
  if (!raw || !sheet) return;

  // Migration: delete legacy Date Used column from existing sheets
  var i2Val = sheet.getRange('I2').getValue().toString().toUpperCase().trim();
  if (i2Val === 'DATE USED' || i2Val === 'DATE') {
    sheet.deleteColumn(9);
  }

  // Migration: remove legacy Used? column (col H) — dashboard now uses PIN instead
  var h2Val = sheet.getRange('H2').getValue().toString().toUpperCase().trim();
  if (h2Val === 'USED?' || h2Val === 'USED') {
    sheet.deleteColumn(8);
  }

  // Column widths: A=margin B=bank C=card D=benefit E=value F=freq G=category H=notes I=pin J=margin
  sheet.setColumnWidth(1,  24);  // A: margin
  sheet.setColumnWidth(2,  200); // B: Bank
  sheet.setColumnWidth(3,  200); // C: Card
  sheet.setColumnWidth(4,  280); // D: Benefit
  sheet.setColumnWidth(5,  105); // E: Value
  sheet.setColumnWidth(6,  125); // F: Frequency
  sheet.setColumnWidth(7,  140); // G: Category
  sheet.setColumnWidth(8,  260); // H: Notes
  sheet.setColumnWidth(9,  80);  // I: PIN
  sheet.setColumnWidth(10, 24);  // J: margin
  // Title and header row heights
  sheet.setRowHeight(1, 76);
  sheet.setRowHeight(2, 42);

  // Self-heal: add PIN column header if missing (lives at I2)
  var pinHdr = sheet.getRange('I2').getValue();
  if (!pinHdr || pinHdr.toString().indexOf('PIN') === -1) {
    sheet.getRange('I2')
      .setValue('📌 PIN')
      .setFontSize(11).setFontWeight('bold').setFontColor('#6b7280')
      .setVerticalAlignment('middle')
      .setNote('Check ✓ to pin this credit to your Dashboard');
    sheet.getRange(1, 10, sheet.getLastRow(), 1).setBackground('#0f1117');
  }

  // Self-heal: fix stale title and header labels from "Benefits" → "Credits"
  var b1Val = sheet.getRange('B1').getValue().toString();
  if (b1Val.toUpperCase().indexOf('BENEFIT') !== -1) {
    sheet.getRange('B1').setValue('✅  CREDITS TRACKER');
  }
  var d2Val = sheet.getRange('D2').getValue().toString().toUpperCase().trim();
  if (d2Val === 'BENEFIT') {
    sheet.getRange('D2').setValue('CREDIT');
  }

  // Build set of active card keys from raw data
  var lastRaw    = raw.getLastRow();
  var activeCards = {};
  if (lastRaw > 1) {
    raw.getRange(2, 1, lastRaw - 1, 2).getValues().forEach(function(r) {
      if (r[0]) activeCards[r[0] + '|' + r[1]] = true;
    });
  }

  var lastBT = sheet.getLastRow();
  if (lastBT < 3) return;

  // Scan column B from row 3 down for card header rows ("💳  BANK  —  CARD")
  var colB     = sheet.getRange(3, 2, lastBT - 2, 1).getValues();
  var sections = [];

  for (var i = 0; i < colB.length; i++) {
    var val = String(colB[i][0] || '');
    if (val.indexOf('💳  ') !== 0 || val.indexOf('  —  ') === -1) continue;
    var inner = val.substring(4);       // strip "💳  " (emoji=2 chars + 2 spaces)
    var sep   = inner.indexOf('  —  ');
    sections.push({
      key:      inner.substring(0, sep).trim() + '|' + inner.substring(sep + 5).trim(),
      sheetRow: i + 3                   // offset: rows 1-2 are headers
    });
  }

  // Assign end row for each section
  for (var i = 0; i < sections.length; i++) {
    sections[i].endRow = (i + 1 < sections.length)
      ? sections[i + 1].sheetRow - 1
      : lastBT;
  }

  // Delete stale sections in reverse order so row numbers stay valid
  sections.slice().reverse().forEach(function(s) {
    if (!activeCards[s.key]) {
      sheet.deleteRows(s.sheetRow, s.endRow - s.sheetRow + 1);
    }
  });
}

// ── Adds benefits from the JSON that are missing from the Benefits Tracker ──
// Runs on open so the BT stays in sync as cards.json is updated.
function syncBenefitsFromJson() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var raw = ss.getSheetByName('Credit Card Raw Data');
  var bt  = ss.getSheetByName('Credits Tracker');
  if (!raw || !bt || raw.getLastRow() < 2 || bt.getLastRow() < 3) return;

  var json = fetchAllCardData();
  if (!json) return;

  var recurringFreqs = { 'Monthly': true, 'Semi-Annual': true, 'Quarterly': true };
  var lastBT = bt.getLastRow();
  var btData = bt.getRange(3, 2, lastBT - 2, 7).getValues(); // cols B-H

  // Build per-card map: existing benefit names + per-benefit row info for value updates
  var cardInfo = {};
  var benefitRows = {};
  btData.forEach(function(row, idx) {
    var bank = row[0], card = row[1];
    if (!bank || !card) return;
    var key = bank + '|' + card;
    if (!cardInfo[key]) cardInfo[key] = { bank: bank, card: card, names: {}, lastRow: 0 };
    cardInfo[key].lastRow = idx + 3; // +3 because data starts at sheet row 3
    if (row[2] && row[5] !== 'Earn Rate') {
      var name = row[2].toString().trim();
      cardInfo[key].names[name.toLowerCase()] = true;
      benefitRows[key + '|' + name] = { sheetRow: idx + 3, valStr: row[3].toString() };
    }
  });

  // Find cards that have benefits in JSON not yet in BT
  var updates = [];
  Object.keys(cardInfo).forEach(function(key) {
    var info     = cardInfo[key];
    var cardData = json[key];
    if (!cardData || !cardData.benefits) return;
    var missing = cardData.benefits.filter(function(b) {
      return b.name && !info.names[b.name.trim().toLowerCase()];
    });
    if (missing.length) updates.push({ bank: info.bank, card: info.card, lastRow: info.lastRow, missing: missing });
  });

  // Insert missing benefits (bottom-to-top so row numbers for cards above stay valid)
  if (updates.length) {
    updates.sort(function(a, b) { return b.lastRow - a.lastRow; });
    updates.forEach(function(u) {
      var insertRow = u.lastRow + 1;
      var m = u.missing.length;
      var bVals = [], bBgs = [], bFgClrs = [], bFgSzs = [], bFgWts = [], bVAlns = [], bWraps = [];
      u.missing.forEach(function(benefit, i) {
        var bg     = i % 2 === 0 ? '#1a1d27' : '#141720';
        var hasVal = benefit.value !== null && benefit.value !== undefined;
        var valStr = hasVal ? '$' + benefit.value : '—';
        var valClr = hasVal ? '#f5a623' : '#6b7280';
        var autoPin = recurringFreqs[benefit.frequency] ? true : false;
        bVals.push(['', u.bank, u.card, benefit.name, valStr, benefit.frequency || '—', benefit.category || '—', '', autoPin, '']);
        bBgs.push(Array(10).fill(bg));
        bFgClrs.push(['#e8eaf0','#6b7280','#9ca3af','#e8eaf0',valClr,'#6b7280','#a78bfa','#e8eaf0','#9ca3af','#e8eaf0']);
        bFgSzs.push([11, 10, 10, 13, 15, 12, 12, 11, 11, 11]);
        bFgWts.push(['normal','normal','normal','normal','bold','normal','normal','normal','normal','normal']);
        bVAlns.push(Array(10).fill('middle'));
        bWraps.push([false, false, false, true, false, false, false, false, false, false]);
      });
      bt.insertRowsBefore(insertRow, m);
      for (var i = 0; i < m; i++) bt.setRowHeight(insertRow + i, 54);
      bt.getRange(insertRow, 1, m, 10)
        .setValues(bVals).setBackgrounds(bBgs).setFontColors(bFgClrs)
        .setFontSizes(bFgSzs).setFontWeights(bFgWts).setVerticalAlignments(bVAlns)
        .setWraps(bWraps).setFontFamily('Arial');
      bt.getRange(insertRow, 9, m, 1).insertCheckboxes().setHorizontalAlignment('center').setVerticalAlignment('middle');
    });
  }

  // Update values for existing credits where JSON has changed (e.g. annual → per-period migration)
  Object.keys(cardInfo).forEach(function(key) {
    var cardData = json[key];
    if (!cardData || !cardData.benefits) return;
    cardData.benefits.forEach(function(b) {
      if (!b.name) return;
      var existing = benefitRows[key + '|' + b.name.trim()];
      if (!existing) return;
      var hasVal = b.value !== null && b.value !== undefined;
      var newValStr = hasVal ? '$' + b.value : '—';
      if (existing.valStr !== newValStr) {
        bt.getRange(existing.sheetRow, 5).setValue(newValStr);
      }
    });
  });

  // Clean up any override JSON previously written to Credits Tracker Notes column (col H)
  if (bt.getLastRow() > 2) {
    var notesVals = bt.getRange(3, 8, bt.getLastRow() - 2, 1).getValues();
    notesVals.forEach(function(cell, idx) {
      if (cell[0] && cell[0].toString().charAt(0) === '{') {
        bt.getRange(idx + 3, 8).setValue('');
      }
    });
  }

  // Write per-benefit override data to Credit Card Raw Data col I (Notes)
  if (raw.getLastRow() > 1) {
    var rawNotes = raw.getRange(2, 1, raw.getLastRow() - 1, 9).getValues();
    rawNotes.forEach(function(r, idx) {
      if (!r[0]) return;
      var cardData = json[r[0] + '|' + r[1]];
      if (!cardData || !cardData.benefits) return;
      var cardOverrides = {};
      cardData.benefits.forEach(function(b) {
        if (b.valueOverrides) cardOverrides[b.name] = b.valueOverrides;
      });
      var newStr = Object.keys(cardOverrides).length ? JSON.stringify(cardOverrides) : '';
      if ((r[8] ? r[8].toString() : '') !== newStr) {
        raw.getRange(idx + 2, 9).setValue(newStr);
      }
    });
  }

}


// ─────────────────────────────────────────
// Run ONCE from the menu to build the shell
// ─────────────────────────────────────────
function setupDashboard() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var RAW = 'Credit Card Raw Data';

  // Create fresh Dashboard tab
  var dash = ss.getSheetByName('Dashboard');
  if (dash) ss.deleteSheet(dash);
  dash = ss.insertSheet('Dashboard', 0);
  dash.setHiddenGridlines(true);
  dash.setTabColor('#4f8ef7');

  // Column widths
  dash.setColumnWidth(1, 24);   // A: left margin
  dash.setColumnWidth(2, 275);  // B: bank / card
  dash.setColumnWidth(3, 180);  // C: points
  dash.setColumnWidth(4, 180);  // D: est. value
  dash.setColumnWidth(5, 155);  // E: annual fee
  dash.setColumnWidth(6, 220);  // F: renewal
  dash.setColumnWidth(7, 170);  // G: reminder + NEW CARD BUTTON (row 1 only)
  dash.setColumnWidth(8, 24);   // H: right margin

  // Hide extra columns
  var maxCols = dash.getMaxColumns();
  if (maxCols > 8) dash.hideColumns(9, maxCols - 8);

  // Full dark background
  dash.getRange('A1:H150')
    .setBackground('#0f1117')
    .setFontColor('#e8eaf0')
    .setFontFamily('Arial');

  // ─── ROW 1: Header bar ───
  dash.setRowHeight(1, 80);
  dash.getRange('A1:H1').setBackground('#1a1d27');

  // Title
  dash.getRange('B1')
    .setValue('💳  CARD DASHBOARD')
    .setFontSize(18).setFontWeight('bold').setFontColor('#e8eaf0')
    .setVerticalAlignment('middle');

  // Date
  dash.getRange('E1:F1').merge()
    .setFormula('="Updated · "&TEXT(NOW(),"mmm d, yyyy")')
    .setFontSize(10).setFontColor('#6b7280')
    .setVerticalAlignment('middle').setHorizontalAlignment('right');

  // G1: empty — real "New Card" button is a drawing object assigned to showEntryForm
  dash.getRange('G1').setValue('').setBackground('#1a1d27');

  // ─── ROW 2: gap ───
  dash.setRowHeight(2, 18);

  // ─── ROW 3: KPI labels ───
  dash.setRowHeight(3, 28);
  dash.getRange('A3:H3').setBackground('#1a1d27');
  dash.getRange('A3:H3').setValues([['','💳  ANNUAL FEES','⭐  TOTAL POINTS','💰  EST. VALUE','🔔  RENEWING SOON','','','']])
    .setFontSize(10).setFontWeight('bold').setFontColor('#6b7280')
    .setVerticalAlignment('bottom');

  // ─── ROW 4: KPI values ───
  dash.setRowHeight(4, 70);
  dash.getRange('A4:H4').setBackground('#1a1d27');

  var q = "'"+RAW+"'";
  dash.getRange('B4')
    .setFormula('=IF(COUNTA('+q+'!A:A)>1,"$"&TEXT(SUM('+q+'!E2:E1000),"#,##0"),"—")')
    .setFontSize(32).setFontWeight('bold').setFontColor('#4f8ef7').setVerticalAlignment('middle');
  dash.getRange('C4')
    .setFormula('=IF(COUNTA('+q+'!A:A)>1,TEXT(SUM('+q+'!C2:C1000),"#,##0"),"—")')
    .setFontSize(32).setFontWeight('bold').setFontColor('#3ecf8e').setVerticalAlignment('middle');
  dash.getRange('D4')
    .setFormula('=IF(COUNTA('+q+'!A:A)>1,"$"&TEXT(SUMPRODUCT('+q+'!C2:C1000,'+q+'!D2:D1000)/100,"#,##0"),"—")')
    .setFontSize(32).setFontWeight('bold').setFontColor('#f5a623').setVerticalAlignment('middle');
  dash.getRange('E4')
    .setFormula('=IF(COUNTA('+q+'!A:A)>1,TEXT(COUNTIFS('+q+'!F2:F1000,">="&TODAY(),'+q+'!F2:F1000,"<="&TODAY()+90),"0"),"—")')
    .setFontSize(32).setFontWeight('bold').setFontColor('#a78bfa').setVerticalAlignment('middle');

  // ─── ROW 5: KPI sub-label (B5 only) ───
  dash.setRowHeight(5, 28);
  dash.getRange('A5:H5').setBackground('#1a1d27');
  dash.getRange('B5')
    .setFormula('=IF(COUNTA('+q+'!A:A)>1,"Across "&(COUNTA('+q+'!A:A)-1)&" card(s)","")')
    .setFontSize(11).setFontColor('#6b7280').setVerticalAlignment('top');
  dash.getRange('C5').clearContent();
  dash.getRange('D5')
    .setFormula('=IF(COUNTA('+q+'!A:A)>1,"Avg "&TEXT(IFERROR(AVERAGEIF('+q+'!D2:D1000,">"&0),0),"0.00")&"¢ per point","")')
    .setFontSize(11).setFontColor('#6b7280').setVerticalAlignment('top');
  dash.getRange('E5').clearContent();
  dash.getRange('B3:E5').setHorizontalAlignment('center');

  // KPI block border
  dash.getRange('B3:E5').setBorder(
    true, true, true, true, false, false,
    '#22263a', SpreadsheetApp.BorderStyle.SOLID
  );

  // ─── ROW 6: gap ───
  dash.setRowHeight(6, 22);

  dash.setFrozenRows(1);

  refreshDashboard();
  SpreadsheetApp.getUi().alert('✅ Dashboard is ready. Use Insert → Drawing to add a "New Card" button, then right-click it → Assign script → showEntryForm.');
}

// Returns true if the gated static layout for `key` should be (re)applied —
// i.e. the stored version differs from LAYOUT_VERSION. Defaults to true if the
// property service can't be read (keeps the old always-apply behaviour as a safe
// fallback in unauthorized contexts).
function layoutNeedsApply_(key) {
  try {
    return PropertiesService.getDocumentProperties().getProperty(key) !== LAYOUT_VERSION;
  } catch (e) {
    return true;
  }
}

// Records that the gated layout for `key` is now at LAYOUT_VERSION. No-op if the
// property service is unavailable (layout simply re-applies next time).
function layoutMarkApplied_(key) {
  try {
    PropertiesService.getDocumentProperties().setProperty(key, LAYOUT_VERSION);
  } catch (e) {}
}

// ─────────────────────────────────────────
// Called on every open — writes live card rows
// ─────────────────────────────────────────
function refreshDashboard() {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var raw  = ss.getSheetByName('Credit Card Raw Data');
  var dash = ss.getSheetByName('Dashboard');
  if (!dash || !raw) return;

  // Static layout (column widths, header row heights, fonts) only needs to be
  // applied on setup or after a LAYOUT_VERSION bump — not on every single edit.
  // Gated behind a document property so onEdit refreshes stay cheap. If the
  // property service is unavailable (e.g. an unauthorized simple-trigger context),
  // fall back to applying it, matching the previous always-on behaviour.
  if (layoutNeedsApply_('dashLayoutV')) {
    // Column widths for 1920×1080
    dash.setColumnWidth(1, 24);  dash.setColumnWidth(2, 275);
    dash.setColumnWidth(3, 180); dash.setColumnWidth(4, 180);
    dash.setColumnWidth(5, 155); dash.setColumnWidth(6, 220);
    dash.setColumnWidth(7, 170); dash.setColumnWidth(8, 24);
    // Title / KPI area row heights + font sizes
    dash.setRowHeight(1, 80); dash.setRowHeight(2, 18);
    dash.setRowHeight(3, 28); dash.setRowHeight(4, 70);
    dash.setRowHeight(5, 28); dash.setRowHeight(6, 28);
    dash.getRange('B1').setFontSize(18);
    dash.getRange('A3:H3').setFontSize(10);
    dash.getRange('B4:E4').setFontSize(32);
    dash.getRange('B5:E5').setFontSize(11);
    // Self-heal: C5 and E5 cleared; D5 shows avg CPP
    dash.getRange('C5').clearContent();
    dash.getRange('E5').clearContent();
    var rawName = "'Credit Card Raw Data'";
    if (!dash.getRange('D5').getFormula()) {
      dash.getRange('D5')
        .setFormula('=IF(COUNTA('+rawName+'!A:A)>1,"Avg "&TEXT(IFERROR(AVERAGEIF('+rawName+'!D2:D1000,">"&0),0),"0.00")&"¢ per point","")')
        .setFontSize(11).setFontColor('#6b7280').setVerticalAlignment('top');
    }
    dash.getRange('B3:E5').setHorizontalAlignment('center');
    layoutMarkApplied_('dashLayoutV');
  }

  var today   = new Date(); today.setHours(0,0,0,0);
  var lastRow = raw.getLastRow();

  // Clear everything from row 7 down
  var clearCount = Math.max(dash.getLastRow() - 6, 80);
  dash.getRange(7, 1, clearCount, 8)
    .clearContent().clearFormat().clearDataValidations()
    .setBackground('#0f1117').setFontColor('#e8eaf0').setFontFamily('Arial');

  // Empty state
  if (lastRow < 2) {
    dash.setRowHeight(7, 56);
    dash.getRange('B7')
      .setValue('No cards yet — click ＋ New Card to add your first card.')
      .setFontSize(13).setFontColor('#6b7280').setVerticalAlignment('middle');
    return;
  }

  var rawData = raw.getRange(2, 1, lastRow - 1, 10).getValues();
  var cards   = rawData.filter(function(r) { return r[0]; });
  if (!cards.length) return;

  // Build override map from Raw Data col I: { "bank|card": { "benefit name": {"12":20} } }
  var rawOverrides = {};
  rawData.forEach(function(r) {
    if (!r[0] || !r[8]) return;
    try { rawOverrides[r[0] + '|' + r[1]] = JSON.parse(r[8]); } catch(e) {}
  });

  // Build used-amounts map from Raw Data col J: { "bank|card": { "benefit name": 15 } }
  var rawUsed = {};
  rawData.forEach(function(r) {
    if (!r[0] || !r[9]) return;
    try { rawUsed[r[0] + '|' + r[1]] = JSON.parse(r[9]); } catch(e) {}
  });

  // Read Credits Tracker — build pinned credits per card (all frequencies)
  var benefitsByCard = {};
  var cardTotals = {};
  var btSheet2 = ss.getSheetByName('Credits Tracker');
  if (btSheet2 && btSheet2.getLastRow() > 2) {
    var btData = btSheet2.getRange(3, 2, btSheet2.getLastRow() - 2, 8).getValues(); // B-I
    btData.forEach(function(br) {
      var bank = br[0], card = br[1];
      if (!bank || !card || !br[2] || br[5] === 'Earn Rate') return;
      if (br[7] !== true) return; // only pinned credits
      var key = bank + '|' + card;
      if (!cardTotals[key]) cardTotals[key] = 0;
      if (!benefitsByCard[key]) benefitsByCard[key] = [];
      var overrideObj = (rawOverrides[key] || {})[br[2]] || null;
      var perPeriod = resolvePerPeriodValue_(br[3], overrideObj);
      if (perPeriod !== null) perPeriod = Math.round(perPeriod * 100) / 100;
      var valDisplay = perPeriod !== null
        ? '$' + (perPeriod % 1 === 0 ? perPeriod : perPeriod.toFixed(2))
        : '—';
      var rd2       = getNextReset_(br[4]);
      var dLeft2    = rd2 ? Math.round((rd2 - today) / 86400000) : null;
      var notesUsed = (rawUsed[key] || {})[br[2]] || 0;
      var remaining = perPeriod !== null ? Math.max(0, Math.round((perPeriod - notesUsed) * 100) / 100) : null;
      // "$X to use" reflects what's still UNused this period, not the full value.
      if (remaining !== null) cardTotals[key] += remaining;
      var daysStr   = dLeft2 === null ? null : dLeft2 < 0 ? 'Overdue' : dLeft2 === 0 ? 'Today' : dLeft2 + ' days left';
      var rstLabel  = rd2
        ? daysStr + '\n' + Utilities.formatDate(rd2, Session.getScriptTimeZone(), 'MMM d')
        : '—';
      benefitsByCard[key].push({ name: br[2], valStr: valDisplay, valueNum: perPeriod, remaining: remaining, frequency: br[4] || '—', resetLabel: rstLabel, daysLeft: dLeft2 });
    });
    Object.keys(benefitsByCard).forEach(function(key) {
      benefitsByCard[key].sort(function(a, b) {
        if (a.valueNum === null && b.valueNum === null) return 0;
        if (a.valueNum === null) return 1;
        if (b.valueNum === null) return -1;
        return b.valueNum - a.valueNum;
      });
    });
  }

  var numCards  = cards.length;
  var maxPoints = Math.max.apply(null, cards.map(function(r) { return parseFloat(r[2]) || 0; })) || 1;

  var byRenewal = cards.slice().sort(function(a, b) {
    if (!a[5] && !b[5]) return 0;
    if (!a[5]) return 1;
    if (!b[5]) return -1;
    return new Date(a[5]) - new Date(b[5]);
  });
  var byPoints = cards.slice().sort(function(a, b) {
    return (parseFloat(b[2]) || 0) - (parseFloat(a[2]) || 0);
  });

  function daysUntil(d) {
    if (!d) return null;
    return Math.round((new Date(d) - today) / 86400000);
  }
  function fmtDate(d) {
    if (!d) return '—';
    return Utilities.formatDate(new Date(d), Session.getScriptTimeZone(), 'MMM d, yyyy');
  }
  function urgencyColor(days) {
    if (days === null) return '#6b7280';
    if (days <= 30)    return '#ff5c5c';
    if (days <= 90)    return '#f5a623';
    return '#3ecf8e';
  }
  function urgencyDot(days) {
    if (days === null) return '';
    if (days < 0)      return '⚠️';
    if (days <= 30)    return '🔴';
    if (days <= 90)    return '🟡';
    return '🟢';
  }

  var r = 7;

  // ── Section header row ──
  dash.setRowHeight(r, 34);
  dash.getRange(r, 1, 1, 8)
    .setBackground('#0f1117').setFontColor('#6b7280')
    .setFontSize(11).setFontWeight('bold').setVerticalAlignment('middle');
  dash.getRange(r, 2).setValue('📊  POINTS BY CARD');
  dash.getRange(r, 5).setValue('🗓  RENEWING SOON');
  r++;

  // ── Side-by-side panel — build arrays and write in one batch ──
  for (var i = 0; i < numCards; i++) { dash.setRowHeight(r + i, 58); }

  var pVals = [], pBgs = [], pFgClrs = [], pFgSzs = [], pFgWts = [], pWraps = [], pVAlns = [], pHAlns = [];

  for (var i = 0; i < numCards; i++) {
    var rowBg   = i % 2 === 0 ? '#1a1d27' : '#141720';
    var lr      = byRenewal[i];
    var lDay    = daysUntil(lr[5]);
    var lDot    = urgencyDot(lDay);
    var lDayStr = lDay === null ? '—' : lDay < 0 ? 'Overdue' : lDay + ' days left';
    var rr      = byPoints[i];
    var rPts    = parseFloat(rr[2]) || 0;

    pVals.push([
      '',
      rr[0] + '\n' + rr[1],
      rPts > 0 ? rPts.toLocaleString() : '—',
      '',
      lr[0] + '\n' + lr[1],
      (lDot ? lDot + '  ' : '') + lDayStr + '\n' + fmtDate(lr[5]),
      lr[4] > 0 ? '$' + lr[4].toLocaleString() + '/yr' : 'No Fee',
      ''
    ]);
    pBgs.push([rowBg, rowBg, rowBg, '#0f1117', rowBg, rowBg, rowBg, rowBg]);
    pFgClrs.push(['#e8eaf0','#e8eaf0','#4f8ef7','#e8eaf0','#e8eaf0',urgencyColor(lDay),'#6b7280','#e8eaf0']);
    pFgSzs.push([11, 13, 14, 11, 13, 12, 11, 11]);
    pFgWts.push(['normal','normal','bold','normal','normal','normal','normal','normal']);
    pWraps.push([false, true, true, false, true, true, false, false]);
    pVAlns.push(Array(8).fill('middle'));
    pHAlns.push(['left', 'left', 'center', 'left', 'left', 'center', 'left', 'left']);
  }

  var pRange = dash.getRange(r, 1, numCards, 8);
  pRange.setValues(pVals).setBackgrounds(pBgs).setFontColors(pFgClrs)
    .setFontSizes(pFgSzs).setFontWeights(pFgWts).setWraps(pWraps)
    .setVerticalAlignments(pVAlns).setHorizontalAlignments(pHAlns).setFontFamily('Arial');
  r += numCards;

  // ── Gap ──
  dash.setRowHeight(r, 28); r++;

  // ── TOP UNUSED BENEFITS header ──
  dash.setRowHeight(r, 32);
  dash.getRange(r, 2).setValue('🔁  RECURRING CREDITS TO USE')
    .setFontSize(11).setFontWeight('bold').setFontColor('#6b7280').setVerticalAlignment('middle');
  r++;

  // ── Column headers ──
  dash.setRowHeight(r, 36);
  dash.getRange(r, 1, 1, 8).setBackground('#1a1d27');
  dash.getRange(r, 2, 1, 5)
    .setValues([['CREDIT', 'VALUE', 'FREQUENCY', 'RESETS', 'USED']])
    .setFontSize(10).setFontWeight('bold').setFontColor('#6b7280').setVerticalAlignment('middle')
    .setHorizontalAlignments([['left', 'center', 'center', 'center', 'center']]);
  r++;

  // ── Per-card unused benefits ──
  cards.forEach(function(row) {
    if (!row[0]) return;
    var bank = row[0], cardName = row[1];
    var key = bank + '|' + cardName;
    var topBenefits = benefitsByCard[key] || [];
    if (topBenefits.length === 0) return; // skip cards with nothing pinned

    var total = cardTotals[key] || 0;
    dash.setRowHeight(r, 34);
    dash.getRange(r, 1, 1, 8).setBackground('#22263a');
    dash.getRange(r, 2)
      .setValue('💳  ' + bank + '  —  ' + cardName)
      .setFontSize(12).setFontWeight('bold').setFontColor('#4f8ef7').setVerticalAlignment('middle');
    if (total > 0) {
      dash.getRange(r, 5)
        .setValue('$' + Math.round(total) + ' to use')
        .setFontSize(11).setFontWeight('bold').setFontColor('#f5a623')
        .setVerticalAlignment('middle').setHorizontalAlignment('right');
    }
    r++;

    if (topBenefits.length === 0) {
      dash.setRowHeight(r, 38);
      dash.getRange(r, 1, 1, 8).setBackground('#1a1d27');
      dash.getRange(r, 2)
        .setValue('—  No recurring credits pinned')
        .setFontSize(11).setFontColor('#6b7280').setVerticalAlignment('middle');
      r++;
    } else {
      var m = topBenefits.length;
      var bVals = [], bBgs = [], bFgClrs = [], bFgSzs = [], bFgWts = [], bVAlns = [], bWraps = [], bFmts = [], bHAlns = [];
      topBenefits.forEach(function(b, i) {
        var bg     = i % 2 === 0 ? '#1a1d27' : '#141720';
        var valClr = b.valueNum !== null ? '#f5a623' : '#6b7280';
        var rstClr = b.daysLeft === null ? '#6b7280' : b.daysLeft <= 5 ? '#ff5c5c' : b.daysLeft <= 14 ? '#f5a623' : '#6b7280';
        // Reflect persisted usage (Raw Data col J): fully-used credits show checked.
        var used   = b.valueNum !== null && b.remaining !== null && b.remaining <= 0;
        bVals.push(['', b.name, b.valueNum !== null ? b.valueNum : '', b.frequency, b.resetLabel || '—', used, '', '']);
        bBgs.push(Array(8).fill(bg));
        bFgClrs.push(['#e8eaf0','#e8eaf0',valClr,'#6b7280',rstClr,'#e8eaf0','#e8eaf0','#e8eaf0']);
        bFgSzs.push([11, 13, 13, 11, 12, 11, 11, 11]);
        bFgWts.push(['normal','normal','bold','normal','normal','normal','normal','normal']);
        bVAlns.push(Array(8).fill('middle'));
        bWraps.push([false, true, false, false, true, false, false, false]);
        bFmts.push(['@', '@', b.valueNum !== null ? '"$"#,##0.##' : '@', '@', '@', '', '@', '@']);
        bHAlns.push(['left', 'left', 'center', 'center', 'center', 'center', 'left', 'left']);
      });
      for (var i = 0; i < m; i++) dash.setRowHeight(r + i, 64);
      dash.getRange(r, 1, m, 8)
        .setValues(bVals).setNumberFormats(bFmts).setBackgrounds(bBgs).setFontColors(bFgClrs)
        .setFontSizes(bFgSzs).setFontWeights(bFgWts).setWraps(bWraps)
        .setVerticalAlignments(bVAlns).setHorizontalAlignments(bHAlns).setFontFamily('Arial');
      dash.getRange(r, 6, m, 1)
        .insertCheckboxes().setHorizontalAlignment('center').setVerticalAlignment('middle');
      r += m;
    }

    // Small gap between cards
    dash.setRowHeight(r, 10);
    dash.getRange(r, 1, 1, 8).setBackground('#0f1117');
    r++;
  });
}

// ── Returns bank→cards map from GitHub JSON for the form dropdown ──
function getCardList() {
  try {
    var json   = fetchAllCardData();
    if (!json) return JSON.stringify({ 'Other': ['Other'] });
    var result = {};
    for (var key in json) {
      if (key.indexOf('|') === -1) continue; // skip _meta, programs, etc.
      var parts = key.split('|');
      var bank  = parts[0];
      var card  = parts[1];
      if (!result[bank]) result[bank] = [];
      result[bank].push(card);
    }
    result['Other'] = ['Other'];
    return JSON.stringify(result);
  } catch(e) {
    Logger.log('getCardList error: ' + e.message);
    return JSON.stringify({ 'Other': ['Other'] });
  }
}

// ── Fetches and parses the full cards JSON from GitHub ──
// Result is cached for 6h via CacheService so repeated calls within a flow
// (e.g. getCardList → getCardPreview → newCard) don't each hit GitHub.
function fetchAllCardData() {
  try {
    var cache = null, cached = null;
    try { cache = CacheService.getScriptCache(); cached = cache.get('cardsJson'); } catch (ce) {}
    if (cached) return JSON.parse(cached);
    var text = UrlFetchApp.fetch(CARD_DATA_URL).getContentText();
    if (cache) { try { cache.put('cardsJson', text, 21600); } catch (pe) {} }
    return JSON.parse(text);
  } catch(e) {
    Logger.log('fetchAllCardData error: ' + e.message);
    return null;
  }
}

// ── Returns a single card's data object ──
function fetchCardData(bank, cardName) {
  var json = fetchAllCardData();
  if (!json) return null;
  return json[bank + '|' + cardName] || null;
}

// ── Returns card data for the form preview step (benefits + earn rates + fee) ──
function getCardPreview(bank, cardName) {
  try {
    var json = fetchAllCardData();
    if (!json) return null;
    var cardData = json[bank + '|' + cardName];
    if (!cardData) return null;
    return JSON.stringify(cardData);
  } catch(e) {
    Logger.log('getCardPreview error: ' + e.message);
    return null;
  }
}

// ── Creates Benefits Tracker sheet if it doesn't exist ──
function setupBenefitsTracker() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Credits Tracker');
  if (sheet) return sheet;

  sheet = ss.insertSheet('Credits Tracker');
  sheet.setHiddenGridlines(true);
  sheet.setTabColor('#3ecf8e');

  // Column widths: A=margin B=bank C=card D=benefit E=value F=freq G=cat H=notes I=pin J=margin
  sheet.setColumnWidth(1,  24);  // A: margin
  sheet.setColumnWidth(2,  200); // B: Bank
  sheet.setColumnWidth(3,  200); // C: Card
  sheet.setColumnWidth(4,  280); // D: Benefit
  sheet.setColumnWidth(5,  105); // E: Value
  sheet.setColumnWidth(6,  125); // F: Frequency
  sheet.setColumnWidth(7,  140); // G: Category
  sheet.setColumnWidth(8,  260); // H: Notes
  sheet.setColumnWidth(9,  80);  // I: PIN
  sheet.setColumnWidth(10, 24);  // J: margin

  // Hide extra columns
  var maxCols = sheet.getMaxColumns();
  if (maxCols > 10) sheet.hideColumns(11, maxCols - 10);

  // Dark background
  sheet.getRange('A1:J300')
    .setBackground('#0f1117')
    .setFontColor('#e8eaf0')
    .setFontFamily('Arial');

  // Row 1: Title
  sheet.setRowHeight(1, 76);
  sheet.getRange('A1:J1').setBackground('#1a1d27');
  sheet.getRange('B1')
    .setValue('✅  CREDITS TRACKER')
    .setFontSize(18).setFontWeight('bold').setFontColor('#e8eaf0')
    .setVerticalAlignment('middle');
  sheet.getRange('G1:I1').merge()
    .setFormula('="Updated · "&TEXT(NOW(),"mmm d, yyyy")')
    .setFontSize(12).setFontColor('#6b7280')
    .setVerticalAlignment('middle').setHorizontalAlignment('right');

  // Row 2: Column headers
  sheet.setRowHeight(2, 42);
  sheet.getRange('A2:J2').setBackground('#1a1d27');
  sheet.getRange('B2:I2')
    .setValues([['BANK','CARD','CREDIT','VALUE','FREQUENCY','CATEGORY','NOTES','📌 PIN']])
    .setFontSize(11).setFontWeight('bold').setFontColor('#6b7280')
    .setVerticalAlignment('middle');
  sheet.getRange('I2').setNote('Check ✓ to pin this benefit to your Dashboard');

  sheet.setFrozenRows(2);
  return sheet;
}

// ── Writes one card's benefits into Benefits Tracker (batched for speed) ──
function addCardBenefits(bank, cardName, cardData) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Credits Tracker');
  if (!sheet) return;

  // Skip if card already exists
  var lastRow = sheet.getLastRow();
  if (lastRow > 2) {
    var existing = sheet.getRange(3, 2, lastRow - 2, 2).getValues();
    for (var i = 0; i < existing.length; i++) {
      if (existing[i][0] === bank && existing[i][1] === cardName) return;
    }
  }

  var benefits      = cardData.benefits  || [];
  var earnRates     = cardData.earnRates || [];
  var startRow      = Math.max(lastRow + 1, 3);

  // Auto-select up to 3 credits to pin by default, priority: Monthly → Semi-Annual → Annual → Quarterly
  var autoSelected = {};
  var autoCount = 0;
  ['Monthly', 'Semi-Annual', 'Annual', 'Quarterly'].forEach(function(freq) {
    if (autoCount >= 3) return;
    benefits
      .filter(function(b) { return b.frequency === freq; })
      .sort(function(a, b) {
        return ((b.value !== null && b.value !== undefined) ? parseFloat(b.value) || 0 : 0)
             - ((a.value !== null && a.value !== undefined) ? parseFloat(a.value) || 0 : 0);
      })
      .forEach(function(b) { if (autoCount < 3) { autoSelected[b.name] = true; autoCount++; } });
  });

  // ── Card group header row (single row — no benefit from batching) ──
  sheet.setRowHeight(startRow, 40);
  sheet.getRange(startRow, 1, 1, 10).setBackground('#22263a');
  sheet.getRange(startRow, 2)
    .setValue('💳  ' + bank + '  —  ' + cardName)
    .setFontSize(14).setFontWeight('bold').setFontColor('#4f8ef7')
    .setVerticalAlignment('middle');
  sheet.getRange(startRow, 5)
    .setValue(cardData.rewardsProgram || '')
    .setFontSize(10).setFontColor('#6b7280').setVerticalAlignment('middle');
  startRow++;

  // ── Earn Rates section — build 2D arrays, write in one batch ──
  if (earnRates.length > 0) {
    sheet.setRowHeight(startRow, 30);
    sheet.getRange(startRow, 1, 1, 10).setBackground('#1a1d27');
    sheet.getRange(startRow, 2)
      .setValue('⚡  EARN RATES')
      .setFontSize(10).setFontWeight('bold').setFontColor('#6b7280')
      .setVerticalAlignment('middle');
    startRow++;

    var n = earnRates.length;
    var eVals = [], eBgs = [], eFgClrs = [], eFgSzs = [], eFgWts = [], eVAlns = [], eWraps = [];
    earnRates.forEach(function(rate, i) {
      var bg = i % 2 === 0 ? '#1a1d27' : '#141720';
      eVals.push(['', bank, cardName, rate.category, rate.multiplier + 'x', 'per $1 spent', 'Earn Rate', '', '', '']);
      eBgs.push(Array(10).fill(bg));
      eFgClrs.push(['#e8eaf0','#6b7280','#9ca3af','#e8eaf0','#a78bfa','#6b7280','#6b7280','#e8eaf0','#e8eaf0','#e8eaf0']);
      eFgSzs.push([11, 10, 10, 12, 15, 10, 10, 11, 11, 11]);
      eFgWts.push(['normal','normal','normal','normal','bold','normal','normal','normal','normal','normal']);
      eVAlns.push(Array(10).fill('middle'));
      eWraps.push([false, false, false, true, false, false, false, false, false, false]);
    });
    for (var i = 0; i < n; i++) sheet.setRowHeight(startRow + i, 50);
    sheet.getRange(startRow, 1, n, 10)
      .setValues(eVals).setBackgrounds(eBgs).setFontColors(eFgClrs)
      .setFontSizes(eFgSzs).setFontWeights(eFgWts).setVerticalAlignments(eVAlns)
      .setWraps(eWraps).setFontFamily('Arial');
    startRow += n;
  }

  // ── Benefits section — build 2D arrays, write in one batch ──
  sheet.setRowHeight(startRow, 30);
  sheet.getRange(startRow, 1, 1, 10).setBackground('#1a1d27');
  sheet.getRange(startRow, 2)
    .setValue('✨  CREDITS')
    .setFontSize(10).setFontWeight('bold').setFontColor('#6b7280')
    .setVerticalAlignment('middle');
  startRow++;

  if (benefits.length > 0) {
    var m = benefits.length;
    var bVals = [], bBgs = [], bFgClrs = [], bFgSzs = [], bFgWts = [], bVAlns = [], bWraps = [];
    benefits.forEach(function(benefit, i) {
      var bg       = i % 2 === 0 ? '#1a1d27' : '#141720';
      var hasVal   = benefit.value !== null && benefit.value !== undefined;
      var valStr   = hasVal ? '$' + benefit.value : '—';
      var valColor = hasVal ? '#f5a623' : '#6b7280';
      var autoPin = autoSelected[benefit.name] === true;
      bVals.push(['', bank, cardName, benefit.name, valStr, benefit.frequency || '—', benefit.category || '—', '', autoPin, '']);
      bBgs.push(Array(10).fill(bg));
      bFgClrs.push(['#e8eaf0','#6b7280','#9ca3af','#e8eaf0',valColor,'#6b7280','#a78bfa','#e8eaf0','#9ca3af','#e8eaf0']);
      bFgSzs.push([11, 10, 10, 13, 15, 12, 12, 11, 11, 11]);
      bFgWts.push(['normal','normal','normal','normal','bold','normal','normal','normal','normal','normal']);
      bVAlns.push(Array(10).fill('middle'));
      bWraps.push([false, false, false, true, false, false, false, false, false, false]);
    });
    for (var i = 0; i < m; i++) sheet.setRowHeight(startRow + i, 54);
    sheet.getRange(startRow, 1, m, 10)
      .setValues(bVals).setBackgrounds(bBgs).setFontColors(bFgClrs)
      .setFontSizes(bFgSzs).setFontWeights(bFgWts).setVerticalAlignments(bVAlns)
      .setWraps(bWraps).setFontFamily('Arial');
    sheet.getRange(startRow, 9, m, 1).insertCheckboxes().setHorizontalAlignment('center').setVerticalAlignment('middle');
  }
}

// ── Creates Annual Fee Analyzer sheet (run once; auto-called from onOpen) ──
function setupAnnualFeeAnalyzer() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Annual Fee Analyzer');
  if (sheet) { refreshAnnualFeeAnalyzer(); return sheet; }

  sheet = ss.insertSheet('Annual Fee Analyzer');
  sheet.setHiddenGridlines(true);
  sheet.setTabColor('#f5a623');

  sheet.setColumnWidth(1, 24);   // A: margin
  sheet.setColumnWidth(2, 280);  // B: card / benefit name
  sheet.setColumnWidth(3, 130);  // C: value used
  sheet.setColumnWidth(4, 180);  // D: coverage bar
  sheet.setColumnWidth(5, 140);  // E: % covered
  sheet.setColumnWidth(6, 225);  // F: verdict
  sheet.setColumnWidth(7, 24);   // G: margin

  var maxCols = sheet.getMaxColumns();
  if (maxCols > 7) sheet.hideColumns(8, maxCols - 7);

  sheet.getRange('A1:G300')
    .setBackground('#0f1117').setFontColor('#e8eaf0').setFontFamily('Arial');

  // Row 1: Title
  sheet.setRowHeight(1, 76);
  sheet.getRange('A1:G1').setBackground('#1a1d27');
  sheet.getRange('B1')
    .setValue('💰  ANNUAL FEE ANALYZER')
    .setFontSize(18).setFontWeight('bold').setFontColor('#e8eaf0')
    .setVerticalAlignment('middle');
  sheet.getRange('E1:F1').merge()
    .setFormula('="Updated · "&TEXT(NOW(),"mmm d, yyyy")')
    .setFontSize(12).setFontColor('#6b7280')
    .setVerticalAlignment('middle').setHorizontalAlignment('right');

  // Row 2: Column headers
  sheet.setRowHeight(2, 36);
  sheet.getRange('A2:G2').setBackground('#1a1d27');
  sheet.getRange('B2:F2')
    .setValues([['CARD', 'POTENTIAL VALUE', 'COVERAGE', '% COVERED', 'VERDICT']])
    .setFontSize(11).setFontWeight('bold').setFontColor('#6b7280').setVerticalAlignment('middle');

  sheet.setFrozenRows(2);
  refreshAnnualFeeAnalyzer();
  return sheet;
}

// ── Populates Annual Fee Analyzer with live benefit usage data ──
function refreshAnnualFeeAnalyzer() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Annual Fee Analyzer');
  if (!sheet) return;

  var raw = ss.getSheetByName('Credit Card Raw Data');
  var bt  = ss.getSheetByName('Credits Tracker');
  if (!raw) return;

  // Column widths for 1920×1080 — only on setup / LAYOUT_VERSION bump (see refreshDashboard)
  if (layoutNeedsApply_('afaLayoutV')) {
    sheet.setColumnWidth(1, 24);  // A: margin
    sheet.setColumnWidth(2, 280); // B: card/benefit name
    sheet.setColumnWidth(3, 130); // C: value used
    sheet.setColumnWidth(4, 180); // D: coverage bar
    sheet.setColumnWidth(5, 140); // E: % covered
    sheet.setColumnWidth(6, 225); // F: verdict
    sheet.setColumnWidth(7, 24);  // G: margin
    layoutMarkApplied_('afaLayoutV');
  }

  // Read card list (9 cols to include col I override data)
  var rawRows = raw.getLastRow() > 1
    ? raw.getRange(2, 1, raw.getLastRow() - 1, 9).getValues()
    : [];
  var cards = rawRows.filter(function(r) { return r[0]; }).map(function(r) {
    return { bank: r[0], cardName: r[1], annualFee: parseFloat(r[4]) || 0, benefits: [], potentialValue: 0 };
  });

  // Build override map from Raw Data col I
  var rawOverridesAFA = {};
  rawRows.forEach(function(r) {
    if (!r[0] || !r[8]) return;
    try { rawOverridesAFA[r[0] + '|' + r[1]] = JSON.parse(r[8]); } catch(e) {}
  });

  // Title rows — keep in sync on every refresh
  sheet.setRowHeight(1, 76); sheet.setRowHeight(2, 36);
  sheet.getRange('B1').setFontSize(18);
  sheet.getRange('B2:F2').setFontSize(11);

  // Clear from row 3 down
  var clearCount = Math.max(sheet.getLastRow() - 2, 100);
  sheet.getRange(3, 1, clearCount, 7)
    .clearContent().clearFormat()
    .setBackground('#0f1117').setFontColor('#e8eaf0').setFontFamily('Arial');

  if (!cards.length) {
    sheet.setRowHeight(3, 56);
    sheet.getRange('B3')
      .setValue('No cards yet — add cards to get started.')
      .setFontSize(13).setFontColor('#6b7280').setVerticalAlignment('middle');
    return;
  }

  // Index cards for lookup
  var cardIndex = {};
  cards.forEach(function(c) { cardIndex[c.bank + '|' + c.cardName] = c; });

  // Read Credits Tracker — cols B-I (bank, card, credit, value, freq, category, notes, PIN)
  if (bt && bt.getLastRow() > 2) {
    var btRows = bt.getRange(3, 2, bt.getLastRow() - 2, 8).getValues();
    var periodsPerYear = { 'Monthly': 12, 'Quarterly': 4, 'Semi-Annual': 2 };
    btRows.forEach(function(r) {
      var bank = r[0], card = r[1];
      if (!bank || !card) return;
      if (r[5] === 'Earn Rate' || !r[2]) return;
      var c = cardIndex[bank + '|' + card];
      if (!c) return;
      var overrideObj = (rawOverridesAFA[bank + '|' + card] || {})[r[2]] || null;
      var perPeriod = resolvePerPeriodValue_(r[3], overrideObj);
      var annualVal = perPeriod !== null
        ? Math.round(perPeriod * (periodsPerYear[r[4]] || 1) * 100) / 100
        : null;
      var pinned = r[7] === true;
      if (pinned && annualVal !== null) c.potentialValue += annualVal;
      c.benefits.push({ name: r[2], value: perPeriod, frequency: r[4] || '', pinned: pinned });
    });
  }

  var r = 3;

  cards.forEach(function(c) {
    var af     = c.annualFee;
    var potVal = Math.round(c.potentialValue);
    var pct    = af > 0 ? Math.min(100, Math.round(potVal / af * 100)) : 100;
    var gap    = Math.round(af - potVal);

    var verdictText, verdictClr, usedClr, sparkColor;
    if (af === 0) {
      verdictText = 'No Annual Fee';
      verdictClr = '#6b7280'; usedClr = '#6b7280'; sparkColor = '#6b7280';
    } else if (potVal >= af) {
      verdictText = '✓  Worth It  +$' + (potVal - af);
      verdictClr = '#3ecf8e'; usedClr = '#3ecf8e'; sparkColor = '#3ecf8e';
    } else if (pct >= 50) {
      verdictText = '⚠  $' + gap + ' gap remaining';
      verdictClr = '#f5a623'; usedClr = '#f5a623'; sparkColor = '#f5a623';
    } else {
      verdictText = '✕  $' + gap + ' gap remaining';
      verdictClr = '#ff5c5c'; usedClr = '#ff5c5c'; sparkColor = '#ff5c5c';
    }

    // ── Card header ──
    sheet.setRowHeight(r, 40);
    sheet.getRange(r, 1, 1, 7).setBackground('#22263a');
    sheet.getRange(r, 2)
      .setValue('💳  ' + c.bank + '  —  ' + c.cardName)
      .setFontSize(14).setFontWeight('bold').setFontColor('#4f8ef7').setVerticalAlignment('middle');
    if (af > 0) {
      sheet.getRange(r, 6)
        .setValue('$' + af + '/yr')
        .setFontSize(12).setFontColor('#6b7280').setVerticalAlignment('middle').setHorizontalAlignment('right');
    }
    r++;

    // ── Summary row ──
    sheet.setRowHeight(r, 56);
    sheet.getRange(r, 1, 1, 7).setBackground('#1a1d27');
    sheet.getRange(r, 2)
      .setValue(af > 0 ? '$' + potVal + ' value' : 'No Annual Fee')
      .setFontSize(22).setFontWeight('bold').setFontColor(usedClr).setVerticalAlignment('middle');
    if (af > 0) {
      sheet.getRange(r, 3)
        .setValue('of $' + af + '/yr')
        .setFontSize(12).setFontColor('#6b7280').setVerticalAlignment('middle');
      sheet.getRange(r, 4)
        .setFormula('=SPARKLINE({' + pct + ',' + Math.max(0, 100 - pct) + '},{"charttype","bar";"color1","' + sparkColor + '";"color2","#22263a";"bgcolor","#1a1d27"})')
        .setVerticalAlignment('middle');
      sheet.getRange(r, 5)
        .setValue(pct + '%')
        .setFontSize(18).setFontWeight('bold').setFontColor(usedClr).setVerticalAlignment('middle');
    }
    sheet.getRange(r, 6)
      .setValue(verdictText)
      .setFontSize(12).setFontWeight('bold').setFontColor(verdictClr).setVerticalAlignment('middle');
    r++;

    // ── Benefit rows (batched) ──
    if (c.benefits.length > 0) {
      var m = c.benefits.length;
      var bVals = [], bBgs = [], bFgClrs = [], bFgSzs = [], bFgWts = [], bVAlns = [];
      c.benefits.forEach(function(b, i) {
        var bg      = i % 2 === 0 ? '#1a1d27' : '#141720';
        var checkCl = b.pinned ? '#4f8ef7' : '#6b7280';
        var nameClr = b.pinned ? '#e8eaf0' : '#6b7280';
        var valStr  = b.value !== null ? '$' + b.value : '—';
        var valClr  = (b.pinned && b.value !== null) ? '#f5a623' : '#6b7280';
        var statusV = b.pinned ? 'Pinned' : '—';
        var statusC = b.pinned ? '#4f8ef7' : '#6b7280';
        bVals.push(['', b.pinned ? '📌' : '○', b.name, valStr, b.frequency || '—', statusV, '']);
        bBgs.push(Array(7).fill(bg));
        bFgClrs.push(['#e8eaf0', checkCl, nameClr, valClr, '#6b7280', statusC, '#e8eaf0']);
        bFgSzs.push([11, 14, 13, 14, 11, 11, 11]);
        bFgWts.push(['normal', 'bold', 'normal', 'bold', 'normal', 'normal', 'normal']);
        bVAlns.push(Array(7).fill('middle'));
      });
      for (var i = 0; i < m; i++) sheet.setRowHeight(r + i, 44);
      sheet.getRange(r, 1, m, 7)
        .setValues(bVals).setBackgrounds(bBgs).setFontColors(bFgClrs)
        .setFontSizes(bFgSzs).setFontWeights(bFgWts).setVerticalAlignments(bVAlns)
        .setFontFamily('Arial');
      r += m;
    } else {
      sheet.setRowHeight(r, 38);
      sheet.getRange(r, 1, 1, 7).setBackground('#1a1d27');
      sheet.getRange(r, 2)
        .setValue('No credits tracked for this card.')
        .setFontSize(12).setFontColor('#6b7280').setVerticalAlignment('middle');
      r++;
    }

    // ── Spacer ──
    sheet.setRowHeight(r, 18);
    sheet.getRange(r, 1, 1, 7).setBackground('#0f1117');
    r++;
  });
}

