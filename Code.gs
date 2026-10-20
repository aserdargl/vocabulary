var SPREADSHEET_ID = '1cUYfuFg7OSsWjCtLA7jiB77bGJITR6V0jF2ZUnyrXtc';

function getSpreadsheet_() {
  try {
    var active = SpreadsheetApp.getActiveSpreadsheet();
    if (active) return active;
  } catch (e) {}
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🔊 Pronunciation')
    .addItem('📱 Open Vocabulary App (Full Screen)', 'uygulamaEkraniniAc')
    .addSeparator()
    .addItem('🔊 Open Audio Sidebar', 'sesPaneliniAc')
    .addItem('🧹 Clean Speaker Links', 'hoparlorLinkleriniTemizle')
    .addToUi();
  try { hoparlorLinkleriniTemizle(); } catch (e) {}
}

function onEdit(e) {
  try {
    if (!e || !e.range) return;
    var sheet = e.range.getSheet();
    if (sheet.getName() !== 'Sheet1') return;
    var layout = getSheetLayout_(sheet);
    if (layout.audio < 1) return;
    var startRow = e.range.getRow();
    var numRows = e.range.getNumRows();
    if (startRow + numRows - 1 < layout.startRow) return;

    var wordColLetter = String.fromCharCode(64 + layout.word);
    for (var i = 0; i < numRows; i++) {
      var r = startRow + i;
      if (r < layout.startRow) continue;
      var sesCell = sheet.getRange(r, layout.audio);
      var f = sesCell.getFormula();
      if (!f || f.indexOf('HYPERLINK') !== -1) {
        sesCell.setFormula('=IF(ISBLANK(' + wordColLetter + r + '), "", "🔊")');
      }
    }
  } catch (err) {}
}

function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'json') {
    return ContentService
      .createTextOutput(JSON.stringify(verileriGetir()))
      .setMimeType(ContentService.MimeType.JSON);
  }
  if (e && e.parameter && e.parameter.action === 'add' && e.parameter.words) {
    var parsed = JSON.parse(e.parameter.words);
    var res = kelimeEkle(parsed);
    return ContentService
      .createTextOutput(JSON.stringify(res))
      .setMimeType(ContentService.MimeType.JSON);
  }
  return HtmlService.createHtmlOutput(uygulamaHtmlGetir())
    .setTitle('Vocabulary')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents || '{}');
    if (body.action === 'addWords' && body.words) {
      return ContentService
        .createTextOutput(JSON.stringify(kelimeEkle(body.words)))
        .setMimeType(ContentService.MimeType.JSON);
    }
    if (body.action === 'setStatus' && body.row) {
      return ContentService
        .createTextOutput(JSON.stringify(durumGuncelle(body.row, body.status)))
        .setMimeType(ContentService.MimeType.JSON);
    }
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: 'Unknown action' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function uygulamaEkraniniAc() {
  var html = HtmlService.createHtmlOutput(uygulamaHtmlGetir())
    .setWidth(1240)
    .setHeight(780);
  SpreadsheetApp.getUi().showModalDialog(html, 'Vocabulary');
}

function hoparlorLinkleriniTemizle() {
  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName('Sheet1');
  if (!sheet) return { ok: false };
  var layout = getSheetLayout_(sheet);
  if (layout.audio < 1) return { ok: true, count: 0 };
  var maxRows = Math.max(sheet.getMaxRows(), layout.startRow);
  var numRows = maxRows - layout.startRow + 1;
  var range = sheet.getRange(layout.startRow, layout.audio, numRows, 1);
  var existing = range.getFormulas();
  var updated = [];
  var changed = false;
  var wordColLetter = String.fromCharCode(64 + layout.word);
  for (var i = 0; i < existing.length; i++) {
    var f = existing[i][0];
    var r = layout.startRow + i;
    var targetFormula = '=IF(ISBLANK(' + wordColLetter + r + '), "", "🔊")';
    if (f !== targetFormula) {
      updated.push([targetFormula]);
      changed = true;
    } else {
      updated.push([f]);
    }
  }
  if (changed) {
    range.setFormulas(updated);
  }
  return { ok: true, count: numRows };
}

function getSheetLayout_(sheet) {
  var maxR = Math.min(sheet.getLastRow(), 6);
  var maxC = Math.min(sheet.getLastColumn(), 15);
  if (maxR < 1 || maxC < 2) {
    return { headerRow: 5, startRow: 6, date: 2, word: 3, pron: 4, audio: 5, type: 6, level: 7, meaning: 8, example: 9, status: 10, known: 11, unknown: 12 };
  }
  var topVals = sheet.getRange(1, 1, maxR, maxC).getDisplayValues();
  for (var r = 0; r < topVals.length; r++) {
    var rowArr = topVals[r];
    var wordCol = -1;
    for (var c = 0; c < rowArr.length; c++) {
      var h = (rowArr[c] || '').toString().trim().toLowerCase();
      if (h === 'word' || h === 'kelime') { wordCol = c + 1; break; }
    }
    if (wordCol !== -1) {
      var map = { headerRow: r + 1, startRow: r + 2, date: 2, word: wordCol, pron: 4, audio: -1, type: 6, level: 7, meaning: 8, example: 9, status: 10, known: 11, unknown: 12 };
      for (var k = 0; k < rowArr.length; k++) {
        var hk = (rowArr[k] || '').toString().trim().toLowerCase();
        var colNum = k + 1;
        if (hk === 'date' || hk === 'tarih') map.date = colNum;
        else if (hk === 'word' || hk === 'kelime') map.word = colNum;
        else if (hk.indexOf('pronun') !== -1 || hk.indexOf('okunu') !== -1) map.pron = colNum;
        else if (hk === 'audio' || hk.indexOf('sesli') !== -1) map.audio = colNum;
        else if (hk === 'type' || hk === 'tür' || hk === 'tur') map.type = colNum;
        else if (hk === 'level' || hk === 'seviye') map.level = colNum;
        else if (hk.indexOf('meaning') !== -1 || hk.indexOf('anlam') !== -1) map.meaning = colNum;
        else if (hk.indexOf('example') !== -1 || hk.indexOf('örnek') !== -1 || hk.indexOf('ornek') !== -1) map.example = colNum;
        else if (hk === 'status' || hk === 'durum') map.status = colNum;
        else if (hk === 'known' || hk === 'biliyorum') map.known = colNum;
        else if (hk === 'unknown' || hk === 'bilmiyorum') map.unknown = colNum;
      }
      return map;
    }
  }
  return { headerRow: 5, startRow: 6, date: 2, word: 3, pron: 4, audio: 5, type: 6, level: 7, meaning: 8, example: 9, status: 10, known: 11, unknown: 12 };
}

function verileriGetir() {
  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName('Sheet1');
  if (!sheet) return { ok: false, words: [] };

  var layout = getSheetLayout_(sheet);
  var lastRow = sheet.getLastRow();
  if (lastRow < layout.startRow) {
    return { ok: true, words: [], sheetUrl: ss.getUrl(), webAppUrl: getWebAppUrl_() };
  }

  var numRows = lastRow - layout.startRow + 1;
  var maxCol = Math.max(sheet.getLastColumn(), 12);
  var displayVals = sheet.getRange(layout.startRow, 1, numRows, maxCol).getDisplayValues();
  var rawVals = sheet.getRange(layout.startRow, 1, numRows, maxCol).getValues();

  var words = [];
  for (var i = 0; i < numRows; i++) {
    var dRow = displayVals[i];
    var rRow = rawVals[i];
    var kelime = (dRow[layout.word - 1] || '').toString().trim();
    if (!kelime) continue;

    var kVal = rRow[layout.known - 1];
    var uVal = rRow[layout.unknown - 1];
    var kDisp = dRow[layout.known - 1];
    var uDisp = dRow[layout.unknown - 1];
    var biliyorum = kVal === true || kDisp === 'TRUE';
    var bilmiyorum = uVal === true || uDisp === 'TRUE';
    var durumKod = biliyorum ? 'biliyorum' : (bilmiyorum ? 'bilmiyorum' : 'beklemede');

    words.push({
      row: layout.startRow + i,
      tarih: (dRow[layout.date - 1] || '').toString().trim(),
      kelime: kelime,
      okunus: (dRow[layout.pron - 1] || '').toString().trim(),
      tur: (dRow[layout.type - 1] || '').toString().trim(),
      seviye: (dRow[layout.level - 1] || '').toString().trim() || 'B1',
      anlam: (dRow[layout.meaning - 1] || '').toString().trim(),
      ornek: (dRow[layout.example - 1] || '').toString().trim(),
      durumMetni: (dRow[layout.status - 1] || '').toString().trim(),
      biliyorum: biliyorum,
      bilmiyorum: bilmiyorum,
      durum: durumKod
    });
  }

  return {
    ok: true,
    words: words,
    sheetUrl: ss.getUrl(),
    webAppUrl: getWebAppUrl_()
  };
}

function getWebAppUrl_() {
  try {
    return ScriptApp.getService().getUrl() || '';
  } catch (e) {
    return '';
  }
}

function durumGuncelle(satirNo, yeniDurum) {
  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName('Sheet1');
  var r = parseInt(satirNo, 10);
  if (!sheet || isNaN(r) || r < 2) return { ok: false, error: 'Invalid row' };

  var layout = getSheetLayout_(sheet);
  var biliyorum = yeniDurum === 'biliyorum';
  var bilmiyorum = yeniDurum === 'bilmiyorum';
  sheet.getRange(r, layout.known).setValue(biliyorum);
  sheet.getRange(r, layout.unknown).setValue(bilmiyorum);

  return {
    ok: true,
    row: r,
    durum: yeniDurum,
    biliyorum: biliyorum,
    bilmiyorum: bilmiyorum
  };
}

function kelimeEkle(kelimeler) {
  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName('Sheet1');
  if (!sheet) return { ok: false, error: 'Sheet1 not found' };

  var layout = getSheetLayout_(sheet);
  var list = Array.isArray(kelimeler) ? kelimeler : [kelimeler];
  var lastRow = Math.max(sheet.getLastRow(), layout.startRow);
  var wordVals = sheet.getRange(layout.startRow, layout.word, Math.max(lastRow - layout.startRow + 1, 1), 1).getDisplayValues();

  var existingMap = {};
  var firstEmptyRow = lastRow + 1;
  var foundEmpty = false;

  for (var i = 0; i < wordVals.length; i++) {
    var w = (wordVals[i][0] || '').toString().trim();
    if (w) {
      existingMap[w.toLowerCase()] = true;
    } else if (!foundEmpty) {
      firstEmptyRow = i + layout.startRow;
      foundEmpty = true;
    }
  }

  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'GMT+3', 'yyyy-MM-dd');
  var added = [];
  var targetRow = firstEmptyRow;
  var wordColLetter = String.fromCharCode(64 + layout.word);

  for (var j = 0; j < list.length; j++) {
    var item = list[j];
    if (!item || !item.kelime) continue;
    var cleanWord = item.kelime.toString().trim();
    if (!cleanWord) continue;
    if (existingMap[cleanWord.toLowerCase()] && !item.allowDuplicate) continue;

    var tarih = (item.tarih || today).toString().trim();
    var okunus = (item.okunus || '').toString().trim();
    var tur = (item.tur || '').toString().trim();
    var seviye = (item.seviye || 'B2').toString().trim().toUpperCase();
    var anlam = (item.anlam || '').toString().trim();
    var ornek = (item.ornek || '').toString().trim();
    var biliyorum = item.durum === 'biliyorum';
    var bilmiyorum = item.durum === 'bilmiyorum';

    if (targetRow > layout.startRow) {
      sheet.getRange(layout.startRow, 2, 1, 11).copyTo(
        sheet.getRange(targetRow, 2, 1, 11),
        SpreadsheetApp.CopyPasteType.PASTE_FORMAT,
        false
      );
      if (layout.status > 0) {
        sheet.getRange(layout.startRow, layout.status).copyTo(sheet.getRange(targetRow, layout.status));
      }
    }

    sheet.getRange(targetRow, layout.date).setValue(tarih);
    sheet.getRange(targetRow, layout.word).setValue(cleanWord);
    sheet.getRange(targetRow, layout.pron).setValue(okunus);
    if (layout.audio > 0) {
      sheet.getRange(targetRow, layout.audio).setFormula('=IF(ISBLANK(' + wordColLetter + targetRow + '), "", "🔊")');
    }
    sheet.getRange(targetRow, layout.type).setValue(tur);
    sheet.getRange(targetRow, layout.level).setValue(seviye);
    sheet.getRange(targetRow, layout.meaning).setValue(anlam);
    sheet.getRange(targetRow, layout.example).setValue(ornek);

    var cbKnown = sheet.getRange(targetRow, layout.known);
    cbKnown.insertCheckboxes();
    cbKnown.setValue(biliyorum);

    var cbUnknown = sheet.getRange(targetRow, layout.unknown);
    cbUnknown.insertCheckboxes();
    cbUnknown.setValue(bilmiyorum);

    existingMap[cleanWord.toLowerCase()] = true;
    added.push({ row: targetRow, kelime: cleanWord });
    targetRow++;
  }

  return { ok: true, addedCount: added.length, added: added };
}

function sesPaneliniAc() {
  var html = HtmlService.createHtmlOutput(
    '<div style="font-family:sans-serif;padding:14px;text-align:center;">' +
    '<div style="font-size:13px;color:#475569;margin-bottom:8px;">Click any <b>🔊</b> icon in the table to hear the word:</div>' +
    '<div id="kelime" style="font-size:22px;font-weight:bold;color:#4f46e5;margin:14px 0;">Ready</div>' +
    '<button onclick="tekrarOku()" style="background:#4f46e5;color:#fff;border:none;padding:10px 18px;border-radius:8px;cursor:pointer;font-weight:bold;width:100%;margin-bottom:8px;">🔊 Listen Again</button>' +
    '<button onclick="google.script.run.uygulamaEkraniniAc()" style="background:#0f5132;color:#fff;border:none;padding:10px 18px;border-radius:8px;cursor:pointer;font-weight:bold;width:100%;">📱 Open Vocabulary App</button>' +
    '<script>' +
    'var sonKelime = "";' +
    'function oku(metin) {' +
    '  if (!metin) return;' +
    '  document.getElementById("kelime").innerText = metin;' +
    '  var ses = new Audio("https://ssl.gstatic.com/dictionary/static/sounds/20200429/" + encodeURIComponent(metin.toLowerCase().trim()) + "--_us_1.mp3");' +
    '  ses.play().catch(function(){' +
    '    speechSynthesis.cancel();' +
    '    var u = new SpeechSynthesisUtterance(metin);' +
    '    u.lang = "en-US";' +
    '    speechSynthesis.speak(u);' +
    '  });' +
    '}' +
    'function tekrarOku() { oku(sonKelime); }' +
    'function kontrolEt() {' +
    '  google.script.run' +
    '    .withSuccessHandler(function(k){' +
    '      if (k) { sonKelime = k; oku(k); }' +
    '      setTimeout(kontrolEt, 350);' +
    '    })' +
    '    .withFailureHandler(function(){' +
    '      setTimeout(kontrolEt, 600);' +
    '    })' +
    '    .seciliKelimeyiAl();' +
    '}' +
    'kontrolEt();' +
    '<\/script></div>'
  ).setTitle('🔊 Pronunciation');
  try { hoparlorLinkleriniTemizle(); } catch (e) {}
  SpreadsheetApp.getUi().showSidebar(html);
}

function seciliKelimeyiAl() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var cell = sheet.getActiveCell();
  if (!cell || cell.getRow() < 6) return '';

  var sheetName = sheet.getName();
  var isHavuz = (sheetName === 'Unknown Words Pool' || sheetName === 'Bilmediğim Kelimeler Havuzu');
  var sesCol = isHavuz ? 8 : 5;
  var wordCol = isHavuz ? 6 : 3;

  var val = cell.getValue().toString().trim();
  if (cell.getColumn() !== sesCol && val !== '🔊') return '';

  var row = cell.getRow();
  if (!isHavuz && cell.getColumn() === 5) {
    var f = cell.getFormula();
    if (!f || f.indexOf('HYPERLINK') !== -1) {
      cell.setFormula('=IF(ISBLANK(C' + row + '), "", "🔊")');
    }
  }

  var word = sheet.getRange(row, wordCol).getValue().toString().trim();
  sheet.getRange(row, wordCol).activate();
  SpreadsheetApp.flush();

  return /^[A-Za-z\s'\-]+$/.test(word) ? word : '';
}

function uygulamaHtmlGetir() {
  return [
    '<!DOCTYPE html>',
    '<html lang="en">',
    '<head>',
    '<meta charset="UTF-8">',
    '<meta name="viewport" content="width=device-width, initial-scale=1.0">',
    '<title>Vocabulary</title>',
    '<style>',
    '  :root { --fs: 12px; }',
    '  * { box-sizing: border-box; margin: 0; padding: 0; }',
    '  body { font-family: Roboto, Arial, sans-serif; font-size: var(--fs); color: #1f2937; background: #ffffff; overflow-x: auto; padding-bottom: 36px; }',
    '  /* Compact Toolbar */',
    '  .toolbar { display: flex; align-items: center; justify-content: space-between; padding: 6px 10px; background: #f8fafc; border-bottom: 1px solid #cbd5e1; gap: 8px; flex-wrap: wrap; position: sticky; top: 0; z-index: 20; }',
    '  .toolbar-left { display: flex; align-items: center; gap: 8px; flex-wrap: wrap; }',
    '  .mini-stats { display: flex; align-items: center; gap: 6px; font-size: 11.5px; color: #475569; }',
    '  .mini-stat { padding: 2px 7px; border-radius: 4px; background: #e2e8f0; font-weight: 600; }',
    '  .mini-stat.unk { background: #fee2e2; color: #991b1b; }',
    '  .mini-stat.knw { background: #dcfce7; color: #166534; }',
    '  .toolbar-right { display: flex; align-items: center; gap: 6px; flex-wrap: wrap; }',
    '  .inp { border: 1px solid #cbd5e1; border-radius: 4px; padding: 3px 8px; font-size: 12px; height: 26px; outline: none; background: #fff; }',
    '  .inp:focus { border-color: #137333; }',
    '  .btn { border: 1px solid #cbd5e1; background: #fff; color: #1f2937; border-radius: 4px; padding: 0 9px; height: 26px; font-size: 11.5px; font-weight: 600; cursor: pointer; display: inline-flex; align-items: center; gap: 4px; text-decoration: none; user-select: none; }',
    '  .btn:hover { background: #f1f5f9; }',
    '  .btn-green { background: #135c37; color: #fff; border-color: #135c37; }',
    '  .btn-green:hover { background: #0e4629; }',
    '  /* Iconic Speaker On/Off Button */',
    '  .spk-toggle { width: 30px; height: 26px; border-radius: 4px; border: 1px solid #cbd5e1; background: #fff; cursor: pointer; display: inline-flex; align-items: center; justify-content: center; font-size: 14px; transition: all 0.15s; user-select: none; padding: 0; }',
    '  .spk-toggle.on { background: #e6f4ea; border-color: #81c995; color: #137333; }',
    '  .spk-toggle.off { background: #f1f5f9; border-color: #cbd5e1; color: #94a3b8; opacity: 0.8; }',
    '  .spk-toggle:hover { transform: scale(1.04); }',
    '  /* Font Size Control */',
    '  .font-ctrl { display: inline-flex; align-items: center; border: 1px solid #cbd5e1; border-radius: 4px; background: #fff; height: 26px; overflow: hidden; user-select: none; }',
    '  .font-btn { border: none; background: transparent; color: #1f2937; padding: 0 8px; height: 100%; font-size: 11.5px; font-weight: 700; cursor: pointer; }',
    '  .font-btn:hover { background: #e2e8f0; }',
    '  .font-val { font-size: 10.5px; color: #475569; padding: 0 5px; border-left: 1px solid #e2e8f0; border-right: 1px solid #e2e8f0; min-width: 34px; text-align: center; font-weight: 600; }',
    '  /* Tight Spreadsheet Table */',
    '  .table-wrap { padding: 6px 8px; }',
    '  table.sheet-table { width: 100%; border-collapse: collapse; table-layout: auto; font-size: var(--fs); border: 1px solid #bdc1c6; }',
    '  table.sheet-table th { background: #135c37; color: #ffffff; font-weight: 700; font-size: calc(var(--fs) - 0.5px); padding: 5px 7px; border: 1px solid #0e4629; text-align: left; white-space: nowrap; position: sticky; top: 38px; z-index: 10; }',
    '  table.sheet-table th.center, table.sheet-table td.center { text-align: center; }',
    '  table.sheet-table td { padding: 3px 7px; border: 1px solid #dadce0; vertical-align: middle; line-height: 1.3; color: #202124; font-size: var(--fs); }',
    '  table.sheet-table tr:nth-child(even) td { background: #fcfcfc; }',
    '  table.sheet-table tr:hover td { background: #f1f8f4; }',
    '  table.sheet-table tr.row-known td { background: #f4fbf7; }',
    '  th.col-row { width: 26px; text-align: center; }',
    '  td.col-row { width: 26px; color: #5f6368; font-size: calc(var(--fs) - 1px); background: #f8f9fa !important; text-align: center; user-select: none; }',
    '  .col-date { width: 76px; white-space: nowrap; color: #3c4043; font-size: calc(var(--fs) - 0.5px); }',
    '  .col-lvl { width: 44px; text-align: center; font-weight: 700; }',
    '  .col-type { white-space: nowrap; width: 78px; font-size: calc(var(--fs) - 0.5px); color: #374151; }',
    '  .col-word { font-weight: 700; color: #111827; white-space: nowrap; cursor: pointer; transition: color 0.1s, background-color 0.15s; }',
    '  .col-word:hover { color: #135c37; text-decoration: underline; }',
    '  .col-word.speaking { background-color: #dcfce7 !important; color: #135c37; }',
    '  body.audio-muted .col-word { cursor: default; }',
    '  body.audio-muted .col-word:hover { color: #111827; text-decoration: none; }',
    '  .col-phonetic { color: #4b5563; font-size: calc(var(--fs) - 0.5px); white-space: nowrap; }',
    '  .col-meaning { min-width: 125px; }',
    '  .col-example { min-width: 190px; }',
    '  .col-status { width: 122px; text-align: center; white-space: nowrap; }',
    '  .status-pill { display: inline-block; padding: 2px 8px; border-radius: 99px; font-size: calc(var(--fs) - 1.5px); font-weight: 600; background: #e8eaed; color: #3c4043; min-width: 112px; text-align: center; }',
    '  .status-pill.bilmiyorum { background: #fce8e6; color: #c5221f; }',
    '  .status-pill.biliyorum { background: #e6f4ea; color: #137333; }',
    '  .col-cb { width: 52px; text-align: center; }',
    '  input.sheet-cb { width: 15px; height: 15px; cursor: pointer; accent-color: #135c37; vertical-align: middle; }',
    '  /* Pool Layout */',
    '  .pool-layout { display: grid; grid-template-columns: 185px 1fr; gap: 14px; align-items: start; }',
    '  /* Bottom Sheet Tabs */',
    '  .bottom-bar { position: fixed; bottom: 0; left: 0; right: 0; height: 30px; background: #f1f3f4; border-top: 1px solid #dadce0; display: flex; align-items: center; padding: 0 12px; gap: 2px; z-index: 30; }',
    '  .sheet-tab { border: 1px solid transparent; border-top: none; background: transparent; padding: 0 14px; height: 29px; font-size: 12px; font-weight: 500; color: #5f6368; cursor: pointer; display: inline-flex; align-items: center; gap: 6px; }',
    '  .sheet-tab:hover { background: #e8eaed; }',
    '  .sheet-tab.active { background: #ffffff; color: #137333; font-weight: 700; border-left-color: #dadce0; border-right-color: #dadce0; border-bottom: 2px solid #137333; }',
    '  /* Compact Add Row Bar */',
    '  .add-bar { display: none; background: #f8fafc; border: 1px solid #cbd5e1; border-radius: 4px; padding: 6px 10px; margin: 6px 8px 0; align-items: center; gap: 6px; flex-wrap: wrap; }',
    '  .add-bar.open { display: flex; }',
    '  .toast { position: fixed; bottom: 36px; right: 14px; background: #202124; color: #fff; padding: 5px 10px; border-radius: 4px; font-size: 11.5px; opacity: 0; pointer-events: none; transition: opacity 0.15s; z-index: 100; }',
    '  .toast.show { opacity: 1; }',
    '  /* Flashcards View */',
    '  .fc-wrap { max-width: 560px; margin: 18px auto 32px; padding: 0 12px; user-select: none; }',
    '  .fc-top { display: flex; align-items: center; justify-content: space-between; margin-bottom: 8px; font-size: 12px; color: #475569; font-weight: 600; }',
    '  .fc-progress-track { width: 100%; height: 5px; background: #e2e8f0; border-radius: 99px; overflow: hidden; margin-bottom: 14px; }',
    '  .fc-progress-bar { height: 100%; background: #135c37; width: 0%; transition: width 0.2s ease; }',
    '  .fc-scene { perspective: 1000px; width: 100%; height: 270px; cursor: pointer; }',
    '  .fc-card { width: 100%; height: 100%; position: relative; transform-style: preserve-3d; transition: transform 0.38s cubic-bezier(0.4, 0, 0.2, 1); }',
    '  .fc-card.flipped { transform: rotateY(180deg); }',
    '  .fc-face { position: absolute; inset: 0; backface-visibility: hidden; -webkit-backface-visibility: hidden; border: 1px solid #cbd5e1; border-radius: 10px; background: #ffffff; box-shadow: 0 4px 14px rgba(15, 23, 42, 0.06); display: flex; flex-direction: column; justify-content: space-between; padding: 16px 20px; }',
    '  .fc-face.fc-back { transform: rotateY(180deg); background: #f8faf9; border-color: #86efac; }',
    '  .fc-badges { display: flex; align-items: center; justify-content: space-between; gap: 8px; }',
    '  .fc-badge-group { display: flex; align-items: center; gap: 6px; }',
    '  .fc-lvl-badge { background: #135c37; color: #fff; font-weight: 700; font-size: 11px; padding: 2px 8px; border-radius: 4px; }',
    '  .fc-type-badge { background: #f1f5f9; color: #475569; font-weight: 600; font-size: 11px; padding: 2px 8px; border-radius: 4px; border: 1px solid #e2e8f0; }',
    '  .fc-center { text-align: center; margin: auto 0; }',
    '  .fc-word { font-size: calc(var(--fs) * 2.2); font-weight: 800; color: #0f172a; letter-spacing: -0.3px; margin-bottom: 8px; }',
    '  .fc-meaning { font-size: calc(var(--fs) * 1.55); font-weight: 700; color: #135c37; margin-bottom: 10px; }',
    '  .fc-example { font-size: calc(var(--fs) * 1.05); color: #334155; background: #ffffff; border: 1px solid #e2e8f0; border-radius: 6px; padding: 8px 12px; line-height: 1.4; }',
    '  .fc-hint { text-align: center; font-size: 11px; color: #94a3b8; }',
    '  .fc-controls { display: flex; align-items: center; justify-content: center; gap: 8px; margin-top: 14px; flex-wrap: wrap; }',
    '  .fc-btn { border: 1px solid #cbd5e1; background: #fff; color: #1f2937; border-radius: 6px; padding: 0 12px; height: 32px; font-size: 12px; font-weight: 600; cursor: pointer; display: inline-flex; align-items: center; gap: 5px; }',
    '  .fc-btn:hover { background: #f1f5f9; }',
    '  .fc-btn-known { background: #135c37; color: #fff; border-color: #135c37; }',
    '  .fc-btn-known:hover { background: #0e4629; }',
    '  .fc-empty { text-align: center; padding: 36px 20px; border: 1px dashed #cbd5e1; border-radius: 8px; color: #64748b; background: #f8fafc; }',
    '  @media (max-width: 900px) { .pool-layout { grid-template-columns: 1fr; } }',
    '</style>',
    '</head>',
    '<body>',
    '  <div class="toolbar">',
    '    <div class="toolbar-left">',
    '      <div class="mini-stats">',
    '        <span class="mini-stat" id="stTotal">Total: -</span>',
    '        <span class="mini-stat unk" id="stUnk">In Pool: -</span>',
    '        <span class="mini-stat knw" id="stKnw">Known: -</span>',
    '      </div>',
    '    </div>',
    '    <div class="toolbar-right">',
    '      <button class="spk-toggle on" id="audioToggleBtn" onclick="sesAcKapat()" title="Audio ON: Click any word to hear pronunciation (Click to mute)">🔊</button>',
    '      <div class="font-ctrl" title="Decrease / Increase Font Size">',
    '        <button class="font-btn" onclick="yaziBoyutuDegistir(-1)">A-</button>',
    '        <span class="font-val" id="fontValLabel">12px</span>',
    '        <button class="font-btn" onclick="yaziBoyutuDegistir(1)">A+</button>',
    '      </div>',
    '      <input type="text" id="searchInput" class="inp" placeholder="Search (word / meaning)..." style="width:155px;" oninput="tabloyuCiz()">',
    '      <select id="levelFilter" class="inp" onchange="tabloyuCiz()">',
    '        <option value="ALL">Level: All</option>',
    '        <option value="A1">A1</option>',
    '        <option value="A2">A2</option>',
    '        <option value="B1">B1</option>',
    '        <option value="B2">B2</option>',
    '        <option value="C1">C1</option>',
    '        <option value="C2">C2</option>',
    '      </select>',
    '      <button class="btn" onclick="sekmeDegistir(\'flashcards\')" title="Study Unknown Words Pool with Flashcards">🃏 Flashcards</button>',
    '      <button class="btn btn-green" onclick="ekleBarAcKapat()">+ Add Word</button>',
    '      <button class="btn" onclick="verileriYukle()" title="Refresh">🔄</button>',
    '      <a id="openWebAppBtn" href="#" target="_blank" class="btn" style="display:none;">🌐 Open in New Tab</a>',
    '    </div>',
    '  </div>',
    '',
    '  <div class="add-bar" id="addBar">',
    '    <select id="inSeviye" class="inp">',
    '      <option value="A2">A2</option>',
    '      <option value="B1">B1</option>',
    '      <option value="B2" selected>B2</option>',
    '      <option value="C1">C1</option>',
    '      <option value="C2">C2</option>',
    '    </select>',
    '    <select id="inTur" class="inp">',
    '      <option value="Adjective">Adjective</option>',
    '      <option value="Verb">Verb</option>',
    '      <option value="Noun">Noun</option>',
    '      <option value="Adverb">Adverb</option>',
    '    </select>',
    '    <input type="text" id="inKelime" class="inp" placeholder="Word (e.g. Reluctant)" style="width:140px;">',
    '    <input type="text" id="inAnlam" class="inp" placeholder="Meaning (Turkish)" style="width:180px;">',
    '    <input type="text" id="inOrnek" class="inp" placeholder="Example Sentence" style="flex:1;min-width:190px;">',
    '    <button class="btn btn-green" id="saveWordBtn" onclick="yeniKelimeKaydet()">Save</button>',
    '    <button class="btn" onclick="ekleBarAcKapat()">Close</button>',
    '  </div>',
    '',
    '  <div class="table-wrap" id="sheet1View">',
    '    <table class="sheet-table">',
    '      <thead>',
    '        <tr>',
    '          <th class="col-row">#</th>',
    '          <th class="col-date">Date</th>',
    '          <th class="center">Level</th>',
    '          <th>Type</th>',
    '          <th>Word</th>',
    '          <th>Meaning</th>',
    '          <th>Example Sentence</th>',
    '          <th class="center">Status</th>',
    '          <th class="center">Known</th>',
    '          <th class="center">Unknown</th>',
    '        </tr>',
    '      </thead>',
    '      <tbody id="sheet1Body"></tbody>',
    '    </table>',
    '  </div>',
    '',
    '  <div class="table-wrap" id="poolView" style="display:none;">',
    '    <div class="pool-layout">',
    '      <table class="sheet-table">',
    '        <thead>',
    '          <tr>',
    '            <th>Level</th>',
    '            <th class="center">Unknown Words</th>',
    '          </tr>',
    '        </thead>',
    '        <tbody id="poolSummaryBody"></tbody>',
    '      </table>',
    '      <table class="sheet-table">',
    '        <thead>',
    '          <tr>',
    '            <th class="center">Level</th>',
    '            <th>Type</th>',
    '            <th>Word</th>',
    '            <th>Meaning</th>',
    '            <th>Example Sentence</th>',
    '            <th class="col-date">Date</th>',
    '            <th class="center">Known</th>',
    '          </tr>',
    '        </thead>',
    '        <tbody id="poolBody"></tbody>',
    '      </table>',
    '    </div>',
    '  </div>',
    '',
    '  <div id="flashcardView" style="display:none;">',
    '    <div class="fc-wrap" id="fcContainer"></div>',
    '  </div>',
    '',
    '  <div class="bottom-bar">',
    '    <button class="sheet-tab active" id="tabBtn-sheet1" onclick="sekmeDegistir(\'sheet1\')">Sheet1</button>',
    '    <button class="sheet-tab" id="tabBtn-pool" onclick="sekmeDegistir(\'pool\')">Unknown Words Pool</button>',
    '    <button class="sheet-tab" id="tabBtn-flashcards" onclick="sekmeDegistir(\'flashcards\')">🃏 Flashcards</button>',
    '  </div>',
    '',
    '  <div class="toast" id="toastBox"></div>',
    '',
    '<script>',
    '  var TUM_KELIMELER = [];',
    '  var AKTIF_SEKME = "sheet1";',
    '  var YAZI_BOYUTU = 12;',
    '  var SES_ACIK = true;',
    '',
    '  function yaziBoyutuUygula() {',
    '    document.documentElement.style.setProperty("--fs", YAZI_BOYUTU + "px");',
    '    var lbl = document.getElementById("fontValLabel");',
    '    if (lbl) lbl.innerText = YAZI_BOYUTU + "px";',
    '  }',
    '',
    '  function yaziBoyutuDegistir(delta) {',
    '    YAZI_BOYUTU = Math.max(9, Math.min(20, +(YAZI_BOYUTU + delta * 0.5).toFixed(1)));',
    '    yaziBoyutuUygula();',
    '    try { localStorage.setItem("kelime_tablo_fs", YAZI_BOYUTU); } catch (e) {}',
    '  }',
    '',
    '  function sesDurumuUygula() {',
    '    var btn = document.getElementById("audioToggleBtn");',
    '    if (!btn) return;',
    '    if (SES_ACIK) {',
    '      btn.className = "spk-toggle on";',
    '      btn.innerText = "🔊";',
    '      btn.title = "Audio ON: Click any word to hear pronunciation (Click to mute)";',
    '      document.body.classList.remove("audio-muted");',
    '    } else {',
    '      btn.className = "spk-toggle off";',
    '      btn.innerText = "🔇";',
    '      btn.title = "Audio OFF: Click to enable word pronunciation on click";',
    '      document.body.classList.add("audio-muted");',
    '      if ("speechSynthesis" in window) window.speechSynthesis.cancel();',
    '    }',
    '  }',
    '',
    '  function sesAcKapat() {',
    '    SES_ACIK = !SES_ACIK;',
    '    sesDurumuUygula();',
    '    try { localStorage.setItem("kelime_tablo_ses", SES_ACIK ? "1" : "0"); } catch (e) {}',
    '    bildirimGoster(SES_ACIK ? "🔊 Audio On" : "🔇 Audio Off");',
    '  }',
    '',
    '  try {',
    '    var kayitliFs = parseFloat(localStorage.getItem("kelime_tablo_fs"));',
    '    if (!isNaN(kayitliFs) && kayitliFs >= 9 && kayitliFs <= 20) { YAZI_BOYUTU = kayitliFs; }',
    '    var kayitliSes = localStorage.getItem("kelime_tablo_ses");',
    '    if (kayitliSes === "0") { SES_ACIK = false; }',
    '  } catch (e) {}',
    '  yaziBoyutuUygula();',
    '  sesDurumuUygula();',
    '',
    '  function bildirimGoster(m) {',
    '    var t = document.getElementById("toastBox");',
    '    t.innerText = m;',
    '    t.classList.add("show");',
    '    setTimeout(function(){ t.classList.remove("show"); }, 1400);',
    '  }',
    '',
    '  function kelimeOku(kelime, cellEl) {',
    '    if (!SES_ACIK || !kelime) return;',
    '    var temiz = kelime.toString().trim();',
    '    if (cellEl) {',
    '      cellEl.classList.add("speaking");',
    '      setTimeout(function(){ cellEl.classList.remove("speaking"); }, 550);',
    '    }',
    '    var ses = new Audio("https://ssl.gstatic.com/dictionary/static/sounds/20200429/" + encodeURIComponent(temiz.toLowerCase()) + "--_us_1.mp3");',
    '    ses.play().catch(function(){ tarayiciIleOku(temiz); });',
    '  }',
    '',
    '  function tarayiciIleOku(metin) {',
    '    if (!SES_ACIK || !("speechSynthesis" in window)) return;',
    '    window.speechSynthesis.cancel();',
    '    var u = new SpeechSynthesisUtterance(metin);',
    '    u.lang = "en-US";',
    '    u.rate = 0.95;',
    '    window.speechSynthesis.speak(u);',
    '  }',
    '',
    '  function escHtml(s) {',
    '    return (s || "").toString().replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");',
    '  }',
    '',
    '  function escJs(s) {',
    '    return (s || "").toString().replace(/\\\\/g, "\\\\\\\\").replace(/\'/g, "\\\\\'").replace(/"/g, \'\\\\"\');',
    '  }',
    '',
    '  function verileriYukle() {',
    '    document.getElementById("sheet1Body").innerHTML = \'<tr><td colspan="10" class="center" style="padding:16px;color:#64748b;">Loading...</td></tr>\';',
    '    google.script.run.withSuccessHandler(function(res) {',
    '      if (!res || !res.ok) return;',
    '      TUM_KELIMELER = res.words || [];',
    '      if (res.webAppUrl && window.self !== window.top) {',
    '        var wb = document.getElementById("openWebAppBtn");',
    '        wb.href = res.webAppUrl;',
    '        wb.style.display = "inline-flex";',
    '      }',
    '      tabloyuCiz();',
    '    }).verileriGetir();',
    '  }',
    '',
    '  var FC_INDEX = 0;',
    '  var FC_FLIPPED = false;',
    '  var FC_SHUFFLED_KEYS = null;',
    '',
    '  function sekmeDegistir(s) {',
    '    AKTIF_SEKME = s;',
    '    document.getElementById("tabBtn-sheet1").classList.toggle("active", s === "sheet1");',
    '    document.getElementById("tabBtn-pool").classList.toggle("active", s === "pool");',
    '    var fcTab = document.getElementById("tabBtn-flashcards");',
    '    if (fcTab) fcTab.classList.toggle("active", s === "flashcards");',
    '    document.getElementById("sheet1View").style.display = s === "sheet1" ? "block" : "none";',
    '    document.getElementById("poolView").style.display = s === "pool" ? "block" : "none";',
    '    document.getElementById("flashcardView").style.display = s === "flashcards" ? "block" : "none";',
    '    FC_FLIPPED = false;',
    '    tabloyuCiz();',
    '  }',
    '',
    '  function fcDestesiniGetir(poolList) {',
    '    if (!FC_SHUFFLED_KEYS || !FC_SHUFFLED_KEYS.length) return poolList;',
    '    var byKey = {};',
    '    for (var i = 0; i < poolList.length; i++) { byKey[poolList[i].kelime.toLowerCase()] = poolList[i]; }',
    '    var result = [], used = {};',
    '    for (var j = 0; j < FC_SHUFFLED_KEYS.length; j++) {',
    '      var k = FC_SHUFFLED_KEYS[j];',
    '      if (byKey[k]) { result.push(byKey[k]); used[k] = true; }',
    '    }',
    '    for (var m = 0; m < poolList.length; m++) {',
    '      var km = poolList[m].kelime.toLowerCase();',
    '      if (!used[km]) result.push(poolList[m]);',
    '    }',
    '    return result;',
    '  }',
    '',
    '  function fcCiz(poolList) {',
    '    var box = document.getElementById("fcContainer");',
    '    if (!box) return;',
    '    var deck = fcDestesiniGetir(poolList);',
    '    if (!deck.length) {',
    '      box.innerHTML = \'<div class="fc-empty"><div style="font-size:28px;margin-bottom:8px;">🎉</div><div style="font-weight:700;color:#1e293b;margin-bottom:6px;">No Unknown Words in Pool</div><div style="font-size:12px;margin-bottom:14px;">Mark words as <b>Unknown</b> in Sheet1 to study them with flashcards.</div><button class="fc-btn" onclick="sekmeDegistir(\\\'sheet1\\\')">Go to Sheet1</button></div>\';',
    '      return;',
    '    }',
    '    if (FC_INDEX >= deck.length) FC_INDEX = 0;',
    '    if (FC_INDEX < 0) FC_INDEX = deck.length - 1;',
    '    var card = deck[FC_INDEX];',
    '    var pct = Math.round(((FC_INDEX + 1) / deck.length) * 100);',
    '    box.innerHTML =',
    '      \'<div class="fc-top"><span>Unknown Pool Flashcards — <b>\' + (FC_INDEX + 1) + \' / \' + deck.length + \'</b></span><button class="btn" onclick="fcKaristir()" title="Shuffle cards">🔀 Shuffle</button></div>\' +',
    '      \'<div class="fc-progress-track"><div class="fc-progress-bar" style="width:\' + pct + \'%;"></div></div>\' +',
    '      \'<div class="fc-scene" onclick="fcCevir()">\' +',
    '        \'<div class="fc-card\' + (FC_FLIPPED ? \' flipped\' : \'\') + \'" id="fcCardInner">\' +',
    '          \'<div class="fc-face fc-front">\' +',
    '            \'<div class="fc-badges"><div class="fc-badge-group"><span class="fc-lvl-badge">\' + escHtml(card.seviye) + \'</span><span class="fc-type-badge">\' + escHtml(turTemizle(card.tur)) + \'</span></div><button class="btn" onclick="event.stopPropagation(); kelimeOku(\\\'\' + escJs(card.kelime) + \'\\\')" title="Listen pronunciation">🔊 Listen</button></div>\' +',
    '            \'<div class="fc-center"><div class="fc-word">\' + escHtml(card.kelime) + \'</div></div>\' +',
    '            \'<div class="fc-hint">Click card or press Space to flip ↻</div>\' +',
    '          \'</div>\' +',
    '          \'<div class="fc-face fc-back">\' +',
    '            \'<div class="fc-badges"><div class="fc-badge-group"><span class="fc-lvl-badge">\' + escHtml(card.seviye) + \'</span><span class="fc-type-badge">\' + escHtml(turTemizle(card.tur)) + \'</span><strong style="font-size:13px;color:#0f172a;margin-left:4px;">\' + escHtml(card.kelime) + \'</strong></div><button class="btn" onclick="event.stopPropagation(); kelimeOku(\\\'\' + escJs(card.kelime) + \'\\\')" title="Listen pronunciation">🔊</button></div>\' +',
    '            \'<div class="fc-center"><div class="fc-meaning">\' + escHtml(card.anlam) + \'</div>\' + (card.ornek ? \'<div class="fc-example">\' + escHtml(card.ornek) + \'</div>\' : \'\') + \'</div>\' +',
    '            \'<div class="fc-hint">Click card to flip back ↻</div>\' +',
    '          \'</div>\' +',
    '        \'</div>\' +',
    '      \'</div>\' +',
    '      \'<div class="fc-controls">\' +',
    '        \'<button class="fc-btn" onclick="fcDegistir(-1)">◀ Prev</button>\' +',
    '        \'<button class="fc-btn" onclick="fcCevir()">↻ Flip</button>\' +',
    '        \'<button class="fc-btn fc-btn-known" onclick="fcBiliyorumIsaretle(\' + card.row + \')">✅ Known (Remove from Pool)</button>\' +',
    '        \'<button class="fc-btn" onclick="fcDegistir(1)">Next ▶</button>\' +',
    '      \'</div>\';',
    '  }',
    '',
    '  function fcCevir() {',
    '    FC_FLIPPED = !FC_FLIPPED;',
    '    var el = document.getElementById("fcCardInner");',
    '    if (el) el.classList.toggle("flipped", FC_FLIPPED);',
    '  }',
    '',
    '  function fcDegistir(delta) {',
    '    FC_INDEX += delta;',
    '    FC_FLIPPED = false;',
    '    tabloyuCiz();',
    '    var deck = document.querySelector("#fcContainer .fc-word");',
    '    if (deck && SES_ACIK) { kelimeOku(deck.textContent); }',
    '  }',
    '',
    '  function fcKaristir() {',
    '    var unk = TUM_KELIMELER.filter(function(w){ return w.durum === "bilmiyorum"; }).map(function(w){ return w.kelime.toLowerCase(); });',
    '    for (var i = unk.length - 1; i > 0; i--) {',
    '      var r = Math.floor(Math.random() * (i + 1));',
    '      var tmp = unk[i]; unk[i] = unk[r]; unk[r] = tmp;',
    '    }',
    '    FC_SHUFFLED_KEYS = unk;',
    '    FC_INDEX = 0;',
    '    FC_FLIPPED = false;',
    '    tabloyuCiz();',
    '    bildirimGoster("🔀 Shuffled");',
    '  }',
    '',
    '  function fcBiliyorumIsaretle(row) {',
    '    FC_FLIPPED = false;',
    '    kutucukDegisti(row, "biliyorum", true);',
    '  }',
    '',
    '  document.addEventListener("keydown", function(e) {',
    '    if (AKTIF_SEKME !== "flashcards") return;',
    '    var tag = (e.target && e.target.tagName) ? e.target.tagName.toUpperCase() : "";',
    '    if (tag === "INPUT" || tag === "SELECT" || tag === "TEXTAREA") return;',
    '    if (e.code === "Space") { e.preventDefault(); fcCevir(); }',
    '    else if (e.code === "ArrowRight") { e.preventDefault(); fcDegistir(1); }',
    '    else if (e.code === "ArrowLeft") { e.preventDefault(); fcDegistir(-1); }',
    '  });',
    '',
    '  function durumMetniGetir(durum) {',
    '    if (durum === "bilmiyorum") return "In Pool (Unknown)";',
    '    if (durum === "biliyorum") return "Learned (Known)";',
    '    return "Pending";',
    '  }',
    '',
    '  function turTemizle(tur) {',
    '    var s = (tur || "").toString().trim();',
    '    var m = s.match(/\\(([^)]+)\\)/);',
    '    return m ? m[1] : s;',
    '  }',
    '',
    '  function tabloyuCiz() {',
    '    var total = TUM_KELIMELER.length, unk = 0, knw = 0;',
    '    var lvlCounts = { B1: 0, B2: 0, C1: 0, C2: 0 };',
    '    for (var i = 0; i < TUM_KELIMELER.length; i++) {',
    '      var w = TUM_KELIMELER[i];',
    '      if (w.durum === "bilmiyorum") {',
    '        unk++;',
    '        lvlCounts[w.seviye] = (lvlCounts[w.seviye] || 0) + 1;',
    '      } else if (w.durum === "biliyorum") knw++;',
    '    }',
    '    document.getElementById("stTotal").innerText = "Total: " + total;',
    '    document.getElementById("stUnk").innerText = "In Pool: " + unk;',
    '    document.getElementById("stKnw").innerText = "Known: " + knw;',
    '    var fcTabBtn = document.getElementById("tabBtn-flashcards");',
    '    if (fcTabBtn) fcTabBtn.innerText = "🃏 Flashcards (" + unk + ")";',
    '',
    '    var q = (document.getElementById("searchInput").value || "").toLowerCase().trim();',
    '    var lvl = document.getElementById("levelFilter").value;',
    '',
    '    var s1Rows = [];',
    '    for (var j = 0; j < TUM_KELIMELER.length; j++) {',
    '      var x = TUM_KELIMELER[j];',
    '      if (lvl !== "ALL" && x.seviye !== lvl) continue;',
    '      if (q && (x.kelime + " " + x.anlam + " " + x.ornek).toLowerCase().indexOf(q) === -1) continue;',
    '      s1Rows.push(',
    '        \'<tr class="\' + (x.durum === "biliyorum" ? "row-known" : "") + \'">\' +',
    '          \'<td class="col-row">\' + x.row + \'</td>\' +',
    '          \'<td class="col-date">\' + escHtml(x.tarih) + \'</td>\' +',
    '          \'<td class="col-lvl">\' + escHtml(x.seviye) + \'</td>\' +',
    '          \'<td class="col-type">\' + escHtml(turTemizle(x.tur)) + \'</td>\' +',
    '          \'<td class="col-word" onclick="kelimeOku(\\\'\' + escJs(x.kelime) + \'\\\', this)">\' + escHtml(x.kelime) + \'</td>\' +',
    '          \'<td class="col-meaning">\' + escHtml(x.anlam) + \'</td>\' +',
    '          \'<td class="col-example">\' + escHtml(x.ornek) + \'</td>\' +',
    '          \'<td class="col-status"><span class="status-pill \' + x.durum + \'">\' + durumMetniGetir(x.durum) + \'</span></td>\' +',
    '          \'<td class="col-cb"><input type="checkbox" class="sheet-cb" \' + (x.biliyorum ? "checked" : "") + \' onchange="kutucukDegisti(\' + x.row + \', \\\'biliyorum\\\', this.checked)"></td>\' +',
    '          \'<td class="col-cb"><input type="checkbox" class="sheet-cb" \' + (x.bilmiyorum ? "checked" : "") + \' onchange="kutucukDegisti(\' + x.row + \', \\\'bilmiyorum\\\', this.checked)"></td>\' +',
    '        \'</tr>\'',
    '      );',
    '    }',
    '    document.getElementById("sheet1Body").innerHTML = s1Rows.join("") || \'<tr><td colspan="10" class="center" style="padding:14px;">No words found.</td></tr>\';',
    '',
    '    var sumLvls = ["A1", "A2", "B1", "B2", "C1", "C2"];',
    '    var sumHtml = [];',
    '    for (var k = 0; k < sumLvls.length; k++) {',
    '      var sl = sumLvls[k];',
    '      if (lvlCounts[sl] !== undefined || sl === "B1" || sl === "B2" || sl === "C1" || sl === "C2") {',
    '        sumHtml.push(\'<tr><td>\' + sl + \'</td><td class="center">\' + (lvlCounts[sl] || 0) + \'</td></tr>\');',
    '      }',
    '    }',
    '    document.getElementById("poolSummaryBody").innerHTML = sumHtml.join("");',
    '',
    '    var poolList = TUM_KELIMELER.filter(function(w){',
    '      if (w.durum !== "bilmiyorum") return false;',
    '      if (lvl !== "ALL" && w.seviye !== lvl) return false;',
    '      if (q && (w.kelime + " " + w.anlam + " " + w.ornek).toLowerCase().indexOf(q) === -1) return false;',
    '      return true;',
    '    }).sort(function(a, b){ return (a.seviye || "").localeCompare(b.seviye || ""); });',
    '',
    '    var pRows = [];',
    '    for (var m = 0; m < poolList.length; m++) {',
    '      var p = poolList[m];',
    '      pRows.push(',
    '        \'<tr>\' +',
    '          \'<td class="col-lvl">\' + escHtml(p.seviye) + \'</td>\' +',
    '          \'<td class="col-type">\' + escHtml(turTemizle(p.tur)) + \'</td>\' +',
    '          \'<td class="col-word" onclick="kelimeOku(\\\'\' + escJs(p.kelime) + \'\\\', this)">\' + escHtml(p.kelime) + \'</td>\' +',
    '          \'<td class="col-meaning">\' + escHtml(p.anlam) + \'</td>\' +',
    '          \'<td class="col-example">\' + escHtml(p.ornek) + \'</td>\' +',
    '          \'<td class="col-date">\' + escHtml(p.tarih) + \'</td>\' +',
    '          \'<td class="col-cb"><input type="checkbox" class="sheet-cb" onchange="kutucukDegisti(\' + p.row + \', \\\'biliyorum\\\', this.checked)" title="Mark as Known"></td>\' +',
    '        \'</tr>\'',
    '      );',
    '    }',
    '    document.getElementById("poolBody").innerHTML = pRows.join("") || \'<tr><td colspan="7" class="center" style="padding:14px;">No words in pool.</td></tr>\';',
    '',
    '    fcCiz(poolList);',
    '  }',
    '',
    '  function kutucukDegisti(row, tip, isChecked) {',
    '    var hedef = null;',
    '    for (var i = 0; i < TUM_KELIMELER.length; i++) {',
    '      if (TUM_KELIMELER[i].row === row) { hedef = TUM_KELIMELER[i]; break; }',
    '    }',
    '    if (!hedef) return;',
    '    var yeniDurum = isChecked ? tip : "beklemede";',
    '    hedef.durum = yeniDurum;',
    '    hedef.biliyorum = yeniDurum === "biliyorum";',
    '    hedef.bilmiyorum = yeniDurum === "bilmiyorum";',
    '    tabloyuCiz();',
    '    bildirimGoster("Saved: " + hedef.kelime);',
    '    google.script.run.durumGuncelle(row, yeniDurum);',
    '  }',
    '',
    '  function ekleBarAcKapat() {',
    '    document.getElementById("addBar").classList.toggle("open");',
    '  }',
    '',
    '  function yeniKelimeKaydet() {',
    '    var kelime = document.getElementById("inKelime").value.trim();',
    '    var anlam = document.getElementById("inAnlam").value.trim();',
    '    if (!kelime || !anlam) { bildirimGoster("Word and Meaning are required."); return; }',
    '    var btn = document.getElementById("saveWordBtn");',
    '    btn.disabled = true;',
    '    google.script.run.withSuccessHandler(function(){',
    '      btn.disabled = false;',
    '      document.getElementById("inKelime").value = "";',
    '      document.getElementById("inAnlam").value = "";',
    '      document.getElementById("inOrnek").value = "";',
    '      verileriYukle();',
    '    }).kelimeEkle({',
    '      kelime: kelime,',
    '      okunus: "",',
    '      tur: document.getElementById("inTur").value,',
    '      seviye: document.getElementById("inSeviye").value,',
    '      anlam: anlam,',
    '      ornek: document.getElementById("inOrnek").value.trim(),',
    '      durum: "bilmiyorum"',
    '    });',
    '  }',
    '',
    '  verileriYukle();',
    '<\/script>',
    '</body>',
    '</html>'
  ].join('\n');
}
