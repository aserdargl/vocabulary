var SPREADSHEET_ID = '1cUYfuFg7OSsWjCtLA7jiB77bGJITR6V0jF2ZUnyrXtc';

function getSpreadsheet_() {
  try {
    var active = SpreadsheetApp.getActiveSpreadsheet();
    if (active) return active;
  } catch (e) {}
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function getMainSheet_(ss) {
  var sh = ss.getSheetByName('Words');
  if (sh) return sh;
  sh = ss.getSheetByName('Sheet1');
  if (sh) {
    try { sh.setName('Words'); } catch (e) {}
    return sh;
  }
  return ss.getSheets()[0] || null;
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
    var shName = sheet.getName();
    if (shName !== 'Words' && shName !== 'Sheet1') return;
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
  if (e && e.parameter && e.parameter.action === 'recordAttempt' && e.parameter.word) {
    var isCorr = e.parameter.isCorrect === 'true' || e.parameter.isCorrect === '1';
    var entryParam = null;
    if (e.parameter.mistakes !== undefined || e.parameter.correct !== undefined) {
      entryParam = {
        mistakes: parseInt(e.parameter.mistakes, 10) || 0,
        correct: parseInt(e.parameter.correct, 10) || 0
      };
    }
    var resStat = istatistikKaydet(e.parameter.word, isCorr, e.parameter.mode || 'Study', entryParam);
    return ContentService
      .createTextOutput(JSON.stringify(resStat))
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
    if (body.action === 'recordAttempt' && body.word) {
      return ContentService
        .createTextOutput(JSON.stringify(istatistikKaydet(body.word, !!body.isCorrect, body.mode || 'Study', body.entry)))
        .setMimeType(ContentService.MimeType.JSON);
    }
    if (body.action === 'syncStats' && body.stats) {
      return ContentService
        .createTextOutput(JSON.stringify(istatistikTopluSenkronizeEt(body.stats)))
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
  var sheet = getMainSheet_(ss);
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
  var sheet = getMainSheet_(ss);
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

  var statsMap = {};
  try {
    statsMap = istatistikleriOku_(ss, words);
  } catch (e) {}

  try {
    knownHavuzSayfasiniHazirla_(ss);
  } catch (e) {}

  return {
    ok: true,
    words: words,
    stats: statsMap,
    sheetUrl: ss.getUrl(),
    webAppUrl: getWebAppUrl_()
  };
}

var STATS_SHEET_NAME = 'Mistake Stats';

function istatistikSayfasiniHazirla_(ss) {
  var sh = ss.getSheetByName(STATS_SHEET_NAME);
  var yeniMi = false;
  if (!sh) {
    sh = ss.insertSheet(STATS_SHEET_NAME);
    yeniMi = true;
  }
  var col5Header = sh.getRange(1, 5).getDisplayValue();
  if (yeniMi || col5Header !== 'Mistakes') {
    var headers = [[
      '#',
      'Level',
      'Word',
      'Meaning',
      'Mistakes',
      'Correct',
      'Total',
      'Error Rate',
      'Mistake Bar',
      'Last Practice'
    ]];
    sh.getRange(1, 1, 1, 10).setValues(headers)
      .setBackground('#135c37')
      .setFontColor('#ffffff')
      .setFontWeight('bold');
    sh.setFrozenRows(1);
    try {
      sh.setColumnWidth(1, 32);
      sh.setColumnWidth(2, 48);
      sh.setColumnWidth(3, 110);
      sh.setColumnWidth(4, 165);
      sh.setColumnWidth(5, 68);
      sh.setColumnWidth(6, 65);
      sh.setColumnWidth(7, 56);
      sh.setColumnWidth(8, 76);
      sh.setColumnWidth(9, 135);
      sh.setColumnWidth(10, 130);
    } catch (e) {}
  }
  return sh;
}

function istatistikleriOku_(ss, wordsList) {
  var sh = istatistikSayfasiniHazirla_(ss);
  var lastR = sh.getLastRow();
  var map = {};
  if (lastR >= 2) {
    var vals = sh.getRange(2, 1, lastR - 1, 10).getValues();
    for (var i = 0; i < vals.length; i++) {
      var r = vals[i];
      var w = (r[2] || '').toString().trim();
      if (!w) continue;
      var k = w.toLowerCase();
      map[k] = {
        word: w,
        level: (r[1] || 'B2').toString().trim(),
        meaning: (r[3] || '').toString().trim(),
        mistakes: parseInt(r[4], 10) || 0,
        correct: parseInt(r[5], 10) || 0,
        lastPractice: (r[9] || '').toString().trim()
      };
    }
  }
  if (lastR < 2 && wordsList && wordsList.length > 0) {
    istatistikTablosunuYaz_(ss, map, wordsList);
  }
  return map;
}

function istatistikTablosunuYaz_(ss, statsMap, wordsList) {
  var sh = istatistikSayfasiniHazirla_(ss);
  var kelimeSozluk = {};
  if (wordsList && wordsList.length) {
    for (var i = 0; i < wordsList.length; i++) {
      var wObj = wordsList[i];
      var wk = (wObj.kelime || '').toString().trim().toLowerCase();
      if (wk) kelimeSozluk[wk] = wObj;
    }
  } else {
    var s1 = getMainSheet_(ss);
    if (s1) {
      var layout = getSheetLayout_(s1);
      var lr = s1.getLastRow();
      if (lr >= layout.startRow) {
        var dVals = s1.getRange(layout.startRow, 1, lr - layout.startRow + 1, Math.max(s1.getLastColumn(), 12)).getDisplayValues();
        for (var j = 0; j < dVals.length; j++) {
          var dr = dVals[j];
          var kw = (dr[layout.word - 1] || '').toString().trim();
          if (kw) {
            kelimeSozluk[kw.toLowerCase()] = {
              kelime: kw,
              seviye: (dr[layout.level - 1] || 'B2').toString().trim(),
              anlam: (dr[layout.meaning - 1] || '').toString().trim()
            };
          }
        }
      }
    }
  }

  var birlesikAnahtarlar = {};
  for (var k1 in kelimeSozluk) birlesikAnahtarlar[k1] = true;
  for (var k2 in statsMap) birlesikAnahtarlar[k2] = true;

  var satirlar = [];
  for (var key in birlesikAnahtarlar) {
    var st = statsMap[key] || { mistakes: 0, correct: 0, lastPractice: '' };
    var meta = kelimeSozluk[key] || {};
    var mCount = parseInt(st.mistakes, 10) || 0;
    var cCount = parseInt(st.correct, 10) || 0;
    var tot = mCount + cCount;
    var errPct = tot > 0 ? Math.round((mCount / tot) * 100) : 0;
    satirlar.push({
      word: meta.kelime || st.word || key,
      level: meta.seviye || st.level || 'B2',
      meaning: meta.anlam || st.meaning || '',
      mistakes: mCount,
      correct: cCount,
      total: tot,
      errPct: errPct,
      lastPractice: st.lastPractice || '-'
    });
  }

  satirlar.sort(function(a, b) {
    if (b.mistakes !== a.mistakes) return b.mistakes - a.mistakes;
    if (b.errPct !== a.errPct) return b.errPct - a.errPct;
    if (b.total !== a.total) return b.total - a.total;
    return a.word.localeCompare(b.word);
  });

  var mevcutSonSatir = sh.getLastRow();
  if (mevcutSonSatir >= 2) {
    sh.getRange(2, 1, mevcutSonSatir - 1, 10).clearContent();
  }

  if (satirlar.length === 0) return;

  var outVals = [];
  var barFormulas = [];
  var maxRowNum = satirlar.length + 1;

  for (var idx = 0; idx < satirlar.length; idx++) {
    var it = satirlar[idx];
    var rNum = idx + 2;
    outVals.push([
      idx + 1,
      it.level,
      it.word,
      it.meaning,
      it.mistakes,
      it.correct,
      it.total,
      it.errPct + '%',
      '',
      it.lastPractice
    ]);
    barFormulas.push([
      '=IF(E' + rNum + '>0, SPARKLINE(E' + rNum + ', {"charttype","bar";"max",MAX(1,MAX($E$2:$E$' + maxRowNum + '));"color1","#dc2626"}), "")'
    ]);
  }

  sh.getRange(2, 1, outVals.length, 10).setValues(outVals);
  sh.getRange(2, 9, barFormulas.length, 1).setFormulas(barFormulas);
  try { SpreadsheetApp.flush(); } catch (e) {}
}

function istatistikKaydet(kelime, dogruMu, mod, clientEntry) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(5000); } catch (e) {}
  try {
    var ss = getSpreadsheet_();
    var temiz = (kelime || '').toString().trim();
    if (!temiz) return { ok: false, error: 'Empty word' };
    var key = temiz.toLowerCase();
    var map = istatistikleriOku_(ss);
    var mevcut = map[key] || { word: temiz, level: 'B2', meaning: '', mistakes: 0, correct: 0, lastPractice: '' };

    if (dogruMu) {
      mevcut.correct = (mevcut.correct || 0) + 1;
    } else {
      mevcut.mistakes = (mevcut.mistakes || 0) + 1;
    }
    if (clientEntry) {
      if ((clientEntry.mistakes || 0) > mevcut.mistakes) mevcut.mistakes = parseInt(clientEntry.mistakes, 10) || mevcut.mistakes;
      if ((clientEntry.correct || 0) > mevcut.correct) mevcut.correct = parseInt(clientEntry.correct, 10) || mevcut.correct;
      if (clientEntry.level || clientEntry.seviye) mevcut.level = clientEntry.level || clientEntry.seviye;
      if (clientEntry.meaning || clientEntry.anlam) mevcut.meaning = clientEntry.meaning || clientEntry.anlam;
    }

    var tz = 'GMT+3';
    try { tz = Session.getScriptTimeZone() || 'GMT+3'; } catch (e) {}
    var tarihStr = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm');
    mevcut.lastPractice = (mod || 'Study').toUpperCase() + ' • ' + (dogruMu ? '✅' : '❌') + ' (' + tarihStr + ')';
    map[key] = mevcut;

    istatistikTablosunuYaz_(ss, map);
    return { ok: true, word: temiz, entry: mevcut, stats: map };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function istatistikTopluSenkronizeEt(incomingStats) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(5000); } catch (e) {}
  try {
    var ss = getSpreadsheet_();
    var map = istatistikleriOku_(ss);
    if (incomingStats && typeof incomingStats === 'object') {
      for (var k in incomingStats) {
        var inc = incomingStats[k];
        if (!inc) continue;
        var key = k.toLowerCase().trim();
        var cur = map[key] || {
          word: inc.word || inc.kelime || k,
          level: inc.level || inc.seviye || 'B2',
          meaning: inc.meaning || inc.anlam || '',
          mistakes: 0,
          correct: 0,
          lastPractice: ''
        };
        cur.mistakes = Math.max(cur.mistakes || 0, parseInt(inc.mistakes, 10) || 0);
        cur.correct = Math.max(cur.correct || 0, parseInt(inc.correct, 10) || 0);
        if (inc.lastPractice && inc.lastPractice !== '-' && (!cur.lastPractice || cur.lastPractice === '-')) {
          cur.lastPractice = inc.lastPractice;
        }
        map[key] = cur;
      }
    }
    istatistikTablosunuYaz_(ss, map);
    return { ok: true, stats: map };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
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
  var sheet = getMainSheet_(ss);
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
  var sheet = getMainSheet_(ss);
  if (!sheet) return { ok: false, error: 'Words sheet not found' };

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

function knownHavuzSayfasiniHazirla_(ss) {
  if (ss.getSheetByName('Known Words Pool') || ss.getSheetByName('Bildiğim Kelimeler Havuzu')) return;
  var unkSheet = ss.getSheetByName('Unknown Words Pool') || ss.getSheetByName('Bilmediğim Kelimeler Havuzu');
  if (!unkSheet) return;
  var knwSheet = unkSheet.copyTo(ss);
  knwSheet.setName('Known Words Pool');
  try { knwSheet.setTabColor('#059669'); } catch (e) {}
  var maxR = Math.min(knwSheet.getMaxRows(), 60);
  var maxC = Math.min(knwSheet.getMaxColumns(), 15);
  var range = knwSheet.getRange(1, 1, maxR, maxC);
  var formulas = range.getFormulas();
  var vals = range.getValues();
  for (var r = 0; r < maxR; r++) {
    for (var c = 0; c < maxC; c++) {
      var f = formulas[r][c];
      if (f) {
        var newF = f
          .replace(/Sheet1!\$?L/g, function(m) { return m.indexOf('$') !== -1 ? 'Sheet1!$K' : 'Sheet1!K'; })
          .replace(/Unknown/gi, 'Known');
        if (newF !== f) knwSheet.getRange(r + 1, c + 1).setFormula(newF);
      } else if (typeof vals[r][c] === 'string' && vals[r][c]) {
        var s = vals[r][c];
        if (/unknown/i.test(s) || /bilmediğim/i.test(s)) {
          var newS = s.replace(/Unknown/g, 'Known').replace(/unknown/g, 'known').replace(/Bilmediğim/g, 'Bildiğim');
          knwSheet.getRange(r + 1, c + 1).setValue(newS);
        }
      }
    }
  }
}

function seciliKelimeyiAl() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var cell = sheet.getActiveCell();
  if (!cell || cell.getRow() < 6) return '';

  var sheetName = sheet.getName();
  var isHavuz = (
    sheetName === 'Unknown Words Pool' ||
    sheetName === 'Bilmediğim Kelimeler Havuzu' ||
    sheetName === 'Known Words Pool' ||
    sheetName === 'Bildiğim Kelimeler Havuzu'
  );
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
    '  /* Compact Single-Row Toolbar */',
    '  .toolbar { display: flex; align-items: center; justify-content: space-between; padding: 5px 8px; background: #f8fafc; border-bottom: 1px solid #cbd5e1; gap: 6px; flex-wrap: nowrap; white-space: nowrap; overflow-x: auto; position: sticky; top: 0; z-index: 20; }',
    '  .toolbar-left { display: flex; align-items: center; gap: 6px; flex-wrap: nowrap; white-space: nowrap; flex-shrink: 0; }',
    '  .mini-stats { display: flex; align-items: center; gap: 4px; font-size: 11px; color: #475569; flex-wrap: nowrap; white-space: nowrap; }',
    '  .mini-stat { padding: 2px 6px; border-radius: 4px; background: #e2e8f0; font-weight: 600; white-space: nowrap; }',
    '  .mini-stat.unk { background: #fee2e2; color: #991b1b; }',
    '  .mini-stat.knw { background: #dcfce7; color: #166534; }',
    '  .toolbar-right { display: flex; align-items: center; gap: 5px; flex-wrap: nowrap; white-space: nowrap; flex-shrink: 0; }',
    '  .inp { border: 1px solid #cbd5e1; border-radius: 4px; padding: 3px 7px; font-size: 11.5px; height: 26px; outline: none; background: #fff; }',
    '  .inp:focus { border-color: #137333; }',
    '  .btn { border: 1px solid #cbd5e1; background: #fff; color: #1f2937; border-radius: 4px; padding: 0 8px; height: 26px; font-size: 11.5px; font-weight: 600; cursor: pointer; display: inline-flex; align-items: center; gap: 4px; text-decoration: none; user-select: none; white-space: nowrap; flex-shrink: 0; }',
    '  .btn:hover { background: #f1f5f9; }',
    '  .btn-green { background: #135c37; color: #fff; border-color: #135c37; }',
    '  .btn-green:hover { background: #0e4629; }',
    '  .toggle-known-btn.show-mode { background: #f0fdf4; border-color: #86efac; color: #166534; }',
    '  .toggle-known-btn.show-mode:hover { background: #dcfce7; }',
    '  .toggle-known-btn.hide-mode { background: #fffbeb; border-color: #fcd34d; color: #92400e; }',
    '  .toggle-known-btn.hide-mode:hover { background: #fef3c7; }',
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
    '  .col-word { font-weight: 700; color: #111827; white-space: nowrap; cursor: pointer; transition: color 0.1s, background-color 0.15s; position: relative; }',
    '  .col-word:hover { color: #135c37; text-decoration: underline; }',
    '  .col-word.speaking { background-color: #dcfce7 !important; color: #135c37; }',
    '  body.audio-muted .col-word { cursor: default; }',
    '  body.audio-muted .col-word:hover { color: #111827; text-decoration: none; }',
    '  .word-hover-ex { display: none; position: absolute; left: 6px; top: calc(100% + 22px); z-index: 50; background: #fffef7; color: #0f172a; border: 2px solid #135c37; border-left: 5px solid #135c37; font-weight: 600; font-size: calc(var(--fs) + 3.5px); padding: 10px 14px; border-radius: 8px; box-shadow: 0 10px 26px rgba(15, 23, 42, 0.24); white-space: normal; min-width: 280px; max-width: 440px; line-height: 1.45; pointer-events: none; text-decoration: none !important; }',
    '  .col-word:hover .word-hover-ex { display: block; }',
    '  .ex-tr { display: block; margin-top: 4px; color: #135c37; font-weight: 600; font-size: 0.92em; }',
    '  .col-phonetic { color: #4b5563; font-size: calc(var(--fs) - 0.5px); white-space: nowrap; }',
    '  .col-meaning { min-width: 125px; }',
    '  .col-example { min-width: 190px; }',
    '  .col-status { width: 122px; text-align: center; white-space: nowrap; }',
    '  .status-pill { display: inline-block; padding: 2px 8px; border-radius: 99px; font-size: calc(var(--fs) - 1.5px); font-weight: 600; background: #e8eaed; color: #3c4043; min-width: 112px; text-align: center; }',
    '  .status-pill.bilmiyorum { background: #fce8e6; color: #c5221f; }',
    '  .status-pill.biliyorum { background: #e6f4ea; color: #137333; }',
    '  th.col-mark, td.col-mark { width: 40px; min-width: 40px; max-width: 44px; text-align: center; padding: 2px 3px !important; user-select: none; }',
    '  .mark-split { position: relative; width: 34px; height: 24px; border-radius: 5px; border: 1.5px solid #cbd5e1; background: #ffffff; display: inline-block; vertical-align: middle; overflow: hidden; user-select: none; }',
    '  .mark-split::after { content: ""; position: absolute; inset: 0; background: linear-gradient(to bottom right, transparent calc(50% - 0.75px), #94a3b8 calc(50% - 0.75px), #94a3b8 calc(50% + 0.75px), transparent calc(50% + 0.75px)); pointer-events: none; z-index: 2; }',
    '  .mark-half { position: absolute; inset: 0; cursor: pointer; font-weight: 900; line-height: 1; transition: background-color 0.12s ease, color 0.12s ease; color: #94a3b8; }',
    '  .mark-top { clip-path: polygon(0 0, 100% 0, 0 100%); font-size: 10.5px; padding-top: 2px; padding-left: 3px; text-align: left; }',
    '  .mark-bot { clip-path: polygon(100% 0, 100% 100%, 0 100%); font-size: 10px; display: flex; align-items: flex-end; justify-content: flex-end; padding-bottom: 2px; padding-right: 3px; }',
    '  .mark-top:hover { background: #dcfce7; color: #15803d; }',
    '  .mark-bot:hover { background: #fee2e2; color: #b91c1c; }',
    '  .mark-split.biliyorum { border-color: #16a34a; }',
    '  .mark-split.biliyorum .mark-top { background: #bbf7d0; color: #14532d; }',
    '  .mark-split.biliyorum .mark-bot { background: #f8fafc; color: #cbd5e1; }',
    '  .mark-split.biliyorum .mark-bot:hover { background: #fee2e2; color: #b91c1c; }',
    '  .mark-split.bilmiyorum { border-color: #dc2626; }',
    '  .mark-split.bilmiyorum .mark-bot { background: #fecaca; color: #991b1b; }',
    '  .mark-split.bilmiyorum .mark-top { background: #f8fafc; color: #cbd5e1; }',
    '  .mark-split.bilmiyorum .mark-top:hover { background: #dcfce7; color: #15803d; }',
    '  /* Pool Layout */',
    '  .pool-layout { display: grid; grid-template-columns: 185px 1fr; gap: 14px; align-items: start; }',
    '  /* Bottom Sheet Tabs */',
    '  .bottom-bar { position: fixed; bottom: 0; left: 0; right: 0; height: 34px; background: #f8fafc; border-top: 1px solid #cbd5e1; display: flex; align-items: center; padding: 0 12px; gap: 5px; z-index: 30; flex-wrap: nowrap; white-space: nowrap; }',
    '  .sheet-tab { border: 1px solid #cbd5e1; background: #ffffff; padding: 0 11px; height: 25px; border-radius: 6px; font-size: 11.5px; font-weight: 700; color: #334155; cursor: pointer; display: inline-flex; align-items: center; gap: 5px; white-space: nowrap; flex-shrink: 0; transition: all 0.15s ease; }',
    '  .sheet-tab:hover { filter: brightness(0.96); transform: translateY(-1px); }',
    '  .sheet-tab.active { box-shadow: 0 2px 6px rgba(15, 23, 42, 0.18); transform: translateY(-1px); }',
    '  #tabBtn-sheet1 { background: #dcfce7; color: #14532d; border-color: #86efac; }',
    '  #tabBtn-sheet1.active { background: #15803d; color: #ffffff; border-color: #14532d; }',
    '  #tabBtn-pool { background: #ffe4e6; color: #881337; border-color: #fda4af; }',
    '  #tabBtn-pool.active { background: #e11d48; color: #ffffff; border-color: #be123c; }',
    '  #tabBtn-knownPool { background: #d1fae5; color: #065f46; border-color: #6ee7b7; }',
    '  #tabBtn-knownPool.active { background: #059669; color: #ffffff; border-color: #047857; }',
    '  #tabBtn-stats { background: #fef3c7; color: #78350f; border-color: #fcd34d; }',
    '  #tabBtn-stats.active { background: #d97706; color: #ffffff; border-color: #b45309; }',
    '  #tabBtn-flashcards { background: #e0e7ff; color: #1e3a8a; border-color: #a5b4fc; }',
    '  #tabBtn-flashcards.active { background: #4f46e5; color: #ffffff; border-color: #4338ca; }',
    '  #tabBtn-cloze { background: #f3e8ff; color: #581c87; border-color: #d8b4fe; }',
    '  #tabBtn-cloze.active { background: #9333ea; color: #ffffff; border-color: #7e22ce; }',
    '  #tabBtn-quiz { background: #e0f2fe; color: #0c4a6e; border-color: #7dd3fc; }',
    '  #tabBtn-quiz.active { background: #0284c7; color: #ffffff; border-color: #0369a1; }',
    '  #tabBtn-listen { background: #fae8ff; color: #701a75; border-color: #f0abfc; }',
    '  #tabBtn-listen.active { background: #c026d3; color: #ffffff; border-color: #a21caf; }',
    '  #tabBtn-match { background: #ccfbf1; color: #134e4a; border-color: #5eead4; }',
    '  #tabBtn-match.active { background: #0d9488; color: #ffffff; border-color: #0f766e; }',
    '  /* Compact Add Row Bar */',
    '  .add-bar { display: none; background: #f8fafc; border: 1px solid #cbd5e1; border-radius: 4px; padding: 6px 10px; margin: 6px 8px 0; align-items: center; gap: 6px; flex-wrap: nowrap; white-space: nowrap; overflow-x: auto; }',
    '  .add-bar.open { display: flex; }',
    '  .toast { position: fixed; bottom: 36px; right: 14px; background: #202124; color: #fff; padding: 5px 10px; border-radius: 4px; font-size: 11.5px; opacity: 0; pointer-events: none; transition: opacity 0.15s; z-index: 100; }',
    '  .toast.show { opacity: 1; }',
    '  /* Flashcards View */',
    '  .fc-wrap { width: 100%; max-width: 100%; margin: 10px auto 32px; padding: 0 14px; user-select: none; }',
    '  .fc-top, .fc-progress-track, .fc-scene, .fc-below-card, .study-box, .fc-empty { max-width: 960px; margin-left: auto; margin-right: auto; }',
    '  .fc-top { display: flex; align-items: center; justify-content: space-between; margin-bottom: 8px; font-size: 12px; color: #475569; font-weight: 600; flex-wrap: nowrap; white-space: nowrap; gap: 8px; }',
    '  .fc-progress-track { width: 100%; height: 5px; background: #e2e8f0; border-radius: 99px; overflow: hidden; margin-bottom: 14px; }',
    '  .fc-progress-bar { height: 100%; background: #135c37; width: 0%; transition: width 0.2s ease; }',
    '  .fc-scene { perspective: 1000px; width: 100%; height: 300px; cursor: pointer; }',
    '  .fc-card { width: 100%; height: 100%; position: relative; transform-style: preserve-3d; transition: transform 0.38s cubic-bezier(0.4, 0, 0.2, 1); }',
    '  .fc-card.flipped { transform: rotateY(180deg); }',
    '  .fc-face { position: absolute; inset: 0; backface-visibility: hidden; -webkit-backface-visibility: hidden; border: 1px solid #cbd5e1; border-radius: 10px; background: #ffffff; box-shadow: 0 4px 14px rgba(15, 23, 42, 0.06); display: flex; flex-direction: column; justify-content: space-between; padding: 16px 20px; }',
    '  .fc-face.fc-back { transform: rotateY(180deg); background: #f8faf9; border-color: #86efac; }',
    '  .fc-badges { display: flex; align-items: center; justify-content: space-between; gap: 8px; }',
    '  .fc-badge-group { display: flex; align-items: center; gap: 6px; }',
    '  .fc-lvl-badge { background: #135c37; color: #fff; font-weight: 700; font-size: 11px; padding: 2px 8px; border-radius: 4px; }',
    '  .fc-type-badge { background: #f1f5f9; color: #475569; font-weight: 600; font-size: 11px; padding: 2px 8px; border-radius: 4px; border: 1px solid #e2e8f0; }',
    '  .fc-center { text-align: center; margin: auto 0; }',
    '  .fc-word-box { display: inline-block; padding: 6px 14px; border-radius: 8px; }',
    '  .fc-word { font-size: calc(var(--fs) * 2.2); font-weight: 800; color: #0f172a; letter-spacing: -0.3px; margin-bottom: 6px; display: inline-block; transition: color 0.15s; }',
    '  .fc-word-box:hover .fc-word { color: #135c37; }',
    '  .fc-word-box.no-hover-ex:hover .fc-word { color: #0f172a; }',
    '  .fc-hover-example { font-size: calc(var(--fs) + 3.5px); font-weight: 600; color: #0f172a; background: #fffef7; border: 2px solid #135c37; border-left: 5px solid #135c37; border-radius: 8px; padding: 10px 14px; box-shadow: 0 8px 20px rgba(15, 23, 42, 0.14); line-height: 1.45; max-width: 480px; margin: 24px auto 0; pointer-events: none; opacity: 0; visibility: hidden; transform: translateY(-4px); transition: opacity 0.18s ease, transform 0.18s ease, visibility 0.18s; text-align: left; }',
    '  .fc-word-box:hover .fc-hover-example { opacity: 1; visibility: visible; transform: translateY(0); }',
    '  .fc-meaning { font-size: calc(var(--fs) * 1.55); font-weight: 700; color: #135c37; margin-bottom: 10px; }',
    '  .fc-example { font-size: calc(var(--fs) * 1.05); color: #334155; background: #ffffff; border: 1px solid #e2e8f0; border-radius: 6px; padding: 8px 12px; line-height: 1.4; }',
    '  .fc-hint { text-align: center; font-size: 11px; color: #94a3b8; }',
    '  .fc-controls { display: flex; align-items: center; justify-content: center; gap: 6px; margin: 12px auto 0; width: fit-content; flex-wrap: nowrap; white-space: nowrap; }',
    '  .fc-btn { flex: 0 0 auto; border: 1px solid #cbd5e1; background: #fff; color: #1f2937; border-radius: 5px; padding: 0 9px; height: 28px; font-size: 11.5px; font-weight: 600; cursor: pointer; display: inline-flex; align-items: center; justify-content: center; gap: 4px; white-space: nowrap; flex-shrink: 0; }',
    '  .fc-btn:hover { background: #f1f5f9; }',
    '  .fc-btn-known { background: #135c37; color: #fff; border-color: #135c37; }',
    '  .fc-btn-known:hover { background: #0e4629; }',
    '  .fc-btn-toggle-on { background: #dcfce7; color: #14532d; border-color: #22c55e; font-weight: 700; }',
    '  .fc-btn-toggle-on:hover { background: #bbf7d0; }',
    '  .fc-below-card { margin-top: 12px; border: 2px solid #86efac; border-left: 5px solid #135c37; border-radius: 10px; background: #f8faf9; box-shadow: 0 4px 14px rgba(15, 23, 42, 0.06); padding: 14px 18px; text-align: center; }',
    '  .fc-below-title { font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.6px; color: #64748b; margin-bottom: 6px; }',
    '  .fc-empty { text-align: center; padding: 36px 20px; border: 1px dashed #cbd5e1; border-radius: 8px; color: #64748b; background: #f8fafc; }',
    '  .study-modes-bar { display: flex; align-items: center; justify-content: center; gap: 5px; margin: 0 auto 12px; width: fit-content; flex-wrap: nowrap; white-space: nowrap; background: #f1f5f9; padding: 5px; border-radius: 8px; border: 1px solid #cbd5e1; }',
    '  .mode-btn { flex: 0 0 auto; border: 1px solid transparent; background: transparent; color: #334155; font-size: 11.5px; font-weight: 700; padding: 0 9px; height: 27px; border-radius: 5px; cursor: pointer; display: inline-flex; align-items: center; justify-content: center; gap: 4px; transition: all 0.15s; white-space: nowrap; }',
    '  .mode-btn:hover { background: #ffffff; color: #0f172a; }',
    '  .mode-btn.active { background: #135c37; color: #ffffff; border-color: #0e4629; box-shadow: 0 2px 6px rgba(19, 92, 55, 0.25); }',
    '  .srs-badge { display: inline-flex; align-items: center; gap: 3px; background: #fef9c3; color: #854d0e; border: 1px solid #fde047; font-size: 11px; font-weight: 700; padding: 2px 8px; border-radius: 999px; }',
    '  .study-box { border: 1px solid #cbd5e1; border-radius: 10px; background: #ffffff; box-shadow: 0 4px 14px rgba(15, 23, 42, 0.06); padding: 18px 20px; display: flex; flex-direction: column; gap: 14px; }',
    '  .cloze-sentence { font-size: calc(var(--fs) + 5px); font-weight: 700; color: #0f172a; line-height: 1.55; background: #fffef7; border: 2px solid #135c37; border-left: 5px solid #135c37; padding: 14px 16px; border-radius: 8px; text-align: center; }',
    '  .cloze-blank { display: inline-block; min-width: 95px; border-bottom: 3px solid #135c37; color: #135c37; padding: 0 6px; margin: 0 3px; font-weight: 800; letter-spacing: 1px; }',
    '  .cloze-blank.filled { background: #dcfce7; border-radius: 4px; border-bottom-color: #15803d; }',
    '  .cloze-hint { font-size: calc(var(--fs) + 1px); color: #334155; text-align: center; background: #f8fafc; padding: 8px 12px; border-radius: 6px; border: 1px solid #e2e8f0; }',
    '  .choice-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 8px; }',
    '  .choice-btn { border: 1.5px solid #cbd5e1; background: #ffffff; color: #0f172a; border-radius: 8px; padding: 10px 12px; font-size: calc(var(--fs) + 1px); font-weight: 700; cursor: pointer; text-align: left; transition: all 0.15s; display: flex; align-items: center; justify-content: space-between; }',
    '  .choice-btn:hover { border-color: #135c37; background: #f0fdf4; }',
    '  .choice-btn.correct { background: #dcfce7; border-color: #16a34a; color: #14532d; }',
    '  .choice-btn.wrong { background: #fee2e2; border-color: #ef4444; color: #7f1d1d; }',
    '  .type-row { display: flex; gap: 8px; align-items: center; }',
    '  .type-inp { flex: 1; height: 36px; border: 1.5px solid #cbd5e1; border-radius: 6px; padding: 0 12px; font-size: calc(var(--fs) + 1px); font-weight: 600; outline: none; }',
    '  .type-inp:focus { border-color: #135c37; }',
    '  .match-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; }',
    '  .match-col { display: flex; flex-direction: column; gap: 8px; }',
    '  .match-item { border: 1.5px solid #cbd5e1; background: #ffffff; color: #0f172a; border-radius: 8px; padding: 10px 12px; font-size: calc(var(--fs) + 0.5px); font-weight: 600; cursor: pointer; text-align: left; transition: all 0.15s; min-height: 42px; display: flex; align-items: center; justify-content: space-between; }',
    '  .match-item:hover { border-color: #135c37; background: #f8fafc; }',
    '  .match-item.selected { border-color: #135c37; background: #dcfce7; color: #0f172a; box-shadow: 0 0 0 2px rgba(19, 92, 55, 0.2); }',
    '  .match-item.matched { opacity: 0.38; pointer-events: none; background: #f1f5f9; border-color: #cbd5e1; text-decoration: line-through; }',
    '  #statsView { max-width: 980px; }',
    '  #statsView table.sheet-table { width: auto; min-width: 820px; max-width: 960px; table-layout: fixed; }',
    '  #statsView table.sheet-table th, #statsView table.sheet-table td { padding: 3px 6px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }',
    '  .stats-kpi-bar { display: flex; align-items: center; justify-content: space-between; gap: 8px; margin-bottom: 8px; padding: 6px 10px; background: #f8fafc; border: 1px solid #cbd5e1; border-radius: 6px; max-width: 960px; flex-wrap: nowrap; white-space: nowrap; }',
    '  .stats-kpis { display: flex; align-items: center; gap: 8px; flex-wrap: nowrap; white-space: nowrap; }',
    '  .stats-kpi { padding: 3px 8px; border-radius: 5px; font-size: 11.5px; font-weight: 700; background: #ffffff; border: 1px solid #cbd5e1; color: #1e293b; }',
    '  .stats-kpi.err { background: #fef2f2; border-color: #fecaca; color: #991b1b; }',
    '  .stats-kpi.ok { background: #f0fdf4; border-color: #bbf7d0; color: #166534; }',
    '  .err-bar-wrap { display: flex; align-items: center; gap: 6px; width: 100%; }',
    '  .err-bar-track { flex: 1; height: 10px; background: #f1f5f9; border: 1px solid #e2e8f0; border-radius: 99px; overflow: hidden; }',
    '  .err-bar-fill { height: 100%; background: linear-gradient(90deg, #ef4444, #b91c1c); border-radius: 99px; transition: width 0.2s ease; }',
    '  .err-bar-pct { font-size: 10.5px; font-weight: 700; color: #64748b; min-width: 38px; text-align: right; }',
    '  .err-count-pill { display: inline-block; min-width: 26px; padding: 1px 6px; border-radius: 99px; font-weight: 800; font-size: 11px; background: #fee2e2; color: #991b1b; text-align: center; }',
    '  .err-count-pill.zero { background: #f1f5f9; color: #64748b; font-weight: 600; }',
    '  .corr-count-pill { display: inline-block; min-width: 26px; padding: 1px 6px; border-radius: 99px; font-weight: 700; font-size: 11px; background: #dcfce7; color: #166534; text-align: center; }',
    '  @media (max-width: 900px) { .pool-layout { grid-template-columns: 1fr; } }',
    '</style>',
    '</head>',
    '<body>',
    '  <div class="toolbar">',
    '    <div class="toolbar-left">',
    '      <div class="mini-stats">',
    '        <span class="mini-stat" id="stTotal" style="cursor:pointer;" onclick="sekmeDegistir(\'sheet1\')" title="Click to open Main Vocabulary Table (Words)">Total: -</span>',
    '        <span class="mini-stat unk" id="stUnk" style="cursor:pointer;" onclick="sekmeDegistir(\'pool\')" title="Click to open Unknown Words Pool Table">In Pool: -</span>',
    '        <span class="mini-stat knw" id="stKnw" style="cursor:pointer;" onclick="sekmeDegistir(\'knownPool\')" title="Click to open Known Words Pool Table">Known: -</span>',
    '      </div>',
    '      <button class="btn toggle-known-btn show-mode" id="toggleKnownBtn" onclick="bilinenleriGosterAcKapat()" title="Bildiğim kelimeleri tabloda göster / gizle">👁️ Bildiklerimi Gizle</button>',
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
    '      <select id="streakWeeksSelect" class="inp" onchange="hedefHaftaDegistir(this.value)" title="Kaç hafta üst üste hatasız bilinen kelimeler otomatik Biliyorum sayılsın? (1-12 Hafta)">',
    '        <option value="1">Streak: 1W</option>',
    '        <option value="2">Streak: 2W</option>',
    '        <option value="3">Streak: 3W</option>',
    '        <option value="4" selected>Streak: 4W</option>',
    '        <option value="5">Streak: 5W</option>',
    '        <option value="6">Streak: 6W</option>',
    '        <option value="7">Streak: 7W</option>',
    '        <option value="8">Streak: 8W</option>',
    '        <option value="9">Streak: 9W</option>',
    '        <option value="10">Streak: 10W</option>',
    '        <option value="11">Streak: 11W</option>',
    '        <option value="12">Streak: 12W</option>',
    '      </select>',
    '      <button class="btn" id="topBtn-sheet1" onclick="sekmeDegistir(\'sheet1\')" title="Ana Kelime Tablosu (Words)">📋 Words</button>',
    '      <button class="btn" id="topBtn-pool" onclick="sekmeDegistir(\'pool\')" title="Bilmediğim Kelimeler Havuzu Tablosu (Unknown Words Pool)">📕 Havuz</button>',
    '      <button class="btn" id="topBtn-knownPool" onclick="sekmeDegistir(\'knownPool\')" title="Bildiğim Kelimeler Havuzu Tablosu (Known Words Pool)">📗 Known</button>',
    '      <button class="btn" onclick="calismaModuAc(\'flashcards\')" title="3D Flashcards + 3-Step Spaced Repetition">🃏 Flashcards</button>',
    '      <button class="btn" onclick="calismaModuAc(\'cloze\')" title="Fill in the blank in Example Sentences">📝 Cloze</button>',
    '      <button class="btn" onclick="calismaModuAc(\'quiz\')" title="Active Recall (Turkish -> English Quiz & Typing)">🎯 Quiz</button>',
    '      <button class="btn" onclick="calismaModuAc(\'listen\')" title="Listen to audio & find/type the word">🎧 Listen</button>',
    '      <button class="btn" onclick="calismaModuAc(\'match\')" title="Match 5 English words with Turkish meanings">🧩 Match</button>',
    '      <button class="btn" id="topBtn-stats" onclick="sekmeDegistir(\'stats\')" title="Most Missed Words & Mistake Bar Chart">📊 Stats</button>',
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
    '          <th class="col-mark" title="Üst (✓): Biliyorum | Alt (✕): Bilmiyorum">✓/✕</th>',
    '          <th>Word</th>',
    '          <th>Meaning</th>',
    '          <th>Example Sentence</th>',
    '          <th class="center">Status</th>',
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
    '            <th class="col-mark" title="Tek tık: ✓ Biliyorum | Çift tık: ✕ Bilmiyorum">✓/✕</th>',
    '            <th>Word</th>',
    '            <th>Meaning</th>',
    '            <th>Example Sentence</th>',
    '            <th class="col-date">Date</th>',
    '          </tr>',
    '        </thead>',
    '        <tbody id="poolBody"></tbody>',
    '      </table>',
    '    </div>',
    '  </div>',
    '',
    '  <div class="table-wrap" id="knownPoolView" style="display:none;">',
    '    <div class="pool-layout">',
    '      <table class="sheet-table">',
    '        <thead>',
    '          <tr>',
    '            <th>Level</th>',
    '            <th class="center">Known Words</th>',
    '          </tr>',
    '        </thead>',
    '        <tbody id="knownPoolSummaryBody"></tbody>',
    '      </table>',
    '      <table class="sheet-table">',
    '        <thead>',
    '          <tr>',
    '            <th class="center">Level</th>',
    '            <th>Type</th>',
    '            <th class="col-mark" title="Tek tık: ✓ Biliyorum | Çift tık: ✕ Bilmiyorum">✓/✕</th>',
    '            <th>Word</th>',
    '            <th>Meaning</th>',
    '            <th>Example Sentence</th>',
    '            <th class="col-date">Date</th>',
    '          </tr>',
    '        </thead>',
    '        <tbody id="knownPoolBody"></tbody>',
    '      </table>',
    '    </div>',
    '  </div>',
    '',
    '  <div class="table-wrap" id="statsView" style="display:none;">',
    '    <div class="stats-kpi-bar">',
    '      <div class="stats-kpis">',
    '        <span class="stats-kpi" id="kpiTotalAttempts">Total Exercises: 0</span>',
    '        <span class="stats-kpi err" id="kpiTotalMistakes">❌ Total Mistakes: 0</span>',
    '        <span class="stats-kpi ok" id="kpiTotalCorrect">✅ Total Correct: 0</span>',
    '        <span class="stats-kpi" id="kpiAccuracy">🎯 Accuracy: -</span>',
    '      </div>',
    '      <div style="display:flex;gap:6px;align-items:center;">',
    '        <label for="statsStreakSelect" style="font-size:11px;font-weight:700;color:#334155;white-space:nowrap;">🏆 Auto-Known Streak:</label>',
    '        <select id="statsStreakSelect" class="inp" style="font-weight:700;color:#135c37;border-color:#86efac;background:#f0fdf4;" onchange="hedefHaftaDegistir(this.value)" title="Select consecutive zero-mistake weeks (1-12) before a word is automatically marked Known">',
    '          <option value="1">1 Week</option>',
    '          <option value="2">2 Weeks</option>',
    '          <option value="3">3 Weeks</option>',
    '          <option value="4" selected>4 Weeks</option>',
    '          <option value="5">5 Weeks</option>',
    '          <option value="6">6 Weeks</option>',
    '          <option value="7">7 Weeks</option>',
    '          <option value="8">8 Weeks</option>',
    '          <option value="9">9 Weeks</option>',
    '          <option value="10">10 Weeks</option>',
    '          <option value="11">11 Weeks</option>',
    '          <option value="12">12 Weeks</option>',
    '        </select>',
    '      </div>',
    '    </div>',
    '    <table class="sheet-table">',
    '      <thead>',
    '        <tr>',
    '          <th style="width:28px;" class="center">#</th>',
    '          <th style="width:44px;" class="center">Level</th>',
    '          <th style="width:115px;">Word</th>',
    '          <th style="width:170px;">Meaning</th>',
    '          <th style="width:66px;" class="center">Mistakes</th>',
    '          <th style="width:62px;" class="center">Correct</th>',
    '          <th style="width:50px;" class="center">Total</th>',
    '          <th style="width:72px;" class="center">Error Rate</th>',
    '          <th style="width:140px;">Mistake Bar</th>',
    '          <th id="thStreakCol" style="width:72px;" class="center">4W Streak</th>',
    '          <th style="width:120px;">Last Practice</th>',
    '        </tr>',
    '      </thead>',
    '      <tbody id="statsBody"></tbody>',
    '    </table>',
    '  </div>',
    '',
    '  <div id="flashcardView" style="display:none;">',
    '    <div class="fc-wrap" id="fcContainer"></div>',
    '  </div>',
    '',
    '  <div class="bottom-bar">',
    '    <button class="sheet-tab active" id="tabBtn-sheet1" onclick="sekmeDegistir(\'sheet1\')">Words</button>',
    '    <button class="sheet-tab" id="tabBtn-pool" onclick="sekmeDegistir(\'pool\')">Unknown Words Pool</button>',
    '    <button class="sheet-tab" id="tabBtn-knownPool" onclick="sekmeDegistir(\'knownPool\')">Known Words Pool</button>',
    '    <button class="sheet-tab" id="tabBtn-stats" onclick="sekmeDegistir(\'stats\')">📊 Mistake Stats</button>',
    '    <button class="sheet-tab" id="tabBtn-flashcards" onclick="calismaModuAc(\'flashcards\')">🃏 Flashcards</button>',
    '    <button class="sheet-tab" id="tabBtn-cloze" onclick="calismaModuAc(\'cloze\')">📝 Cloze</button>',
    '    <button class="sheet-tab" id="tabBtn-quiz" onclick="calismaModuAc(\'quiz\')">🎯 Quiz</button>',
    '    <button class="sheet-tab" id="tabBtn-listen" onclick="calismaModuAc(\'listen\')">🎧 Listen</button>',
    '    <button class="sheet-tab" id="tabBtn-match" onclick="calismaModuAc(\'match\')">🧩 Match</button>',
    '  </div>',
    '',
    '  <div class="toast" id="toastBox"></div>',
    '',
    '<script>',
    '  var TUM_KELIMELER = [];',
    '  var AKTIF_SEKME = "sheet1";',
    '  var YAZI_BOYUTU = 12;',
    '  var SES_ACIK = true;',
    '  var BILINENLERI_GOSTER = true;',
    '  var HEDEF_HAFTA = 4;',
    '',
    '  function hedefHaftaAl() {',
    '    var n = parseInt(HEDEF_HAFTA, 10);',
    '    if (isNaN(n) || n < 1) return 4;',
    '    if (n > 12) return 12;',
    '    return n;',
    '  }',
    '  function hedefHaftaSecicileriGuncelle() {',
    '    var h = hedefHaftaAl();',
    '    var s1 = document.getElementById("streakWeeksSelect"); if (s1) s1.value = String(h);',
    '    var s2 = document.getElementById("statsStreakSelect"); if (s2) s2.value = String(h);',
    '    var th = document.getElementById("thStreakCol"); if (th) th.innerText = h + "W Streak";',
    '  }',
    '  function hedefHaftaDegistir(val) {',
    '    var n = parseInt(val, 10);',
    '    if (isNaN(n) || n < 1) n = 1;',
    '    if (n > 12) n = 12;',
    '    HEDEF_HAFTA = n;',
    '    try { localStorage.setItem("kelime_hedef_hafta", String(HEDEF_HAFTA)); } catch (e) {}',
    '    hedefHaftaSecicileriGuncelle();',
    '    if (typeof tercihleriKaydet === "function") tercihleriKaydet();',
    '    tabloyuCiz();',
    '    bildirimGoster("🏆 Auto-Known Streak: " + HEDEF_HAFTA + (HEDEF_HAFTA === 1 ? " Week" : " Weeks"));',
    '  }',
    '',
    '  function bilinenleriGosterDurumuUygula() {',
    '    var btn = document.getElementById("toggleKnownBtn");',
    '    if (!btn) return;',
    '    if (BILINENLERI_GOSTER) {',
    '      btn.className = "btn toggle-known-btn show-mode";',
    '      btn.innerText = "👁️ Bildiklerimi Gizle";',
    '      btn.title = "Şu an bildiğin kelimeler tabloda gösteriliyor. Gizlemek için tıkla.";',
    '    } else {',
    '      btn.className = "btn toggle-known-btn hide-mode";',
    '      btn.innerText = "🙈 Bildiklerimi Göster";',
    '      btn.title = "Şu an bildiğin kelimeler tabloda gizli. Göstermek için tıkla.";',
    '    }',
    '  }',
    '',
    '  function bilinenleriGosterAcKapat() {',
    '    BILINENLERI_GOSTER = !BILINENLERI_GOSTER;',
    '    bilinenleriGosterDurumuUygula();',
    '    try { localStorage.setItem("kelime_tablo_bilinen_goster", BILINENLERI_GOSTER ? "1" : "0"); } catch (e) {}',
    '    if (typeof tercihleriKaydet === "function") tercihleriKaydet();',
    '    tabloyuCiz();',
    '    bildirimGoster(BILINENLERI_GOSTER ? "👁️ Bildiğim kelimeler gösteriliyor" : "🙈 Bildiğim kelimeler gizlendi");',
    '  }',
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
    '    var kayitliBilinen = localStorage.getItem("kelime_tablo_bilinen_goster");',
    '    if (kayitliBilinen === "0") { BILINENLERI_GOSTER = false; }',
    '    var kayitliHafta = parseInt(localStorage.getItem("kelime_hedef_hafta"), 10);',
    '    if (!isNaN(kayitliHafta) && kayitliHafta >= 1 && kayitliHafta <= 12) { HEDEF_HAFTA = kayitliHafta; }',
    '  } catch (e) {}',
    '  yaziBoyutuUygula();',
    '  sesDurumuUygula();',
    '  bilinenleriGosterDurumuUygula();',
    '  hedefHaftaSecicileriGuncelle();',
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
    '      if (res.stats) { try { localStorage.setItem("kelime_istatistik_v1", JSON.stringify(res.stats)); } catch (e) {} }',
    '      if (res.webAppUrl && window.self !== window.top) {',
    '        var wb = document.getElementById("openWebAppBtn");',
    '        wb.href = res.webAppUrl;',
    '        wb.style.display = "inline-flex";',
    '      }',
    '      tabloyuCiz();',
    '    }).verileriGetir();',
    '  }',
    '',
    '  function istatistikMapAl() { try { return JSON.parse(localStorage.getItem("kelime_istatistik_v1") || "{}"); } catch (e) { return {}; } }',
    '  function alistirmaSonucuKaydet(kelime, dogruMu, mod) {',
    '    var temiz = (kelime || "").toString().trim(); if (!temiz) return;',
    '    var key = temiz.toLowerCase();',
    '    var hedef = TUM_KELIMELER.filter(function(w){ return (w.kelime || "").toLowerCase() === key; })[0];',
    '    var map = istatistikMapAl();',
    '    var kayit = map[key] || { kelime: hedef ? hedef.kelime : temiz, seviye: hedef ? hedef.seviye : "", anlam: hedef ? hedef.anlam : "", mistakes: 0, correct: 0, lastPractice: "" };',
    '    if (dogruMu) kayit.correct = (Number(kayit.correct) || 0) + 1; else kayit.mistakes = (Number(kayit.mistakes) || 0) + 1;',
    '    map[key] = kayit;',
    '    try { localStorage.setItem("kelime_istatistik_v1", JSON.stringify(map)); } catch (e) {}',
    '    istatistikTablosuCiz();',
    '    google.script.run.withSuccessHandler(function(res){',
    '      if (res && res.ok && res.stats) { try { localStorage.setItem("kelime_istatistik_v1", JSON.stringify(res.stats)); } catch (e) {} istatistikTablosuCiz(); }',
    '    }).istatistikKaydet(kayit.kelime, !!dogruMu, mod || AKTIF_MOD, kayit.seviye, kayit.anlam);',
    '  }',
    '  function istatistikTablosuCiz() {',
    '    var hedefHafta = hedefHaftaAl();',
    '    hedefHaftaSecicileriGuncelle();',
    '    var map = istatistikMapAl(), kelimeMap = {}, tum = {}, liste = [], topHata = 0, topDogru = 0, hataliKelime = 0, maxMistakes = 1;',
    '    for (var i = 0; i < TUM_KELIMELER.length; i++) { var w = TUM_KELIMELER[i]; var wk = (w.kelime || "").toLowerCase().trim(); if (wk) { kelimeMap[wk] = w; tum[wk] = true; } }',
    '    var mKeys = Object.keys(map); for (var mk = 0; mk < mKeys.length; mk++) tum[mKeys[mk].toLowerCase().trim()] = true;',
    '    var keys = Object.keys(tum);',
    '    for (var k = 0; k < keys.length; k++) {',
    '      var key = keys[k], st = map[key] || {}, ref = kelimeMap[key];',
    '      var m = Number(st.mistakes) || 0, c = Number(st.correct) || 0;',
    '      topHata += m; topDogru += c; if (m > 0) hataliKelime++; if (m > maxMistakes) maxMistakes = m;',
    '      var tot = m + c, errRate = tot > 0 ? Math.round((m / tot) * 100) : 0;',
    '      var cw = m > 0 ? 0 : Math.min(hedefHafta, (Array.isArray(st.cleanWeeks) ? st.cleanWeeks.length : (c > 0 ? 1 : 0)));',
    '      liste.push({ kelime: ref ? ref.kelime : (st.kelime || st.word || key), seviye: ref ? ref.seviye : (st.seviye || st.level || "-"), anlam: ref ? ref.anlam : (st.anlam || st.meaning || ""), mistakes: m, correct: c, total: tot, errRate: errRate, cw: cw, lastPractice: st.lastPractice || "-" });',
    '    }',
    '    liste.sort(function(a, b) { if (b.mistakes !== a.mistakes) return b.mistakes - a.mistakes; if (b.errRate !== a.errRate) return b.errRate - a.errRate; if (b.total !== a.total) return b.total - a.total; return a.kelime.localeCompare(b.kelime); });',
    '    var topDeneme = topHata + topDogru, acc = topDeneme > 0 ? Math.round((topDogru / topDeneme) * 100) + "%" : "—";',
    '    var elA = document.getElementById("kpiTotalAttempts"), elE = document.getElementById("kpiTotalMistakes"), elC = document.getElementById("kpiTotalCorrect"), elAc = document.getElementById("kpiAccuracy");',
    '    if (elA) elA.innerText = "Total Attempts: " + topDeneme; if (elE) elE.innerText = "❌ Mistakes: " + topHata; if (elC) elC.innerText = "✅ Correct: " + topDogru; if (elAc) elAc.innerText = "🎯 Accuracy: " + acc;',
    '    var tabStats = document.getElementById("tabBtn-stats"); if (tabStats) tabStats.innerText = topDeneme > 0 ? "📊 Mistake Stats (❌" + topHata + " / ✅" + topDogru + ")" : "📊 Mistake Stats";',
    '    var tbody = document.getElementById("statsBody"); if (!tbody) return;',
    '    if (!liste.length) { tbody.innerHTML = \'<tr><td colspan="11" class="center" style="padding:20px;color:#64748b;">No words found.</td></tr>\'; return; }',
    '    var rows = [];',
    '    for (var r = 0; r < liste.length; r++) {',
    '      var it = liste[r], barPct = it.mistakes > 0 ? Math.max(6, Math.round((it.mistakes / maxMistakes) * 100)) : 0;',
    '      var stTxt = (it.cw >= hedefHafta && it.mistakes === 0 ? "🏆 " : "") + it.cw + "/" + hedefHafta + " Wk";',
    '      rows.push(\'<tr><td class="col-row">\' + (r + 1) + \'</td><td class="col-lvl">\' + escHtml(it.seviye) + \'</td><td class="col-word" onclick="kelimeOku(\\\'\' + escJs(it.kelime) + \'\\\', this)">\' + escHtml(it.kelime) + \'</td><td title="\' + escHtml(it.anlam) + \'">\' + escHtml(it.anlam) + \'</td><td class="center"><span class="err-count-pill\' + (it.mistakes === 0 ? \' zero\' : \'\') + \'">\' + it.mistakes + \'</span></td><td class="center"><span class="corr-count-pill">\' + it.correct + \'</span></td><td class="center" style="font-weight:700;">\' + it.total + \'</td><td class="center" style="font-weight:700;">\' + it.errRate + \'%</td><td><div class="err-bar-wrap"><div class="err-bar-track"><div class="err-bar-fill" style="width:\' + barPct + \'%;"></div></div><span class="err-bar-pct">\' + it.mistakes + \' err</span></div></td><td class="center" style="font-weight:700;">\' + stTxt + \'</td><td class="col-date">\' + escHtml(it.lastPractice) + \'</td></tr>\');',
    '    }',
    '    tbody.innerHTML = rows.join("");',
    '  }',
    '',
    '  var FC_INDEX = 0;',
    '  var FC_FLIPPED = false;',
    '  var FC_SHUFFLED_KEYS = null;',
    '  var AKTIF_MOD = "flashcards";',
    '  var SECENEKLER_CACHE = { key: "", list: [] };',
    '  var MATCH_DURUM = { seciliEn: null, seciliTr: null, eslesenler: {}, trSira: [], keySet: "" };',
    '',
    '  function srsMapAl() { try { return JSON.parse(localStorage.getItem("kelime_srs_sayac_v1") || "{}"); } catch (e) { return {}; } }',
    '  function srsGetir(kelime) { var map = srsMapAl(); return map[(kelime || "").toString().trim().toLowerCase()] || 0; }',
    '  function srsArtir(card) {',
    '    if (!card || !card.kelime) return false;',
    '    var map = srsMapAl();',
    '    var k = card.kelime.toString().trim().toLowerCase();',
    '    var yeni = (map[k] || 0) + 1;',
    '    if (yeni >= 3) {',
    '      map[k] = 0;',
    '      try { localStorage.setItem("kelime_srs_sayac_v1", JSON.stringify(map)); } catch (e) {}',
    '      kutucukDegisti(card.row, "biliyorum", true);',
    '      bildirimGoster("🎉 Mastered (3/3): " + card.kelime + " → Known!");',
    '      return true;',
    '    } else {',
    '      map[k] = yeni;',
    '      try { localStorage.setItem("kelime_srs_sayac_v1", JSON.stringify(map)); } catch (e) {}',
    '      bildirimGoster("⭐ Progress (" + yeni + "/3): " + card.kelime);',
    '      return false;',
    '    }',
    '  }',
    '',
    '  function tercihleriKaydet() {',
    '    try {',
    '      var addBarEl = document.getElementById("addBar");',
    '      var lvlEl = document.getElementById("levelFilter");',
    '      var qEl = document.getElementById("searchInput");',
    '      var prefs = {',
    '        aktifSekme: AKTIF_SEKME,',
    '        aktifMod: AKTIF_MOD,',
    '        fcIndex: FC_INDEX,',
    '        fcShuffledKeys: FC_SHUFFLED_KEYS,',
    '        fcAltKartAcik: !!window.FC_ALT_KART_ACIK,',
    '        fcHoverOrnekAcik: window.FC_HOVER_ORNEK_ACIK !== false,',
    '        yaziBoyutu: YAZI_BOYUTU,',
    '        sesAcik: SES_ACIK,',
    '        bilinenleriGoster: BILINENLERI_GOSTER,',
    '        hedefHafta: hedefHaftaAl(),',
    '        levelFilter: lvlEl ? lvlEl.value : "ALL",',
    '        searchQuery: qEl ? qEl.value : "",',
    '        addBarOpen: addBarEl ? addBarEl.classList.contains("open") : false',
    '      };',
    '      localStorage.setItem("kelime_ui_tercihler_v1", JSON.stringify(prefs));',
    '    } catch (e) {}',
    '  }',
    '',
    '  function tercihleriYukleVeUygula() {',
    '    try {',
    '      var raw = localStorage.getItem("kelime_ui_tercihler_v1");',
    '      if (!raw) return;',
    '      var p = JSON.parse(raw); if (!p) return;',
    '      if (typeof p.yaziBoyutu === "number") { YAZI_BOYUTU = p.yaziBoyutu; yaziBoyutuUygula(); }',
    '      if (typeof p.sesAcik === "boolean") { SES_ACIK = p.sesAcik; sesDurumuUygula(); }',
    '      if (typeof p.bilinenleriGoster === "boolean") { BILINENLERI_GOSTER = p.bilinenleriGoster; bilinenleriGosterDurumuUygula(); }',
    '      if (typeof p.hedefHafta === "number" && p.hedefHafta >= 1 && p.hedefHafta <= 12) { HEDEF_HAFTA = p.hedefHafta; }',
    '      hedefHaftaSecicileriGuncelle();',
    '      if (typeof p.fcAltKartAcik === "boolean") { window.FC_ALT_KART_ACIK = p.fcAltKartAcik; }',
    '      if (typeof p.fcHoverOrnekAcik === "boolean") { window.FC_HOVER_ORNEK_ACIK = p.fcHoverOrnekAcik; }',
    '      if (typeof p.fcIndex === "number" && p.fcIndex >= 0) { FC_INDEX = p.fcIndex; }',
    '      if (Array.isArray(p.fcShuffledKeys) && p.fcShuffledKeys.length > 0) { FC_SHUFFLED_KEYS = p.fcShuffledKeys; }',
    '      var lvlEl = document.getElementById("levelFilter"); if (lvlEl && p.levelFilter) lvlEl.value = p.levelFilter;',
    '      var qEl = document.getElementById("searchInput"); if (qEl && typeof p.searchQuery === "string") qEl.value = p.searchQuery;',
    '      var addBarEl = document.getElementById("addBar"); if (addBarEl && p.addBarOpen) addBarEl.classList.add("open");',
    '      if (p.aktifMod) AKTIF_MOD = p.aktifMod;',
    '      if (p.aktifSekme) AKTIF_SEKME = p.aktifSekme;',
    '      document.getElementById("sheet1View").style.display = AKTIF_SEKME === "sheet1" ? "block" : "none";',
    '      document.getElementById("poolView").style.display = AKTIF_SEKME === "pool" ? "block" : "none";',
    '      var kpV = document.getElementById("knownPoolView"); if (kpV) kpV.style.display = AKTIF_SEKME === "knownPool" ? "block" : "none";',
    '      document.getElementById("flashcardView").style.display = AKTIF_SEKME === "flashcards" ? "block" : "none";',
    '      var stV = document.getElementById("statsView"); if (stV) stV.style.display = AKTIF_SEKME === "stats" ? "block" : "none";',
    '      sekmeButonlariniGuncelle();',
    '    } catch (e) {}',
    '  }',
    '',
    '  function sekmeButonlariniGuncelle() {',
    '    var s = AKTIF_SEKME, m = AKTIF_MOD;',
    '    document.getElementById("tabBtn-sheet1").classList.toggle("active", s === "sheet1");',
    '    document.getElementById("tabBtn-pool").classList.toggle("active", s === "pool");',
    '    var tkp = document.getElementById("tabBtn-knownPool"); if (tkp) tkp.classList.toggle("active", s === "knownPool");',
    '    var ts = document.getElementById("tabBtn-stats"); if (ts) ts.classList.toggle("active", s === "stats");',
    '    var studyTabs = ["flashcards", "cloze", "quiz", "listen", "match"];',
    '    for (var i = 0; i < studyTabs.length; i++) {',
    '      var el = document.getElementById("tabBtn-" + studyTabs[i]);',
    '      if (el) el.classList.toggle("active", s === "flashcards" && m === studyTabs[i]);',
    '    }',
    '  }',
    '',
    '  function sekmeDegistir(s) {',
    '    AKTIF_SEKME = s;',
    '    sekmeButonlariniGuncelle();',
    '    document.getElementById("sheet1View").style.display = s === "sheet1" ? "block" : "none";',
    '    document.getElementById("poolView").style.display = s === "pool" ? "block" : "none";',
    '    var kpV = document.getElementById("knownPoolView"); if (kpV) kpV.style.display = s === "knownPool" ? "block" : "none";',
    '    document.getElementById("flashcardView").style.display = s === "flashcards" ? "block" : "none";',
    '    var stV = document.getElementById("statsView"); if (stV) stV.style.display = s === "stats" ? "block" : "none";',
    '    FC_FLIPPED = false;',
    '    tercihleriKaydet();',
    '    tabloyuCiz();',
    '  }',
    '',
    '  function calismaModuAc(mod) {',
    '    AKTIF_MOD = mod || "flashcards";',
    '    AKTIF_SEKME = "flashcards";',
    '    FC_FLIPPED = false;',
    '    SECENEKLER_CACHE = { key: "", list: [] };',
    '    if (mod === "match") { MATCH_DURUM = { seciliEn: null, seciliTr: null, eslesenler: {}, trSira: [], keySet: "" }; }',
    '    sekmeButonlariniGuncelle();',
    '    document.getElementById("sheet1View").style.display = "none";',
    '    document.getElementById("poolView").style.display = "none";',
    '    var kpV = document.getElementById("knownPoolView"); if (kpV) kpV.style.display = "none";',
    '    document.getElementById("flashcardView").style.display = "block";',
    '    var stV = document.getElementById("statsView"); if (stV) stV.style.display = "none";',
    '    tercihleriKaydet();',
    '    tabloyuCiz();',
    '  }',
    '',
    '  function modCubuguHtml() {',
    '    var modes = [',
    '      { id: "flashcards", label: "🃏 Flashcards" },',
    '      { id: "cloze", label: "📝 Cloze" },',
    '      { id: "quiz", label: "🎯 Quiz" },',
    '      { id: "listen", label: "🎧 Listen" },',
    '      { id: "match", label: "🧩 Match" }',
    '    ];',
    '    var html = [\'<div class="study-modes-bar">\'];',
    '    html.push(\'<button class="mode-btn" onclick="sekmeDegistir(\\\'sheet1\\\')">📋 Words</button>\');',
    '    html.push(\'<button class="mode-btn" onclick="sekmeDegistir(\\\'pool\\\')">📕 Havuz</button>\');',
    '    html.push(\'<button class="mode-btn" onclick="sekmeDegistir(\\\'knownPool\\\')">📗 Known</button>\');',
    '    for (var i = 0; i < modes.length; i++) {',
    '      var m = modes[i];',
    '      html.push(\'<button class="mode-btn\' + (AKTIF_MOD === m.id ? \' active\' : \'\') + \'" onclick="calismaModuAc(\\\'\' + m.id + \'\\\')">\' + m.label + \'</button>\');',
    '    }',
    '    html.push(\'<button class="mode-btn" onclick="sekmeDegistir(\\\'stats\\\')">📊 Stats</button>\');',
    '    html.push(\'</div>\');',
    '    return html.join("");',
    '  }',
    '',
    '  function secenekleriUret(dogruCard, tumHavuz) {',
    '    var dogruKey = (dogruCard.kelime || "").toLowerCase().trim();',
    '    var havuzAdaylar = (tumHavuz || []).filter(function(w) { return w && w.durum === "bilmiyorum" && (w.kelime || "").toLowerCase().trim() !== dogruKey; });',
    '    var tumBilinmeyenAdaylar = TUM_KELIMELER.filter(function(w) { return w && w.durum === "bilmiyorum" && (w.kelime || "").toLowerCase().trim() !== dogruKey; });',
    '    var hedefSayi = Math.min(4, 1 + tumBilinmeyenAdaylar.length);',
    '    var cacheKey = AKTIF_MOD + ":" + dogruKey;',
    '    if (SECENEKLER_CACHE.key === cacheKey && SECENEKLER_CACHE.list.length === hedefSayi && SECENEKLER_CACHE.list.every(function(w) { return w && w.durum === "bilmiyorum"; })) return SECENEKLER_CACHE.list;',
    '    for (var i = havuzAdaylar.length - 1; i > 0; i--) { var r = Math.floor(Math.random() * (i + 1)); var t = havuzAdaylar[i]; havuzAdaylar[i] = havuzAdaylar[r]; havuzAdaylar[r] = t; }',
    '    var secenekler = [dogruCard], kullanilan = {}; kullanilan[dogruKey] = true;',
    '    for (var j = 0; j < havuzAdaylar.length && secenekler.length < 4; j++) { var hk = (havuzAdaylar[j].kelime || "").toLowerCase().trim(); if (!kullanilan[hk]) { secenekler.push(havuzAdaylar[j]); kullanilan[hk] = true; } }',
    '    if (secenekler.length < 4 && tumBilinmeyenAdaylar.length > havuzAdaylar.length) {',
    '      for (var m = tumBilinmeyenAdaylar.length - 1; m > 0; m--) { var rm = Math.floor(Math.random() * (m + 1)); var tm = tumBilinmeyenAdaylar[m]; tumBilinmeyenAdaylar[m] = tumBilinmeyenAdaylar[rm]; tumBilinmeyenAdaylar[rm] = tm; }',
    '      for (var n = 0; n < tumBilinmeyenAdaylar.length && secenekler.length < 4; n++) { var tk = (tumBilinmeyenAdaylar[n].kelime || "").toLowerCase().trim(); if (!kullanilan[tk]) { secenekler.push(tumBilinmeyenAdaylar[n]); kullanilan[tk] = true; } }',
    '    }',
    '    for (var k = secenekler.length - 1; k > 0; k--) { var rk = Math.floor(Math.random() * (k + 1)); var tk2 = secenekler[k]; secenekler[k] = secenekler[rk]; secenekler[rk] = tk2; }',
    '    SECENEKLER_CACHE = { key: cacheKey, list: secenekler };',
    '    return secenekler;',
    '  }',
    '',
    '  function ornekCumleyiAyir(ornek) {',
    '    var s = (ornek || "").toString().trim();',
    '    var m = s.match(/^(.*?)\\s*\\(([^)]+)\\)\\s*$/);',
    '    return m ? { en: m[1].trim(), tr: m[2].trim() } : { en: s, tr: "" };',
    '  }',
    '',
    '  function cumledeKelimeyiGizle(enCumle, kelime) {',
    '    if (!enCumle) return \'<span class="cloze-blank" id="clozeBlankEl">_______</span>\';',
    '    var temizKelime = (kelime || "").trim();',
    '    var kok = temizKelime.length > 4 ? temizKelime.slice(0, -1) : temizKelime;',
    '    var escKok = kok.replace(/[.*+?^${}()|[\\]\\\\]/g, "\\\\$&");',
    '    var rgx = new RegExp("\\\\b" + escKok + "[a-zA-Z]*\\\\b", "i");',
    '    var guvenli = escHtml(enCumle);',
    '    if (rgx.test(guvenli)) return guvenli.replace(rgx, \'<span class="cloze-blank" id="clozeBlankEl">_______</span>\');',
    '    return guvenli + \' (<span class="cloze-blank" id="clozeBlankEl">_______</span>)\';',
    '  }',
    '',
    '  function fcDestesiniGetir(poolList) {',
    '    var sadeceBilinmeyenler = (poolList || []).filter(function(w) { return w && w.durum === "bilmiyorum"; });',
    '    if (!FC_SHUFFLED_KEYS || !FC_SHUFFLED_KEYS.length) return sadeceBilinmeyenler;',
    '    var byKey = {};',
    '    for (var i = 0; i < sadeceBilinmeyenler.length; i++) { byKey[sadeceBilinmeyenler[i].kelime.toLowerCase()] = sadeceBilinmeyenler[i]; }',
    '    var result = [], used = {};',
    '    for (var j = 0; j < FC_SHUFFLED_KEYS.length; j++) {',
    '      var k = FC_SHUFFLED_KEYS[j];',
    '      if (byKey[k]) { result.push(byKey[k]); used[k] = true; }',
    '    }',
    '    for (var m = 0; m < sadeceBilinmeyenler.length; m++) {',
    '      var km = sadeceBilinmeyenler[m].kelime.toLowerCase();',
    '      if (!used[km]) result.push(sadeceBilinmeyenler[m]);',
    '    }',
    '    return result;',
    '  }',
    '',
    '  function fcCiz(poolList) {',
    '    var box = document.getElementById("fcContainer");',
    '    if (!box) return;',
    '    var deck = fcDestesiniGetir(poolList);',
    '    if (!deck.length) {',
    '      box.innerHTML = modCubuguHtml() + \'<div class="fc-empty"><div style="font-size:28px;margin-bottom:8px;">🎉</div><div style="font-weight:700;color:#1e293b;margin-bottom:6px;">No Unknown Words in Pool</div><div style="font-size:12px;margin-bottom:14px;">Mark words as <b>Unknown</b> in Words to study them.</div><button class="fc-btn" onclick="sekmeDegistir(\\\'sheet1\\\')">Go to Words</button></div>\';',
    '      return;',
    '    }',
    '    if (FC_INDEX >= deck.length) FC_INDEX = 0;',
    '    if (FC_INDEX < 0) FC_INDEX = deck.length - 1;',
    '    if (AKTIF_MOD === "cloze") { clozeModuCiz(box, deck); return; }',
    '    if (AKTIF_MOD === "quiz") { quizModuCiz(box, deck); return; }',
    '    if (AKTIF_MOD === "listen") { listenModuCiz(box, deck); return; }',
    '    if (AKTIF_MOD === "match") { matchModuCiz(box, deck); return; }',
    '    var card = deck[FC_INDEX];',
    '    var pct = Math.round(((FC_INDEX + 1) / deck.length) * 100);',
    '    var srsSayi = srsGetir(card.kelime);',
    '    var altAcik = !!window.FC_ALT_KART_ACIK;',
    '    var hoverAcik = window.FC_HOVER_ORNEK_ACIK !== false;',
    '    box.innerHTML =',
    '      modCubuguHtml() +',
    '      \'<div class="fc-top"><span>Unknown Pool Flashcards — <b>\' + (FC_INDEX + 1) + \' / \' + deck.length + \'</b></span><div style="display:flex;gap:6px;align-items:center;"><span class="srs-badge">⭐ \' + srsSayi + \' / 3</span><button class="fc-btn\' + (hoverAcik ? \' fc-btn-toggle-on\' : \'\') + \'" style="height:26px;font-size:11px;padding:0 9px;" onclick="fcHoverOrnekAcKapat()" title="Flashcards\\\'da kelimenin üstüne gelince altta örnek cümle ve anlamını göster/gizle">\' + (hoverAcik ? \'💬 Üstüne Gelince Örnek: AÇIK\' : \'💬 Üstüne Gelince Örnek: KAPALI\') + \'</button><button class="fc-btn\' + (altAcik ? \' fc-btn-toggle-on\' : \'\') + \'" style="height:26px;font-size:11px;padding:0 9px;" onclick="fcAltKartAcKapat()">\' + (altAcik ? \'📖 Alt Açıklama Kartı: AÇIK\' : \'📖 Alt Açıklama Kartı: KAPALI\') + \'</button><button class="btn" onclick="fcKaristir()" title="Shuffle cards">🔀 Shuffle</button></div></div>\' +',
    '      \'<div class="fc-progress-track"><div class="fc-progress-bar" style="width:\' + pct + \'%;"></div></div>\' +',
    '      \'<div class="fc-scene" onclick="fcCevir()">\' +',
    '        \'<div class="fc-card\' + (FC_FLIPPED ? \' flipped\' : \'\') + \'" id="fcCardInner">\' +',
    '          \'<div class="fc-face fc-front">\' +',
    '            \'<div class="fc-badges"><div class="fc-badge-group"><span class="fc-lvl-badge">\' + escHtml(card.seviye) + \'</span><span class="fc-type-badge">\' + escHtml(turTemizle(card.tur)) + \'</span><span class="srs-badge">⭐ \' + srsSayi + \'/3</span></div><button class="btn" onclick="event.stopPropagation(); kelimeOku(\\\'\' + escJs(card.kelime) + \'\\\')" title="Listen pronunciation">🔊 Listen</button></div>\' +',
    '            \'<div class="fc-center"><div class="fc-word-box\' + (hoverAcik ? \'\' : \' no-hover-ex\') + \'"><div class="fc-word">\' + escHtml(card.kelime) + \'</div>\' + (hoverAcik && card.ornek ? \'<div class="fc-hover-example">\' + ornekBicimlendir(card.ornek) + \'</div>\' : \'\') + \'</div></div>\' +',
    '            \'<div class="fc-hint">\' + (hoverAcik ? \'Hover over word for example sentence • Click card or Space to flip ↻\' : \'Click card or Space to flip ↻\') + \'</div>\' +',
    '          \'</div>\' +',
    '          \'<div class="fc-face fc-back">\' +',
    '            \'<div class="fc-badges"><div class="fc-badge-group"><span class="fc-lvl-badge">\' + escHtml(card.seviye) + \'</span><span class="fc-type-badge">\' + escHtml(turTemizle(card.tur)) + \'</span><strong style="font-size:13px;color:#0f172a;margin-left:4px;">\' + escHtml(card.kelime) + \'</strong></div><button class="btn" onclick="event.stopPropagation(); kelimeOku(\\\'\' + escJs(card.kelime) + \'\\\')" title="Listen pronunciation">🔊</button></div>\' +',
    '            \'<div class="fc-center"><div class="fc-meaning">\' + escHtml(card.anlam) + \'</div>\' + (card.ornek ? \'<div class="fc-example">\' + ornekBicimlendir(card.ornek) + \'</div>\' : \'\') + \'</div>\' +',
    '            \'<div class="fc-hint">Click card to flip back ↻</div>\' +',
    '          \'</div>\' +',
    '        \'</div>\' +',
    '      \'</div>\' +',
    '      (altAcik ? \'<div class="fc-below-card"><div class="fc-below-title">📖 Meaning & Example Sentence (\' + escHtml(card.kelime) + \')</div><div class="fc-meaning">\' + escHtml(card.anlam) + \'</div>\' + (card.ornek ? \'<div class="fc-example">\' + ornekBicimlendir(card.ornek) + \'</div>\' : \'\') + \'</div>\' : \'\') +',
    '      \'<div class="fc-controls">\' +',
    '        \'<button class="fc-btn" onclick="fcDegistir(-1)">◀ Prev</button>\' +',
    '        \'<button class="fc-btn" onclick="fcCevir()">↻ Flip</button>\' +',
    '        \'<button class="fc-btn\' + (hoverAcik ? \' fc-btn-toggle-on\' : \'\') + \'" onclick="fcHoverOrnekAcKapat()" title="Flashcards\\\'da kelimenin üstüne gelince altta örnek cümle ve anlamını göster/gizle">\' + (hoverAcik ? \'💬 Hover Örnek: AÇIK\' : \'💬 Hover Örnek: KAPALI\') + \'</button>\' +',
    '        \'<button class="fc-btn\' + (altAcik ? \' fc-btn-toggle-on\' : \'\') + \'" onclick="fcAltKartAcKapat()">\' + (altAcik ? \'📖 Alt Kart: AÇIK\' : \'📖 Alt Kart: KAPALI\') + \'</button>\' +',
    '        \'<button class="fc-btn" style="background:#fffbeb;border-color:#fde68a;color:#b45309;" onclick="fcSrsArtir(\' + card.row + \')">⭐ +1 Step (\' + srsSayi + \'/3)</button>\' +',
    '        \'<button class="fc-btn fc-btn-known" onclick="fcBiliyorumIsaretle(\' + card.row + \')">✅ Known (Instant)</button>\' +',
    '        \'<button class="fc-btn" onclick="fcDegistir(1)">Next ▶</button>\' +',
    '      \'</div>\';',
    '  }',
    '  try { window.FC_ALT_KART_ACIK = localStorage.getItem("kelime_fc_alt_kart") === "1"; } catch (e) { window.FC_ALT_KART_ACIK = false; }',
    '  try { var kayitliHoverOrnek = localStorage.getItem("kelime_fc_hover_ornek"); window.FC_HOVER_ORNEK_ACIK = kayitliHoverOrnek === null ? true : kayitliHoverOrnek === "1"; } catch (e) { window.FC_HOVER_ORNEK_ACIK = true; }',
    '  function fcHoverOrnekAcKapat() { window.FC_HOVER_ORNEK_ACIK = window.FC_HOVER_ORNEK_ACIK === false ? true : false; try { localStorage.setItem("kelime_fc_hover_ornek", window.FC_HOVER_ORNEK_ACIK ? "1" : "0"); } catch (e) {} tercihleriKaydet(); tabloyuCiz(); bildirimGoster(window.FC_HOVER_ORNEK_ACIK ? "💬 Üstüne gelince örnek cümle: AÇIK" : "💬 Üstüne gelince örnek cümle: KAPALI"); }',
    '  function fcAltKartAcKapat() { window.FC_ALT_KART_ACIK = !window.FC_ALT_KART_ACIK; try { localStorage.setItem("kelime_fc_alt_kart", window.FC_ALT_KART_ACIK ? "1" : "0"); } catch (e) {} tabloyuCiz(); bildirimGoster(window.FC_ALT_KART_ACIK ? "📖 Alt açıklama kartı açıldı" : "📖 Alt açıklama kartı kapatıldı"); }',
    '',
    '  function clozeModuCiz(box, deck) {',
    '    var card = deck[FC_INDEX], pct = Math.round(((FC_INDEX + 1) / deck.length) * 100), srsSayi = srsGetir(card.kelime);',
    '    var ayrilmis = ornekCumleyiAyir(card.ornek), gizliCumle = cumledeKelimeyiGizle(ayrilmis.en, card.kelime), secenekler = secenekleriUret(card, deck);',
    '    var btnHtml = [];',
    '    for (var i = 0; i < secenekler.length; i++) {',
    '      var sc = secenekler[i];',
    '      btnHtml.push(\'<button class="choice-btn" onclick="clozeSecenekSec(this, \\\'\' + escJs(sc.kelime) + \'\\\', \\\'\' + escJs(card.kelime) + \'\\\', \' + card.row + \')">\' + (i + 1) + \'. \' + escHtml(sc.kelime) + \'</button>\');',
    '    }',
    '    box.innerHTML = modCubuguHtml() + \'<div class="fc-top"><span>📝 Cloze — <b>\' + (FC_INDEX + 1) + \' / \' + deck.length + \'</b></span><span class="srs-badge">⭐ \' + srsSayi + \' / 3</span></div><div class="fc-progress-track"><div class="fc-progress-bar" style="width:\' + pct + \'%;"></div></div><div class="study-box"><div class="cloze-sentence">\' + gizliCumle + \'</div><div class="cloze-hint">💡 <b>Meaning:</b> \' + escHtml(card.anlam) + (ayrilmis.tr ? \' <br>🇹🇷 <i>(\' + escHtml(ayrilmis.tr) + \')</i>\' : \'\') + \'</div><div class="choice-grid">\' + btnHtml.join("") + \'</div></div><div class="fc-controls"><button class="fc-btn" onclick="fcDegistir(-1)">◀ Prev</button><button class="fc-btn" onclick="fcDegistir(1)">Skip / Next ▶</button></div>\';',
    '  }',
    '  function clozeSecenekSec(btnEl, secilen, dogru, row) {',
    '    var card = TUM_KELIMELER.filter(function(w){ return w.row === row; })[0];',
    '    if (secilen.toLowerCase() === dogru.toLowerCase()) {',
    '      btnEl.classList.add("correct");',
    '      var blankEl = document.getElementById("clozeBlankEl"); if (blankEl) { blankEl.innerText = dogru; blankEl.classList.add("filled"); }',
    '      kelimeOku(dogru); alistirmaSonucuKaydet(dogru, true, AKTIF_MOD); var mezun = card ? srsArtir(card) : false;',
    '      setTimeout(function() { SECENEKLER_CACHE = { key: "", list: [] }; if (!mezun) FC_INDEX++; tabloyuCiz(); }, 850);',
    '    } else { btnEl.classList.add("wrong"); alistirmaSonucuKaydet(dogru, false, AKTIF_MOD); bildirimGoster("❌ Try again! (Mistake recorded)"); }',
    '  }',
    '  function quizModuCiz(box, deck) {',
    '    var card = deck[FC_INDEX], pct = Math.round(((FC_INDEX + 1) / deck.length) * 100), srsSayi = srsGetir(card.kelime), secenekler = secenekleriUret(card, deck);',
    '    var btnHtml = [];',
    '    for (var i = 0; i < secenekler.length; i++) {',
    '      var sc = secenekler[i];',
    '      btnHtml.push(\'<button class="choice-btn" onclick="clozeSecenekSec(this, \\\'\' + escJs(sc.kelime) + \'\\\', \\\'\' + escJs(card.kelime) + \'\\\', \' + card.row + \')">\' + (i + 1) + \'. \' + escHtml(sc.kelime) + \'</button>\');',
    '    }',
    '    box.innerHTML = modCubuguHtml() + \'<div class="fc-top"><span>🎯 Active Recall (TR → EN) — <b>\' + (FC_INDEX + 1) + \' / \' + deck.length + \'</b></span><span class="srs-badge">⭐ \' + srsSayi + \' / 3</span></div><div class="fc-progress-track"><div class="fc-progress-bar" style="width:\' + pct + \'%;"></div></div><div class="study-box"><div style="text-align:center;padding:12px;"><div style="font-size:calc(var(--fs) + 10px);font-weight:800;color:#135c37;">\' + escHtml(card.anlam) + \'</div></div><div class="choice-grid">\' + btnHtml.join("") + \'</div></div><div class="fc-controls"><button class="fc-btn" onclick="fcDegistir(-1)">◀ Prev</button><button class="fc-btn" onclick="fcDegistir(1)">Skip / Next ▶</button></div>\';',
    '  }',
    '  function listenModuCiz(box, deck) {',
    '    var card = deck[FC_INDEX], pct = Math.round(((FC_INDEX + 1) / deck.length) * 100), srsSayi = srsGetir(card.kelime), secenekler = secenekleriUret(card, deck);',
    '    var btnHtml = [];',
    '    for (var i = 0; i < secenekler.length; i++) {',
    '      var sc = secenekler[i];',
    '      btnHtml.push(\'<button class="choice-btn" onclick="clozeSecenekSec(this, \\\'\' + escJs(sc.kelime) + \'\\\', \\\'\' + escJs(card.kelime) + \'\\\', \' + card.row + \')">\' + (i + 1) + \'. \' + escHtml(sc.anlam) + \'</button>\');',
    '    }',
    '    box.innerHTML = modCubuguHtml() + \'<div class="fc-top"><span>🎧 Listen & Find — <b>\' + (FC_INDEX + 1) + \' / \' + deck.length + \'</b></span><span class="srs-badge">⭐ \' + srsSayi + \' / 3</span></div><div class="fc-progress-track"><div class="fc-progress-bar" style="width:\' + pct + \'%;"></div></div><div class="study-box"><div style="text-align:center;padding:12px;"><button class="fc-btn fc-btn-known" style="font-size:15px;padding:12px 24px;border-radius:999px;" onclick="kelimeOku(\\\'\' + escJs(card.kelime) + \'\\\')">🔊 Click to Listen</button></div><div class="choice-grid">\' + btnHtml.join("") + \'</div></div><div class="fc-controls"><button class="fc-btn" onclick="fcDegistir(-1)">◀ Prev</button><button class="fc-btn" onclick="fcDegistir(1)">Skip / Next ▶</button></div>\';',
    '  }',
    '  function matchModuCiz(box, deck) {',
    '    var grupBoyutu = Math.min(5, deck.length), baslangic = Math.floor(FC_INDEX / 5) * 5; if (baslangic >= deck.length) baslangic = 0;',
    '    var altListe = deck.slice(baslangic, baslangic + grupBoyutu), grupKey = altListe.map(function(w){ return w.kelime.toLowerCase(); }).join("|");',
    '    if (MATCH_DURUM.keySet !== grupKey) {',
    '      var trIndisler = []; for (var i = 0; i < altListe.length; i++) trIndisler.push(i);',
    '      for (var j = trIndisler.length - 1; j > 0; j--) { var r = Math.floor(Math.random() * (j + 1)); var t = trIndisler[j]; trIndisler[j] = trIndisler[r]; trIndisler[r] = t; }',
    '      MATCH_DURUM = { seciliEn: null, seciliTr: null, eslesenler: {}, trSira: trIndisler, keySet: grupKey };',
    '    }',
    '    var enHtml = [], trHtml = [], eslesenSayisi = 0;',
    '    for (var m = 0; m < altListe.length; m++) {',
    '      var wEn = altListe[m], kEn = wEn.kelime.toLowerCase(), isM = !!MATCH_DURUM.eslesenler[kEn]; if (isM) eslesenSayisi++;',
    '      enHtml.push(\'<div class="match-item\' + (isM ? \' matched\' : \'\') + (MATCH_DURUM.seciliEn === kEn ? \' selected\' : \'\') + \'" onclick="matchEnSec(\\\'\' + escJs(kEn) + \'\\\', \\\'\' + escJs(wEn.kelime) + \'\\\')"><span>\' + escHtml(wEn.kelime) + \'</span><span class="srs-badge">⭐ \' + srsGetir(wEn.kelime) + \'/3</span></div>\');',
    '    }',
    '    for (var n = 0; n < MATCH_DURUM.trSira.length; n++) {',
    '      var idx = MATCH_DURUM.trSira[n], wTr = altListe[idx]; if (!wTr) continue; var kTr = wTr.kelime.toLowerCase();',
    '      trHtml.push(\'<div class="match-item\' + (MATCH_DURUM.eslesenler[kTr] ? \' matched\' : \'\') + (MATCH_DURUM.seciliTr === kTr ? \' selected\' : \'\') + \'" onclick="matchTrSec(\\\'\' + escJs(kTr) + \'\\\')">\' + escHtml(wTr.anlam) + \'</div>\');',
    '    }',
    '    box.innerHTML = modCubuguHtml() + \'<div class="fc-top"><span>🧩 5-Word Match — <b>\' + eslesenSayisi + \' / \' + altListe.length + \'</b></span><button class="btn" onclick="FC_INDEX+=5;MATCH_DURUM.keySet=\\\'\\\';tabloyuCiz();">🔄 Next Set</button></div><div class="study-box"><div class="match-grid"><div class="match-col">\' + enHtml.join("") + \'</div><div class="match-col">\' + trHtml.join("") + \'</div></div></div>\';',
    '  }',
    '  function matchEnSec(kEn, orijinal) { if (MATCH_DURUM.eslesenler[kEn]) return; kelimeOku(orijinal); MATCH_DURUM.seciliEn = kEn; matchKontrolEt(); }',
    '  function matchTrSec(kTr) { if (MATCH_DURUM.eslesenler[kTr]) return; MATCH_DURUM.seciliTr = kTr; matchKontrolEt(); }',
    '  function matchKontrolEt() {',
    '    if (MATCH_DURUM.seciliEn && MATCH_DURUM.seciliTr) {',
    '      var en = MATCH_DURUM.seciliEn;',
    '      var card = TUM_KELIMELER.filter(function(w){ return w.kelime.toLowerCase() === en; })[0];',
    '      if (MATCH_DURUM.seciliEn === MATCH_DURUM.seciliTr) {',
    '        MATCH_DURUM.eslesenler[en] = true; MATCH_DURUM.seciliEn = null; MATCH_DURUM.seciliTr = null;',
    '        alistirmaSonucuKaydet(card ? card.kelime : en, true, "Match");',
    '        if (card) srsArtir(card);',
    '      } else {',
    '        MATCH_DURUM.seciliEn = null; MATCH_DURUM.seciliTr = null;',
    '        alistirmaSonucuKaydet(card ? card.kelime : en, false, "Match");',
    '        bildirimGoster("❌ Try again! (Mistake recorded)");',
    '      }',
    '    }',
    '    tabloyuCiz();',
    '  }',
    '  function fcSrsArtir(row) {',
    '    var card = TUM_KELIMELER.filter(function(w){ return w.row === row; })[0]; if (!card) return;',
    '    FC_FLIPPED = false; var mezun = srsArtir(card); if (!mezun) FC_INDEX++; tabloyuCiz();',
    '  }',
    '',
    '  function ornekBicimlendir(s) {',
    '    var temiz = escHtml(s || "");',
    '    return temiz.replace(/\\(([^)]+)\\)\\s*$/, \'<span class="ex-tr">($1)</span>\');',
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
    '    tercihleriKaydet();',
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
    '    tercihleriKaydet();',
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
    '    var knownLvlCounts = { B1: 0, B2: 0, C1: 0, C2: 0 };',
    '    for (var i = 0; i < TUM_KELIMELER.length; i++) {',
    '      var w = TUM_KELIMELER[i];',
    '      if (w.durum === "bilmiyorum") {',
    '        unk++;',
    '        lvlCounts[w.seviye] = (lvlCounts[w.seviye] || 0) + 1;',
    '      } else if (w.durum === "biliyorum") {',
    '        knw++;',
    '        knownLvlCounts[w.seviye] = (knownLvlCounts[w.seviye] || 0) + 1;',
    '      }',
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
    '      if (!BILINENLERI_GOSTER && x.durum === "biliyorum") continue;',
    '      if (lvl !== "ALL" && x.seviye !== lvl) continue;',
    '      if (q && (x.kelime + " " + x.anlam + " " + x.ornek).toLowerCase().indexOf(q) === -1) continue;',
    '      s1Rows.push(',
    '        \'<tr class="\' + (x.durum === "biliyorum" ? "row-known" : "") + \'">\' +',
    '          \'<td class="col-row">\' + x.row + \'</td>\' +',
    '          \'<td class="col-date">\' + escHtml(x.tarih) + \'</td>\' +',
    '          \'<td class="col-lvl">\' + escHtml(x.seviye) + \'</td>\' +',
    '          \'<td class="col-type">\' + escHtml(turTemizle(x.tur)) + \'</td>\' +',
    '          \'<td class="col-mark">\' + durumKutusuHtml(x.row, x.durum) + \'</td>\' +',
    '          \'<td class="col-word" onclick="kelimeOku(\\\'\' + escJs(x.kelime) + \'\\\', this)">\' + escHtml(x.kelime) + (x.ornek ? \'<div class="word-hover-ex">\' + ornekBicimlendir(x.ornek) + \'</div>\' : \'\') + \'</td>\' +',
    '          \'<td class="col-meaning">\' + escHtml(x.anlam) + \'</td>\' +',
    '          \'<td class="col-example">\' + escHtml(x.ornek) + \'</td>\' +',
    '          \'<td class="col-status"><span class="status-pill \' + x.durum + \'">\' + durumMetniGetir(x.durum) + \'</span></td>\' +',
    '        \'</tr>\'',
    '      );',
    '    }',
    '    document.getElementById("sheet1Body").innerHTML = s1Rows.join("") || \'<tr><td colspan="9" class="center" style="padding:14px;">No words found.</td></tr>\';',
    '',
    '    var sumLvls = ["A1", "A2", "B1", "B2", "C1", "C2"];',
    '    var sumHtml = [];',
    '    var knownSumHtml = [];',
    '    for (var k = 0; k < sumLvls.length; k++) {',
    '      var sl = sumLvls[k];',
    '      if (lvlCounts[sl] !== undefined || sl === "B1" || sl === "B2" || sl === "C1" || sl === "C2") {',
    '        sumHtml.push(\'<tr><td>\' + sl + \'</td><td class="center">\' + (lvlCounts[sl] || 0) + \'</td></tr>\');',
    '      }',
    '      if (knownLvlCounts[sl] !== undefined || sl === "B1" || sl === "B2" || sl === "C1" || sl === "C2") {',
    '        knownSumHtml.push(\'<tr><td>\' + sl + \'</td><td class="center">\' + (knownLvlCounts[sl] || 0) + \'</td></tr>\');',
    '      }',
    '    }',
    '    document.getElementById("poolSummaryBody").innerHTML = sumHtml.join("");',
    '    var kpSumEl = document.getElementById("knownPoolSummaryBody");',
    '    if (kpSumEl) kpSumEl.innerHTML = knownSumHtml.join("");',
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
    '          \'<td class="col-mark">\' + durumKutusuHtml(p.row, p.durum) + \'</td>\' +',
    '          \'<td class="col-word" onclick="kelimeOku(\\\'\' + escJs(p.kelime) + \'\\\', this)">\' + escHtml(p.kelime) + (p.ornek ? \'<div class="word-hover-ex">\' + ornekBicimlendir(p.ornek) + \'</div>\' : \'\') + \'</td>\' +',
    '          \'<td class="col-meaning">\' + escHtml(p.anlam) + \'</td>\' +',
    '          \'<td class="col-example">\' + escHtml(p.ornek) + \'</td>\' +',
    '          \'<td class="col-date">\' + escHtml(p.tarih) + \'</td>\' +',
    '        \'</tr>\'',
    '      );',
    '    }',
    '    document.getElementById("poolBody").innerHTML = pRows.join("") || \'<tr><td colspan="7" class="center" style="padding:14px;">No words in pool.</td></tr>\';',
    '',
    '    var knownPoolList = TUM_KELIMELER.filter(function(w){',
    '      if (w.durum !== "biliyorum") return false;',
    '      if (lvl !== "ALL" && w.seviye !== lvl) return false;',
    '      if (q && (w.kelime + " " + w.anlam + " " + w.ornek).toLowerCase().indexOf(q) === -1) return false;',
    '      return true;',
    '    }).sort(function(a, b){ return (a.seviye || "").localeCompare(b.seviye || ""); });',
    '',
    '    var kpRows = [];',
    '    for (var n = 0; n < knownPoolList.length; n++) {',
    '      var kp = knownPoolList[n];',
    '      kpRows.push(',
    '        \'<tr class="row-known">\' +',
    '          \'<td class="col-lvl">\' + escHtml(kp.seviye) + \'</td>\' +',
    '          \'<td class="col-type">\' + escHtml(turTemizle(kp.tur)) + \'</td>\' +',
    '          \'<td class="col-mark">\' + durumKutusuHtml(kp.row, kp.durum) + \'</td>\' +',
    '          \'<td class="col-word" onclick="kelimeOku(\\\'\' + escJs(kp.kelime) + \'\\\', this)">\' + escHtml(kp.kelime) + (kp.ornek ? \'<div class="word-hover-ex">\' + ornekBicimlendir(kp.ornek) + \'</div>\' : \'\') + \'</td>\' +',
    '          \'<td class="col-meaning">\' + escHtml(kp.anlam) + \'</td>\' +',
    '          \'<td class="col-example">\' + escHtml(kp.ornek) + \'</td>\' +',
    '          \'<td class="col-date">\' + escHtml(kp.tarih) + \'</td>\' +',
    '        \'</tr>\'',
    '      );',
    '    }',
    '    var kpBodyEl = document.getElementById("knownPoolBody");',
    '    if (kpBodyEl) kpBodyEl.innerHTML = kpRows.join("") || \'<tr><td colspan="7" class="center" style="padding:14px;">No words in known pool.</td></tr>\';',
    '',
    '    fcCiz(poolList);',
    '    istatistikTablosuCiz();',
    '  }',
    '',
    '  function durumKutusuHtml(row, durum) {',
    '    return \'<div class="mark-split \' + (durum || "beklemede") + \'"><span class="mark-half mark-top" onclick="durumSec(event, \' + row + \', \\\'biliyorum\\\')" title="Üst (✓): Biliyorum">✓</span><span class="mark-half mark-bot" onclick="durumSec(event, \' + row + \', \\\'bilmiyorum\\\')" title="Alt (✕): Bilmiyorum">✕</span></div>\';',
    '  }',
    '  function durumSec(e, row, tip) {',
    '    if (e) { e.preventDefault(); e.stopPropagation(); }',
    '    kutucukDegisti(row, tip, true);',
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
    '    var durumEtiket = yeniDurum === "biliyorum" ? "✓ Biliyorum" : (yeniDurum === "bilmiyorum" ? "✕ Bilmiyorum" : "○ Beklemede");',
    '    bildirimGoster(durumEtiket + ": " + hedef.kelime);',
    '    google.script.run.durumGuncelle(row, yeniDurum);',
    '  }',
    '',
    '  function ekleBarAcKapat() {',
    '    document.getElementById("addBar").classList.toggle("open");',
    '    tercihleriKaydet();',
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
    '  tercihleriYukleVeUygula();',
    '  verileriYukle();',
    '<\/script>',
    '</body>',
    '</html>'
  ].join('\n');
}
