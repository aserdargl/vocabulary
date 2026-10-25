const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const { test } = require('node:test');

const html = fs.readFileSync('index.html', 'utf8');
const script = html.match(/<script>([\s\S]*?)<\/script>/)[1];
const functions = script.match(/^    function \w+\([^\n]*\) \{[\s\S]*?^    \}/gm).join('\n');
const day = 24 * 60 * 60 * 1000;

function createApp(storage = new Map()) {
  const elements = new Map();
  const app = vm.createContext({
    localStorage: {
      getItem: key => storage.get(key) ?? null,
      setItem: (key, value) => storage.set(key, String(value)),
      removeItem: key => storage.delete(key)
    },
    document: { getElementById(id) {
      if (!elements.has(id)) elements.set(id, {
        value: id === 'levelFilter' ? 'ALL' : '', style: {},
        classList: { add() {}, remove() {} }
      });
      return elements.get(id);
    } },
    window: {}, setTimeout: () => 1, clearTimeout() {},
    TUM_KELIMELER: [], SHEET_KELIMELER: [], SON_YEREL_DEGISIM: {},
    MASTERED_REVIEW_DAYS: [7, 14, 30], REVIEW_LEARNING_BATCH: 3,
    REVIEW_DAY_MS: day, HEDEF_SAYI: 5, HEDEF_BIRIM: 'gun',
    FC_SHUFFLED_KEYS: null, FC_INDEX: 0, AKTIF_MOD: 'quiz',
    APPS_SCRIPT_URL: '', QUIZ_GECIS_TIMER: null, QUIZ_CORRECT_DELAY_MS: 1000
  });
  vm.runInContext(functions, app);
  // Keep the production progress, status, storage and answer paths; replace presentation.
  vm.runInContext(`
    var renderCount = 0;
    tabloyuCiz = function() { renderCount++; dortHaftaHatasizKontrolEt(); };
    istatistikTablosuCiz = bildirimGoster = kelimeOku = function() {};
  `, app);
  return app;
}

function addWord(app, level = 'B2', status = 'biliyorum', name = 'Recall') {
  const word = { row: app.TUM_KELIMELER.length + 2, kelime: name,
    seviye: level, durum: status, anlam: 'hatırlamak', ornek: '' };
  app.TUM_KELIMELER.push(word);
  app.SHEET_KELIMELER.push({ ...word });
  return word;
}

test('only due B2/C1/C2 mastered words enter practice; filters still apply', () => {
  const app = createApp();
  for (const level of ['B1', 'B2', 'C1', 'C2']) addWord(app, level, 'biliyorum', level);
  assert.deepEqual(Array.from(app.aktifHavuzuGetir(), w => w.kelime), ['B2', 'C1', 'C2']);
  app.istatistikMapKaydet({ b2: { nextReviewAt: Date.now() + day } });
  assert.deepEqual(Array.from(app.aktifHavuzuGetir(), w => w.kelime), ['C1', 'C2']);
  app.document.getElementById('levelFilter').value = 'C2';
  assert.equal(app.aktifHavuzuGetir()[0].kelime, 'C2');
  app.document.getElementById('searchInput').value = 'absent';
  assert.equal(app.aktifHavuzuGetir().length, 0);
});

test('correct reviews schedule 7, 14, then recurring 30 days and survive reload', () => {
  const storage = new Map();
  let app = createApp(storage);
  const word = addWord(app);
  for (const interval of [7, 14, 30, 30]) {
    const stats = app.istatistikMapAl();
    if (stats.recall) stats.recall.nextReviewAt = Date.now() - 1;
    app.istatistikMapKaydet(stats);
    const before = Date.now();
    app.alistirmaSonucuKaydet(word.kelime, true, 'Quiz');
    const entry = app.istatistikMapAl().recall;
    assert.equal(entry.reviewIntervalDays, interval);
    assert.ok(entry.nextReviewAt >= before + interval * day);
    assert.equal(word.durum, 'biliyorum');
    assert.equal(app.aktifHavuzuGetir().length, 0);
  }
  app = createApp(storage);
  addWord(app);
  assert.equal(app.aktifHavuzuGetir().length, 0);
});

test('early correct answers do not postpone scheduled review', () => {
  const app = createApp();
  const word = addWord(app);
  const nextReviewAt = Date.now() + 7 * day;
  app.istatistikMapKaydet({ recall: { nextReviewAt, reviewIntervalDays: 7 } });
  app.alistirmaSonucuKaydet(word.kelime, true, 'Twins');
  assert.equal(app.istatistikMapAl().recall.nextReviewAt, nextReviewAt);
});

for (const level of ['B2', 'C1', 'C2']) {
  test(`${level}: one error demotes, resets mastery, persists status and blocks immediate remastery`, () => {
    const app = createApp();
    const word = addWord(app, level);
    app.istatistikMapKaydet({ recall: { cleanDays: ['a', 'b', 'c', 'd', 'e'],
      cleanWeeks: ['a', 'b'], mastered4w: true, manualKnown: true,
      autoMasteredByStreak: true, nextReviewAt: Date.now() - 1, reviewIntervalDays: 30 } });
    app.alistirmaSonucuKaydet(word.kelime, false, 'Quiz');
    const entry = app.istatistikMapAl().recall;
    assert.equal(word.durum, 'bilmiyorum');
    assert.equal(app.SHEET_KELIMELER[0].durum, 'bilmiyorum');
    assert.equal(app.kisiselDurumlariAl().recall, 'bilmiyorum');
    assert.equal(entry.mistakes, 1);
    assert.equal(entry.mastered4w, false);
    assert.equal(entry.manualKnown, false);
    assert.equal(entry.cleanDays.length, 0);
    assert.equal(entry.cleanWeeks.length, 0);
    assert.equal(entry.nextReviewAt, undefined);
    assert.equal(app.renderCount, 0, 'answer feedback must stay visible');
    app.alistirmaSonucuKaydet(word.kelime, true, 'Quiz');
    app.dortHaftaHatasizKontrolEt();
    assert.equal(word.durum, 'bilmiyorum');
    assert.equal(app.fcDestesiniGetir(app.aktifHavuzuGetir()).length, 1);
  });
}

test('reviews are interleaved with learning, including a shuffled deck', () => {
  const app = createApp();
  addWord(app, 'B2', 'biliyorum', 'review1');
  addWord(app, 'C2', 'biliyorum', 'review2');
  for (let i = 1; i <= 5; i++) addWord(app, 'B2', 'bilmiyorum', 'learn' + i);
  app.FC_SHUFFLED_KEYS = ['learn5', 'learn4', 'learn3', 'learn2', 'learn1'];
  const names = Array.from(app.fcDestesiniGetir(app.aktifHavuzuGetir()), w => w.kelime);
  assert.deepEqual(names, ['review1', 'learn5', 'learn4', 'learn3', 'review2', 'learn2', 'learn1']);
});

test('new automatic and manual mastery start at seven days', () => {
  const app = createApp();
  const word = addWord(app, 'C1', 'bilmiyorum');
  app.HEDEF_SAYI = 1;
  app.alistirmaSonucuKaydet(word.kelime, true, 'Quiz');
  assert.equal(word.durum, 'biliyorum');
  assert.equal(app.istatistikMapAl().recall.reviewIntervalDays, 7);
  const manual = addWord(app, 'C2', 'bilmiyorum', 'Manual');
  app.kutucukDegisti(manual.row, 'biliyorum', true);
  assert.equal(app.istatistikMapAl().manual.reviewIntervalDays, 7);
});

for (const mode of ['Quiz', 'Cloze', 'Listen', 'Twins', 'Match', 'Flashcards']) {
  test(`${mode}: shared answer recording demotes a forgotten mastered word`, () => {
    const app = createApp();
    const word = addWord(app);
    app.alistirmaSonucuKaydet(word.kelime, false, mode);
    assert.equal(word.durum, 'bilmiyorum');
  });
}

for (const handler of ['quizSecenekSec', 'clozeSecenekSec', 'listenSecenekSec']) {
  test(`${handler}: wrong choice triggers relearning through the real handler`, () => {
    const app = createApp();
    const word = addWord(app);
    app.ikizKarsilastirmaKutusuHtml = () => '';
    app[handler]({ classList: { add() {} } }, 'wrong', word.kelime, word.row);
    assert.equal(word.durum, 'bilmiyorum');
  });
}

for (const [handler, input] of [['quizYazKontrol', 'quizTypeInput'],
  ['clozeYazKontrol', 'clozeTypeInput'], ['listenYazKontrol', 'listenTypeInput']]) {
  test(`${handler}: forgotten typed answer triggers relearning`, () => {
    const app = createApp();
    const word = addWord(app);
    app.ikizKarsilastirmaKutusuHtml = () => '';
    app.document.getElementById(input).value = 'wrong';
    app[handler](word.kelime, word.row);
    assert.equal(word.durum, 'bilmiyorum');
  });
}

test('flashcard forgotten action and incorrect match both trigger relearning', () => {
  const app = createApp();
  const word = addWord(app);
  app.tercihleriKaydet = () => {};
  app.fcHataIsaretle(word.row);
  assert.equal(word.durum, 'bilmiyorum');
  const match = addWord(app, 'C1', 'biliyorum', 'Matchword');
  app.MATCH_DURUM = { seciliEn: 'matchword', seciliTr: 'wrong' };
  app.matchKontrolEt(0);
  assert.equal(match.durum, 'bilmiyorum');
});

test('mastered-only quiz still has four distinct answer options', () => {
  const app = createApp();
  const word = addWord(app);
  for (const name of ['Other', 'Another', 'Fourth']) addWord(app, 'C1', 'biliyorum', name);
  app.window.AKILLI_BENZER_SIKLAR = false;
  app.SECENEKLER_CACHE = { key: '', list: [] };
  const options = app.secenekleriUret(word, app.aktifHavuzuGetir());
  assert.equal(new Set(Array.from(options, w => w.kelime)).size, 4);
});

test('flashcard Known action advances retention instead of restarting at seven days', () => {
  const app = createApp();
  const word = addWord(app);
  app.istatistikMapKaydet({ recall: { nextReviewAt: Date.now() - 1, reviewIntervalDays: 7 } });
  app.fcBiliyorumIsaretle(word.row);
  assert.equal(app.istatistikMapAl().recall.reviewIntervalDays, 14);
});

test('demotion sends the existing server status update as well as local persistence', () => {
  const app = createApp();
  const word = addWord(app);
  const requests = [];
  app.APPS_SCRIPT_URL = 'https://example.test';
  app.fetch = (url, options) => {
    requests.push(JSON.parse(options.body));
    return Promise.resolve({ json: () => Promise.resolve({ ok: true }) });
  };
  app.alistirmaSonucuKaydet(word.kelime, false, 'Quiz');
  assert.equal(requests.find(r => r.action === 'setStatus').status, 'bilmiyorum');
  assert.equal(requests.find(r => r.action === 'recordAttempt').isCorrect, false);
});

test('relearning can complete on a later day and resumes the seven-day schedule', () => {
  const app = createApp();
  const word = addWord(app);
  app.alistirmaSonucuKaydet(word.kelime, false, 'Quiz');
  const stats = app.istatistikMapAl();
  stats.recall.lastMistakeDate = '2000-01-01';
  stats.recall.lastMistakeWeek = '2000-W01';
  app.istatistikMapKaydet(stats);
  app.HEDEF_SAYI = 1;
  app.alistirmaSonucuKaydet(word.kelime, true, 'Quiz');
  assert.equal(word.durum, 'biliyorum');
  assert.equal(app.istatistikMapAl().recall.reviewIntervalDays, 7);
  assert.equal(app.istatistikMapAl().recall.mistakes, 1, 'retain lifetime mistake counts');
});

test('Apps Script serves exactly the updated HTML; both scripts parse', () => {
  const backend = fs.readFileSync('Code.gs', 'utf8');
  new vm.Script(script);
  const context = vm.createContext({});
  vm.runInContext(backend, context);
  assert.equal(context.uygulamaHtmlGetir(), html);
});

test('word sidebar uses exact first practice, lifetime errors and configured daily goal', () => {
  const app = createApp();
  const word = addWord(app, 'B2', 'bilmiyorum');
  const result = app.wordStudySummary(word, { firstPracticedOn: '2026-10-20',
    correct: 8, mistakes: 2, cleanDays: ['2026-10-24', '2026-10-25'] }, new Date(2026, 9, 25, 12));
  assert.equal(result.elapsed, 6);
  assert.equal(result.exactStart, true);
  assert.equal(result.remainingDays, 3);
  assert.equal(result.errorRate, 20);
});

test('word sidebar distinguishes unstarted words and incomplete historical dates', () => {
  const app = createApp();
  const word = addWord(app, 'B2', 'bilmiyorum');
  const fresh = app.wordStudySummary(word, {}, new Date(2026, 9, 25));
  assert.equal(fresh.elapsed, 0);
  assert.equal(fresh.errorRate, null);
  const history = app.wordStudySummary(word, { correct: 2, mistakes: 1,
    lastPractice: 'QUIZ (2026-10-22 12:00)' }, new Date(2026, 9, 25));
  assert.equal(history.elapsed, 4);
  assert.equal(history.exactStart, false);
  assert.equal(history.errorRate, 33.3);
});

test('weekly sidebar countdown follows ISO weeks including year boundaries', () => {
  const app = createApp();
  app.HEDEF_BIRIM = 'hafta';
  app.HEDEF_SAYI = 2;
  const word = addWord(app, 'B2', 'bilmiyorum');
  const sunday = new Date(2027, 0, 3, 12);
  const week = app.haftaAnahtariAl(sunday);
  const result = app.wordStudySummary(word, { cleanWeeks: [week], correct: 1 }, sunday);
  assert.equal(result.remainingDays, 1);
  const reset = app.wordStudySummary(word, { lastMistakeWeek: week, mistakes: 1 }, sunday);
  assert.equal(reset.remainingDays, 8);
});

test('mastered sidebar has no remaining learning days and does not mutate history', () => {
  const app = createApp();
  const word = addWord(app);
  const entry = { correct: 1, lastPractice: 'QUIZ (2026-10-25 12:00)' };
  const original = JSON.stringify(entry);
  const result = app.wordStudySummary(word, entry, new Date(2026, 9, 25));
  assert.equal(result.remainingDays, 0);
  assert.equal(result.mastered, true);
  assert.equal(JSON.stringify(entry), original);
});

test('first practice date persists through mistakes and reload without inventing old start dates', () => {
  const storage = new Map();
  let app = createApp(storage);
  const word = addWord(app, 'B2', 'bilmiyorum');
  app.alistirmaSonucuKaydet(word.kelime, false, 'Quiz');
  const first = app.istatistikMapAl().recall.firstPracticedOn;
  assert.match(first, /^\d{4}-\d{2}-\d{2}$/);
  app = createApp(storage);
  addWord(app, 'B2', 'bilmiyorum');
  app.alistirmaSonucuKaydet(word.kelime, true, 'Flashcards');
  assert.equal(app.istatistikMapAl().recall.firstPracticedOn, first);
  app.istatistikMapKaydet({ recall: { mistakes: 4, correct: 3 } });
  app.alistirmaSonucuKaydet(word.kelime, false, 'Quiz');
  assert.equal(app.istatistikMapAl().recall.firstPracticedOn, undefined);
});

test('Apps Script serves the same updated interface as index.html', () => {
  const source = fs.readFileSync('Code.gs', 'utf8');
  const encoded = source.match(/function uygulamaHtmlGetir\(\) \{\s*return ("(?:[^"\\]|\\.)*");/)[1];
  assert.equal(JSON.parse(encoded), html);
  new vm.Script(script);
});

test('sidebar follows the current word, updates answers immediately and hides in other modes', () => {
  const app = createApp();
  const word = addWord(app, 'B2', 'bilmiyorum', 'SecretAnswer');
  const panel = app.document.getElementById('wordStudyStats');
  panel.dataset = {};
  app.updateWordStudyStats(word);
  assert.equal(panel.hidden, false);
  assert.match(panel.innerHTML, /No answers yet/);
  assert.doesNotMatch(panel.innerHTML, /SecretAnswer/);
  app.alistirmaSonucuKaydet(word.kelime, false, 'Quiz');
  assert.match(panel.innerHTML, /100%/);
  assert.match(panel.innerHTML, /0 correct \/ 1 wrong/);
  const next = addWord(app, 'B2', 'bilmiyorum', 'Next');
  app.updateWordStudyStats(next);
  assert.match(panel.innerHTML, /No answers yet/);
  assert.equal(panel.dataset.wordKey, 'next');
  app.AKTIF_MOD = 'match';
  app.updateWordStudyStats(next);
  assert.equal(panel.hidden, true);
  app.AKTIF_MOD = 'flashcards';
  app.updateWordStudyStats(word);
  assert.equal(panel.hidden, false);
  app.updateWordStudyStats(undefined);
  assert.equal(panel.hidden, true);
});
