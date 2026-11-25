const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const { test } = require('node:test');

const html = fs.readFileSync('index.html', 'utf8');
const script = html.match(/<script>([\s\S]*?)<\/script>/)[1];
const functions = script.match(/^    function \w+\([^\n]*\) \{[\s\S]*?^    \}/gm).join('\n');
const backend = fs.readFileSync('Code.gs', 'utf8');
const day = 86400000;
const initialTime = +new Date(2026, 0, 1, 12);

function createApp(storage = new Map(), startTime = initialTime) {
  let now = startTime;
  const elements = new Map();
  const timers = new Map();
  let timerId = 0;
  const app = vm.createContext({
    Date: class extends Date {
      constructor(...args) { super(...(args.length ? args : [now])); }
      static now() { return now; }
    },
    localStorage: {
      getItem: key => storage.get(key) ?? null,
      setItem: (key, value) => storage.set(key, String(value)),
      removeItem: key => storage.delete(key)
    },
    document: {
      hidden: false,
      body: { classList: { add() {}, remove() {} } },
      getElementById(id) {
        if (!elements.has(id)) elements.set(id, {
          value: id === 'levelFilter' ? 'ALL' : '', style: {},
          classList: { add() {}, remove() {}, toggle() {}, contains: () => false },
          setAttribute() {}, querySelectorAll: () => []
        });
        return elements.get(id);
      }
    },
    window: {},
    setTimeout: (fn, delay) => { timers.set(++timerId, { fn, delay }); return timerId; },
    clearTimeout: id => timers.delete(id),
    TUM_KELIMELER: [], SHEET_KELIMELER: [], SON_YEREL_DEGISIM: {},
    MASTERED_REVIEW_DAYS: [30], REVIEW_LEARNING_BATCH: 3,
    REVIEW_DAY_MS: day, HEDEF_SAYI: 4, HEDEF_BIRIM: 'hafta',
    FC_SHUFFLED_KEYS: null, FC_INDEX: 0, FC_FLIPPED: false, FC_OKUNAN_KART: '',
    MOD_KONUMLARI: {}, DENGELI_DESTE_TUR_ADIM: 0, DENGELI_DESTE_MIGRASYON_GEREKLI: false,
    QUIZ_BEKLEYEN_ILERLEME: null, SON_CIZILEN_SORU_ANAHTARI: '',
    AKTIF_MOD: 'quiz', AKTIF_SEKME: 'sheet1', SES_ACIK: true,
    YAZI_BOYUTU: 13, BILINENLERI_GOSTER: true,
    IKIZ_SOZLUK_HARITASI: {}, CANLI_SENKRON_ARALIGI_MS: 60000, SON_CANLI_SENKRON: 0,
    SECENEKLER_CACHE: { key: '', list: [] }, CALISMA_CEVAPLARI: {},
    APPS_SCRIPT_URL: '', QUIZ_GECIS_TIMER: null, QUIZ_CORRECT_DELAY_MS: 30000,
    OGRENME_SENKRON_TIMER: null, MATCH_DURUM: { seciliEn: null, seciliTr: null, eslesenler: {}, trSira: [], keySet: '' },
    PHRASAL_VERBS_LIST: [], PV_MOD: 'list', PV_INDEX: 0, PV_FLIPPED: false,
    PV_SHUFFLED_KEYS: null, PV_GECIS_TIMER: null, PV_SECENEKLER_CACHE: { key: '', list: [] },
    PV_OKUNAN_KART: '', PV_SON_SORU_ANAHTARI: '', PV_PRATIK_TUMU: false, PV_VERB_FORMS: {},
    PV_HAVUZ_SECILENLER: {}, PV_STARTER_COUNT: 15
  });
  vm.runInContext(functions, app);
  app.renderCount = 0;
  app.tabloyuCiz = () => { app.renderCount++; app.dortHaftaHatasizKontrolEt(); };
  app.istatistikTablosuCiz = app.bildirimGoster = app.kelimeOku = () => {};
  app.advance = days => { now += days * day; };
  app.runTimer = delay => {
    const entry = [...timers].find(([, t]) => t.delay === delay);
    if (!entry) throw new Error('Missing timer ' + delay);
    timers.delete(entry[0]); entry[1].fn();
  };
  return app;
}

function addWord(app, status = 'bilmiyorum', name = 'Recall', level = 'B2') {
  const word = { row: app.TUM_KELIMELER.length + 2, kelime: name, seviye: level,
    tarih: '', tur: 'Verb', okunus: '', durum: status, anlam: 'hatırlamak', ornek: '',
    biliyorum: status === 'biliyorum', bilmiyorum: status === 'bilmiyorum' };
  app.TUM_KELIMELER.push(word);
  app.SHEET_KELIMELER.push({ ...word });
  return word;
}

function learnedWord(app, name = 'Recall', level = 'B2') {
  const word = addWord(app, 'ogrenildi', name, level);
  const map = app.istatistikMapAl();
  map[name.toLowerCase()] = { learningVersion: 2, learningUpdatedAt: app.Date.now(),
    learned: true, manualKnown: false, firstSeenOn: '2025-11-01', trainingStartedOn: '2025-11-01',
    correctDuringGoal: 5, correctAfterGoal: 1, correctDays: ['2025-11-01', '2025-11-02', '2025-11-03'],
    consecutiveCorrect: 3, correct: 5, mistakes: 0,
    nextReviewAt: app.Date.now() - 1, reviewIntervalDays: 30 };
  app.istatistikMapKaydet(map);
  return word;
}

function studyApp(mode = 'flashcards') {
  const app = createApp();
  app.AKTIF_MOD = mode;
  app.AKTIF_SEKME = 'flashcards';
  app.spoken = [];
  app.kelimeOku = word => {
    assert.ok(app.document.getElementById('fcContainer').innerHTML.includes(word));
    app.spoken.push(word);
  };
  app.tabloyuCiz = () => { app.renderCount++; app.dortHaftaHatasizKontrolEt(); app.fcCiz(app.aktifHavuzuGetir()); };
  app.tercihleriKaydet = app.mobilMenuKapat = app.sekmeButonlariniGuncelle = () => {};
  return app;
}

function answer(app, word, correct = true) {
  if (app.AKTIF_SEKME === 'flashcards') app.fcCiz(app.aktifHavuzuGetir());
  app.alistirmaSonucuKaydet(word.kelime, correct, 'Quiz');
}

test('first display starts the Goal before any answer; imported words and distractors do not', () => {
  const app = studyApp('quiz');
  addWord(app); addWord(app, 'beklemede', 'Other');
  assert.deepEqual(Object.keys(app.istatistikMapAl()), []);
  app.tabloyuCiz();
  const stats = app.istatistikMapAl();
  assert.equal(stats.recall.firstSeenOn, '2026-01-01');
  assert.equal(stats.recall.trainingStartedOn, '2026-01-01');
  assert.equal(stats.recall.correct || 0, 0);
  assert.equal(stats.other, undefined);
});

test('first-seen date survives reload, later answers and mistakes', () => {
  const storage = new Map();
  let app = createApp(storage);
  let word = addWord(app);
  app.kelimeGosteriminiKaydet(word);
  app = createApp(storage, initialTime + 35 * day);
  word = addWord(app);
  app.kelimeGosteriminiKaydet(word);
  answer(app, word, false);
  assert.equal(app.istatistikMapAl().recall.firstSeenOn, '2026-01-01');
});

test('same-day cramming never graduates; 3 distinct days during Goal require Goal completion plus 1 final check', () => {
  const app = createApp(); const word = addWord(app);
  for (let i = 0; i < 8; i++) answer(app, word);
  assert.equal(word.durum, 'bilmiyorum');
  assert.equal(app.istatistikMapAl().recall.correctDays.length, 1);
  app.advance(1); answer(app, word);
  app.advance(1); answer(app, word);
  assert.equal(app.istatistikMapAl().recall.correctDays.length, 3);
  app.advance(25); app.dortHaftaHatasizKontrolEt();
  assert.equal(word.durum, 'bilmiyorum');
  app.advance(1); app.dortHaftaHatasizKontrolEt();
  // Goal ended (day 28), but still requires 1 final check after Goal
  assert.equal(word.durum, 'bilmiyorum');
  answer(app, word);
  assert.equal(word.durum, 'ogrenildi');
  assert.equal(app.istatistikMapAl().recall.nextReviewAt, app.Date.now() + 30 * day);
  assert.equal(app.kisiselDurumlariAl().recall, 'ogrenildi');
});

test('4 weeks with no answers: after Goal ends, requires correct answers on 3 different days', () => {
  const app = createApp(); const word = addWord(app);
  app.kelimeGosteriminiKaydet(word);
  app.advance(35); app.dortHaftaHatasizKontrolEt();
  assert.equal(word.durum, 'bilmiyorum');
  // Multiple answers on the same day only count as 1 distinct day
  answer(app, word); answer(app, word); answer(app, word);
  assert.equal(word.durum, 'bilmiyorum');
  assert.equal(app.istatistikMapAl().recall.correctDays.length, 1);
  app.advance(1); answer(app, word);
  assert.equal(word.durum, 'bilmiyorum');
  assert.equal(app.istatistikMapAl().recall.correctDays.length, 2);
  app.advance(1); answer(app, word);
  assert.equal(word.durum, 'ogrenildi');
  assert.equal(app.istatistikMapAl().recall.correctDays.length, 3);
});

test('2 distinct days during Goal plus 1 final check on a 3rd day after Goal graduates', () => {
  const app = createApp(); const word = addWord(app);
  answer(app, word);
  app.advance(2); answer(app, word);
  app.advance(26);
  assert.equal(word.durum, 'bilmiyorum');
  answer(app, word);
  assert.equal(word.durum, 'ogrenildi');
});

test('wrong answer resets consecutive streak and final check; requires 2 consecutive correct answers to graduate', () => {
  const app = createApp(); const word = addWord(app);
  answer(app, word);
  app.advance(1); answer(app, word);
  app.advance(1); answer(app, word);
  app.advance(26);
  answer(app, word, false);
  assert.equal(app.istatistikMapAl().recall.consecutiveCorrect, 0);
  assert.equal(app.istatistikMapAl().recall.correctAfterGoal, 0);
  answer(app, word);
  assert.equal(word.durum, 'bilmiyorum');
  assert.equal(app.istatistikMapAl().recall.consecutiveCorrect, 1);
  answer(app, word);
  assert.equal(word.durum, 'ogrenildi');
});

test('a recent error during Goal prevents graduation on first check after Goal and requires 2 consecutive correct answers', () => {
  const app = createApp(); const word = addWord(app);
  answer(app, word);
  app.advance(1); answer(app, word);
  app.advance(1); answer(app, word);
  answer(app, word, false);
  app.advance(26); app.dortHaftaHatasizKontrolEt();
  assert.equal(word.durum, 'bilmiyorum');
  answer(app, word);
  assert.equal(word.durum, 'bilmiyorum');
  answer(app, word);
  assert.equal(word.durum, 'ogrenildi');
});

test('day Goal uses calendar days; changing Goal duration does not rewrite first-seen date', () => {
  const app = createApp(); const word = addWord(app);
  app.HEDEF_BIRIM = 'gun'; app.HEDEF_SAYI = 2;
  answer(app, word);
  app.advance(1); answer(app, word);
  app.dortHaftaHatasizKontrolEt(); assert.equal(word.durum, 'bilmiyorum');
  app.HEDEF_SAYI = 3;
  app.advance(1); answer(app, word);
  assert.equal(word.durum, 'bilmiyorum');
  app.advance(1);
  assert.equal(word.durum, 'bilmiyorum');
  answer(app, word);
  assert.equal(word.durum, 'ogrenildi');
  assert.equal(app.istatistikMapAl().recall.firstSeenOn, '2026-01-01');
});

for (const mode of ['flashcards', 'quiz', 'cloze', 'listen', 'confusables', 'match']) {
  test(mode + ': manual Known words stay excluded while due Learned words are tested', () => {
    const app = createApp(); app.AKTIF_MOD = mode;
    addWord(app, 'biliyorum', 'Manual');
    const learned = learnedWord(app);
    addWord(app, 'beklemede', 'Daily');
    assert.deepEqual(Array.from(app.aktifHavuzuGetir(), w => w.kelime), ['Recall', 'Daily']);
    if (mode === 'confusables') assert.ok(!Array.from(app.confusablesDestesiniGetir(), w => w.kelime).includes('Manual'));
    app.alistirmaSonucuKaydet(learned.kelime, false, mode);
    assert.equal(learned.durum, 'bilmiyorum');
    assert.equal(app.istatistikMapAl().recall.learned, false);
  });
}

test('learned words of all levels use recurring 30-day monthly retention intervals', () => {
  for (const level of ['A1', 'B1', 'B2', 'C1', 'C2']) {
    const app = createApp(); const word = learnedWord(app, 'Recall', level);
    for (const interval of [30, 30, 30]) {
      const stats = app.istatistikMapAl(); stats.recall.nextReviewAt = app.Date.now() - 1; app.istatistikMapKaydet(stats);
      answer(app, word);
      assert.equal(app.istatistikMapAl().recall.nextReviewAt, app.Date.now() + interval * day);
      assert.equal(word.durum, 'ogrenildi');
      assert.equal(app.aktifHavuzuGetir().length, 0);
    }
  }
});

test('an early correct review never postpones the next review', () => {
  const app = createApp(); const word = learnedWord(app);
  const stats = app.istatistikMapAl(); stats.recall.nextReviewAt = app.Date.now() + 30 * day; app.istatistikMapKaydet(stats);
  answer(app, word);
  assert.equal(app.istatistikMapAl().recall.nextReviewAt, app.Date.now() + 30 * day);
});

test('forgotten Learned word restarts a full Goal from scratch while keeping its original first-seen and lifetime counts', () => {
  const app = createApp(); const word = learnedWord(app);
  answer(app, word, false);
  const stats = app.istatistikMapAl().recall;
  assert.equal(word.durum, 'bilmiyorum');
  assert.equal(stats.trainingStartedOn, '2026-01-01');
  assert.equal(stats.firstSeenOn, '2025-11-01');
  assert.deepEqual(Array.from(stats.correctDays), []);
  assert.equal(stats.consecutiveCorrect, 0);
  assert.equal(stats.correct, 5); assert.equal(stats.mistakes, 1);
  assert.equal(stats.nextReviewAt, undefined);
  answer(app, word);
  app.advance(1); answer(app, word);
  app.advance(1); answer(app, word);
  assert.equal(word.durum, 'bilmiyorum');
  app.advance(26); app.dortHaftaHatasizKontrolEt();
  assert.equal(word.durum, 'bilmiyorum');
  answer(app, word);
  assert.equal(word.durum, 'ogrenildi');
  assert.equal(app.istatistikMapAl().recall.reviewIntervalDays, 30);
});

test('manual Known action is separate from learning and is never demoted by a stray exercise callback', () => {
  const app = createApp(); const word = addWord(app);
  app.fcBiliyorumIsaretle(word.row);
  assert.equal(word.durum, 'biliyorum');
  assert.equal(app.istatistikMapAl().recall.manualKnown, true);
  assert.equal(app.istatistikMapAl().recall.learned, false);
  assert.equal(app.istatistikMapAl().recall.nextReviewAt, undefined);
  answer(app, word, false); assert.equal(word.durum, 'biliyorum');
});

test('legacy automatic mastery migrates to Learned while explicit manual Known stays excluded', () => {
  const app = createApp(); const auto = addWord(app, 'biliyorum'); addWord(app, 'biliyorum', 'Manual');
  app.istatistikMapKaydet({ recall: { autoMasteredByStreak: true, firstPracticedOn: '2025-10-01' }, manual: { manualKnown: true } });
  app.dortHaftaHatasizKontrolEt();
  assert.equal(auto.durum, 'ogrenildi');
  assert.equal(app.TUM_KELIMELER[1].durum, 'biliyorum');
  assert.equal(app.istatistikMapAl().recall.firstSeenOn, '2025-10-01');
});

test('due retention checks are interleaved with learning in a saved shuffled deck', () => {
  const app = createApp(); learnedWord(app, 'Review1'); learnedWord(app, 'Review2');
  for (let i = 1; i <= 5; i++) addWord(app, 'bilmiyorum', 'Learn' + i);
  app.FC_SHUFFLED_KEYS = ['learn5', 'learn4', 'learn3', 'learn2', 'learn1'];
  assert.deepEqual(Array.from(app.fcDestesiniGetir(app.aktifHavuzuGetir()), w => w.kelime),
    ['Review1', 'Learn5', 'Learn4', 'Learn3', 'Review2', 'Learn2', 'Learn1']);
});

for (const mode of ['flashcards', 'quiz']) {
  test(mode + ': new daily words update a saved deck without interrupting the current word or repeating audio', () => {
    const app = studyApp(mode); const first = addWord(app, 'bilmiyorum', 'First'); const current = addWord(app, 'bilmiyorum', 'Current', 'C1');
    const manual = addWord(app, 'biliyorum', 'Manual'); app.FC_SHUFFLED_KEYS = ['first', 'current', 'manual']; app.FC_INDEX = 1;
    app.tabloyuCiz();
    const daily = { row: 5, kelime: 'Daily', seviye: 'B1', durum: 'beklemede', anlam: 'günlük' };
    app.listeyiBirlestirVeHazirla([first, current, manual, daily], true);
    assert.deepEqual(Array.from(app.fcDestesiniGetir(app.aktifHavuzuGetir()), w => w.kelime), ['First', 'Current', 'Daily']);
    assert.equal(app.fcDestesiniGetir(app.aktifHavuzuGetir())[app.FC_INDEX].kelime, 'Current');
    assert.deepEqual(app.spoken, ['Current']); app.fcDegistir(1); assert.equal(app.spoken.at(-1), 'Daily');
  });
  test(mode + ': first word and navigation speak immediately and respect mute', () => {
    const app = studyApp(mode); addWord(app, 'bilmiyorum', 'First'); addWord(app, 'beklemede', 'Second');
    app.tabloyuCiz(); app.tabloyuCiz(); assert.deepEqual(app.spoken, ['First']);
    app.fcDegistir(1); assert.deepEqual(app.spoken, ['First', 'Second']);
    app.SES_ACIK = false; app.fcDegistir(-1); assert.equal(app.spoken.length, 2);
  });
}

test('unchanged background sync preserves typed answers and feedback', () => {
  const app = studyApp('quiz'); const word = addWord(app);
  app.listeyiBirlestirVeHazirla([word], true); app.document.getElementById('fcContainer').innerHTML += '<div>feedback</div>';
  app.document.getElementById('quizTypeInput').value = 'partial'; const count = app.renderCount;
  app.listeyiBirlestirVeHazirla([word], true);
  assert.equal(app.renderCount, count); assert.match(app.document.getElementById('fcContainer').innerHTML, /feedback/);
  assert.equal(app.document.getElementById('quizTypeInput').value, 'partial');
});

test('repeated clicks or typing in one question count once; another encounter counts again', () => {
  const app = studyApp('quiz'); const word = addWord(app); app.tabloyuCiz();
  app.alistirmaSonucuKaydet(word.kelime, true, 'Quiz'); app.alistirmaSonucuKaydet(word.kelime, true, 'Quiz');
  assert.equal(app.istatistikMapAl().recall.correctDuringGoal, 1);
  app.fcDegistir(1); app.alistirmaSonucuKaydet(word.kelime, true, 'Quiz');
  assert.equal(app.istatistikMapAl().recall.correctDuringGoal, 2);
});

test('answer revealed after an error cannot become a successful confirmation in the same encounter', () => {
  const app = studyApp('quiz'); const word = addWord(app); app.tabloyuCiz(); app.advance(35);
  app.alistirmaSonucuKaydet(word.kelime, false, 'Quiz'); app.alistirmaSonucuKaydet(word.kelime, true, 'Quiz');
  assert.equal(app.istatistikMapAl().recall.correctAfterGoal, 0);
});

for (const handler of ['quizSecenekSec', 'clozeSecenekSec', 'listenSecenekSec']) {
  test(handler + ': a wrong answer returns Learned to training', () => {
    const app = createApp(); const word = learnedWord(app); app.ikizKarsilastirmaKutusuHtml = () => '';
    app[handler]({ classList: { add() {} } }, 'wrong', word.kelime, word.row); assert.equal(word.durum, 'bilmiyorum');
  });
}
for (const [handler, input] of [['quizYazKontrol','quizTypeInput'], ['clozeYazKontrol','clozeTypeInput'], ['listenYazKontrol','listenTypeInput']]) {
  test(handler + ': a wrong typed answer returns Learned to training', () => {
    const app = createApp(); const word = learnedWord(app); app.ikizKarsilastirmaKutusuHtml = () => '';
    app.document.getElementById(input).value = 'wrong'; app[handler](word.kelime, word.row); assert.equal(word.durum, 'bilmiyorum');
  });
}

test('quiz smart and random options exclude manual Known words and use unmarked daily words', () => {
  const app = createApp(); const target = addWord(app); addWord(app, 'biliyorum', 'Recalled');
  for (const name of ['Recalling','Recalls','Recorder']) addWord(app, 'beklemede', name);
  for (const smart of [true, false]) {
    app.window.AKILLI_BENZER_SIKLAR = smart; app.SECENEKLER_CACHE = { key: '', list: [] };
    const options = Array.from(app.secenekleriUret(target, app.aktifHavuzuGetir()), w => w.kelime);
    assert.equal(new Set(options).size, 4); assert.ok(!options.includes('Recalled'));
  }
});

test('learning progress shows first-seen, mandatory duration, and separate statuses in the interface', () => {
  const app = createApp(); const word = addWord(app); app.kelimeGosteriminiKaydet(word);
  vm.runInContext(functions.match(/^    function istatistikTablosuCiz\([^\n]*\) \{[\s\S]*?^    \}/m)[0], app);
  app.istatistikTablosuCiz();
  assert.match(app.document.getElementById('statsBody').innerHTML, /01\.01\.2026/);
  assert.match(app.document.getElementById('statsBody').innerHTML, /0\/28 days/);
  assert.match(html, /id="learnedPoolView"/); assert.match(html, /id="knownPoolView"/);
});

test('newer server state wins without restoring deleted review fields or overwriting the earliest first-seen', () => {
  const app = createApp();
  app.istatistikMapKaydet({ recall: { learningVersion: 2, learningUpdatedAt: 10, firstSeenOn: '2025-01-01', learned: true, nextReviewAt: 999 } });
  app.sunucuIstatistikleriniBirlestir({ recall: { learningVersion: 2, learningUpdatedAt: 20, firstSeenOn: '2026-01-01', learned: false, trainingStartedOn: '2026-01-01' } });
  const st = app.istatistikMapAl().recall;
  assert.equal(st.learned, false); assert.equal(st.nextReviewAt, undefined); assert.equal(st.firstSeenOn, '2025-01-01');
});

test('live refresh is throttled to a minute and waits while hidden', () => {
  const app = createApp(); let syncs = 0;
  app.verileriYukle = () => { syncs++; app.SON_CANLI_SENKRON = app.Date.now(); };
  app.canliVerileriGuncelle(); app.canliVerileriGuncelle(); assert.equal(syncs, 1);
  app.advance(1); app.document.hidden = true; app.canliVerileriGuncelle(); assert.equal(syncs, 1);
  app.document.hidden = false; app.canliVerileriGuncelle(); assert.equal(syncs, 2);
});

test('an obsolete audio failure cannot speak over the current word', async () => {
  const app = createApp(); vm.runInContext(functions.match(/^    function kelimeOku\([^\n]*\) \{[\s\S]*?^    \}/m)[0], app);
  const requests = []; const spoken = [];
  app.Audio = class { constructor() { requests.push(this); } play() { return new Promise((_, reject) => { this.fail = reject; }); } pause() {} };
  app.AKTIF_SES = null; app.tarayiciIleOku = word => spoken.push(word);
  app.kelimeOku('First'); app.kelimeOku('Second'); requests[0].fail(); await Promise.resolve(); assert.deepEqual(spoken, []);
  requests[1].fail(); await Promise.resolve(); assert.deepEqual(spoken, ['Second']);
});

test('the Apps Script interface matches index.html and both scripts parse', () => {
  new vm.Script(script); const app = vm.createContext({}); vm.runInContext(backend, app);
  assert.equal(app.uygulamaHtmlGetir(), html);
});

function createServer() {
  const sheets = new Map();
  class Sheet {
    constructor(name) { this.name = name; this.rows = []; }
    getLastRow() { let n = this.rows.length; while (n && !(this.rows[n - 1] || []).some(v => v !== '' && v !== undefined)) n--; return n; }
    getLastColumn() { return 12; }
    getName() { return this.name; }
    setFrozenRows() {} setColumnWidth() {} hideColumns() {}
    getRange(row, col, height = 1, width = 1) {
      const sheet = this;
      const range = {
        getValues: () => Array.from({length: height}, (_, y) => Array.from({length: width}, (_, x) => sheet.rows[row + y - 1]?.[col + x - 1] ?? '')),
        getDisplayValues: () => range.getValues().map(r => r.map(v => typeof v === 'boolean' ? String(v).toUpperCase() : String(v))),
        getDisplayValue: () => range.getDisplayValues()[0][0],
        getValue: () => range.getValues()[0][0],
        setValues(values) {
          assert.equal(values.length, height); values.forEach((r, y) => {
            assert.equal(r.length, width); sheet.rows[row + y - 1] ||= [];
            r.forEach((value, x) => { sheet.rows[row + y - 1][col + x - 1] = value; });
          }); return range;
        },
        setValue: value => range.setValues([[value]]),
        setFormulas: values => range.setValues(values),
        clearContent: () => range.setValues(Array.from({length: height}, () => Array(width).fill(''))),
        setBackground: () => range, setFontColor: () => range, setFontWeight: () => range
      };
      return range;
    }
  }
  const ss = { getSheetByName: name => sheets.get(name) || null,
    insertSheet: name => { const sheet = new Sheet(name); sheets.set(name, sheet); return sheet; },
    getSheets: () => [...sheets.values()], getUrl: () => 'https://example.test/sheet' };
  const words = ss.insertSheet('Words');
  words.rows = [[], ['', '2026-01-01', 'Recall', '', '', 'Verb', 'B2', 'hatırlamak', '', '', false, true]];
  const server = vm.createContext({ SpreadsheetApp: { getActiveSpreadsheet: () => ss, flush() {} },
    LockService: { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) },
    Utilities: { formatDate: () => '2026-01-01 12:00' }, Session: { getScriptTimeZone: () => 'Europe/Istanbul' } });
  vm.runInContext(backend, server);
  server.getSheetLayout_ = () => ({startRow:2, date:2, word:3, pron:4, audio:5, type:6, level:7, meaning:8, example:9, status:10, known:11, unknown:12});
  return { server, ss, words };
}

test('Google Sheets persists first display, Goal progress, learned status and review dates across a fresh client', () => {
  const {server, ss} = createServer(); const client = createApp(); const word = addWord(client);
  client.kelimeGosteriminiKaydet(word);
  const first = server.istatistikTopluSenkronizeEt(client.istatistikMapAl()).stats.recall;
  assert.equal(first.firstSeenOn, '2026-01-01'); assert.equal(first.trainingStartedOn, '2026-01-01');
  answer(client, word);
  client.advance(1); answer(client, word);
  client.advance(1); answer(client, word);
  client.advance(26); answer(client, word);
  server.istatistikTopluSenkronizeEt(client.istatistikMapAl());
  const fresh = createApp();
  fresh.sunucuIstatistikleriniBirlestir(server.istatistikleriOku_(ss));
  fresh.listeyiBirlestirVeHazirla([{...word, durum:'beklemede', biliyorum:false}], true);
  const state = fresh.istatistikMapAl().recall;
  assert.equal(state.learned, true); assert.equal(state.correctDuringGoal, 3);
  assert.equal(state.correctDays.length, 4); assert.equal(state.consecutiveCorrect, 4);
  assert.equal(state.firstSeenOn, '2026-01-01'); assert.equal(state.reviewIntervalDays, 30);
  assert.equal(fresh.TUM_KELIMELER[0].durum, 'ogrenildi');
  assert.equal(ss.getSheetByName('Mistake Stats').rows[1][10], '2026-01-01');
});

test('Google Sheets rejects stale learning snapshots and removes old review fields after demotion', () => {
  const {server, ss} = createServer();
  const learned = {learningVersion:2, learningUpdatedAt:10, learned:true, firstSeenOn:'2025-11-01', nextReviewAt:999, correct:5};
  server.istatistikTopluSenkronizeEt({recall:learned});
  server.istatistikTopluSenkronizeEt({recall:{learningVersion:2, learningUpdatedAt:20, learned:false, firstSeenOn:'2025-11-01', trainingStartedOn:'2026-01-01', mistakes:1}});
  server.istatistikTopluSenkronizeEt({recall:learned});
  const state = server.istatistikleriOku_(ss).recall;
  assert.equal(state.learned, false); assert.equal(state.nextReviewAt, undefined);
  assert.equal(state.correct, 5); assert.equal(state.mistakes, 1);
});

test('retrying the same answer request cannot inflate Google Sheets counters', () => {
  const {server,ss} = createServer();
  const entry = {learningVersion:2, learningUpdatedAt:1, firstSeenOn:'2026-01-01', trainingStartedOn:'2026-01-01', correctDuringGoal:1, correct:1, mistakes:0};
  server.istatistikKaydet('Recall', true, 'Quiz', entry);
  server.istatistikKaydet('Recall', true, 'Quiz', entry);
  const saved = server.istatistikleriOku_(ss).recall;
  assert.equal(saved.correct, 1); assert.equal(saved.correctDuringGoal, 1);
});

test('atomic status changes keep manual Known and Learned separate and reject delayed status writes', () => {
  const {server,ss,words} = createServer();
  server.durumGuncelle(2, 'ogrenildi', {learningVersion:2, learningUpdatedAt:10, learned:true, manualKnown:false, correct:5});
  assert.equal(words.rows[1][10], false); assert.equal(words.rows[1][11], false);
  assert.equal(server.verileriGetir().words[0].durum, 'ogrenildi');
  server.durumGuncelle(2, 'bilmiyorum', {learningVersion:2, learningUpdatedAt:20, learned:false, manualKnown:false, mistakes:1});
  server.durumGuncelle(2, 'ogrenildi', {learningVersion:2, learningUpdatedAt:10, learned:true});
  assert.equal(words.rows[1][11], true); assert.equal(server.istatistikleriOku_(ss).recall.learned, false);
  server.durumGuncelle(2, 'biliyorum', {learningVersion:2, learningUpdatedAt:30, learned:false, manualKnown:true});
  assert.equal(words.rows[1][10], true); assert.equal(server.verileriGetir().words[0].durum, 'biliyorum');
});

test('first display is scheduled for server sync even before the first answer', async () => {
  const app = studyApp(); addWord(app); const requests = [];
  app.APPS_SCRIPT_URL = 'https://example.test';
  app.fetch = (url, options) => { requests.push(JSON.parse(options.body)); return Promise.resolve({json:()=>Promise.resolve({ok:true})}); };
  app.tabloyuCiz(); app.runTimer(300); await Promise.resolve();
  assert.equal(requests[0].action, 'syncStats');
  assert.equal(requests[0].stats.recall.firstSeenOn, '2026-01-01');
});

test('real library rendering keeps manual Known and Learned in different tables with row numbers', () => {
  const app = createApp(); addWord(app, 'biliyorum', 'Manual'); addWord(app, 'biliyorum', 'ManualTwo'); learnedWord(app); addWord(app, 'beklemede', 'Learning');
  vm.runInContext(functions.match(/^    function tabloyuCiz\([^\n]*\) \{[\s\S]*?^    \}/m)[0], app);
  app.updateWorkspaceOverview = app.havuzTablosuCiz = app.updateVisibleResultCount = () => {};
  app.tabloyuCiz();
  assert.match(app.document.getElementById('knownPoolBody').innerHTML, /<tr class="row-known"><td class="col-row">1<\/td>[\s\S]*Manual/);
  assert.match(app.document.getElementById('knownPoolBody').innerHTML, /<tr class="row-known"><td class="col-row">2<\/td>[\s\S]*ManualTwo/);
  assert.doesNotMatch(app.document.getElementById('knownPoolBody').innerHTML, /Recall/);
  assert.match(app.document.getElementById('learnedPoolBody').innerHTML, /<tr><td class="col-row">1<\/td>[\s\S]*Recall/);
  assert.doesNotMatch(app.document.getElementById('learnedPoolBody').innerHTML, /Manual/);
  assert.match(app.document.getElementById('poolBody').innerHTML, /<tr><td class="col-row">1<\/td>[\s\S]*Learning/);
  assert.doesNotMatch(app.document.getElementById('poolBody').innerHTML, /Recall|Manual/);
});


test('stale word status cannot restore Learned after a newer server demotion', () => {
  const app = createApp(); const word = learnedWord(app);
  const old = {...word}; app.istatistikMapKaydet({recall:{learningVersion:2, learningUpdatedAt:app.Date.now()+1, learned:false, manualKnown:false, firstSeenOn:'2025-11-01', trainingStartedOn:'2026-01-01'}});
  app.listeyiBirlestirVeHazirla([old], true);
  assert.equal(app.TUM_KELIMELER[0].durum, 'bilmiyorum');
  assert.equal(app.aktifHavuzuGetir().length, 1);
});

test('editing a Known checkbox in Google Sheets also separates it from Learned state', () => {
  const {server,ss,words} = createServer();
  server.durumGuncelle(2, 'ogrenildi', {learningVersion:2, learningUpdatedAt:10, learned:true, manualKnown:false});
  words.rows[1][10] = true;
  const layout = server.getSheetLayout_(); layout.audio = -1; server.getSheetLayout_ = () => layout;
  server.onEdit({range:{getSheet:()=>words, getColumn:()=>11, getNumColumns:()=>1, getRow:()=>2, getNumRows:()=>1}});
  const state = server.istatistikleriOku_(ss).recall;
  assert.equal(state.manualKnown, true); assert.equal(state.learned, false);
  assert.equal(server.verileriGetir().words[0].durum, 'biliyorum');
});

test('the real quiz handler and automatic next question require 3 different days after the Goal when never answered during Goal', () => {
  const app = studyApp('quiz'); const word = addWord(app); app.ikizKarsilastirmaKutusuHtml = () => '';
  app.tabloyuCiz(); app.advance(35);
  for (let i = 1; i <= 3; i++) {
    if (i > 1) app.advance(1);
    app.quizSecenekSec({classList:{add(){}}},word.kelime,word.kelime,word.row);
    assert.equal(app.istatistikMapAl().recall.correctDays.length, i);
    assert.equal(word.durum, i === 3 ? 'ogrenildi' : 'bilmiyorum');
    app.runTimer(30000);
  }
  assert.equal(app.aktifHavuzuGetir().length, 0);
});

test('an unchanged live sync waits for the final check once the Goal expires', () => {
  const app = createApp(); const word = addWord(app);
  answer(app, word);
  app.advance(1); answer(app, word);
  app.advance(1); answer(app, word);
  app.advance(26); app.listeyiBirlestirVeHazirla([{...word}], true);
  assert.equal(app.TUM_KELIMELER[0].durum, 'bilmiyorum');
  answer(app, app.TUM_KELIMELER[0]);
  assert.equal(app.TUM_KELIMELER[0].durum, 'ogrenildi');
});

for (const mode of ['flashcards', 'quiz']) {
  test(mode + ': opening practice fetches fresh words even immediately after a sync', () => {
    const app = studyApp(mode); addWord(app);
    app.SON_CANLI_SENKRON = app.Date.now();
    let syncs = 0;
    app.verileriYukle = () => { syncs++; };
    app.calismaModuAc(mode);
    assert.equal(syncs, 1);
    assert.deepEqual(app.spoken, ['Recall']);
  });
  test(mode + ': automatic pronunciation starts without downloading audio and cancels the old recording', async () => {
    const app = studyApp(mode); addWord(app);
    const utterances = []; let paused = 0;
    app.window.speechSynthesis = { cancel() {}, speak: u => utterances.push(u) };
    app.SpeechSynthesisUtterance = class { constructor(text) { this.text = text; } };
    vm.runInContext(functions.match(/^    function kelimeOku\([^\n]*\) \{[\s\S]*?^    \}/m)[0], app);
    let fail;
    app.Audio = class { play() { return new Promise((_, reject) => { fail = reject; }); } pause() { paused++; } };
    app.AKTIF_SES = null;
    app.kelimeOku('Previous');
    app.tabloyuCiz();
    assert.equal(paused, 1);
    assert.equal(utterances[0].text, 'Recall');
    assert.equal(utterances[0].lang, 'en-US');
    fail(); await Promise.resolve();
    assert.equal(utterances.length, 1);
    app.tabloyuCiz(); assert.equal(utterances.length, 1);
    app.SES_ACIK = false; app.fcDegistir(1); assert.equal(utterances.length, 1);
  });
}

test('quiz resumes from the exact question after answering several questions, exiting, and reloading', () => {
  const storage = new Map();
  let app = createApp(storage);
  app.AKTIF_MOD = 'quiz';
  app.AKTIF_SEKME = 'flashcards';
  app.ikizKarsilastirmaKutusuHtml = () => '';
  app.mobilMenuKapat = app.sekmeButonlariniGuncelle = app.canliVerileriGuncelle = () => {};
  app.tabloyuCiz = () => { app.renderCount++; app.dortHaftaHatasizKontrolEt(); app.fcCiz(app.aktifHavuzuGetir()); };
  const names = ['Alpha', 'Bravo', 'Charlie', 'Delta', 'Echo', 'Foxtrot'];
  names.forEach(n => addWord(app, 'bilmiyorum', n));

  app.tabloyuCiz();
  for (let i = 0; i < 3; i++) {
    const deck = app.fcDestesiniGetir(app.aktifHavuzuGetir());
    const card = deck[app.FC_INDEX];
    app.quizSecenekSec({ classList: { add() {} } }, card.kelime, card.kelime, card.row);
    app.runTimer(30000);
  }
  assert.equal(app.FC_INDEX, 3);
  assert.equal(app.fcDestesiniGetir(app.aktifHavuzuGetir())[app.FC_INDEX].kelime, 'Delta');

  // Answer question 4 and exit immediately BEFORE the 1-second auto-advance timer fires
  const fourth = app.fcDestesiniGetir(app.aktifHavuzuGetir())[app.FC_INDEX];
  app.quizSecenekSec({ classList: { add() {} } }, fourth.kelime, fourth.kelime, fourth.row);
  app.sekmeDegistir('sheet1');

  // Re-enter quiz in the same session -> should resume at Echo (index 4), not Alpha (index 0)
  app.calismaModuAc('quiz');
  assert.equal(app.FC_INDEX, 4);
  assert.equal(app.fcDestesiniGetir(app.aktifHavuzuGetir())[app.FC_INDEX].kelime, 'Echo');

  // Simulate full reload from localStorage -> should still resume at Echo (index 4)
  app = createApp(storage);
  app.mobilMenuKapat = app.sekmeButonlariniGuncelle = app.canliVerileriGuncelle = app.yaziBoyutuUygula = app.sesDurumuUygula = app.hedefHaftaSecicileriGuncelle = () => {};
  app.tabloyuCiz = () => { app.renderCount++; app.dortHaftaHatasizKontrolEt(); app.fcCiz(app.aktifHavuzuGetir()); };
  names.forEach(n => addWord(app, 'bilmiyorum', n));
  app.tercihleriYukleVeUygula();
  app.calismaModuAc('quiz');
  assert.equal(app.FC_INDEX, 4);
  assert.equal(app.fcDestesiniGetir(app.aktifHavuzuGetir())[app.FC_INDEX].kelime, 'Echo');
});

test('words are distributed equally: low-exposure words catch up before high-exposure words repeat', () => {
  const app = createApp();
  app.AKTIF_MOD = 'quiz';
  app.AKTIF_SEKME = 'flashcards';
  app.ikizKarsilastirmaKutusuHtml = () => '';
  app.mobilMenuKapat = app.sekmeButonlariniGuncelle = app.canliVerileriGuncelle = () => {};
  app.tabloyuCiz = () => { app.renderCount++; app.dortHaftaHatasizKontrolEt(); app.fcCiz(app.aktifHavuzuGetir()); };

  // 4 words seen 4 times, 2 words seen 10 times
  const lowWords = ['Low1', 'Low2', 'Low3', 'Low4'];
  const highWords = ['High1', 'High2'];
  [...highWords, ...lowWords].forEach(n => addWord(app, 'bilmiyorum', n));
  const stats = {};
  lowWords.forEach(n => { stats[n.toLowerCase()] = { learningVersion: 2, correct: 4, mistakes: 0, seenCount: 4 }; });
  highWords.forEach(n => { stats[n.toLowerCase()] = { learningVersion: 2, correct: 10, mistakes: 0, seenCount: 10 }; });
  app.istatistikMapKaydet(stats);

  app.tabloyuCiz();
  const seenInNext24 = [];
  for (let step = 0; step < 24; step++) {
    const deck = app.fcDestesiniGetir(app.aktifHavuzuGetir());
    const card = deck[app.FC_INDEX];
    seenInNext24.push(card.kelime);
    app.quizSecenekSec({ classList: { add() {} } }, card.kelime, card.kelime, card.row);
    app.runTimer(30000);
  }

  // During the 24 catch-up steps (4 words * 6 rounds = 24), High1 and High2 must NOT appear
  assert.ok(!seenInNext24.includes('High1'));
  assert.ok(!seenInNext24.includes('High2'));

  // All 6 words should now have exact equal exposure count of 10!
  const finalStats = app.istatistikMapAl();
  for (const name of [...lowWords, ...highWords]) {
    assert.equal(app.kelimeGorulmeSayisiAl(name, finalStats), 10);
  }

  // Next 6 steps should show every single word (including High1 and High2) exactly once
  const nextRound = [];
  for (let step = 0; step < 6; step++) {
    const deck = app.fcDestesiniGetir(app.aktifHavuzuGetir());
    const card = deck[app.FC_INDEX];
    nextRound.push(card.kelime);
    app.quizSecenekSec({ classList: { add() {} } }, card.kelime, card.kelime, card.row);
    app.runTimer(30000);
  }
  assert.equal(new Set(nextRound).size, 6);
});
