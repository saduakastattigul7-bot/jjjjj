// End-to-end check: 1 teacher + 4 group laptops = 5 separate browsers on the Firebase emulators.
// Run: npm run test:e2e   (starts the Auth + Firestore emulators with firestore.rules)
import http from 'node:http';
import { readFileSync, existsSync, mkdirSync } from 'node:fs';
import { join, extname } from 'node:path';
import { fileURLToPath } from 'node:url';
import assert from 'node:assert/strict';
import { chromium } from 'playwright';

const ROOT = fileURLToPath(new URL('..', import.meta.url));
const PUBLIC = join(ROOT, 'public');
const OUT = join(ROOT, 'tests', 'artifacts');
mkdirSync(OUT, { recursive: true });
const PROJECT = 'demo-qainau';
const EMAIL = 'mugalim@example.com', PASS = 'test-password-123';
const CHROME = process.env.CHROME_PATH || (existsSync('/opt/pw-browsers/chromium-1194/chrome-linux/chrome') ? '/opt/pw-browsers/chromium-1194/chrome-linux/chrome' : undefined);

/* ---------- static server; firebase-config.js is replaced by an emulator config ---------- */
const TEST_CONFIG = `window.FIREBASE_CONFIG = { apiKey: 'demo-api-key', authDomain: '${PROJECT}.firebaseapp.com', projectId: '${PROJECT}', appId: 'demo-app' };
window.QAINAU_OPTIONS = { emulator: true, heartbeatMs: 1500, presenceTimeoutMs: 6000 };`;
const TYPES = { '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8' };
const server = http.createServer((req, res) => {
  const path = decodeURIComponent(new URL(req.url, 'http://x').pathname);
  if (path === '/firebase-config.js') { res.writeHead(200, { 'content-type': TYPES['.js'] }); res.end(TEST_CONFIG); return; }
  const file = join(PUBLIC, path === '/' ? 'index.html' : path);
  if (!file.startsWith(PUBLIC) || !existsSync(file)) { res.writeHead(404); res.end(); return; }
  res.writeHead(200, { 'content-type': TYPES[extname(file)] || 'application/octet-stream' });
  res.end(readFileSync(file));
});
await new Promise((r) => server.listen(5055, '127.0.0.1', r));
const BASE = 'http://127.0.0.1:5055/';

/* ---------- emulator setup: teacher account + teachers/{uid} ---------- */
const signUp = await fetch(`http://127.0.0.1:9099/identitytoolkit.googleapis.com/v1/accounts:signUp?key=demo-api-key`, {
  method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify({ email: EMAIL, password: PASS, returnSecureToken: true })
}).then((r) => r.json());
assert.ok(signUp.localId, 'teacher account created in the Auth emulator');
await fetch(`http://127.0.0.1:8080/v1/projects/${PROJECT}/databases/(default)/documents/teachers/${signUp.localId}`, {
  method: 'PATCH', headers: { 'content-type': 'application/json', authorization: 'Bearer owner' },
  body: JSON.stringify({ fields: { name: { stringValue: 'Маженова Салтанат Төлендіқызы' } } })
}).then((r) => assert.equal(r.status, 200, 'teachers/{uid} document created'));

/* ---------- browsers ---------- */
const browser = await chromium.launch({ executablePath: CHROME });
const FB = join(ROOT, 'node_modules', 'firebase');
async function newDevice(label) {
  const context = await browser.newContext({ viewport: { width: 1400, height: 1000 }, acceptDownloads: true });
  // The Firebase SDK files are the same ones served by www.gstatic.com; they are served from node_modules here.
  await context.route('https://www.gstatic.com/firebasejs/**', (route) => {
    const name = new URL(route.request().url()).pathname.split('/').pop();
    route.fulfill({ status: 200, contentType: 'text/javascript', body: readFileSync(join(FB, name)) });
  });
  await context.route('https://cdn.jsdelivr.net/npm/qrcode-generator@1.4.4/qrcode.js', (route) =>
    route.fulfill({ status: 200, contentType: 'text/javascript', body: readFileSync(join(ROOT, 'node_modules', 'qrcode-generator', 'qrcode.js')) }));
  await context.route('https://fonts.googleapis.com/**', (route) => route.abort());
  const page = await context.newPage();
  page.on('pageerror', (e) => console.log(`[${label}] page error:`, e.message));
  return { label, context, page };
}
const log = (...a) => console.log('•', ...a);
async function until(page, fn, arg, what, timeout = 20000) {
  try { await page.waitForFunction(fn, arg, { timeout, polling: 200 }); }
  catch (e) { await page.screenshot({ path: join(OUT, 'fail.png'), fullPage: true }); throw new Error('Timed out waiting for: ' + what); }
}
const badge = (page) => page.locator('[data-net]').first().innerText();
async function waitSaved(dev) { await until(dev.page, () => { const b = document.querySelector('[data-net]'); return b && b.textContent === 'Сақталды'; }, null, dev.label + ' shows «Сақталды»'); }

/* ---------- deterministic classroom data ---------- */
const NAMES = [
  ['Айзат Серікова', 'Әсел Нұрланқызы', 'Қайрат Жұмабаев', 'Нұрсұлтан Ерланұлы', 'Дана Асқарова'],
  ['Аружан Болатқызы', 'Ерасыл Тимурұлы', 'Томирис Қанатқызы', 'Бекзат Мұратұлы', 'Мадина Сәкенқызы', 'Алихан Дәуренұлы'],
  ['Жансая Ермекқызы', 'Ернұр Талғатұлы', 'Аяулым Бауыржанқызы', 'Дінмұхаммед Арманұлы', 'Сая Берікқызы', 'Арман Қуанышұлы', 'Ұлжан Маратқызы'],
  ['Інжу Айдосқызы', 'Нұрлан Серікұлы', 'Ақбота Ғалымқызы', 'Темірлан Асанұлы']
];
const s1 = (g, i, j) => (i + j + g) % 3 !== 0;
const s2 = (g, i, j) => (i + j) % 2 === 0;
const s3 = (g, q) => ((q + g) % 4 !== 0 ? 'ok' : 'bad');
const s4 = (g, i, k) => (i * 2 + k + g) % 3 !== 1;
const LAB = [[20, 35, 52, 68, 81, 93, 99, 100], [21, 30, 44], [19, 33, 50, 66, 79], [22, 40, 61, 80, 95, 100, 100]];
function expected(g, i, fix = {}) {
  const c = (n, f) => Array.from({ length: n }, (_, k) => f(k)).filter(Boolean).length;
  let a = c(3, (j) => (fix['s1.' + j] ?? s1(g, i, j))), b = c(2, (j) => s2(g, i, j)), q = c(9, (k) => s3(g, k) === 'ok'), e = c(9, (k) => s4(g, i, k));
  const total = a + b + q + e;
  return { s1: a, s2: b, s3: q, s4: e, total, ten: (Math.round(total / 23 * 100) / 10).toFixed(1).replace('.', ',') };
}

/* ---------- 1. teacher opens a lesson ---------- */
const T = await newDevice('Мұғалім');
await T.page.goto(BASE);
await T.page.getByRole('button', { name: 'МҰҒАЛІМ' }).click();
await T.page.fill('#tEmail', EMAIL);
await T.page.fill('#tPass', PASS);
await T.page.getByRole('button', { name: 'КІРУ →' }).click();
await T.page.getByText('МЕНІҢ САБАҚТАРЫМ').waitFor();
await T.page.fill('#newTitle', '8 «А» сынып');
await T.page.getByRole('button', { name: 'ЖАҢА САБАҚ ОТЫРЫСЫН АШУ' }).click();
await T.page.locator('#sessCode').waitFor();
const CODE = (await T.page.locator('#sessCode').innerText()).trim();
assert.match(CODE, /^[A-HJ-NP-Z2-9]{6}$/);
assert.ok(await T.page.locator('.code-card .qr-img svg').count(), 'QR code rendered on the teacher screen');
log('teacher opened lesson', CODE);

/* ---------- 2. four laptops join at the same time ---------- */
const groups = await Promise.all([0, 1, 2, 3].map((g) => newDevice(`${g + 1}-топ`)));
await Promise.all(groups.map(async (dev, g) => {
  await dev.page.goto(BASE + '?code=' + CODE);
  assert.equal(await dev.page.inputValue('#joinCode'), CODE, 'code prefilled from link');
  await dev.page.click(`[data-act="join-g"][data-g="${g}"]`);
  await dev.page.getByRole('button', { name: 'КІРУ →' }).click();
  await dev.page.locator('.name-input').first().waitFor();
}));
log('4 groups joined');

// a 6th device cannot take a group that is already open elsewhere
const intruder = await newDevice('бөгде');
await intruder.page.goto(BASE + '?code=' + CODE);
await intruder.page.click('[data-act="join-g"][data-g="0"]');
await intruder.page.getByRole('button', { name: 'КІРУ →' }).click();
await intruder.page.getByText('басқа құрылғыда ашылған').waitFor();
await intruder.context.close();
log('second device for 1-топ refused');

await until(T.page, () => document.querySelectorAll('.gst.on').length === 4, null, 'teacher sees 4 groups «Қосылған»');

/* ---------- 3. names typed simultaneously; teacher sees them without reload ---------- */
await Promise.all(groups.map(async (dev, g) => {
  for (let i = 0; i < NAMES[g].length; i++) await dev.page.locator('.name-input').nth(i).fill(NAMES[g][i]);
  await dev.page.locator('body').click({ position: { x: 5, y: 5 } });
  await waitSaved(dev);
}));
const total = NAMES.flat().length;
await until(T.page, (n) => document.querySelectorAll('#resTable tbody tr').length === n, total, `teacher table has ${total} students`);
// stages not entered yet → «—», not 0
assert.equal(await T.page.locator('tr[data-row="g0s0"] td').nth(2).innerText(), '—');
log('teacher table filled live:', total, 'students; not-entered stages show «—»');

/* ---------- 4. all four groups assess at the same time ---------- */
async function assess(dev, g) {
  const p = dev.page, n = NAMES[g].length;
  const click = async (sel) => { await p.locator(sel).click(); };
  await click('[data-act="next"]');                                     // → 1/4
  for (let i = 0; i < n; i++) for (let j = 0; j < 3; j++) if (s1(g, i, j)) await click(`[data-act="s1"][data-id="g${g}s${i}"][data-i="${j}"]`);
  await click('[data-act="next"]');                                     // → 2/4
  for (let i = 0; i < n; i++) for (let j = 0; j < 2; j++) if (s2(g, i, j)) await click(`[data-act="s2"][data-id="g${g}s${i}"][data-i="${j}"]`);
  await click('[data-act="next"]');                                     // → 3/4
  for (let q = 0; q < 9; q++) await click(`[data-act="s3"][data-q="${q}"][data-v="${s3(g, q)}"]`);
  await click('[data-act="next"]');                                     // → 4/4
  for (let i = 0; i < n; i++) {
    await click(`[data-act="sel"][data-slot="${i}"]`);
    for (let k = 0; k < 9; k++) if (s4(g, i, k)) await click(`[data-act="s4"][data-id="g${g}s${i}"][data-i="${k}"]`);
  }
  await click('[data-act="next"]');                                     // → summary
  await waitSaved(dev);
}
await Promise.all(groups.map(assess));
log('4 groups finished assessment in parallel');

async function rowText(page, id) { return page.locator(`tr[data-row="${id}"] td`).allInnerTexts(); }
async function checkTeacherRow(g, i, fix) {
  const e = expected(g, i, fix);
  const want = [`${e.s1}/3`, `${e.s2}/2`, `${e.s3}/9`, `${e.s4}/9`, `${e.total}/23`, e.ten];
  await until(T.page, ([id, w]) => {
    const tds = [...document.querySelectorAll(`tr[data-row="${id}"] td`)].map((td) => td.textContent.replace('✎', '').trim());
    return JSON.stringify(tds.slice(2, 8)) === JSON.stringify(w);
  }, [`g${g}s${i}`, want], `row g${g}s${i} = ${want.join(' ')}`);
}
for (let g = 0; g < 4; g++) for (let i = 0; i < NAMES[g].length; i++) await checkTeacherRow(g, i);
const firstRow = await rowText(T.page, 'g2s3');
assert.equal(firstRow[0], '3-топ · ИНЖЕНЕРЛЕР');
assert.equal(firstRow[1], 'ДІНМҰХАММЕД АРМАНҰЛЫ');
log('teacher table matches every student (formula total/23×10):', firstRow.slice(0, 8).join(' | '));
await T.page.screenshot({ path: join(OUT, 'teacher-monitor.png'), fullPage: true });

/* ---------- 5. lab: each group fills its own table at the same time ---------- */
await Promise.all(groups.map(async (dev, g) => {
  const p = dev.page;
  await p.click('[data-act="go"][data-view="lesson"]');
  await p.click('[data-act="lsec"][data-s="4"]');
  for (let i = 0; i < LAB[g].length; i++) await p.fill(`#labT${i}`, String(LAB[g][i]));
  await p.fill('#labQ3', `${g + 1}-топ: су ${LAB[g].at(-1)} °C-та қайнады`);
  await p.locator('body').click({ position: { x: 5, y: 5 } });
  await waitSaved(dev);
}));
await T.page.click('[data-act="go"][data-view="lab"]');
for (let g = 0; g < 4; g++) {
  await until(T.page, ([g, n]) => { const c = document.querySelector(`[data-lab="${g}"] .badge`); return c && c.textContent.startsWith(n + ' / 15'); }, [g, LAB[g].length], `teacher sees ${LAB[g].length} lab values of group ${g + 1}`);
}
assert.match(await T.page.locator('[data-lab="3"]').innerText(), /4-топ: су 100 °C-та қайнады/);
log('teacher sees 4 separate lab tables/graphs');

// «Графикті тазалау» in group 2 clears only group 2
await groups[1].page.click('[data-act="lab-clear"]');
await groups[1].page.click('[data-act="lab-clear-yes"]');
await waitSaved(groups[1]);
await until(T.page, () => document.querySelector('[data-lab="1"] .badge').textContent.startsWith('0 / 15'), null, 'group 2 lab cleared on teacher screen');
for (const g of [0, 2, 3]) assert.ok((await T.page.locator(`[data-lab="${g}"] .badge`).innerText()).startsWith(LAB[g].length + ' / 15'), `group ${g + 1} lab untouched`);
log('«Графикті тазалау» of group 2 did not touch groups 1, 3, 4');
await T.page.screenshot({ path: join(OUT, 'teacher-labs.png'), fullPage: true });

/* ---------- 6. teacher correction is final; the group cannot override it ---------- */
await T.page.click('[data-act="go"][data-view="monitor"]');
await T.page.click('[data-act="fix-open"][data-id="g0s1"]');
const was = s1(0, 1, 0);
await T.page.click(`.fix-box [data-act="fix"][data-k="s1"][data-i="0"][data-v="${was ? '0' : '1'}"]`);
await T.page.click('.fix-box [data-act="modal-no"]');
await checkTeacherRow(0, 1, { 's1.0': !was });
assert.match(await T.page.locator('tr[data-row="g0s1"] td').nth(2).innerText(), /✎/);
const G1 = groups[0].page;
await G1.click('[data-act="go"][data-view="group"]');
await G1.click('[data-act="step"][data-step="1"]');
const lockedChip = G1.locator('[data-act="s1"][data-id="g0s1"][data-i="0"]');
await until(G1, () => document.querySelector('[data-act="s1"][data-id="g0s1"][data-i="0"]')?.classList.contains('locked'), null, 'group 1 sees the teacher lock');
await lockedChip.click();
await G1.getByText('мұғалім түзеткен', { exact: false }).first().waitFor();
assert.equal(await lockedChip.getAttribute('aria-pressed'), String(!was));
await checkTeacherRow(0, 1, { 's1.0': !was });
log('teacher correction shown on group laptop as locked; group click ignored');

/* ---------- 7. one laptop loses internet; others keep working ---------- */
const G3 = groups[2];
await G3.context.setOffline(true);
await G3.page.click('[data-act="go"][data-view="lesson"]');
await G3.page.fill('#labT6', '97');
await G3.page.locator('body').click({ position: { x: 5, y: 5 } });
await until(G3.page, () => /Байланыс жоқ/.test(document.querySelector('[data-net]').textContent), null, 'group 3 shows «Байланыс жоқ»');
assert.ok(!/Сақталды/.test(await badge(G3.page)), 'offline device never claims «Сақталды»');
await T.page.click('[data-act="go"][data-view="monitor"]');
await until(T.page, () => document.querySelector('[data-gst="2"]').classList.contains('lost'), null, 'teacher sees group 3 «Байланыс үзілді»', 20000);
// group 4 keeps saving while group 3 is offline
await groups[3].page.click('[data-act="go"][data-view="group"]');
await groups[3].page.click('[data-act="step"][data-step="0"]');
await groups[3].page.locator('.name-input').nth(4).fill('Бауыржан Ерікұлы');
await groups[3].page.locator('body').click({ position: { x: 5, y: 5 } });
await waitSaved(groups[3]);
await until(T.page, () => !!document.querySelector('tr[data-row="g3s4"]'), null, 'teacher sees group 4 new student while group 3 is offline');
await T.page.click('[data-act="go"][data-view="lab"]');
assert.ok((await T.page.locator('[data-lab="2"] .badge').innerText()).startsWith('5 / 15'), 'unsent offline value not shown as saved on teacher screen');
await G3.context.setOffline(false);
await waitSaved(G3);
await until(T.page, () => document.querySelector('[data-lab="2"] .badge').textContent.startsWith('6 / 15'), null, 'offline change of group 3 synced after reconnect');
assert.ok((await T.page.locator('[data-lab="3"] .badge').innerText()).startsWith(LAB[3].length + ' / 15'), 'group 4 lab intact after group 3 sync');
await T.page.click('[data-act="go"][data-view="monitor"]');
await until(T.page, () => document.querySelector('[data-gst="2"]').classList.contains('on'), null, 'group 3 back to «Қосылған»');
await checkTeacherRow(3, 0);
log('offline group 3: «Байланыс жоқ» → reconnect → synced; groups 1, 2, 4 unaffected');

/* ---------- 8. reload restores the same lesson on the device; UI state stays per device ---------- */
await G1.reload();
await G1.locator('.stepper').waitFor();
assert.equal(await G1.locator('.step.cur b').innerText(), '1/4', 'group 1 is back on its own step after reload');
await G1.click('[data-act="step"][data-step="0"]');
assert.equal(await G1.locator('.name-input').first().inputValue(), NAMES[0][0]);
assert.equal(await T.page.locator('.tab.cur').innerText(), 'БАҚЫЛАУ', 'teacher screen did not follow the group');
await T.page.reload();
await T.page.locator('#resTable').waitFor();
await checkTeacherRow(1, 2);
log('reload restores data on group and teacher devices');

/* ---------- 9. freeing a laptop, CSV, ending the lesson ---------- */
const G4 = groups[3].page;
await T.page.click('[data-act="release"][data-g="3"]');
await T.page.click('[data-act="release-yes"]');
await G4.getByText('топтан ажыратылды', { exact: false }).waitFor();
await G4.click('[data-act="join-g"][data-g="3"]');
await G4.getByRole('button', { name: 'КІРУ →' }).click();
await G4.locator('.stepper').waitFor();
log('teacher freed 4-топ laptop; it re-joined');

const [download] = await Promise.all([T.page.waitForEvent('download'), T.page.click('[data-act="csv"]')]);
const raw = readFileSync(await download.path());
assert.deepEqual([...raw.subarray(0, 3)], [0xef, 0xbb, 0xbf], 'CSV starts with UTF-8 BOM for Excel');
const lines = raw.subarray(3).toString('utf8').trim().split(/\r\n/);
assert.equal(lines.length, total + 2, 'CSV has a header and every student');
assert.match(lines[0], /^Топ;Топ атауы;Оқушының аты-жөні;Үй тапсырмасын қайталау/);
const e0 = expected(2, 3);
assert.ok(lines.some((l) => l.startsWith(`3-топ;ИНЖЕНЕРЛЕР;Дінмұхаммед Арманұлы;${e0.s1};${e0.s2};${e0.s3};${e0.s4};${e0.total};${e0.ten};`)), 'CSV row values');
assert.match(lines.find((l) => l.startsWith('1-топ;АЛЬПИНИСТЕР;Әсел Нұрланқызы;')), /;иә;/, 'CSV marks the teacher-corrected student');
log('CSV downloaded:', lines.length - 1, 'rows');

await T.page.click('[data-act="end-session"]');
await T.page.click('[data-act="end-yes"]');
await Promise.all(groups.map((d) => d.page.getByText('САБАҚ АЯҚТАЛДЫ', { exact: false }).first().waitFor()));
await G1.click('[data-act="step"][data-step="2"]');
const chip = G1.locator('[data-act="s2"][data-id="g0s0"][data-i="1"]');
const before = await chip.getAttribute('aria-pressed');
await chip.click();
await G1.getByText('Сабақ аяқталды: деректерді өзгертуге болмайды', { exact: false }).waitFor();
assert.equal(await chip.getAttribute('aria-pressed'), before);
log('lesson ended: groups see the banner and cannot edit');
await T.page.screenshot({ path: join(OUT, 'teacher-ended.png'), fullPage: true });
await groups[1].page.screenshot({ path: join(OUT, 'group2-ended.png'), fullPage: true });

await browser.close();
server.close();
console.log('\nE2E OK: 5 browsers, live sync, isolation, offline recovery, corrections, CSV, end of lesson.');
