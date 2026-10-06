// Security-rule tests against the Firestore emulator.
// Run: npm run test:rules
import { test, before, after, beforeEach } from 'node:test';
import { readFileSync } from 'node:fs';
import { initializeTestEnvironment, assertSucceeds, assertFails } from '@firebase/rules-unit-testing';
import { doc, getDoc, setDoc, updateDoc, writeBatch, serverTimestamp, deleteField, collection, getDocs, query, where } from 'firebase/firestore';

const CODE = 'ABC234';
let env;

before(async () => {
  env = await initializeTestEnvironment({
    projectId: 'demo-qainau',
    firestore: { rules: readFileSync(new URL('../firestore.rules', import.meta.url), 'utf8'), host: '127.0.0.1', port: 8080 }
  });
});
after(async () => { await env.cleanup(); });

const teacher = () => env.authenticatedContext('teacher1', { email: 't@example.com', firebase: { sign_in_provider: 'password' } }).firestore();
const otherTeacher = () => env.authenticatedContext('teacher2', { email: 't2@example.com', firebase: { sign_in_provider: 'password' } }).firestore();
const device = (id) => env.authenticatedContext(id, { firebase: { sign_in_provider: 'anonymous' } }).firestore();

function blankGroup(g) { return { g, ownerUid: null, claimedAt: null, names: {}, s3: {}, s3Final: {}, done: {}, lab: { temps: {}, ans: {} }, updatedAt: null }; }

async function createSession(db, code = CODE) {
  const b = writeBatch(db);
  b.set(doc(db, 'sessions', code), { teacherUid: 'teacher1', status: 'open', title: '8А', createdAt: serverTimestamp(), endedAt: null });
  for (let g = 0; g < 4; g++) b.set(doc(db, 'sessions', code, 'groups', String(g)), blankGroup(g));
  return b.commit();
}
const claim = (db, g, uid) => updateDoc(doc(db, 'sessions', CODE, 'groups', String(g)), { ownerUid: uid, claimedAt: serverTimestamp() });

beforeEach(async () => {
  await env.clearFirestore();
  await env.withSecurityRulesDisabled(async (ctx) => {
    await setDoc(doc(ctx.firestore(), 'teachers', 'teacher1'), { name: 'Мұғалім' });
    await setDoc(doc(ctx.firestore(), 'teachers', 'teacher2'), { name: 'Басқа мұғалім' });
  });
  await createSession(teacher());
});

test('only a registered teacher can open a lesson', async () => {
  await assertFails(createSession(device('dev1'), 'XYZ234'));
  await assertSucceeds(createSession(teacher(), 'XYZ235'));
});

test('a device claims a free group once; a second device is refused', async () => {
  await assertSucceeds(claim(device('dev1'), 0, 'dev1'));
  await assertSucceeds(claim(device('dev1'), 0, 'dev1'));      // re-claim after reload
  await assertFails(claim(device('dev2'), 0, 'dev2'));
  await assertFails(claim(device('dev2'), 1, 'dev1'));          // cannot claim in someone else's name
});

test('a group edits only its own group, field by field', async () => {
  await claim(device('dev1'), 0, 'dev1');
  await claim(device('dev2'), 1, 'dev2');
  const g0 = (db) => doc(db, 'sessions', CODE, 'groups', '0');
  await assertSucceeds(updateDoc(g0(device('dev1')), { 'names.0': 'Айбек', 'lab.temps.3': 42.5, 's3.2': 'ok', 'done.s1': true, updatedAt: serverTimestamp() }));
  await assertFails(updateDoc(g0(device('dev2')), { 'names.0': 'Хакер', updatedAt: serverTimestamp() }));
  await assertFails(updateDoc(g0(device('dev2')), { 'lab.temps': {}, updatedAt: serverTimestamp() }));
  await assertFails(getDoc(g0(device('dev2'))));
  // the group cannot touch the teacher's corrections or the owner field
  await assertFails(updateDoc(g0(device('dev1')), { 's3Final.0': 'ok', updatedAt: serverTimestamp() }));
  await assertFails(updateDoc(g0(device('dev1')), { ownerUid: 'dev2', updatedAt: serverTimestamp() }));
  // bad values are rejected
  await assertFails(updateDoc(g0(device('dev1')), { 'lab.temps.3': 500, updatedAt: serverTimestamp() }));
  await assertFails(updateDoc(g0(device('dev1')), { 's3.9': 'ok', updatedAt: serverTimestamp() }));
  await assertFails(updateDoc(g0(device('dev1')), { 'names.0': 'x'.repeat(41), updatedAt: serverTimestamp() }));
});

test('students: self-assessment by the group, final marks only by the teacher', async () => {
  await claim(device('dev1'), 0, 'dev1');
  await claim(device('dev2'), 1, 'dev2');
  const s = (db, g = '0') => doc(db, 'sessions', CODE, 'groups', g, 'students', '2');
  await assertSucceeds(setDoc(s(device('dev1')), { self: { s1: { '0': true } }, updatedAt: serverTimestamp() }, { merge: true }));
  await assertSucceeds(setDoc(s(device('dev1')), { self: { s4: { '8': false } }, updatedAt: serverTimestamp() }, { merge: true }));
  await assertFails(setDoc(s(device('dev2')), { self: { s1: { '1': true } }, updatedAt: serverTimestamp() }, { merge: true }));
  await assertFails(getDoc(s(device('dev2'))));
  await assertFails(setDoc(s(device('dev1')), { final: { s1: { '0': true } }, updatedAt: serverTimestamp() }, { merge: true }));
  await assertFails(setDoc(s(device('dev1')), { self: { s1: { '5': true } }, updatedAt: serverTimestamp() }, { merge: true }));
  await assertSucceeds(setDoc(s(teacher()), { final: { s1: { '0': false } }, finalAt: serverTimestamp() }, { merge: true }));
  await assertSucceeds(updateDoc(s(teacher()), { 'final.s1.0': deleteField(), finalAt: serverTimestamp() }));
  await assertFails(setDoc(s(otherTeacher()), { final: { s1: { '0': true } }, finalAt: serverTimestamp() }, { merge: true }));
  await assertSucceeds(getDocs(collection(teacher(), 'sessions', CODE, 'groups', '0', 'students')));
  await assertSucceeds(getDocs(collection(device('dev1'), 'sessions', CODE, 'groups', '0', 'students')));
  await assertFails(getDocs(collection(device('dev1'), 'sessions', CODE, 'groups', '1', 'students')));
});

test('teacher corrects group work and frees a device; other teachers cannot', async () => {
  await claim(device('dev1'), 0, 'dev1');
  const g0 = (db) => doc(db, 'sessions', CODE, 'groups', '0');
  await assertSucceeds(updateDoc(g0(teacher()), { 's3Final.4': 'bad', finalAt: serverTimestamp() }));
  await assertFails(updateDoc(g0(teacher()), { 'names.0': 'Teacher edit' }));
  await assertFails(updateDoc(g0(otherTeacher()), { 's3Final.4': 'ok', finalAt: serverTimestamp() }));
  await assertFails(getDoc(g0(otherTeacher())));
  await assertSucceeds(updateDoc(g0(teacher()), { ownerUid: null }));
  await assertFails(updateDoc(g0(device('dev1')), { 'names.0': 'Кейін', updatedAt: serverTimestamp() }));
  await assertSucceeds(claim(device('dev3'), 0, 'dev3'));
});

test('after the lesson ends, groups cannot change anything', async () => {
  await claim(device('dev1'), 0, 'dev1');
  await assertFails(updateDoc(doc(device('dev1'), 'sessions', CODE), { status: 'ended' }));
  await assertSucceeds(updateDoc(doc(teacher(), 'sessions', CODE), { status: 'ended', endedAt: serverTimestamp() }));
  await assertFails(updateDoc(doc(device('dev1'), 'sessions', CODE, 'groups', '0'), { 'names.0': 'Кеш', updatedAt: serverTimestamp() }));
  await assertFails(setDoc(doc(device('dev1'), 'sessions', CODE, 'groups', '0', 'students', '0'), { self: { s1: { '0': true } }, updatedAt: serverTimestamp() }, { merge: true }));
  await assertFails(claim(device('dev2'), 1, 'dev2'));
  // the teacher can still correct
  await assertSucceeds(setDoc(doc(teacher(), 'sessions', CODE, 'groups', '0', 'students', '0'), { final: { s2: { '1': true } }, finalAt: serverTimestamp() }, { merge: true }));
});

test('presence: only the owning device writes, only the teacher reads', async () => {
  await claim(device('dev1'), 0, 'dev1');
  const p = (db, g = '0') => doc(db, 'sessions', CODE, 'presence', g);
  await assertSucceeds(setDoc(p(device('dev1')), { lastSeen: serverTimestamp(), online: true }));
  await assertFails(setDoc(p(device('dev2')), { lastSeen: serverTimestamp(), online: true }));
  await assertFails(setDoc(p(device('dev1'), '1'), { lastSeen: serverTimestamp(), online: true }));
  await assertSucceeds(getDocs(collection(teacher(), 'sessions', CODE, 'presence')));
  await assertFails(getDocs(collection(device('dev1'), 'sessions', CODE, 'presence')));
});

test('teacher lists only own lessons', async () => {
  await assertSucceeds(getDocs(query(collection(teacher(), 'sessions'), where('teacherUid', '==', 'teacher1'))));
  await assertFails(getDocs(query(collection(otherTeacher(), 'sessions'), where('teacherUid', '==', 'teacher1'))));
  await assertFails(getDocs(collection(device('dev1'), 'sessions')));
  await assertSucceeds(getDoc(doc(device('dev9'), 'sessions', CODE)));
});

test('worst case: a fully filled group still accepts every write the app makes', async () => {
  // The app sends one field per write (e.g. "names.3", "lab.temps.7"); the rules validate only
  // the keys a write changes, so a full document stays far below the 1000-expression limit.
  await claim(device('dev1'), 0, 'dev1');
  const g0 = doc(device('dev1'), 'sessions', CODE, 'groups', '0');
  const fill = (n, f) => Object.fromEntries(Array.from({ length: n }, (_, i) => [String(i), f(i)]));
  await assertSucceeds(updateDoc(g0, { names: fill(7, (i) => 'Оқушы аты-жөні ' + i), updatedAt: serverTimestamp() }));
  await assertSucceeds(updateDoc(g0, { s3: fill(9, (i) => (i % 2 ? 'ok' : 'bad')), done: { s1: true, s2: true, s3: true, s4: true }, updatedAt: serverTimestamp() }));
  await assertSucceeds(updateDoc(g0, { 'lab.temps': fill(15, (i) => 20 + i * 5.5), updatedAt: serverTimestamp() }));
  await assertSucceeds(updateDoc(g0, { 'lab.ans': fill(6, () => 'ұзын жауап '.repeat(100)), updatedAt: serverTimestamp() }));
  await assertSucceeds(updateDoc(g0, { 'lab.temps.14': 99.5, updatedAt: serverTimestamp() }));
  await assertSucceeds(updateDoc(g0, { 'names.6': 'Жаңа есім', updatedAt: serverTimestamp() }));
  await assertSucceeds(updateDoc(g0, { 'lab.ans.5': 'жауап', updatedAt: serverTimestamp() }));
  await assertFails(updateDoc(g0, { 'lab.ans.5': 'x'.repeat(2001), updatedAt: serverTimestamp() }));
  await assertFails(updateDoc(g0, { 'lab.temps.2': 'жүз', updatedAt: serverTimestamp() }));
  await assertSucceeds(updateDoc(g0, { 'lab.temps': {}, updatedAt: serverTimestamp() }));           // «Графикті тазалау»
  await assertSucceeds(updateDoc(g0, { lab: { temps: {}, ans: {} }, updatedAt: serverTimestamp() })); // «Қайта бастау»
  await assertSucceeds(updateDoc(doc(teacher(), 'sessions', CODE, 'groups', '0'), { s3Final: fill(9, () => 'ok'), finalAt: serverTimestamp() }));
  const s = doc(device('dev1'), 'sessions', CODE, 'groups', '0', 'students', '6');
  await assertSucceeds(setDoc(s, { self: { s1: fill(3, () => true), s2: fill(2, () => false), s4: fill(9, () => true) }, updatedAt: serverTimestamp() }));
  await assertSucceeds(setDoc(doc(teacher(), 'sessions', CODE, 'groups', '0', 'students', '6'), { final: { s1: fill(3, () => false), s2: fill(2, () => true), s4: fill(9, () => false) }, finalAt: serverTimestamp() }, { merge: true }));
  await assertFails(setDoc(s, { self: { s4: { '0': 'yes' } }, updatedAt: serverTimestamp() }, { merge: true }));
});
