'use strict';
// store.cjs — 共有マスターの読み書き（Electron 非依存・単体テストできるように分けてある）
//   loadConfig(dataRoot)                   : config/current.json を読む
//   loadMaster(dataRoot, year, month)      : 全科の予定（master\当直予定表_R8年N月.json）
//   saveDept(dataRoot, owner, ent, opt)    : 自科のエントリだけ差し替えて書き戻す
//
// ⚠ 手順は tools/master_store.py と同じ約束で動かすこと（Python と Node の両方が同じ1本を触る）:
//   ロックを取る → 読み直す → 自科だけ差し替える → 履歴へコピー → 一時ファイル経由で差し替え →
//   読み直して自科の rev が上がっているか確かめる → ロックを外す
//   同時編集の検出は updated（秒精度）ではなく **per-dept の rev** で行う。
//   20261009
const path = require('path');
const fs = require('fs');
const os = require('os');
const crypto = require('crypto');

const SCHEMA = 'hospital.duty-master/1';
const DIRNAME = 'master';
const HISTORY = '_履歴';
// これより古いロックは落ちたものとみなして横取り。20260926: 60秒 → 5分（tools/master_store.py と同じ 300 秒）。
//   ロックの時刻は共有のサーバーの時計、比べる now はこの PC の時計なので、時計がずれていると生きているロックまで
//   横取りしていた。横取りの前に目印（token）と時刻が見たときのままかを確かめるのは従来どおり
const LOCK_STALE_MS = 300 * 1000;
const LOCK_WAIT_MS = 20 * 1000;

function loadConfig(dataRoot) {
  const raw = fs.readFileSync(path.join(dataRoot, 'config', 'current.json'), 'utf8');
  return JSON.parse(raw.replace(/^﻿/, ''));
}

function masterPath(dataRoot, year, month) {
  return path.join(dataRoot, DIRNAME, `当直予定表_R${year - 2018}年${month}月.json`);
}

function emptyMaster(year, month) {
  return { schema: SCHEMA, year, month, rev: 0, updated: nowIso(), confirmed: null, depts: {} };
}

function nowIso() {
  const d = new Date(), p = (n) => String(n).padStart(2, '0');
  return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}T`
    + `${p(d.getHours())}:${p(d.getMinutes())}:${p(d.getSeconds())}`;
}

function sleep(ms) { Atomics.wait(new Int32Array(new SharedArrayBuffer(4)), 0, 0, ms); }

// ファイルの有無。無い（ENOENT）ときだけ null。それ以外の失敗（SMB の一時的な切断・権限など）は投げる
//   20260923: fs.existsSync はどんな失敗でも false を返すので、共有が一瞬見えないだけで「マスターが無い」
//   と誤認し、自科だけの新しいマスターで他の科を消しかねなかった
function statOrNull(file) {
  try { return fs.statSync(file); } catch (e) {
    if (e && e.code === 'ENOENT') return null;
    throw e;
  }
}

// マスターが見えないときの知らせ（利用者向けの平易な文。main が hidden:true を付けて renderer へ渡す。20260926）
const HIDDEN_MSG = '共有の当直予定表が見えません。少し待ってから開き直してください。続くときは管理者へ';
function hiddenError() {
  const e = new Error(HIDDEN_MSG);
  e.code = 'MASTER_HIDDEN';
  return e;
}

// 書き換えの最中を掴むことがあるので、読めなければ少し待って読み直す
//   20260926: マスターが無くても _履歴\ に版があれば、空の予定表を返さずに止める（refuseFreshMaster と同じ条件）。
//   空のまま画面に出すと、その科は「まだ何も入っていない」ように見え、そこから出すと他の人の入力を上書きしかねない
function loadMaster(dataRoot, year, month, retries = 6) {
  const file = masterPath(dataRoot, year, month);
  let last = null;
  for (let i = 0; i < retries; i++) {
    try {
      if (!statOrNull(file)) {
        if (!hasHistory(file)) return emptyMaster(year, month);
        last = hiddenError(); sleep(40 * (i + 1)); continue;   // 一瞬見えないだけかもしれないので読み直す
      }
    } catch (e) { last = e; sleep(40 * (i + 1)); continue; }
    try {
      const m = JSON.parse(fs.readFileSync(file, 'utf8').replace(/^﻿/, ''));
      if (m.schema !== SCHEMA) throw new Error(`schemaが違います: ${m.schema}`);
      if (!m.depts) m.depts = {};
      return m;
    } catch (e) { last = e; sleep(40 * (i + 1)); }
  }
  if (last && last.code === 'MASTER_HIDDEN') throw last;
  throw new Error(`マスターが読めません（${last && last.message}）。`
    + `master\\${HISTORY}\\ の直近の版から戻してください`);
}

function deptRev(master, owner) {
  const e = master && master.depts && master.depts[owner];
  return e ? Number(e.rev || 0) : null;
}

// ロックの中身（持ち主の目印）と更新時刻。無い・読めないときは null
function lockInfo(lock) {
  try {
    const st = fs.statSync(lock);
    return { token: fs.readFileSync(lock, 'utf8'), mtimeMs: st.mtimeMs };
  } catch { return null; }
}

// ロックには持ち主の目印（PC名 pid 時刻 乱数）を書く。外すときは目印が自分のものの時だけ消す
//   （20260923: 60秒を超えて横取りされたあとに、他人のロックを消してしまわないように）
//   waitMs は単体テスト用（既定 LOCK_WAIT_MS）
function acquireLock(file, waitMs = LOCK_WAIT_MS) {
  const lock = file + '.lock';
  fs.mkdirSync(path.dirname(lock), { recursive: true });
  const token = `${os.hostname()} ${process.pid} ${nowIso()} ${crypto.randomBytes(6).toString('hex')}`;
  const t0 = Date.now();
  for (;;) {
    try {
      const fd = fs.openSync(lock, 'wx');
      try { fs.writeSync(fd, token); } finally { fs.closeSync(fd); }
      return () => {
        const cur = lockInfo(lock);
        if (cur && cur.token === token) { try { fs.unlinkSync(lock); } catch { /* 既に無い */ } }
      };
    } catch (e) {
      if (e.code !== 'EEXIST') throw new Error(`ロックを作れません: ${e.message}`);
      const seen = lockInfo(lock);
      const age = seen ? Date.now() - seen.mtimeMs : 0;
      if (age > LOCK_STALE_MS) {
        // 落ちたロックの横取り: 消す直前にもう一度見て、見たときと同じもの（目印と時刻）だけを消す
        const again = lockInfo(lock);
        if (again && again.token === seen.token && again.mtimeMs === seen.mtimeMs) {
          try { fs.unlinkSync(lock); } catch { /* 競合 */ }
        }
        continue;
      }
      if (Date.now() - t0 > waitMs) {
        throw new Error('他の科が保存中です。少し待ってもう一度お試しください');
      }
      sleep(20 + Math.floor(Math.random() * 100));
    }
  }
}

// 上書きの前に _履歴\ へ複製する（改名ではなく複製＝他人が書いたファイルでも通る）
function backup(file) {
  if (!fs.existsSync(file)) return;
  const dir = path.join(path.dirname(file), HISTORY);
  try { fs.mkdirSync(dir, { recursive: true }); } catch { return; }
  const d = new Date(), p = (n) => String(n).padStart(2, '0');
  const stamp = `${d.getFullYear()}${p(d.getMonth() + 1)}${p(d.getDate())}_`
    + `${p(d.getHours())}${p(d.getMinutes())}${p(d.getSeconds())}`;
  const base = path.basename(file, '.json');
  let dest = path.join(dir, `${base}_${stamp}.json`);
  for (let i = 2; fs.existsSync(dest) && i < 100; i++) {
    dest = path.join(dir, `${base}_${stamp}_${i}.json`);
  }
  try { fs.copyFileSync(file, dest); } catch { /* best-effort */ }
}

// 差し替え用の一時ファイル名: 同じフォルダの ~<名前>_<拡張子>_<pid>.tmp（tools/master_store.py の tmp_name と同じ）。
//   20261009: 「….json.user002」のような二重拡張子は、院内の EDR がランサムウェアの疑いとして通報する。拡張子を1つにする
function tmpName(file) {
  const ext = path.extname(file);
  const stem = path.basename(file, ext);
  return path.join(path.dirname(file), `~${stem}_${ext.replace(/^\./, '')}_${process.pid}.tmp`);
}

// 一時ファイル → 差し替え。Windows は他が開いている間は差し替えられないので待って繰り返す。
//   20260923: 通らなかったときの「直接上書き」は廃止（書きかけを読まれる・途中で落ちるとマスターが壊れる）。
//   一時ファイルを消してエラーにする＝マスターは前の版のまま。利用者はもう一度提出する
function writeAtomic(file, obj) {
  fs.mkdirSync(path.dirname(file), { recursive: true });
  const data = JSON.stringify(obj);
  const tmp = tmpName(file);
  let last = null;
  try {
    fs.writeFileSync(tmp, data, 'utf8');
    for (let i = 0; i < 10; i++) {
      try { fs.renameSync(tmp, file); return 'replace'; } catch (e) { last = e; sleep(50 * (i + 1)); }
    }
  } catch (e) {
    last = e;
  }
  try { fs.unlinkSync(tmp); } catch { /* 消えていてよい */ }
  throw new Error('保存できませんでした。少し待ってもう一度お試しください'
    + `（共有の当直予定表を差し替えられません: ${(last && (last.code || last.message)) || '不明'}）`);
}

// マスターが無いのに _履歴\ に版が残っている＝本来あるはずのマスターが見えていない（消えた／共有の不調）。
// ここで新しく作ると他の科の予定が全部消えるので、作らずに止める（20260923）
//   hasHistory: _履歴\ に版（.json）があるか。_履歴\ が無い（ENOENT）ときだけ false。読めなければ投げる
function hasHistory(file) {
  const dir = path.join(path.dirname(file), HISTORY);
  let names = [];
  try { names = fs.readdirSync(dir); } catch (e) {
    if (e && e.code === 'ENOENT') return false;
    throw new Error(`共有の当直予定表の履歴が読めません（${e.message}）。少し待ってもう一度お試しください`);
  }
  return names.some((n) => /\.json$/i.test(n));
}
function refuseFreshMaster(file) {
  if (hasHistory(file)) {
    throw new Error(`共有の当直予定表（${path.basename(file)}）が見つかりません。`
      + `master\\${HISTORY}\\ に過去の版があるので、新しくは作りません。管理者に連絡してください`);
  }
}

// renderer から受け取ってよいのは duties と roster だけ（20260925）。
//   baseline（決裁のときの割当）・changes・submitters・confirmed・rev・by などは管理側と store が持つもので、
//   saveDept の Object.assign で renderer に上書きさせない。配列でないものは渡さない（＝前の版を引き継ぐ）
function pickDeptEntry(ent) {
  const e = ent && typeof ent === 'object' ? ent : {};
  const out = {};
  if (Array.isArray(e.duties)) out.duties = e.duties;
  if (Array.isArray(e.roster)) out.roster = e.roster;
  return out;
}

class Conflict extends Error {
  constructor(owner, rev, updated, by) {
    super(`${owner} は ${updated} に ${by || '別の誰か'} が更新しています（rev ${rev}）`);
    this.name = 'Conflict';
    Object.assign(this, { owner, rev, updated, by });
  }
}

// 自科のエントリだけ差し替えて書き戻す。expectRev は読み込んだときの deptRev（null可・false で照合しない）
function saveDept(dataRoot, owner, ent, opt) {
  const o = opt || {};
  const { year, month } = o;
  const file = masterPath(dataRoot, year, month);
  const expect = o.expectRev;
  const missingMsg = () => `R${year - 2018}年${month}月 の共有の当直予定表がありません（${file}）`;
  if (o.mustExist && !statOrNull(file)) throw new Error(missingMsg());   // ロックの前に（master\ を作らない）
  for (let attempt = 0; attempt < 3; attempt++) {
    const release = acquireLock(file);
    try {
      // 無い（ENOENT）ときだけ新しく作る。mustExist（別の月への書き戻し）は作らない
      if (!statOrNull(file)) {
        if (o.mustExist) throw new Error(missingMsg());
        refuseFreshMaster(file);
      }
      const m = loadMaster(dataRoot, year, month);
      const cur = m.depts[owner] || {};
      if (expect !== false && Object.keys(cur).length) {
        if (Number(cur.rev || 0) !== Number(expect || 0)) {
          throw new Conflict(owner, cur.rev, cur.updated, cur.by);
        }
      }
      const next = Object.assign({}, cur, ent);
      next.duties = ent.duties || cur.duties || [];
      next.roster = ent.roster || cur.roster || [];
      next.changes = ent.changes || cur.changes || [];
      if (next.baseline === undefined) next.baseline = cur.baseline || null;
      next.updated = nowIso();
      next.by = o.by || '';
      // この科に出したことのあるログイン名（記録だけ。画面には出さない。20260923）
      const who = String(next.by).split(' / ')[0].trim();
      const prevWho = String(cur.by || '').split(' / ')[0].trim();          // 一覧を持つ前の版は前回の by から始める
      const seen = Array.isArray(cur.submitters) ? cur.submitters.slice()
        : (prevWho && !['app', 'browser'].includes(prevWho) ? [prevWho] : []);
      if (who && !seen.includes(who)) seen.push(who);
      next.submitters = seen;
      next.rev = Number(cur.rev || 0) + 1;
      m.depts[owner] = next;
      m.rev = Number(m.rev || 0) + 1;
      m.updated = next.updated;
      backup(file);
      writeAtomic(file, m);
      const back = loadMaster(dataRoot, year, month).depts[owner] || {};
      if (Number(back.rev || 0) === next.rev) return { ok: true, file, rev: next.rev, master: m };
    } finally { release(); }
    // 書いたはずの版が読み戻せない（ロックを守らない書き手に上書きされた等）。expect はそのまま（20261009。
    //   tools/master_store.py の 20260926 と同じ）: 自科の rev が読んだときのままなら書き直し、
    //   他の人の新しい版になっていれば次の周で Conflict。以前は expect = false で照合を外し、他の人の提出に黙って重ねかねなかった
    sleep(50 + Math.floor(Math.random() * 150));
  }
  throw new Error('マスターに書けませんでした（他の科が保存中です）');
}

module.exports = { loadConfig, masterPath, loadMaster, deptRev, saveDept, Conflict, SCHEMA,
  pickDeptEntry, acquireLock, LOCK_STALE_MS, tmpName };   // acquireLock・LOCK_STALE_MS・tmpName は単体テスト用（tools/test_app_lib.mjs）
