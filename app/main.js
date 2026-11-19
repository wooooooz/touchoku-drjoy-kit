'use strict';
// main.js — Electron main process for 当直勤務表 入力アプリ.
//   管理当直 (管理当直アプリ/app/main.js) の実証済みの作りを踏襲:
//   PORTABLE_EXECUTABLE_DIR / settings.json(dataRoot,lockDir末尾一致) / （1時間自動クローズは 20260925 に廃止）/
//   renderer はファイルシステムに触らず IPC 経由。（デスクトップショートカットは 20260922 に廃止）
//   20261106
//   20260925 追記: 1つだけ起動・ピン留めからの起動を案内・違う月のマスターも読める・renderer が書けるのは duties/roster だけ
//   20260926 追記: td:loadMaster が決裁済みか（confirmed）も返す（開きっぱなしの画面が③確定・解除に追いつくため）。
//   マスターが見えない（無いのに _履歴 に版がある）ときは hidden:true と平易な文を返す
const { app, BrowserWindow, ipcMain, dialog, shell, screen } = require('electron');
const path = require('path');
const os = require('os');
const fs = require('fs');
const { execFileSync } = require('child_process');
const store = require('./lib/store.cjs');
const rights = require('./lib/rights.cjs');

const SETTINGS_NAME = 'touchoku-settings.json';

function appFolder() {
  if (process.env.PORTABLE_EXECUTABLE_DIR) return process.env.PORTABLE_EXECUTABLE_DIR;
  const exeDir = path.dirname(app.getPath('exe'));
  if (process.platform === 'darwin' && /\.app[\\/]Contents[\\/]MacOS$/i.test(exeDir)) {
    return path.resolve(exeDir, '..', '..', '..');
  }
  return exeDir;
}

function settingsHere() {
  const here = appFolder();
  const candidates = [
    path.join(here, SETTINGS_NAME),
    path.join(process.cwd(), SETTINGS_NAME),
    path.join(__dirname, SETTINGS_NAME),
  ];
  for (const c of candidates) {
    try {
      const s = JSON.parse(fs.readFileSync(c, 'utf8'));
      if (s && s.dataRoot) return { here, settings: s, found: true };
    } catch { /* try next */ }
  }
  return { here, settings: {}, found: false };
}

// 起動係（app/launcher）が展開した写し %LOCALAPPDATA%\Touchoku\app\<版>\ から直接起動されたか
//   （タスクバーのピン留め・exe のダブルクリック。20260925）。このとき PORTABLE_EXECUTABLE_DIR が無く、
//   設定も共有の月フォルダも分からないので「設定を読み込めませんでした: ENOENT」になっていた
function runFromLocalCopy() {
  if (process.env.PORTABLE_EXECUTABLE_DIR) return false;
  const la = process.env.LOCALAPPDATA || path.join(os.homedir(), 'AppData', 'Local');
  const base = normDir(path.join(la, 'Touchoku', 'app')) + '/';
  return normDir(path.dirname(app.getPath('exe'))).startsWith(base);
}

function dataRoot() {
  const { here, settings } = settingsHere();
  if (!settings.dataRoot) return here;
  return path.isAbsolute(settings.dataRoot)
    ? settings.dataRoot
    : path.resolve(here, settings.dataRoot);
}

// 起動場所ロック（末尾3階層の一致。絶対一致はマウント差で成立しない=管理当直の実測）
const TAIL_DEPTH = 3;
function normDir(p) { return String(p).replace(/[\\/]+/g, '/').replace(/\/+$/, '').toLowerCase(); }
function tailOf(p, depth) { return normDir(p).split('/').filter(Boolean).slice(-depth).join('/'); }
function locationAllowed() {
  const { here, settings } = settingsHere();
  if (!settings.lockDir) return true;
  const a = normDir(here), b = normDir(settings.lockDir);
  if (a === b) return true;
  if (tailOf(a, TAIL_DEPTH) === tailOf(b, TAIL_DEPTH)) return true;
  dialog.showErrorBox(
    '起動できません',
    'このアプリは所定の共有フォルダからのみ起動できます。\n' +
    'ファイルがコピーまたは移動された可能性があります。\n\n' +
    '共有フォルダ内のアプリ（当直勤務表入力）から起動してください。'
  );
  return false;
}

// 貴科の医師リストを直せる人（部門管理者。20261028 管理者）: 本番の 名簿\roster_editors.json（毎朝3:30に作り直す）。
//   月フォルダの隣の 名簿\ を読む。読めなければ誰も管理者でない扱い（× と区分の切り替えはできない＝安全側）
function loadEditors() {
  try {
    const p = path.join(path.dirname(dataRoot()), '名簿', 'roster_editors.json');
    const d = JSON.parse(fs.readFileSync(p, 'utf8').replace(/^﻿/, ''));
    return { ok: true, editors: d.editors || {} };
  } catch (_) {
    return { ok: false, editors: {} };
  }
}
function myUser() {
  try { return os.userInfo().username; } catch { return process.env.USERNAME || ''; }
}

ipcMain.handle('td:loadConfig', () => {
  const cfg = store.loadConfig(dataRoot());
  // この PC の人が部門管理者か（renderer が × と区分の切り替えの可否に使う。20261028）
  const ed = loadEditors();
  cfg.rosterRights = { ok: ed.ok, user: myUser(), editor: rights.editorOf(ed.editors, myUser()) };
  // 決裁済みの月はマスターの confirmed が立つ（管理ツールの③）→ 変更は交代申請（承認されたものだけ15時に反映。20260924）
  //   20260909 は config\確定.txt を見ていた。20260912 の xlsx一本化でマスターへ移した
  //   20260923: 決裁後も各科は直して出し直せる（閲覧のみにはしない）。confirmed は帯の表示と提出前の確認にだけ使う
  delete cfg.closed;
  try { cfg.confirmed = (store.loadMaster(dataRoot(), cfg.year, cfg.month) || {}).confirmed || null; }
  catch (_) { cfg.confirmed = null; }
  return cfg;
});

// 年月 → その月の置き場。対象月ならこの月フォルダ、違う月は隣の月フォルダ（…\R8年N月）
function monthRoot(year, month, cfg) {
  const root = dataRoot();
  const c = cfg || store.loadConfig(root);
  if (year === Number(c.year) && month === Number(c.month)) return { dir: root, other: false };
  return { dir: path.join(path.dirname(root), `R${year - 2018}年${month}月`), other: true };
}
function validYm(year, month) {
  return Number.isInteger(year) && year >= 2020 && year <= 2100 && Number.isInteger(month) && month >= 1 && month <= 12;
}

// ---- 共有マスター（全科の予定）: アプリは起動のたびにここから読む。ローカルに写しは持たない ----
//   20260925: 違う月（Excel の当直交代）も読める。その月のマスターが無ければ master:null・missing:true
//   （取り込む前に「書き出したあとに変わっていないか」を確かめるため）
//   20260926: confirmed = その月のマスターの confirmed（決裁済みの印・無ければ null）。renderer は科の切替と提出の直前に
//   これを見て、起動したあとに③確定（または解除）されていたら画面を切り替えて知らせる。
//   hidden = マスターが見えない（store.loadMaster の MASTER_HIDDEN）。renderer は error をそのまま出す
ipcMain.handle('td:loadMaster', (_e, q) => {
  try {
    q = q || {};
    const year = Number(q.year), month = Number(q.month);
    if (!validYm(year, month)) return { ok: false, error: `年月が正しくありません（${q.year}年${q.month}月）` };
    const { dir, other } = monthRoot(year, month);
    if (other) {
      try { fs.statSync(store.masterPath(dir, year, month)); } catch (err) {
        if (err && err.code === 'ENOENT') {
          // 無くても _履歴\ に版があれば「見えない」だけ＝loadMaster が MASTER_HIDDEN で投げる（下の catch で hidden:true。20260926）
          store.loadMaster(dir, year, month);
          return { ok: true, master: null, missing: true, dir, confirmed: null };
        }
        throw err;
      }
    }
    const m = store.loadMaster(dir, year, month);
    return { ok: true, master: m, confirmed: (m && m.confirmed) || null };
  } catch (err) {
    return { ok: false, error: String((err && err.message) || err), hidden: !!(err && err.code === 'MASTER_HIDDEN') };
  }
});

// 誰が出したか（Windows のログイン名／PC 名）。当直交代申請.xlsx の「入力者」になる（20260923 管理者: 公式の交代記録）
function whoAmI() {
  let u = '';
  try { u = os.userInfo().username; } catch { u = process.env.USERNAME || ''; }
  return `${u || '?'} / ${os.hostname()}`;
}

// 自科のエントリだけ差し替えて書き戻す（他科は読み直した最新のまま）
//   20260923: by は renderer の申告を使わず、常にこの PC のログイン名（whoAmI）を記録する（交代記録の「入力者」）。
//   対象月と違う月（Excel の当直交代）は、隣の月フォルダ（…\R8年N月）のマスターへ書く。無ければ作らずに断る
ipcMain.handle('td:saveDept', (_e, q) => {
  try {
    q = q || {};
    const year = Number(q.year), month = Number(q.month);
    if (!validYm(year, month)) {
      return { ok: false, error: `年月が正しくありません（${q.year}年${q.month}月）` };
    }
    const owner = typeof q.owner === 'string' ? q.owner : '';
    if (!owner || !q.ent || typeof q.ent !== 'object') return { ok: false, error: '科名または保存する内容がありません' };
    // renderer が書けるのは duties と roster だけ（baseline・submitters・changes などは上書きさせない。20260925）
    const ent = store.pickDeptEntry(q.ent);
    if (!ent.duties && !ent.roster) return { ok: false, error: '保存する内容がありません' };
    const root = dataRoot();
    const cfg = store.loadConfig(root);
    if (!(cfg.frames || []).some((f) => (f.owner || f.dept) === owner)) {
      return { ok: false, error: `この月の枠に無い科です（${owner}）` };
    }
    const mr = monthRoot(year, month, cfg);
    const dir = mr.dir, mustExist = mr.other;
    // 部門管理者（当直表を作る人。Dr.JOY の所属管理者・全体管理者・追加の作成者）だけができること（20261028 管理者）:
    //   決裁前の月 = 当直予定とメンバーの変更すべて。決裁後の月 = 貴科の医師リストで医師を外す・区分を変える
    //   （交代申請と、交代で入る医師を足すことは誰でもできる）
    {
      let mst = null;
      try { mst = store.loadMaster(dir, year, month) || null; } catch (_) { mst = null; }
      const before = ((mst && mst.depts) || {})[owner] || {};
      const editor = rights.editorOf(loadEditors().editors, myUser());
      if (!(mst && mst.confirmed)) {
        if (rights.entryChanged(before, ent) && !rights.canEdit(owner, editor, cfg.deptMap)) {
          return { ok: false, error: 'この科の当直予定を入力・変更できるのは、この科の部門管理者（庶務・部長など、当直表を作る方）だけです。'
            + '変えたいときは、科の庶務の方か事務担当に頼んでください' };
        }
      } else if (Array.isArray(ent.roster)) {
        const df = rights.rosterDiff(before.roster || [], ent.roster);
        if ((df.removed.length || df.changed.length) && !rights.canEditRoster(owner, editor, cfg.shozoku, cfg.deptMap)) {
          return { ok: false, error: '「貴科リスト」で医師を外すこと、区分と「Dr.JOY所属」を変えることは、その科の部門管理者（庶務・部長など、当直表を作る方）だけができます。'
            + '変えたいときは、科の庶務の方か事務担当に頼んでください' };
        }
      }
    }
    if (mustExist) {
      try { fs.statSync(store.masterPath(dir, year, month)); } catch (err) {
        if (err && err.code === 'ENOENT') {
          return { ok: false, error: `R${year - 2018}年${month}月 の共有の当直予定表がありません（${dir}）。その月にはアプリで出した予定がないので、記録できません` };
        }
        throw err;
      }
    }
    const r = store.saveDept(dir, owner, ent,
      { year, month, by: whoAmI(), expectRev: q.expectRev, mustExist });
    // confirmed: 保存した月が決裁済みか（＝交代申請。renderer の文言に使う。20260924）
    return { ok: true, rev: r.rev, file: r.file, confirmed: !!(r.master && r.master.confirmed) };
  } catch (err) {
    if (err && err.name === 'Conflict') {
      return { ok: false, conflict: true, error: err.message, rev: err.rev, updated: err.updated, by: err.by };
    }
    return { ok: false, error: String((err && err.message) || err) };
  }
});

// ---- 前月の 全日タイプ／昼夜タイプ（20260922 管理者決定）----
// まだ何も入っていない月は、前月の実データから同じ規則でタイプを求めて初期表示に使う。
// 隣の月フォルダ（…\R8年N月）の master と config を読む。どちらも無ければ null を返し、
// renderer は枠マスタの固定値（share_day_night）のまま。
const nameKey = (s) => String(s || '').replace(/[ 　]/g, '');

// 枠・日付・duties から グループ（科×役割）ごとの 全日タイプ(true)／昼夜タイプ(false) を求める。
// 判定は土日祝だけ（平日は夜しか無い）。renderer の mismatchDays と同じ規則。
function typesFromDuties(frames, days, duties) {
  const fr = new Map(frames.map((f) => [f.frame_id, f]));
  const bandsOf = (f, d) => (f.shape === '全日' ? (d.holiday ? ['昼', '夜'] : ['夜']) : [f.shape]);
  const groups = new Map();                       // key -> {frames:[], cells:Map}
  for (const f of frames) {
    const k = nameKey(f.dept) + '|' + f.role;
    if (!groups.has(k)) groups.set(k, { frames: [], cells: new Map() });
    groups.get(k).frames.push(f);
  }
  for (const r of duties || []) {
    const f = fr.get(r.frame_id);
    if (!f || (r.band !== '昼' && r.band !== '夜')) continue;
    const g = groups.get(nameKey(f.dept) + '|' + f.role);
    if (!g) continue;
    const ck = `${r.date}|${r.band}`;
    if (!g.cells.has(ck)) g.cells.set(ck, new Set());
    g.cells.get(ck).add((r.oncall ? 'O|' : 'D|') + r.staff_login);
  }
  const eq = (a, b) => {
    const x = a ? [...a].sort() : [], y = b ? [...b].sort() : [];
    return x.length === y.length && x.every((v, i) => v === y[i]);
  };
  const out = {};
  for (const [key, g] of groups) {
    // 昼と夜の両方が立つ日（= 土日祝）がある組だけがスイッチの対象
    const capable = days.some((d) => {
      const m = {};
      for (const f of g.frames) {
        if (f.shape === '昼' && !d.holiday) continue;
        for (const b of bandsOf(f, d)) { if (m[b]) return false; m[b] = f; }
      }
      return !!(m['昼'] && m['夜']);
    });
    if (!capable) continue;
    let same = true;
    for (const d of days) {
      if (!d.holiday) continue;
      if (!eq(g.cells.get(`${d.date}|昼`), g.cells.get(`${d.date}|夜`))) { same = false; break; }
    }
    out[key] = same;
  }
  return out;
}

// config が無い前月のための代用（土日だけを休みとみなす。祝日は拾えない）
function weekendDays(year, month) {
  const p = (n) => String(n).padStart(2, '0');
  const last = new Date(year, month, 0).getDate();
  const out = [];
  for (let day = 1; day <= last; day++) {
    const dt = new Date(year, month - 1, day), w = dt.getDay();
    out.push({ date: `${year}-${p(month)}-${p(day)}`, day, holiday: w === 0 || w === 6 });
  }
  return out;
}

ipcMain.handle('td:loadPrevTypes', (_e, q) => {
  try {
    const year = Number(q.year), month = Number(q.month), owner = String(q.owner || '');
    const py = month === 1 ? year - 1 : year, pm = month === 1 ? 12 : month - 1;
    const dir = path.join(path.dirname(dataRoot()), `R${py - 2018}年${pm}月`);
    const mfile = store.masterPath(dir, py, pm);
    if (!fs.existsSync(mfile)) return { ok: true, types: null, year: py, month: pm };
    const master = store.loadMaster(dir, py, pm);
    const ent = (master.depts || {})[owner];
    if (!ent || !(ent.duties || []).length) return { ok: true, types: null, year: py, month: pm };
    let frames = null, days = null;
    try { const c = store.loadConfig(dir); frames = c.frames; days = c.days; } catch (_) { /* 無ければ代用 */ }
    if (!frames) { try { frames = store.loadConfig(dataRoot()).frames; } catch (_) { return { ok: true, types: null }; } }
    if (!days || !days.length) days = weekendDays(py, pm);
    return { ok: true, types: typesFromDuties(frames, days, ent.duties), year: py, month: pm };
  } catch (err) {
    return { ok: false, error: String((err && err.message) || err) };
  }
});

// ---- 当直表xlsx の書き出し・取り込み（Python の CLI ヘルパーに任せる）----
//   意匠も読取規則も tools/duty_sheet.py が正本。アプリは xlsx を直に触らない。
function helperPath() {
  const here = appFolder();
  for (const c of [path.join(here, 'touchoku_xlsx.exe'),
                   path.join(dataRoot(), 'touchoku_xlsx.exe')]) {
    if (fs.existsSync(c)) return { exe: c, args: [] };
  }
  const dev = path.resolve(__dirname, '..', 'tools', 'touchoku_xlsx.py');   // 開発時
  if (fs.existsSync(dev)) return { exe: 'python', args: ['-X', 'utf8', dev] };
  return null;
}

const HELPER_TIMEOUT_MS = 120 * 1000;
function runHelper(args) {
  const h = helperPath();
  if (!h) return { ok: false, error: '当直表ツール（touchoku_xlsx）が見つかりません' };
  try {
    // 止まったヘルパーで main が固まらないように上限を付ける（20260923。60分の自動クローズも止まるため）
    const out = execFileSync(h.exe, h.args.concat(args), { encoding: 'utf8', windowsHide: true, timeout: HELPER_TIMEOUT_MS });
    return { ok: true, stdout: out };
  } catch (err) {
    if (err && err.code === 'ETIMEDOUT') {
      return { ok: false, error: `当直表ツールが ${HELPER_TIMEOUT_MS / 1000} 秒で終わらなかったので止めました。もう一度お試しください` };
    }
    const msg = String((err && (err.stderr || err.message)) || err).trim().split('\n').pop();
    return { ok: false, error: msg };
  }
}

ipcMain.handle('td:exportXlsx', async (e, q) => {
  const win = BrowserWindow.fromWebContents(e.sender);
  const r = await dialog.showSaveDialog(win, {
    title: '当直表をExcelに書き出す',
    // 保存先の初期値はダウンロード フォルダ（マニュアルが「ダウンロードします」と書くのに合わせる。管理者 20260922）
    defaultPath: path.join(app.getPath('downloads'), q.name || '当直表.xlsx'),
    filters: [{ name: 'Excel', extensions: ['xlsx'] }],
  });
  if (r.canceled || !r.filePath) return { canceled: true };
  const tmp = path.join(app.getPath('temp'), `td_draft_${process.pid}.json`);
  try {
    fs.writeFileSync(tmp, JSON.stringify(q.draft || {}), 'utf8');
    const res = runHelper(['export', '--owner', q.owner, '--ym', q.ym, '--out', r.filePath,
                           '--from', tmp, '--root', path.dirname(dataRoot()),
                           '--config', path.join(dataRoot(), 'config', 'current.json')]);
    return res.ok ? { ok: true, file: r.filePath } : res;
  } catch (err) {
    return { ok: false, error: String((err && err.message) || err) };
  } finally {
    try { fs.unlinkSync(tmp); } catch (_) { /* 消えていてよい */ }
  }
});

// 当直表 xlsx → ヘルパーで読む（ダイアログ経由でも、ドラッグ＆ドロップで渡された path でも同じ）
function importFromPath(filePath) {
  if (!filePath || !/\.xlsx$/i.test(filePath)) return { ok: false, error: '当直表の xlsx を渡してください' };
  if (!fs.existsSync(filePath)) return { ok: false, error: `ファイルがありません: ${filePath}` };
  const tmp = path.join(app.getPath('temp'), `td_import_${process.pid}.json`);
  try {
    const res = runHelper(['import', '--in', filePath, '--out', tmp,
                           '--config', path.join(dataRoot(), 'config', 'current.json')]);
    if (!res.ok) return res;
    const data = JSON.parse(fs.readFileSync(tmp, 'utf8'));
    return { ok: true, file: path.basename(filePath), result: data };
  } catch (err) {
    return { ok: false, error: String((err && err.message) || err) };
  } finally {
    try { fs.unlinkSync(tmp); } catch (_) { /* 消えていてよい */ }
  }
}

// Excel で開いているあいだは、同じフォルダに「~$<ファイル名>」の目印ができる（名前が長いと先頭2文字を ~$ に置き換えた形）。
// 取り込みはディスクに保存された内容しか読まないので、上書き保存を忘れていないか確かめる（管理者 20260922）
function openInExcel(filePath) {
  const dir = path.dirname(filePath), base = path.basename(filePath);
  return ['~$' + base, '~$' + base.slice(2)].some((n) => { try { return fs.existsSync(path.join(dir, n)); } catch (_) { return false; } });
}
async function confirmSaved(win, filePath) {
  if (!filePath || !openInExcel(filePath)) return true;
  const r = await dialog.showMessageBox(win, {
    type: 'warning', title: 'Excel で開いたままです', noLink: true,
    message: 'このファイルは Excel で開いたままです。上書き保存しましたか？',
    detail: '保存していない分は取り込まれません。',
    buttons: ['取り込む', 'やめる'], defaultId: 1, cancelId: 1,
  });
  return r.response === 0;
}

ipcMain.handle('td:importXlsx', async (e) => {
  const win = BrowserWindow.fromWebContents(e.sender);
  const r = await dialog.showOpenDialog(win, {
    title: '当直表のExcelを取り込む',
    filters: [{ name: 'Excel', extensions: ['xlsx'] }],
    properties: ['openFile'],
  });
  if (r.canceled || !r.filePaths.length) return { canceled: true };
  if (!(await confirmSaved(win, r.filePaths[0]))) return { canceled: true };
  return importFromPath(r.filePaths[0]);
});

// ドラッグ＆ドロップ（20260914 ユーザー指示）: renderer が webUtils で取った path を渡してくる
ipcMain.handle('td:importXlsxPath', async (e, q) => {
  if (!(await confirmSaved(BrowserWindow.fromWebContents(e.sender), q && q.path))) return { canceled: true };
  const r = importFromPath(q && q.path);
  if (process.env.TD_DEBUG_DIR) {
    try { fs.appendFileSync(path.join(process.env.TD_DEBUG_DIR, 'renderer.log'),
      `[importXlsxPath] ${q && q.path} → ok=${r.ok} ${r.error || ''} ${r.result ? 'duties=' + (r.result.duties || []).length : ''}\n`, 'utf8'); } catch (_) { /* ignore */ }
  }
  return r;
});


// ---- デスクトップのショートカット: 20260922 管理者の指示で廃止（お節介が過ぎる）----
//   月フォルダごとに exe が違うので、ショートカットは先月の exe を指したままになる問題もあった。
//   以前の版（既定ON）が置いたショートカット `当直勤務表入力.lnk` だけは、起動時に黙って片付ける
//   （向き先がこのアプリの exe のときだけ。利用者が自分で作ったものは触らない）。
const SC_NAME = '当直勤務表入力.lnk';
function scCleanup() {
  try {
    const p = path.join(app.getPath('desktop'), SC_NAME);
    if (!fs.existsSync(p)) return;
    let cur = ''; try { cur = shell.readShortcutLink(p).target || ''; } catch (_) {}
    if (/当直勤務表入力(_R\d+年\d+月|\(Win\))?\.exe$/i.test(cur)) fs.unlinkSync(p);
  } catch (_) { /* ignore */ }
}

// 未提出の変更（20260922 管理者指示）: renderer が入力の変化を知らせてくる。
//   dirty のまま閉じようとしたら「提出ボタンが押されていません」で止める。
//   （60分の自動クローズは 20260925 に廃止。起動係が共有の exe をすぐ閉じるので、開いたままでも更新できる）
let dirty = false, forceClose = false;
// 早見表を画像（PNG）として保存（20261106 管理者）。renderer が渡す位置は CSS px なので、拡大率を掛けて写す
ipcMain.handle('td:saveImage', async (e, q) => {
  const win = BrowserWindow.fromWebContents(e.sender);
  const z = e.sender.getZoomFactor() || 1;
  const rc = q && q.rect || {};
  const rect = { x: Math.max(0, Math.floor(rc.x * z)), y: Math.max(0, Math.floor(rc.y * z)),
                 width: Math.ceil(rc.width * z), height: Math.ceil(rc.height * z) };
  if (!(rect.width > 0 && rect.height > 0)) return { ok: false, error: '写す範囲がありません' };
  try {
    const img = await e.sender.capturePage(rect);
    const name = String((q && q.name) || '画像.png').replace(/[\\/:*?"<>|]/g, '_');
    const r = await dialog.showSaveDialog(win, {
      title: '画像として保存',
      defaultPath: path.join(app.getPath('downloads'), name),
      filters: [{ name: 'PNG 画像', extensions: ['png'] }],
    });
    if (r.canceled || !r.filePath) return { canceled: true };
    fs.writeFileSync(r.filePath, img.toPNG());
    return { ok: true, file: r.filePath };
  } catch (err) {
    return { ok: false, error: String(err && err.message || err) };
  }
});
ipcMain.on('td:setDirty', (_e, flag) => { dirty = !!flag; });

// 最初の窓の大きさ（20260927）: 画面の作業域（タスクバーを除く）に収める。
//   院内の約半分のノート PC は 1920×1200・表示の拡大 150% ＝ 作業域が約 1280×768。決め打ちの 860 だと下の提出ボタンが画面の外に出ていた。
//   作業域が狭い PC（幅 1440 未満）は最大化して開く。広い PC は 1600×1000 まで・作業域の 92% まで
function initialBounds() {
  let wa = { x: 0, y: 0, width: 1280, height: 860 };
  try { wa = screen.getPrimaryDisplay().workArea; } catch (_) { /* 取れなければ従来の大きさ */ }
  const width = Math.min(1600, Math.round(wa.width * 0.92));
  const height = Math.min(1000, Math.round(wa.height * 0.92));
  return { width, height, maximize: wa.width < 1440 || wa.height < 820 };
}

function createWindow() {
  const ib = initialBounds();
  const win = new BrowserWindow({
    width: ib.width,
    height: ib.height,
    center: true,
    backgroundColor: '#eef0f5',
    minWidth: 900,
    minHeight: 600,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
    },
  });
  win.setMenuBarVisibility(false);
  if (ib.maximize) win.maximize();
  win.on('close', (e) => {
    if (!dirty || forceClose) return;
    e.preventDefault();
    dialog.showMessageBox(win, {
      type: 'warning', title: '提出されていません', noLink: true,
      message: '提出ボタンが押されていません',
      detail: '変更した内容はまだ共有の当直予定表に保存されていません。閉じると失われます。それでも閉じますか？',
      buttons: ['はい、閉じる', 'いいえ（戻る）'], defaultId: 1, cancelId: 1,
    }).then((r) => { if (r.response === 0) { forceClose = true; win.close(); } });
  });
  win.loadFile(path.join(__dirname, 'renderer', 'index.html'));
  win.webContents.once('did-finish-load', () => setTimeout(scCleanup, 800));
  // 開発・検証用（20260914）: TD_DEBUG_DIR を与えると renderer のコンソールを renderer.log に、
  // 8秒後の画面を shot.png に落とす。GUI を目で見られない環境（自動検証）で使う。本番の起動には影響しない
  const dbg = process.env.TD_DEBUG_DIR;
  if (dbg) {
    try { fs.mkdirSync(dbg, { recursive: true }); } catch (_) { /* ignore */ }
    const logf = path.join(dbg, 'renderer.log');
    const put = (s) => { try { fs.appendFileSync(logf, s + '\n', 'utf8'); } catch (_) { /* ignore */ } };
    put(`[start ${new Date().toISOString()}] dataRoot=${dataRoot()}`);
    win.webContents.on('console-message', (_e, level, message, line, sourceId) =>
      put(`[console L${level}] ${message} (${path.basename(String(sourceId))}:${line})`));
    win.webContents.on('render-process-gone', (_e, d) => put(`[gone] ${JSON.stringify(d)}`));
    // TD_DEBUG_OWNER に科名を与えると 3 秒後にその科を選んだことにする（例外は console.error に流す）
    win.webContents.once('did-finish-load', () => setTimeout(() => {
      const o = process.env.TD_DEBUG_OWNER || '';
      win.webContents.executeJavaScript(`
        window.addEventListener('error', e => console.error('[error] ' + e.message + ' @' + e.filename + ':' + e.lineno));
        window.addEventListener('unhandledrejection', e => console.error('[rejection] ' + (e.reason && (e.reason.stack || e.reason))));
        (async () => { try {
          const sel = document.querySelector('#deptSel');
          if (${JSON.stringify(o)}) { sel.value = ${JSON.stringify(o)}; sel.dispatchEvent(new Event('change')); }
          await new Promise(r => setTimeout(r, 2500));
          const cal = document.querySelector('#gridSec');
          console.log('[probe] owner=' + sel.value + ' gridSec html length=' + (cal ? cal.innerHTML.length : -1)
            + ' cells=' + document.querySelectorAll('#gridSec select').length
            + ' members=' + document.querySelectorAll('#memberSec tbody tr').length
            + ' footer=' + ((document.querySelector('#chkSummary') || {}).textContent || '').trim().slice(0, 200)
            + '\\n[probe-detail]\\n' + ((document.querySelector('#chkSummary') || {}).title || '').slice(0, 1500));
          // 昼夜共通スイッチの初期状態（20260922: 月ごとにデータから決める の検証用）
          console.log('[probe-share] ' + ([...document.querySelectorAll('#shareBar .shsw')]
            .map(b => b.textContent.trim().replace(/\\s+/g, ' ')).join(' / ') || '(なし)'));
          const imp = ${JSON.stringify(process.env.TD_DEBUG_IMPORT || '')};
          const dbgImport = (window.__tdDebug || {}).doImport;
          console.log('[probe] doImport available=' + (typeof dbgImport === 'function'));
          if (imp && typeof dbgImport === 'function') {        // ドラッグ＆ドロップと同じ経路（path を直接渡す）
            window.alert = m => console.log('[alert] ' + m);     // モーダルは自動検証を止めるので記録だけにする
            window.confirm = m => { console.log('[confirm] ' + m); return true; };
            await dbgImport(imp);
            await new Promise(r => setTimeout(r, 1500));
            const ov = document.querySelector('#importOv');
            console.log('[probe-import] overlay=' + (ov ? ov.style.display : 'none')
              + ' ok=' + ((document.querySelector('#impOk') || {}).textContent || '').trim().slice(0, 200)
              + ' rejs=' + document.querySelectorAll('#impRejs .rej').length);
          }
        } catch (err) { console.error('[probe] ' + (err && (err.stack || err))); } })();
      `).catch(err => put(`[exec] ${err && err.message}`));
      // TD_DEBUG_JS にファイルを与えると 6 秒後にその中身を renderer で実行する（画面操作の再現用）
      const extra = process.env.TD_DEBUG_JS;
      if (extra) setTimeout(() => {
        let src = ''; try { src = fs.readFileSync(extra, 'utf8'); } catch (e) { put(`[debug-js] ${e.message}`); return; }
        win.webContents.executeJavaScript(src).catch(err => put(`[debug-js] ${err && err.message}`));
      }, 3000);
    }, 3000));
    setTimeout(async () => {
      try {
        const img = await win.webContents.capturePage();
        fs.writeFileSync(path.join(dbg, 'shot.png'), img.toPNG());
        put('[shot] saved');
      } catch (err) { put(`[shot] failed: ${err && err.message}`); }
    }, Number(process.env.TD_DEBUG_SHOT_MS) || 8000);
  }
  return win;
}

// ---- 1つだけ起動（20260925）: 2つ目の起動は、開いている画面を前に出して終わる ----
//   ロックは Electron の userData 単位＝どの月の exe から起動しても同じアプリなので、月をまたいでも1つ。
//   別の月の exe から起動されたときは、開いている画面を前に出し「先にこの画面を閉じて」と知らせる
let mainWin = null;
const gotSingle = app.requestSingleInstanceLock({ dir: appFolder(), hasSettings: settingsHere().found });
if (!gotSingle) {
  app.quit();
} else {
  app.on('second-instance', (_e, _argv, _cwd, extra) => {
    const win = mainWin && !mainWin.isDestroyed() ? mainWin : null;
    if (!win) return;
    if (win.isMinimized()) win.restore();
    win.show(); win.focus();
    const d = extra && extra.dir ? String(extra.dir) : '';
    if (d && extra.hasSettings && normDir(d) !== normDir(appFolder())) {
      dialog.showMessageBox(win, {
        type: 'info', title: '当直勤務表入力', noLink: true,
        message: `${path.basename(appFolder())} の画面がすでに開いています`,
        detail: `${path.basename(d)} を開くには、先にこの画面を閉じてから、もう一度開いてください。`,
        buttons: ['OK'],
      }).catch(() => {});
    }
  });
}

app.whenReady().then(() => {
  if (!gotSingle) return;
  // 起動係が展開した写しを直接起動された（タスクバーのピン留めなど）→ 設定が無いので、正しい起動のしかたを知らせて終わる
  if (!settingsHere().found && runFromLocalCopy()) {
    dialog.showMessageBoxSync({
      type: 'warning', title: '当直勤務表入力', noLink: true, buttons: ['OK'],
      message: 'このアプリは共有フォルダの 当直勤務表入力_R8年N月.exe から起動してください' +
        '（タスクバーのピン留めからは起動できません）。',
    });
    app.quit(); return;
  }
  if (!locationAllowed()) { app.quit(); return; }
  mainWin = createWindow();
  // 55分の警告と60分の自動クローズは廃止（管理者 20260925）。もとは開きっぱなしで共有の exe が使用中になり
  // 更新を配れないのを避けるためだったが、起動係（app/launcher）が各 PC の写しを起動して共有の exe をすぐ閉じるので不要になった
});
app.on('window-all-closed', () => app.quit());
