// drjoy-at0069-import.js  20261108（--type で種別を選べる）
// Import a 勤務予定 xlsx through Dr.JOY's own import screen (at0069), without
// screenshots. This WRITES to Dr.JOY (the schedule), through the screen
// as the policy requires (internal APIs are for reading only).
//
// Flow: open at0069 -> make sure the type is 勤務予定 (IMPORT_SCHEDULE) ->
// remember the existing import log ids -> hand the file to the hidden
// input[type=file] (data-testid="at0069-input-file"; the visible インポート
// button would open the OS file dialog) -> accept the confirmation (native
// dialog, or a modal's OK button) -> poll the import log with the reload
// button until a NEW log for this file name reaches FINISH -> print counts.
// When rows failed, the error report is taken from the page's fetch of
// Firebase Storage if the row offers it.
//
// ⚠ Not yet exercised against a live import (2026-09-26): the shape of the
//   confirmation dialog is unknown. Watch the first live run.
//
// Usage:
//   node drjoy-at0069-import.js --file <xlsx> [--type IMPORT_SCHEDULE|SUB_JOB] [--headless] [--wait-timeout <sec>]
//   (--type 省略=勤務予定。SUB_JOB=兼業・副業予定で .csv も可。種別の値は画面の選択肢から読んだ)
//   node drjoy-at0069-import.js --dry-run [--headless]     (open the screen and read the log only)
// Output: last lines
//   OK: <file> total=<n> success=<n> error=<n>
//   ERRORS=<saved error report path>          (only when rows failed and the report was saved)
// Exit codes: 0 all rows ok / 1 error (import may not have run; see output)
//             2 login timeout / 5 finished with failed rows

'use strict';
const fs = require('fs');
const path = require('path');
const { log, sleep, launch, waitForScreen } = require('./lib/drjoy-common');

const AT0069_URL = 'https://app.drjoy.jp/at/at0069';
const LIST_API = '/api/dr/at/import/list';
const STORAGE_RE = /firebasestorage\.googleapis\.com|storage\.googleapis\.com/;

function parseArgs(argv) {
  const opts = { file: null, type: 'IMPORT_SCHEDULE', headless: false, dryRun: false, waitTimeoutSec: 900, loginTimeoutSec: 300 };
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if (a === '--file') opts.file = path.resolve(argv[++i]);
    else if (a === '--type') opts.type = argv[++i];
    else if (a === '--headless') opts.headless = true;
    else if (a === '--dry-run') opts.dryRun = true;
    else if (a === '--wait-timeout') opts.waitTimeoutSec = Number(argv[++i]);
    else if (a === '--login-timeout') opts.loginTimeoutSec = Number(argv[++i]);
    else throw new Error(`Unknown argument: ${a}`);
  }
  if (!opts.dryRun) {
    if (!opts.file) throw new Error('--file is required (or --dry-run)');
    if (!fs.existsSync(opts.file)) throw new Error(`File not found: ${opts.file}`);
    const ext = path.extname(opts.file).toLowerCase();
    if (ext !== '.xlsx' && !(opts.type === 'SUB_JOB' && ext === '.csv')) throw new Error('Only .xlsx can be imported (SUB_JOB also accepts .csv)');
  }
  return opts;
}

const describe = (l) => `${l.status} 受付 ${l.importReceptionDate} 読み込み ${l.totalCount} 成功 ${l.successCount} 失敗 ${l.errorCount} ${path.basename(l.importFilePath || '')}`;

async function main() {
  const opts = parseArgs(process.argv.slice(2));
  const context = await launch(opts.headless);
  try {
    const page = context.pages()[0] || (await context.newPage());
    const typeSelect = page.getByTestId('at0069-select-import-type');
    const reloadBtn = page.getByTestId('at0069-btn-reload');
    const isList = (r) => r.request().method() === 'GET' && r.status() === 200 && r.url().includes(LIST_API);
    const readLogs = async (r) => ((await r.json()).importLogs) || [];
    // Wait and click together: if the click fails, the pending wait must not become an unhandled rejection
    // (2026-10-01 first live run: the whole script died there right after handing over the file).
    const fetchLogs = async () => {
      const [resp] = await Promise.all([page.waitForResponse(isList, { timeout: 30000 }), reloadBtn.click({ timeout: 20000 })]);
      return readLogs(resp);
    };

    // Writes still in flight (the upload and its POSTs). The screen must not be
    // reopened while one is pending, or the import could be cut off.
    let pendingWrites = 0;
    const isWrite = (req) => req.method() !== 'GET' && (req.url().includes('/api/dr/at/') || STORAGE_RE.test(req.url()));
    page.on('request', (req) => { if (isWrite(req)) pendingWrites++; });
    const writeDone = (req) => { if (isWrite(req)) pendingWrites = Math.max(0, pendingWrites - 1); };
    page.on('requestfinished', writeDone);
    page.on('requestfailed', writeDone);

    // Poll the log. 2026-10-28: three imports finished on Dr.JOY while the poll kept
    // reading nothing ("not in the log yet") because a failed reload click was
    // swallowed into an empty list. Now a failed click is logged, and the screen is
    // reopened (after the upload is done) to read the log from its first fetch.
    let reloadFailNoted = false;
    const pollLogs = async () => {
      try {
        return await fetchLogs();
      } catch (e) {
        if (!reloadFailNoted) {
          log(`更新ボタンで履歴を読めませんでした（${e.message.split('\n')[0]}）。以後は画面を開き直して読みます。`);
          reloadFailNoted = true;
        }
      }
      for (let i = 0; pendingWrites > 0 && i < 120; i++) await sleep(1000);
      if (pendingWrites > 0) throw new Error(`upload still in flight (${pendingWrites} request(s)); not reopening the screen yet`);
      const [resp] = await Promise.all([
        page.waitForResponse(isList, { timeout: 60000 }),
        page.goto(AT0069_URL, { waitUntil: 'domcontentloaded' }),
      ]);
      return readLogs(resp);
    };

    const firstList = page.waitForResponse(isList, { timeout: 60000 }).catch(() => null);
    await page.goto(AT0069_URL, { waitUntil: 'domcontentloaded' });
    const ready = await waitForScreen(page, {
      url: AT0069_URL, urlPart: '/at/at0069', readyLocator: typeSelect,
      timeoutSec: opts.loginTimeoutSec, headless: opts.headless,
    });
    if (!ready) {
      log(`${opts.loginTimeoutSec} 秒待ってもデータインポート画面が出ませんでした。`);
      process.exitCode = 2;
      return;
    }
    log('データインポート画面 (at0069) を確認しました。');
    if ((await typeSelect.inputValue()) !== opts.type) {
      await typeSelect.selectOption(opts.type);
      await sleep(2000);
    }
    const firstResp = await firstList;
    const before = firstResp ? await readLogs(firstResp) : await fetchLogs();
    const knownIds = new Set(before.map((l) => l.id));
    log(`取り込み履歴 ${before.length} 件（直近: ${before[0] ? describe(before[0]) : 'なし'}）`);

    if (opts.dryRun) {
      const input = page.getByTestId('at0069-input-file');
      log(`dry-run: ファイル欄 ${await input.count()} 個（accept=${await input.getAttribute('accept')}）。取り込みはしません。`);
      return;
    }

    // Accept whatever confirmation appears.
    page.on('dialog', async (d) => {
      log(`確認ダイアログ（${d.type()}）: ${d.message().replace(/\s+/g, ' ').slice(0, 120)} → OK`);
      await d.accept().catch(() => {});
    });

    const baseName = path.basename(opts.file);
    const isNewLog = (l) => !knownIds.has(l.id) && path.basename(l.importFilePath || '') .endsWith(baseName);

    log(`ファイルを渡します: ${baseName}`);
    await page.getByTestId('at0069-input-file').setInputFiles(opts.file);

    // A modal confirmation (not a native dialog) would need a click. Give the
    // page a moment, then press the modal's primary button if one is showing.
    await sleep(3000);
    const modal = page.locator('.modal.show, [role="dialog"]:visible').first();
    if (await modal.isVisible().catch(() => false)) {
      const text = (await modal.innerText().catch(() => '')).replace(/\s+/g, ' ').slice(0, 160);
      const ok = modal.locator('button', { hasText: /^\s*(OK|ＯＫ|はい|インポート|実行|確認)\s*$/ }).first();
      if (await ok.count()) {
        log(`確認の画面: ${text} → 「${(await ok.innerText()).trim()}」を押します。`);
        await ok.click();
      } else {
        log(`確認の画面に押すボタンが見つかりません: ${text}`);
      }
    }

    // 20261108: the import log list this script reads is the 勤務予定 list; a 兼業・副業予定 (SUB_JOB)
    // import never shows up there (confirmed: the import ran, the list stayed unchanged). For other types,
    // stop here and let the caller verify the result with an export.
    if (opts.type !== 'IMPORT_SCHEDULE') {
      log(`SUBMITTED: ${opts.file} (type ${opts.type}). 取り込み履歴では確認できないので、書き出し（at0029）で結果を確かめてください。`);
      await sleep(20000);
      return;
    }

    // Wait for our log to appear and finish.
    const deadline = Date.now() + opts.waitTimeoutSec * 1000;
    let mine = null;
    let lastState = '';
    let lastError = '';
    while (Date.now() < deadline) {
      await sleep(15000);
      const logs = await pollLogs().catch((e) => {
        const msg = e.message.split('\n')[0];
        if (msg !== lastError) { log(`履歴を読めませんでした: ${msg}`); lastError = msg; }
        return [];
      });
      mine = logs.find(isNewLog) || null;
      const state = mine ? mine.status : '(not in the log yet)';
      if (state !== lastState) { log(`状態: ${mine ? describe(mine) : state}`); lastState = state; }
      if (mine && mine.status === 'FINISH') break;
      if (mine && /ERR|FAIL/i.test(mine.status || '')) break;
    }
    if (!mine) throw new Error(`The import did not appear in the log within ${opts.waitTimeoutSec} seconds${lastError ? ` (last read error: ${lastError})` : ''}; check at0069 by hand before retrying.`);
    if (mine.status !== 'FINISH') throw new Error(`The import did not finish (${describe(mine)}); check at0069 by hand before retrying.`);

    log(`OK: ${opts.file} total=${mine.totalCount} success=${mine.successCount} error=${mine.errorCount}`);
    if (!mine.errorCount) return;

    // Rows failed: try to save the error report next to the imported file.
    process.exitCode = 5;
    try {
      const logs = await pollLogs();
      const idx = logs.findIndex((l) => l.id === mine.id);
      const row = page.locator('table tbody tr').nth(idx);
      const trigger = row.locator('a, button, i, span').filter({ has: page.locator('[class*="danger"], [class*="error"], [class*="red"]') })
        .or(row.locator('a[class*="danger"], button[class*="danger"], i[class*="danger"], span[class*="danger"], [class*="error"]')).first();
      const media = page.waitForResponse((r) => STORAGE_RE.test(r.url()) && r.status() === 200 && !/json/.test(r.headers()['content-type'] || ''), { timeout: 30000 });
      await trigger.click({ timeout: 10000 });
      const body = await (await media).body();
      const out = opts.file.replace(/\.xlsx$/i, '_エラー報告.txt');
      fs.writeFileSync(out, body);
      log(`ERRORS=${out}`);
    } catch (err) {
      log(`エラー報告を自動で落とせませんでした（${err.message.split('\n')[0]}）。at0069 の該当行から手で落としてください: ${mine.errorFilePath}`);
    }
  } finally {
    await context.close();
  }
}

main().catch((err) => {
  log(`ERROR: ${err.message}`);
  process.exitCode = 1;
});
