// drjoy-request-list.js  20260926
// Fetch Dr.JOY's request list (申請一覧, including pending RS_NEW) for whole
// months and save it as JSON, without screenshots. Playwright port of
// the earlier request-list script (same output shape and
// file name, so drjoy_kotai_sync.py --req reads it unchanged).
//
// The at0022 (承認) screen calls POST /api/dr/at/admin_request/list_request
// with a bearer token; cookies alone get 401. The token is read from the
// page's own request headers and kept in memory only (never printed or saved).
// Read-only: the screen's 承認 / 一括承認 / 非承認 buttons are never touched;
// at most the search form's "リセット" is clicked once to make the page load
// the list.
//
// Usage:
//   node drjoy-request-list.js --month 2026-11 [--month 2026-12] [--out-dir <dir>] [--headless]
// Output: <out-dir>\DrJOY申請一覧_YYYY-MM_<UTC stamp>.json per month, and a line
//   OK: <full path> (<n> 件 / totalRecords <t>)
// Exit codes: 0 ok / 1 error / 2 login timeout

'use strict';
const fs = require('fs');
const path = require('path');
const { DEFAULT_OUT_DIR, log, sleep, launch, waitForScreen } = require('./lib/drjoy-common');

const AT0022_URL = 'https://app.drjoy.jp/at/at0022';
const LIST_URL = 'https://app.drjoy.jp/api/dr/at/admin_request/list_request';
const PAGE_SIZE = 100;

function parseArgs(argv) {
  const opts = { months: [], outDir: DEFAULT_OUT_DIR, headless: false, loginTimeoutSec: 300 };
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if (a === '--month') {
      const m = /^(\d{4})-?(\d{2})$/.exec(argv[++i] || '');
      if (!m) throw new Error(`--month must be YYYY-MM: ${argv[i]}`);
      opts.months.push(`${m[1]}-${m[2]}`);
    } else if (a === '--out-dir') opts.outDir = argv[++i];
    else if (a === '--headless') opts.headless = true;
    else if (a === '--login-timeout') opts.loginTimeoutSec = Number(argv[++i]);
    else throw new Error(`Unknown argument: ${a}`);
  }
  if (opts.months.length === 0) throw new Error('--month is required');
  if (!fs.existsSync(opts.outDir)) throw new Error(`Output folder not found: ${opts.outDir}`);
  return opts;
}

async function main() {
  const opts = parseArgs(process.argv.slice(2));
  const context = await launch(opts.headless);
  try {
    const page = context.pages()[0] || (await context.newPage());

    // Remember the bearer token from any Dr.JOY API call the page makes.
    let token = null;
    page.on('request', (r) => {
      if (!r.url().startsWith('https://app.drjoy.jp/api/')) return;
      const h = r.headers().authorization;
      if (h && /^Bearer\s/i.test(h)) token = h;
    });

    await page.goto(AT0022_URL, { waitUntil: 'domcontentloaded' });
    const resetBtn = page.locator('button', { hasText: /^\s*リセット\s*$/ }).first();
    const ready = await waitForScreen(page, {
      url: AT0022_URL, urlPart: '/at/at0022', readyLocator: resetBtn,
      timeoutSec: opts.loginTimeoutSec, headless: opts.headless,
    });
    if (!ready) {
      log(`${opts.loginTimeoutSec} 秒待っても承認の画面が出ませんでした。`);
      process.exitCode = 2;
      return;
    }
    log('承認の画面 (at0022) を確認しました。');
    await page.waitForLoadState('networkidle').catch(() => {});
    if (!token) {
      log('一覧を読み直させるため「リセット」を1回押します。');
      await resetBtn.click();
      for (let i = 0; i < 20 && !token; i++) await sleep(500);
    }
    if (!token) throw new Error('Could not pick up the API token from the page.');

    const post = async (month, lastDay, p) => {
      const r = await context.request.post(LIST_URL, {
        headers: { Authorization: token, 'Content-Type': 'application/json; charset=utf-8', Accept: 'application/json' },
        data: {
          keyword: '', officeUserIds: '', departmentId: '', type: '', status: '',
          startDate: `${month}-01T00:00:00+0900`, endDate: `${month}-${lastDay}T00:00:00+0900`,
          page: p, size: PAGE_SIZE, dayOff: false, showLockDeleteUser: false, jobTitle: '',
          shouldOnlyDisplayConfiguredRequestOfUsers: false, option: 'OVERTIME', newLogicJobTitle: true, sort: '-editDate',
        },
      });
      if (r.status() !== 200) throw new Error(`list_request HTTP ${r.status()}`);
      return r.json();
    };

    for (const month of opts.months) {
      const [yy, mm] = month.split('-').map(Number);
      const lastDay = String(new Date(yy, mm, 0).getDate()).padStart(2, '0');
      // lastPage is unreliable: page until empty or totalRecords reached. The
      // order is -editDate, so edits during the run shift pages; dedupe by
      // requestId and make a second pass if still short.
      const seen = new Map();
      let total = null;
      for (let pass = 0; pass < 2 && (total === null || seen.size < total); pass++) {
        for (let p = 0; p < 400; p++) {
          const r = await post(month, lastDay, p);
          total = r.totalRecords;
          const items = r.requests || [];
          if (items.length === 0) break;
          for (const q of items) seen.set(q.requestId, q);
          if (seen.size >= total) break;
          await sleep(250);
        }
      }
      const stamp = new Date().toISOString().replace(/[-:]/g, '').slice(0, 13);
      const file = path.join(opts.outDir, `DrJOY申請一覧_${month}_${stamp}.json`);
      fs.writeFileSync(file, JSON.stringify({ month, fetched: new Date().toISOString(), totalRecords: total, requests: [...seen.values()] }));
      if (seen.size !== total) log(`注意: ${month} は ${seen.size} 件で、totalRecords ${total} と合いません。`);
      log(`OK: ${file} (${seen.size} 件 / totalRecords ${total})`);
    }
  } finally {
    await context.close();
  }
}

main().catch((err) => {
  log(`ERROR: ${err.message}`);
  process.exitCode = 1;
});
