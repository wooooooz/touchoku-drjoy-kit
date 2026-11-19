// drjoy-stay-alert-export.js  20260926
// Export Dr.JOY's 滞在アラート list (at0052 アラート管理, 種別=滞在アラート) as CSV
// without screenshots. Used by the 働き方レポート (P11, 置きっぱなし).
//
// Flow on the screen: 種別 = BEACON_STAY_ALERT, 月度 = --month (the screen only
// offers last, this and next month), "csv" -> a modal asks for the format
// (CSV(Shift_JIS) is the default) -> OK -> POST
// /api/dr/at/alert/over_time_summary/download returns a downloadPath -> the page
// fetches it from Firebase Storage, makes a blob download and deletes the
// temporary file. As with at0006, the file is taken from the page's fetch
// (browser downloads stay off) and the double encoding is undone.
// The threshold ("同一エリア内にN時間以上") is a Dr.JOY setting, not an option here.
//
// Usage:
//   node drjoy-stay-alert-export.js --month 2026-09 [--out-dir <dir>] [--headless]
// Output: <out-dir>\滞在アラート_YYYY_M.csv (Shift_JIS; " (N)" added when taken)
//   OK: <full path> (<n> 行, Shift_JIS)
// Exit codes: 0 ok / 1 error / 2 login timeout / 3 CSV rejected by validation

'use strict';
const fs = require('fs');
const path = require('path');
const { DEFAULT_OUT_DIR, log, launch, waitForScreen, toShiftJisBytes, uniquePath } = require('./lib/drjoy-common');

const AT0052_URL = 'https://app.drjoy.jp/at/at0052';
const DOWNLOAD_API = '/api/dr/at/alert/over_time_summary/download';
const STORAGE_RE = /firebasestorage\.googleapis\.com|storage\.googleapis\.com/;
const REQUIRED_HEADERS = ['氏名', '回数'];

function parseArgs(argv) {
  const opts = { month: null, outDir: DEFAULT_OUT_DIR, headless: false, loginTimeoutSec: 300 };
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if (a === '--month') {
      const m = /^(\d{4})-?(\d{2})$/.exec(argv[++i] || '');
      if (!m) throw new Error(`--month must be YYYY-MM: ${argv[i]}`);
      opts.month = { ym: m[1] + m[2], year: Number(m[1]), month: Number(m[2]) };
    } else if (a === '--out-dir') opts.outDir = argv[++i];
    else if (a === '--headless') opts.headless = true;
    else if (a === '--login-timeout') opts.loginTimeoutSec = Number(argv[++i]);
    else throw new Error(`Unknown argument: ${a}`);
  }
  if (!opts.month) throw new Error('--month is required');
  if (!fs.existsSync(opts.outDir)) throw new Error(`Output folder not found: ${opts.outDir}`);
  return opts;
}

async function main() {
  const opts = parseArgs(process.argv.slice(2));
  const { ym, year, month } = opts.month;
  const expectedName = `滞在アラート_${year}_${month}.csv`;
  log(`対象: 滞在アラート ${year}年${month}月（Shift_JIS）`);

  const context = await launch(opts.headless);
  try {
    const page = context.pages()[0] || (await context.newPage());
    await page.goto(AT0052_URL, { waitUntil: 'domcontentloaded' });
    const typeSelect = page.getByTestId('at0052-select-type');
    const ready = await waitForScreen(page, {
      url: AT0052_URL, urlPart: '/at/at0052', readyLocator: typeSelect,
      timeoutSec: opts.loginTimeoutSec, headless: opts.headless,
    });
    if (!ready) {
      log(`${opts.loginTimeoutSec} 秒待ってもアラート管理の画面が出ませんでした。`);
      process.exitCode = 2;
      return;
    }
    log('アラート管理の画面 (at0052) を確認しました。');
    await page.waitForLoadState('networkidle').catch(() => {});

    await typeSelect.selectOption('BEACON_STAY_ALERT');
    const monthSelect = page.getByTestId('at0052-select-month-beacon');
    await monthSelect.waitFor({ state: 'visible', timeout: 15000 });
    const offered = await monthSelect.locator('option').evaluateAll((os) => os.map((o) => o.value));
    if (!offered.includes(ym)) throw new Error(`The screen does not offer ${ym} (only ${offered.join(', ')}; it keeps last, this and next month).`);
    await monthSelect.selectOption(ym);
    await page.waitForTimeout(2000);

    // Make sure the request asks for this alert type, month and Shift_JIS.
    await page.route(`**${DOWNLOAD_API}*`, async (route) => {
      let body;
      try {
        body = JSON.parse(route.request().postData() || '');
      } catch {
        log('CSV の要求が JSON ではありませんでした。送らずに止めます。');
        await route.abort();
        return;
      }
      const sent = { ...body, option: 'BEACON_STAY_ALERT', year, month: String(month), encode: 1 };
      log(`CSV の要求: option=${sent.option} year=${sent.year} month=${sent.month} encode=${sent.encode} accountStatus="${sent.accountStatus}"`);
      await route.continue({ postData: JSON.stringify(sent) });
    });

    const info = page.waitForResponse((r) => r.url().includes(DOWNLOAD_API), { timeout: 60000 });
    const media = page.waitForResponse(
      (r) => STORAGE_RE.test(r.url()) && r.status() === 200 && !/json/.test(r.headers()['content-type'] || ''),
      { timeout: 90000 },
    );
    await page.getByTestId('at0052-btn-export-csv').click();
    const dlg = page.locator('.modal.show, ngb-modal-window').filter({ hasText: 'ファイルの形式を選択してください' }).first();
    await dlg.waitFor({ state: 'visible', timeout: 15000 });
    const sjis = dlg.locator('label', { hasText: 'CSV(Shift_JIS)' }).first();
    if (await sjis.count()) await sjis.click();
    await dlg.locator('button', { hasText: /^\s*OK\s*$/ }).first().click();

    const infoResp = await info;
    if (infoResp.status() !== 200) throw new Error(`CSV request failed: HTTP ${infoResp.status()}`);
    const { downloadPath } = await infoResp.json();
    if (!downloadPath) throw new Error('Dr.JOY returned no downloadPath.');
    const resp = await media;
    if (!decodeURIComponent(resp.url()).includes(downloadPath)) throw new Error('Received a different file from Storage; nothing was saved.');
    const fileName = path.basename(downloadPath);
    const bytes = toShiftJisBytes(await resp.body(), REQUIRED_HEADERS[0]);

    const lines = new TextDecoder('shift_jis').decode(bytes).split(/\r?\n/).filter((l) => l.length > 0);
    const problems = [];
    if (fileName !== expectedName) problems.push(`file name "${fileName}" is not ${expectedName}`);
    for (const h of REQUIRED_HEADERS) if (!(lines[0] || '').includes(h)) problems.push(`header "${h}" not found when decoded as Shift_JIS`);
    if (problems.length > 0) {
      const rejected = path.join(opts.outDir, `rejected-stay-alert-${Date.now()}.csv`);
      fs.writeFileSync(rejected, bytes);
      log(`NG: ${problems.join(' / ')}`);
      log(`下流に拾われないよう ${path.basename(rejected)} として保存しました。`);
      process.exitCode = 3;
      return;
    }
    const tmp = path.join(opts.outDir, `.stay-alert-downloading-${Date.now()}`);
    fs.writeFileSync(tmp, bytes);
    const target = uniquePath(opts.outDir, fileName);
    fs.renameSync(tmp, target);
    log(`OK: ${target} (${lines.length} 行, Shift_JIS)`);
  } finally {
    await context.close();
  }
}

main().catch((err) => {
  log(`ERROR: ${err.message}`);
  process.exitCode = 1;
});
