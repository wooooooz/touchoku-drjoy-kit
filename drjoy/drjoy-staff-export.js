// drjoy-staff-export.js  20260926
// Export the Dr.JOY staff list (at0006) as CSV without screenshots.
// Drives a dedicated Chrome profile through Playwright (pipe transport, no
// remote-debugging port). The first run needs a manual login in the opened
// window; later runs reuse the saved session.
//
// Usage:
//   node drjoy-staff-export.js                 export to %USERPROFILE%\Downloads
//   node drjoy-staff-export.js --inspect       dump the at0006 form controls (no export)
//   node drjoy-staff-export.js --login-only    open the window, wait for login, then quit
//   options: --out-dir <dir>  --login-timeout <sec>  --headless  --min-rows <n>
//
// Exit codes: 0 ok / 1 error / 2 login timeout / 3 CSV rejected by validation

'use strict';
const fs = require('fs');
const path = require('path');
const { DEFAULT_OUT_DIR, log, sleep, launch, waitForScreen, toShiftJisBytes } = require('./lib/drjoy-common');

const STAFF_URL = 'https://app.drjoy.jp/at/at0006';
const REQUIRED_HEADERS = ['職員コード', 'ログインID'];
const MAX_ATTEMPTS = 3;
const DOWNLOAD_TIMEOUT_MS = 30000;
// Temporary file the server writes to Firebase Storage for each export.
const GENERATED_CSV_RE = /import_staffs_[^/?]*\.csv\?/;

function parseArgs(argv) {
  const opts = {
    outDir: DEFAULT_OUT_DIR,
    loginTimeoutSec: 300,
    minRows: 500,
    headless: false,
    inspect: false,
    loginOnly: false,
  };
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if (a === '--out-dir') opts.outDir = argv[++i];
    else if (a === '--login-timeout') opts.loginTimeoutSec = Number(argv[++i]);
    else if (a === '--min-rows') opts.minRows = Number(argv[++i]);
    else if (a === '--headless') opts.headless = true;
    else if (a === '--inspect') opts.inspect = true;
    else if (a === '--login-only') opts.loginOnly = true;
    else throw new Error(`Unknown argument: ${a}`);
  }
  return opts;
}

// Dump every form control on the page with a nearby label, so the expected
// defaults (encoding, account status, show locked/deleted) can be asserted
// later. Text only; no screenshots.
async function inspectControls(page) {
  const rows = await page.evaluate(() => {
    const labelOf = (el) => {
      if (el.id) {
        const l = document.querySelector(`label[for="${CSS.escape(el.id)}"]`);
        if (l) return l.innerText.trim();
      }
      const wrap = el.closest('label');
      if (wrap) return wrap.innerText.trim();
      let p = el.parentElement;
      for (let i = 0; i < 3 && p; i++, p = p.parentElement) {
        const t = p.innerText && p.innerText.trim();
        if (t) return t.slice(0, 60).replace(/\s+/g, ' ');
      }
      return '';
    };
    return [...document.querySelectorAll('input, select, button')]
      .filter((el) => el.offsetParent !== null)
      .map((el) => ({
        tag: el.tagName.toLowerCase(),
        type: el.type || '',
        name: el.name || '',
        id: el.id || '',
        ngModel: el.getAttribute('ng-model') || el.getAttribute('formcontrolname') || '',
        checked: el.type === 'checkbox' || el.type === 'radio' ? el.checked : undefined,
        value:
          el.tagName === 'SELECT'
            ? (el.selectedOptions[0] ? el.selectedOptions[0].text.trim() : '')
            : el.tagName === 'BUTTON' ? el.innerText.trim() : (el.type === 'password' ? '***' : el.value),
        cls: el.className && typeof el.className === 'string' ? el.className.slice(0, 60) : '',
        label: labelOf(el),
      }));
  });
  console.log(JSON.stringify(rows, null, 2));
}

// Next free "staffs (N).csv" name; Tool 20 only picks files matching that pattern.
function nextStaffsPath(outDir) {
  let max = 0;
  for (const f of fs.readdirSync(outDir)) {
    const m = /^staffs \((\d+)\)\.csv$/i.exec(f);
    if (m) max = Math.max(max, Number(m[1]));
  }
  return path.join(outDir, `staffs (${max + 1}).csv`);
}

function validateCsv(file, minRows) {
  const buf = fs.readFileSync(file);
  const text = new TextDecoder('shift_jis').decode(buf);
  const lines = text.split(/\r?\n/).filter((l) => l.length > 0);
  const header = lines[0] || '';
  const problems = [];
  if (buf.length >= 3 && buf[0] === 0xef && buf[1] === 0xbb && buf[2] === 0xbf) {
    problems.push('UTF-8 BOM found; the screen encoding is probably UTF-8 (expected Shift_JIS)');
  }
  for (const h of REQUIRED_HEADERS) {
    if (!header.includes(h)) problems.push(`header "${h}" not found when decoded as Shift_JIS`);
  }
  if (lines.length < minRows) problems.push(`only ${lines.length} lines (expected >= ${minRows}); a filter may be active`);
  return { lines: lines.length, problems };
}

async function main() {
  const opts = parseArgs(process.argv.slice(2));
  if (!opts.inspect && !opts.loginOnly && !fs.existsSync(opts.outDir)) {
    throw new Error(`Output folder not found: ${opts.outDir}`);
  }

  const context = await launch(opts.headless);
  try {
    const page = context.pages()[0] || (await context.newPage());
    await page.goto(STAFF_URL, { waitUntil: 'domcontentloaded' });

    const exportBtn = page.locator('button', { hasText: 'エクスポート' }).first();
    const ready = await waitForScreen(page, {
      url: STAFF_URL, urlPart: '/at/at0006', readyLocator: exportBtn,
      timeoutSec: opts.loginTimeoutSec, headless: opts.headless,
    });
    if (!ready) {
      log(`${opts.loginTimeoutSec} 秒待ってもスタッフ画面が出ませんでした。`);
      process.exitCode = 2;
      return;
    }
    log('スタッフ画面 (at0006) を確認しました。');

    if (opts.loginOnly) {
      log('ログイン済みです。セッションはプロファイルに保存されました。');
      return;
    }
    if (opts.inspect) {
      await page.waitForLoadState('networkidle').catch(() => {});
      await inspectControls(page);
      return;
    }

    // Let the list settle so the export reflects the default filters.
    await page.waitForLoadState('networkidle').catch(() => {});

    // Take the CSV from the page's own fetch of the generated file instead of
    // the browser download. The fetch occasionally goes through a 302 to a
    // signed storage.googleapis.com URL that fails CORS; clicking again succeeds.
    const tmp = path.join(opts.outDir, `.staffs-downloading-${Date.now()}.csv`);
    let saved = false;
    for (let attempt = 1; attempt <= MAX_ATTEMPTS && !saved; attempt++) {
      log(`エクスポートを押します（${attempt}/${MAX_ATTEMPTS} 回目）。`);
      const pending = page
        .waitForResponse(
          (r) => r.request().method() === 'GET' && r.status() === 200 && GENERATED_CSV_RE.test(decodeURIComponent(r.url())),
          { timeout: DOWNLOAD_TIMEOUT_MS },
        )
        .catch(() => null);
      await exportBtn.click();
      const response = await pending;
      if (!response) {
        log(`${DOWNLOAD_TIMEOUT_MS / 1000} 秒待っても CSV が届きませんでした。`);
        await sleep(3000);
        continue;
      }
      try {
        fs.writeFileSync(tmp, toShiftJisBytes(await response.body(), REQUIRED_HEADERS[0]));
        saved = true;
      } catch (err) {
        log(`受け取りに失敗しました: ${err.message}`);
        await sleep(3000);
      }
    }
    if (!saved) throw new Error(`Export failed after ${MAX_ATTEMPTS} attempts.`);

    const { lines, problems } = validateCsv(tmp, opts.minRows);
    if (problems.length > 0) {
      const rejected = path.join(opts.outDir, `rejected-staffs-${Date.now()}.csv`);
      fs.renameSync(tmp, rejected);
      log(`NG: ${problems.join(' / ')}`);
      log(`Tool 20 に拾われないよう ${path.basename(rejected)} として保存しました。`);
      process.exitCode = 3;
      return;
    }

    const target = nextStaffsPath(opts.outDir);
    fs.renameSync(tmp, target);
    log(`OK: ${target} (${lines} 行, Shift_JIS)`);
  } finally {
    await context.close();
  }
}

main().catch((err) => {
  log(`ERROR: ${err.message}`);
  process.exitCode = 1;
});
