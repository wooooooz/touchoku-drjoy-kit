// drjoy-week-patterns.js  20261031
// Read-only: open at0055 (週パターン) and save the list of week patterns the page fetches
// (GET /api/dr/at/setting/get_list_basic_schedule_pattern?getAllFlag=true) as JSON.
// Used by Touchoku_app\tools\week_pattern.py to find "<氏名>_自動登録" (patterns made from 兼業申請 on 2026-07-17).
// Nothing in Dr.JOY is changed.
//
// Usage:  node drjoy-week-patterns.js --out <json> [--headless]
// Output: last line "OK: <json> n=<count>"
// Exit codes: 0 ok / 1 error (headless and not logged in: "Not logged in") / 2 login timeout

'use strict';
const fs = require('fs');
const path = require('path');
const { log, sleep, launch, waitForScreen } = require('./lib/drjoy-common');

const URL = 'https://app.drjoy.jp/at/at0055';
const API = '/api/dr/at/setting/get_list_basic_schedule_pattern';

function parseArgs(argv) {
  const opts = { out: null, headless: false, loginTimeoutSec: 300 };
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if (a === '--out') opts.out = path.resolve(argv[++i]);
    else if (a === '--headless') opts.headless = true;
    else throw new Error(`Unknown argument: ${a}`);
  }
  if (!opts.out) throw new Error('--out is required');
  return opts;
}

async function main() {
  const opts = parseArgs(process.argv.slice(2));
  const context = await launch(opts.headless);
  try {
    const page = context.pages()[0] || await context.newPage();
    let list = null;
    page.on('response', async (r) => {
      const u = r.url();
      if (!u.includes(API) || !u.includes('getAllFlag=true')) return;
      try { list = (await r.json()).basicSchedulePatterns || null; } catch (e) { /* not json */ }
    });
    await page.goto(URL, { waitUntil: 'domcontentloaded' });
    const ready = page.locator('a.add-link, [data-testid^="at0055-"]').first();
    const ok = await waitForScreen(page, { url: URL, urlPart: 'at0055', readyLocator: ready,
      timeoutSec: opts.loginTimeoutSec, headless: opts.headless });
    if (!ok) { log('Login timed out'); process.exitCode = 2; return; }
    for (let i = 0; i < 30 && !list; i++) await sleep(1000);
    if (!list) throw new Error('the week pattern list was not fetched');
    const rows = list.filter((p) => !p.deleted).map((p) => ({ id: p.id, name: p.patternName, code: p.patternCode,
      permission: p.permission, creatorId: p.creatorId, lastUpdated: p.lastUpdated }));
    fs.writeFileSync(opts.out, JSON.stringify(rows, null, 1), 'utf8');
    log(`OK: ${opts.out} n=${rows.length}`);
  } finally {
    await context.close();
  }
}

main().catch((e) => { console.error(e.message); process.exitCode = 1; });
