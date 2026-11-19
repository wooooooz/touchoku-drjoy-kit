// drjoy-at0023-probe.js  20261001
// Read-only probe of the 残業｜自己研鑽 screen (at0023) in the automation profile: open it, wait, and dump the
// page text plus every button/link (text, data-testid, disabled) to a file, so a later script can be written
// against real selectors. Nothing is clicked. No screenshots.
//
// Usage: node drjoy-at0023-probe.js --out <file> [--url <at0023 url>] [--headless]
'use strict';
const fs = require('fs');
const path = require('path');
const { log, sleep, launch, waitForScreen } = require('./lib/drjoy-common');

function parseArgs(argv) {
  const o = { out: null, url: 'https://app.drjoy.jp/at/at0023', headless: false };
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if (a === '--out') o.out = path.resolve(argv[++i]);
    else if (a === '--url') o.url = argv[++i];
    else if (a === '--headless') o.headless = true;
    else if (a === '--date') o.date = argv[++i];   // YYYY-MM-DD: step back with 前日 until the title shows it
    else throw new Error(`Unknown argument: ${a}`);
  }
  if (!o.out) throw new Error('--out is required');
  return o;
}

(async () => {
  const o = parseArgs(process.argv.slice(2));
  const ctx = await launch(o.headless);
  try {
    const page = ctx.pages()[0] || (await ctx.newPage());
    await page.goto(o.url, { waitUntil: 'domcontentloaded' });
    // Dump whatever is showing after a fixed wait (the screen may redirect; the dump shows where it went).
    await sleep(15000);
    if ((await page.locator('input[type="password"]').count()) > 0) throw new Error('Not logged in');
    if (o.date) {
      const [, m, d] = o.date.split('-').map(Number);
      const want = `${m}月${d}日`;
      const title = page.getByTestId('at0023-current-day-title');
      for (let k = 0; k < 400; k++) {
        const t = (await title.innerText()).trim();
        if (t.startsWith(want)) { log(`day: ${t}`); break; }
        if (k === 399) throw new Error(`could not reach ${want} (now ${t})`);
        await page.getByTestId('at0023-btn-yesterday').click();
        await page.waitForFunction((prev) => {
          const e = document.querySelector('[data-testid="at0023-current-day-title"]');
          return e && e.innerText.trim() !== prev;
        }, t, { timeout: 15000 });
      }
      await sleep(4000);
    }
    const dump = await page.evaluate(() => {
      const els = [...document.querySelectorAll('button, a, [role=button], input[type=button], [data-testid]')];
      const items = els.map((e) => ({
        tag: e.tagName, text: (e.innerText || e.value || '').trim().slice(0, 60),
        testid: e.getAttribute('data-testid') || '', disabled: !!e.disabled || e.getAttribute('aria-disabled') === 'true',
      })).filter((x) => x.text || x.testid);
      return { url: location.href, title: document.title, text: document.querySelector('main')?.innerText || document.body.innerText, items };
    });
    fs.writeFileSync(o.out, JSON.stringify(dump, null, 1), 'utf8');
    log(`OK: ${o.out} (${dump.items.length} elements)`);
  } finally {
    await ctx.close();
  }
})().catch((e) => { console.error(e.message); process.exit(1); });
