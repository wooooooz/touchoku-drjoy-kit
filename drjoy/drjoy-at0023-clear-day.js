// drjoy-at0023-clear-day.js  20261001
// Clear one day of the signed-in user's own 出勤簿 on at0023: press the punch table's 編集, which (after the
// confirmation) removes the day's requests including approved ones, then remove every punch row and save.
// Used to undo test data (2026-10-01: 管理者's 9/6 and 9/13 test punches and requests).
//
// Without --apply: press 編集, read the confirmation, and press its cancel/いいえ (or leave the page). Nothing is
// saved, so Dr.JOY is unchanged. With --apply: confirm, remove the punch rows, press 保存, and read the day again.
// No screenshots, no clipboard.
//
// Usage: node drjoy-at0023-clear-day.js --date YYYY-MM-DD --out <dump.json> [--apply] [--headless]
'use strict';
const fs = require('fs');
const path = require('path');
const { log, sleep, launch } = require('./lib/drjoy-common');

function parseArgs(argv) {
  const o = { date: null, out: null, apply: false, headless: false };
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if (a === '--date') o.date = argv[++i];
    else if (a === '--out') o.out = path.resolve(argv[++i]);
    else if (a === '--apply') o.apply = true;
    else if (a === '--headless') o.headless = true;
    else throw new Error(`Unknown argument: ${a}`);
  }
  if (!o.date || !o.out) throw new Error('--date --out are required');
  return o;
}

const visibleText = (page) => page.evaluate(() => [...document.querySelectorAll('.modal.show, [role=dialog], .swal2-popup, .modal-content')]
  .filter((e) => e.offsetParent !== null).map((e) => e.innerText.trim().slice(0, 300)));
const buttonsIn = (page) => page.evaluate(() => [...document.querySelectorAll('.modal.show button, [role=dialog] button, .swal2-popup button, .modal-content button')]
  .filter((e) => e.offsetParent !== null).map((e) => ({ text: e.innerText.trim(), testid: e.getAttribute('data-testid') || '' })));
const dayState = (page) => page.evaluate(() => ({
  punches: [...document.querySelectorAll('[data-testid^="at0023-log-input-time-"]')].map((e) => e.value),
  requests: [...document.querySelectorAll('[data-testid^="at0023-ot-row-"]')].map((r) => r.innerText.replace(/\s+/g, ' ').trim()),
  logButtons: [...document.querySelectorAll('[data-testid^="at0023-log-btn"]')].filter((e) => e.offsetParent !== null).map((e) => (e.getAttribute('data-testid') + ':' + e.innerText.trim())),
}));

// Click the first visible dialog button whose trimmed text is one of `names`; returns its text or ''.
// (The buttons' text carries surrounding whitespace, so Playwright's anchored hasText did not match.)
const clickDialogButton = (page, names) => page.evaluate((names) => {
  const b = [...document.querySelectorAll('.modal.show button, [role=dialog] button, .swal2-popup button, .modal-content button')]
    .find((e) => e.offsetParent !== null && names.includes(e.innerText.trim()));
  if (!b) return '';
  b.click();
  return b.innerText.trim();
}, names);

async function pressReal(page, testid) {
  const el = page.getByTestId(testid).first();
  await el.scrollIntoViewIfNeeded();
  await el.dispatchEvent('mousedown');
  await el.dispatchEvent('mouseup');
  await el.click();
}

async function gotoDay(page, want) {
  const title = page.getByTestId('at0023-current-day-title');
  for (let k = 0; ; k++) {
    const t = (await title.innerText()).trim();
    if (t.startsWith(want)) return t;
    if (k > 400) throw new Error(`could not reach ${want}`);
    await page.getByTestId('at0023-btn-yesterday').click();
    await page.waitForFunction((prev) => document.querySelector('[data-testid="at0023-current-day-title"]')?.innerText.trim() !== prev, t, { timeout: 15000 });
  }
}

(async () => {
  const o = parseArgs(process.argv.slice(2));
  const dump = { date: o.date, apply: o.apply, steps: [] };
  const save = () => fs.writeFileSync(o.out, JSON.stringify(dump, null, 1), 'utf8');
  const ctx = await launch(o.headless);
  try {
    const page = ctx.pages()[0] || (await ctx.newPage());
    page.on('dialog', async (d) => { dump.steps.push({ step: 'native-dialog', text: d.message() }); if (o.apply) await d.accept(); else await d.dismiss(); });
    await page.goto('https://app.drjoy.jp/at/at0023', { waitUntil: 'domcontentloaded' });
    await sleep(15000);
    if ((await page.locator('input[type="password"]').count()) > 0) throw new Error('Not logged in');
    const [, m, d] = o.date.split('-').map(Number);
    const want = `${m}月${d}日`;
    log(`day: ${await gotoDay(page, want)}`);
    await sleep(4000);
    dump.steps.push({ step: 'before', state: await dayState(page) });

    const submitText = async () => (await page.getByTestId('at0023-log-btn-submit').first().innerText()).trim();

    // Dry run reads only. Pressing 編集 is itself a write: on a day with no punches it un-confirms the day on the
    // server without any confirmation (2026-10-01, 9/20 and 9/22 had to be confirmed again).
    if (!o.apply) { save(); log('dry run: read only, nothing pressed'); return; }

    // 1. a confirmed day shows 編集: press it; the confirmation says it deletes the day's requests
    if ((await submitText()) === '編集') {
      await pressReal(page, 'at0023-log-btn-submit');
      await sleep(2500);
      dump.steps.push({ step: 'after-edit', modal: await visibleText(page), buttons: await buttonsIn(page) });
      if (!o.apply) {
        const cancelled = await clickDialogButton(page, ['いいえ', 'キャンセル', '閉じる']);
        dump.steps.push({ step: 'cancelled', pressed: cancelled, state: await dayState(page) });
        save(); log('dry run: nothing saved'); return;
      }
      const pressed = await clickDialogButton(page, ['はい', 'OK']);
      if (!pressed) throw new Error('confirmation has no はい/OK button');
      await sleep(3000);
      dump.steps.push({ step: 'after-yes', modal: await visibleText(page), state: await dayState(page) });
    }
    if (!o.apply) { dump.steps.push({ step: 'end-dry', state: await dayState(page) }); save(); log('dry run: nothing saved'); return; }

    // 2. now unconfirmed (確定): remove every punch row, last first (the remove icons are hover-only, so DOM click)
    for (let k = 0; k < 10; k++) {
      const n = await page.evaluate(() => {
        const bs = [...document.querySelectorAll('[data-testid^="at0023-log-btn-remove-"]')];
        if (!bs.length) return 0;
        bs[bs.length - 1].click();
        return bs.length;
      });
      if (!n) break;
      await sleep(1000);
      // "削除すると未同期のデータがある場合反映されません。削除しますか？" -> はい
      const yes = await clickDialogButton(page, ['はい', 'OK']);
      await sleep(1500);
      dump.steps.push({ step: `remove-${k}`, before: n, pressed: yes, state: await dayState(page), modal: await visibleText(page) });
      const after = await page.evaluate(() => document.querySelectorAll('[data-testid^="at0023-log-btn-remove-"]').length);
      if (after >= n) break;   // the screen would not drop the row (e.g. the last one): stop trying
    }

    // 3. press 確定 and accept its confirmation, if any
    const label = await submitText();
    dump.steps.push({ step: 'before-confirm', label, state: await dayState(page) });
    if (label === '確定' || label === '保存') {
      await pressReal(page, 'at0023-log-btn-submit');
      await sleep(3000);
      dump.steps.push({ step: 'after-confirm', modal: await visibleText(page), buttons: await buttonsIn(page) });
      const okText = await clickDialogButton(page, ['はい', 'OK', '確定', '保存']);
      if (okText) { await sleep(3000); dump.steps.push({ step: 'after-confirm-ok', pressed: okText, modal: await visibleText(page) }); }
    }

    await page.reload({ waitUntil: 'domcontentloaded' });
    await sleep(12000);
    await gotoDay(page, want);
    await sleep(4000);
    dump.steps.push({ step: 'after-reload', state: await dayState(page) });
    save();
    const st = dump.steps[dump.steps.length - 1].state;
    log(`after: punches=${JSON.stringify(st.punches)} requests=${st.requests.length}`);
  } finally {
    save();   // always write what happened, also on errors
    await ctx.close();
  }
})().catch((e) => { console.error(e.message); process.exit(1); });
