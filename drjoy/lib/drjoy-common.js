// lib/drjoy-common.js  20261028（お知らせのポップアップを隠す） / 20261104（使用中のプロファイルは空くまで待つ）
// Shared helpers for the Dr.JOY Playwright scripts: launching the dedicated
// Chrome profile, waiting for a screen (and for a manual login), undoing the
// double encoding of generated CSV files, and small file utilities.
'use strict';
const fs = require('fs');
const os = require('os');
const path = require('path');
const zlib = require('zlib');
const { chromium } = require('playwright-core');

const PROFILE_DIR = path.join(process.env.LOCALAPPDATA || os.tmpdir(), 'DrJoyAutomation', 'chrome-profile');
const DEFAULT_OUT_DIR = path.join(os.homedir(), 'Downloads');

function log(msg) {
  const t = new Date().toTimeString().slice(0, 8);
  console.log(`[${t}] ${msg}`);
}

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

// Every script shares PROFILE_DIR, and Chrome refuses a profile that another
// Chrome still has open (it exits at once with code 21). On 2026-11-04 the 15:30
// KotaiSync started while a request-list run from another session held the
// profile, and failed. Wait for the other run to finish instead (20261104).
const PROFILE_WAIT_MIN = 15;
const PROFILE_RETRY_SEC = 20;

// Downloads stay off: letting Playwright accept the page's blob download made
// Chrome exit ~0.2 s later on most runs (2026-09-23). Files are taken from the
// page's own fetch of the generated file instead.
async function launch(headless) {
  fs.mkdirSync(PROFILE_DIR, { recursive: true });
  log(`Profile: ${PROFILE_DIR}`);
  const started = Date.now();
  let context;
  let tries = 1;
  for (; ; tries++) {
    try {
      context = await chromium.launchPersistentContext(PROFILE_DIR, {
        channel: 'chrome',
        headless,
        acceptDownloads: false,
        viewport: null,
        args: ['--no-first-run', '--no-default-browser-check'],
      });
      break;
    } catch (e) {
      const msg = String((e && e.message) || e);
      if (!/exitCode=21\b/.test(msg)) throw e;
      if (Date.now() - started >= PROFILE_WAIT_MIN * 60 * 1000) {
        throw new Error(`プロファイルが ${PROFILE_WAIT_MIN} 分たっても空きませんでした（別の Chrome が使用中・終了コード 21）。\n${msg}`);
      }
      if (tries === 1) log(`プロファイルを別の Chrome が使っています。${PROFILE_RETRY_SEC} 秒おきに、最長 ${PROFILE_WAIT_MIN} 分待ちます。`);
      await sleep(PROFILE_RETRY_SEC * 1000);
    }
  }
  if (tries > 1) log(`プロファイルが空いたので起動しました（${Math.round((Date.now() - started) / 1000)} 秒待った）。`);
  // Dr.JOY's announcement modal (app-modal-drjoy-info) covered every screen on
  // 2026-10-27 and took the clicks meant for the export / create buttons, so the
  // staff, department-history and beacon exports all failed. Nobody reads it in
  // this profile: keep it hidden on every page of the context (20261028).
  await context.addInitScript(() => {
    const add = () => {
      const s = document.createElement('style');
      s.textContent = 'app-modal-drjoy-info{display:none!important;pointer-events:none!important}';
      (document.head || document.documentElement).appendChild(s);
    };
    if (document.documentElement) add();
    else document.addEventListener('DOMContentLoaded', add);
  });
  return context;
}

// Wait until `readyLocator` is visible on `url`. If a login form is showing,
// wait for the user (headed) or fail (headless); once logged in, go back to
// `url` because Dr.JOY usually redirects to the home screen after login.
// Returns true when ready, false on timeout.
async function waitForScreen(page, { url, urlPart, readyLocator, timeoutSec, headless }) {
  const deadline = Date.now() + timeoutSec * 1000;
  let promptedLogin = false;
  while (Date.now() < deadline) {
    if (await readyLocator.isVisible().catch(() => false)) return true;

    const onLoginForm = (await page.locator('input[type="password"]').count().catch(() => 0)) > 0;
    if (onLoginForm) {
      if (headless) throw new Error('Not logged in. Run once without --headless (or with --login-only) to log in.');
      if (!promptedLogin) {
        log('Dr.JOY のログイン画面です。開いたウィンドウで user003 としてログインしてください。');
        promptedLogin = true;
      }
    } else if (!page.url().includes(urlPart)) {
      await page.goto(url, { waitUntil: 'domcontentloaded' }).catch(() => {});
    }
    await sleep(3000);
  }
  return false;
}

// Generated CSV files are stored "double-encoded": each Shift_JIS byte became
// one character (Windows-1252, with undefined bytes kept as U+0080-U+009F),
// and that text was saved as UTF-8. The page unwraps it before offering the
// download; do the same here. Verified byte-identical to a real download.
const CP1252_REVERSE = new Map([
  [0x20ac, 0x80], [0x201a, 0x82], [0x0192, 0x83], [0x201e, 0x84], [0x2026, 0x85],
  [0x2020, 0x86], [0x2021, 0x87], [0x02c6, 0x88], [0x2030, 0x89], [0x0160, 0x8a],
  [0x2039, 0x8b], [0x0152, 0x8c], [0x017d, 0x8e], [0x2018, 0x91], [0x2019, 0x92],
  [0x201c, 0x93], [0x201d, 0x94], [0x2022, 0x95], [0x2013, 0x96], [0x2014, 0x97],
  [0x02dc, 0x98], [0x2122, 0x99], [0x0161, 0x9a], [0x203a, 0x9b], [0x0153, 0x9c],
  [0x017e, 0x9e], [0x0178, 0x9f],
]);

// `probe` is a header word expected near the top of the file; if it already
// decodes as plain Shift_JIS, the body is returned unchanged.
function toShiftJisBytes(body, probe) {
  if (new TextDecoder('shift_jis').decode(body.subarray(0, 400)).includes(probe)) return body;

  const text = new TextDecoder('utf-8', { fatal: true }).decode(body);
  const out = Buffer.alloc(text.length);
  let n = 0;
  for (const ch of text) {
    const c = ch.codePointAt(0);
    const b = c < 0x100 ? c : CP1252_REVERSE.get(c);
    if (b === undefined) throw new Error(`Unexpected character U+${c.toString(16)} in the generated CSV`);
    out[n++] = b;
  }
  return out.subarray(0, n);
}

// "name.ext", then "name (1).ext", "name (2).ext", ... like Chrome, but
// without Chrome's limit of 100.
function uniquePath(outDir, fileName) {
  const ext = path.extname(fileName);
  const stem = fileName.slice(0, fileName.length - ext.length);
  let candidate = path.join(outDir, fileName);
  for (let i = 1; fs.existsSync(candidate); i++) candidate = path.join(outDir, `${stem} (${i})${ext}`);
  return candidate;
}

// Read one entry from a zip (xlsx) buffer. Enough for small XML parts.
function readZipEntry(buf, entryName) {
  let eocd = -1;
  for (let i = buf.length - 22; i >= Math.max(0, buf.length - 65557); i--) {
    if (buf.readUInt32LE(i) === 0x06054b50) { eocd = i; break; }
  }
  if (eocd < 0) throw new Error('Not a zip file (end of central directory not found)');
  const count = buf.readUInt16LE(eocd + 10);
  let p = buf.readUInt32LE(eocd + 16);
  for (let k = 0; k < count; k++) {
    if (buf.readUInt32LE(p) !== 0x02014b50) throw new Error('Broken zip central directory');
    const method = buf.readUInt16LE(p + 10);
    const compSize = buf.readUInt32LE(p + 20);
    const nameLen = buf.readUInt16LE(p + 28);
    const extraLen = buf.readUInt16LE(p + 30);
    const commentLen = buf.readUInt16LE(p + 32);
    const localOff = buf.readUInt32LE(p + 42);
    const name = buf.toString('utf8', p + 46, p + 46 + nameLen);
    if (name === entryName) {
      const dataStart = localOff + 30 + buf.readUInt16LE(localOff + 26) + buf.readUInt16LE(localOff + 28);
      const data = buf.subarray(dataStart, dataStart + compSize);
      if (method === 0) return data;
      if (method === 8) return zlib.inflateRawSync(data);
      throw new Error(`Unsupported zip compression method ${method}`);
    }
    p += 46 + nameLen + extraLen + commentLen;
  }
  return null;
}

module.exports = { PROFILE_DIR, DEFAULT_OUT_DIR, log, sleep, launch, waitForScreen, toShiftJisBytes, uniquePath, readZipEntry };
