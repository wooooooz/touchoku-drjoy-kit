// drjoy-at0029-export.js  20261017（勤務実績 work-record を追加） / 20261028（下限 560・所属履歴に仮登録も）
// Create and download an export on Dr.JOY's data-download screen (at0029)
// without screenshots, using the dedicated Chrome profile.
//
// The page's own "作成" (create) button is clicked; the create request body is
// rewritten on the way out (month / period / account status), which is more
// reliable than driving the select2 and read-only date pickers. The job is
// then polled with the page's reload button until READY, its download button
// is clicked, and the file is taken from the page's fetch of Firebase Storage
// (browser downloads stay off).
//
// Usage:
//   node drjoy-at0029-export.js --type working-plan --month 2026-10
//   node drjoy-at0029-export.js --type beacon --from 2026-10-01 --to 2026-10-07
//   node drjoy-at0029-export.js --type payroll --month 2026-09     (給与連携項目一覧, 医師, 1日〜月末)
//   node drjoy-at0029-export.js --type department-history --year 2026   (スタッフ所属履歴, Excel, 本登録＋仮登録)
//   node drjoy-at0029-export.js --type work-record --month 2026-09     (勤務実績, 医師, 本登録, Shift_JIS, HH:MM)
//   options: --out-dir <dir>  --headless  --login-timeout <sec>  --ready-timeout <sec>
//            --min-sheets <n> (working-plan, default 560)
//
// Output: the file is saved under its Dr.JOY name in --out-dir (default
// Downloads), with " (N)" added when the name is taken. The last line is
//   OK: <full path> (...)
// Exit codes: 0 ok / 1 error / 2 login timeout / 3 file rejected by validation

'use strict';
const fs = require('fs');
const path = require('path');
const { DEFAULT_OUT_DIR, log, sleep, launch, waitForScreen, toShiftJisBytes, uniquePath, readZipEntry } = require('./lib/drjoy-common');

const AT0029_URL = 'https://app.drjoy.jp/at/at0029';
const STORAGE_RE = /firebasestorage\.googleapis\.com|storage\.googleapis\.com/;
const MAX_DOWNLOAD_ATTEMPTS = 3;
const POLL_INTERVAL_MS = 10000;
// The template to use for 給与連携項目一覧. Several retired copies with ×
// marks exist next to it; only the exact name is accepted.
const PAYROLL_TEMPLATE = '給与連携項目一覧';

const TYPES = {
  // 勤務予定: xlsx, one sheet per person. Account status must include 仮登録
  // (PROVISIONAL); without it about 100 new staff are missing and importing
  // plans built from it has deleted real schedules before.
  'working-plan': {
    uiValue: 'working_plan',
    createApi: '/api/dr/at/staff/export_working_plan',
    listApi: '/api/dr/at/staff/get_working_plan',
    listKey: 'listDataCSV',
    idKey: 'workCSVId',
    downloadApi: '/api/dr/at/staff/get_download_csv',
    overrides: (o) => ({ month: o.month, unitFormat: '2', accountStatus: ['PROVISIONAL', 'VALID'] }),
    matches: (item, o) => item.month === o.month && String(item.unitFormat) === '2',
    describe: (o) => `勤務予定 ${o.month}（1ファイル複数シート・本登録＋仮登録）`,
  },
  // ビーコン滞在ログ: Shift_JIS CSV, stored double-encoded like at0006.
  beacon: {
    uiValue: 'Beacon_Stay_Log',
    createApi: '/api/dr/at/beacon/create_beacon_stay_log',
    listApi: '/api/dr/at/beacon/list_beacon_stay_log',
    listKey: 'listCsv',
    idKey: 'workCSVId',
    downloadApi: '/api/dr/at/staff/get_download_csv',
    overrides: (o) => ({ fromDate: o.from, toDate: o.to, encoding: 1 }),
    matches: (item, o) => item.period === `${slashDate(o.from)}~${slashDate(o.to)}`,
    describe: (o) => `ビーコン滞在ログ ${o.from}〜${o.to}（Shift_JIS）`,
  },
  // 給与連携項目一覧: the "Template" type with the 給与連携項目一覧 template,
  // doctors only, one combined xlsx for the 1st to the last day of the month
  // (本登録 only, as the monthly allowance work has always used).
  payroll: {
    uiValue: 'Template',
    createApi: '/api/dr/at/staff/post_template_file',
    listApi: '/api/dr/at/staff/get_customize_download_logs',
    listKey: 'customizeDownloadLogs',
    idKey: 'customizeDownloadLogId',
    downloadApi: '/api/dr/at/staff/get_customize_download',
    prepare: async (page) => {
      const tpl = page.getByTestId('at0029-select-template');
      const labels = (await tpl.locator('option').allInnerTexts()).map((s) => s.trim());
      if (!labels.includes(PAYROLL_TEMPLATE)) throw new Error(`Template "${PAYROLL_TEMPLATE}" not found (options: ${labels.join(' | ')})`);
      await tpl.selectOption({ label: PAYROLL_TEMPLATE });
      await page.getByTestId('at0029-select-job-type').selectOption({ label: '医師' });
      await sleep(1500);
    },
    overrides: (o) => ({
      fromDate: o.from, toDate: o.to, month: o.month, jobTypeId: 'J0001', jobTypeName: '医師',
      encoding: 1, timeFormat: 0, exportUnit: 'COMBINE_ONE_FILE', accountStatus: ['VALID'],
    }),
    // The create body must carry the exact template; checked before it is sent.
    checkBody: (b) => (b.templateId && b.templateId !== 'ALL' ? null : `templateId is "${b.templateId}"`),
    matches: (item, o) => item.fromDate === o.from && item.toDate === o.to && item.jobType === '医師',
    describe: (o) => `給与連携項目一覧 ${o.from}〜${o.to}（医師・1ファイルに結合・Shift_JIS・HH:MM）`,
  },
  // 勤務実績: one Shift_JIS CSV, one row per doctor, monthly totals computed by
  // Dr.JOY (残業時間（25％）/ 法定休日時間（35％）etc.). Doctors, 本登録 only.
  // The 働き方レポート deck reads its overtime from this file (20261017).
  // API and body captured from the page on 2026-10-17.
  'work-record': {
    uiValue: 'Work_Record',
    createApi: '/api/dr/at/staff/post_create_work_csv',
    listApi: '/api/dr/at/staff/get_work_csv',
    listKey: 'listDataCSV',
    idKey: 'workCSVId',
    downloadApi: '/api/dr/at/staff/get_download_csv',
    overrides: (o) => ({
      month: o.month, jobTypeId: 'J0001', encoding: 1, timeFormat: 0, accountStatus: ['VALID'],
      listDeptIdSelected: ['ALL'], departmentId: ['ALL'], contractType: 'WF_ALL', workingPattern: 'ALL', closingDayStatus: 0,
    }),
    matches: (item, o) => item.month === o.month && item.jobType === '医師',
    describe: (o) => `勤務実績 ${o.month}（医師・本登録・Shift_JIS・HH:MM）`,
  },
  // スタッフ所属履歴: one xlsx per calendar year; everyone (all departments,
  // job types and contract types), 本登録 and 仮登録 (2026-10-28: 仮登録 初期研 such as
  // 夏木 二郎 were missing, so their rotations could not be seen). Holds the planned department
  // changes (e.g. 初期研 rotations imported once a year), which staffs.csv
  // (today's department only) does not.
  'department-history': {
    uiValue: 'department_histories',
    createApi: '/api/dr/at/staff/create_data_department_histories',
    listApi: '/api/dr/at/staff/get_data_department_histories_csv',
    listKey: 'listDataCSV',
    idKey: 'workCSVId',
    downloadApi: '/api/dr/at/staff/get_download_data_department_histories_csv',
    overrides: (o) => ({
      year: o.year, fileFormat: 2, jobTypeId: 'ALL', departmentId: ['ALL'], contractType: 'WF_ALL', accountStatus: ['PROVISIONAL', 'VALID'],
    }),
    matches: (item, o) => String(item.year) === o.year,
    describe: (o) => `スタッフ所属履歴 ${o.year}年（Excel・すべて・本登録＋仮登録）`,
  },
};

const slashDate = (d) => `${d.slice(0, 4)}/${d.slice(4, 6)}/${d.slice(6, 8)}`;

function normMonth(s) {
  const m = /^(\d{4})-?(\d{2})$/.exec(s || '');
  if (!m) throw new Error(`--month must be YYYY-MM or YYYYMM: ${s}`);
  return m[1] + m[2];
}

function normDate(s, name) {
  const m = /^(\d{4})-?(\d{2})-?(\d{2})$/.exec(s || '');
  if (!m) throw new Error(`${name} must be YYYY-MM-DD or YYYYMMDD: ${s}`);
  return m[1] + m[2] + m[3];
}

function parseArgs(argv) {
  const opts = {
    type: null, month: null, from: null, to: null, year: null,
    outDir: DEFAULT_OUT_DIR, headless: false,
    // 620 -> 560 (2026-10-28): 39 accounts were removed from Dr.JOY on 10/21-23 and
    // a complete export (with 仮登録) fell to 608; 本登録 only is now about 510-520.
    loginTimeoutSec: 300, readyTimeoutSec: 600, minSheets: 560,
  };
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if (a === '--type') opts.type = argv[++i];
    else if (a === '--month') opts.month = normMonth(argv[++i]);
    else if (a === '--from') opts.from = normDate(argv[++i], '--from');
    else if (a === '--to') opts.to = normDate(argv[++i], '--to');
    else if (a === '--year') {
      opts.year = argv[++i];
      if (!/^\d{4}$/.test(opts.year || '')) throw new Error(`--year must be YYYY: ${opts.year}`);
    }
    else if (a === '--out-dir') opts.outDir = argv[++i];
    else if (a === '--headless') opts.headless = true;
    else if (a === '--login-timeout') opts.loginTimeoutSec = Number(argv[++i]);
    else if (a === '--ready-timeout') opts.readyTimeoutSec = Number(argv[++i]);
    else if (a === '--min-sheets') opts.minSheets = Number(argv[++i]);
    else throw new Error(`Unknown argument: ${a}`);
  }
  if (!TYPES[opts.type]) throw new Error(`--type must be one of: ${Object.keys(TYPES).join(', ')}`);
  if (['working-plan', 'payroll', 'work-record'].includes(opts.type) && !opts.month) throw new Error(`--month is required for ${opts.type}`);
  if (opts.type === 'payroll') {
    const y = Number(opts.month.slice(0, 4));
    const m = Number(opts.month.slice(4, 6));
    opts.from = `${opts.month}01`;
    opts.to = `${opts.month}${String(new Date(y, m, 0).getDate()).padStart(2, '0')}`;
  }
  if (opts.type === 'department-history' && !opts.year) opts.year = String(new Date().getFullYear());
  if (opts.type === 'beacon') {
    if (!opts.from || !opts.to) throw new Error('--from and --to are required for beacon');
    if (opts.from > opts.to) throw new Error('--from is after --to');
  }
  if (!fs.existsSync(opts.outDir)) throw new Error(`Output folder not found: ${opts.outDir}`);
  return opts;
}

function countSheets(buf, problems) {
  try {
    const wb = readZipEntry(buf, 'xl/workbook.xml');
    if (!wb) { problems.push('xl/workbook.xml not found (not an xlsx?)'); return 0; }
    return (wb.toString('utf8').match(/<sheet\s/g) || []).length;
  } catch (err) {
    problems.push(`cannot read the xlsx: ${err.message}`);
    return 0;
  }
}

function validateWorkingPlan(buf, fileName, opts) {
  const problems = [];
  const monthTag = `${opts.month.slice(0, 4)}年${opts.month.slice(4, 6)}月度`;
  if (!fileName.includes(monthTag)) problems.push(`file name "${fileName}" does not contain ${monthTag}`);
  const sheets = countSheets(buf, problems);
  if (sheets > 0 && sheets < opts.minSheets) {
    problems.push(`only ${sheets} sheets (expected >= ${opts.minSheets}); 仮登録 may be missing`);
  }
  return { summary: `${sheets} シート, ${buf.length} bytes`, problems };
}

function validatePayroll(buf, fileName, opts) {
  const problems = [];
  const prefix = `${opts.from}_${opts.to}_給与連携項目一覧`;
  if (!fileName.startsWith(prefix)) problems.push(`file name "${fileName}" does not start with ${prefix}`);
  const sheets = countSheets(buf, problems);
  if (sheets < 1 && problems.length === 0) problems.push('no sheet in the xlsx');
  return { summary: `${sheets} シート, ${buf.length} bytes`, problems };
}

function validateDepartmentHistory(buf, fileName, opts) {
  const problems = [];
  const tag = `スタッフ所属履歴_${opts.year}年`;
  if (!fileName.startsWith(tag)) problems.push(`file name "${fileName}" does not start with ${tag}`);
  const sheets = countSheets(buf, problems);
  if (sheets < 1 && problems.length === 0) problems.push('no sheet in the xlsx');
  return { summary: `${sheets} シート, ${buf.length} bytes`, problems };
}

function validateWorkRecord(buf, fileName, opts) {
  const problems = [];
  const tag = `${opts.month}月度`;
  if (!fileName.includes(tag)) problems.push(`file name "${fileName}" does not contain ${tag}`);
  if (!fileName.includes('_医師_')) problems.push(`file name "${fileName}" does not contain _医師_`);
  const lines = new TextDecoder('shift_jis').decode(buf).split(/\r?\n/).filter((l) => l.length > 0);
  const header = lines[1] || '';
  // The deck stops when these are missing (dtx18_stats.load_work_record_ot).
  for (const h of ['氏名', '職員コード', '残業時間（25％）', '法定休日時間（35％）']) {
    if (!header.includes(h)) problems.push(`header "${h}" not found in row 2 (Shift_JIS)`);
  }
  const rows = Math.max(0, lines.length - 2);
  if (rows < 300) problems.push(`only ${rows} doctors (expected >= 300)`);
  return { summary: `${rows} 名, ${header.split(',').length} 列, Shift_JIS`, problems };
}

function validateBeacon(buf) {
  const problems = [];
  const lines = new TextDecoder('shift_jis').decode(buf).split(/\r?\n/).filter((l) => l.length > 0);
  const header = lines[0] || '';
  for (const h of ['開始時刻', '終了時刻', '職員コード', '滞在エリア']) {
    if (!header.includes(h)) problems.push(`header "${h}" not found when decoded as Shift_JIS`);
  }
  return { summary: `${lines.length} 行, Shift_JIS`, problems };
}

async function main() {
  const opts = parseArgs(process.argv.slice(2));
  const type = TYPES[opts.type];
  log(`対象: ${type.describe(opts)}`);

  const context = await launch(opts.headless);
  try {
    const page = context.pages()[0] || (await context.newPage());
    await page.goto(AT0029_URL, { waitUntil: 'domcontentloaded' });

    const typeSelect = page.getByTestId('at0029-select-type-csv');
    const ready = await waitForScreen(page, {
      url: AT0029_URL, urlPart: '/at/at0029', readyLocator: typeSelect,
      timeoutSec: opts.loginTimeoutSec, headless: opts.headless,
    });
    if (!ready) {
      log(`${opts.loginTimeoutSec} 秒待ってもデータダウンロード画面が出ませんでした。`);
      process.exitCode = 2;
      return;
    }
    log('データダウンロード画面 (at0029) を確認しました。');
    await page.waitForLoadState('networkidle').catch(() => {});

    const reloadBtn = page.getByTestId('at0029-btn-reload');
    const isList = (r) => r.request().method() === 'GET' && r.status() === 200 && r.url().includes(type.listApi);
    const readList = async (r) => ((await r.json())[type.listKey]) || [];
    const fetchList = async () => {
      const pending = page.waitForResponse(isList, { timeout: 30000 });
      await reloadBtn.click();
      return readList(await pending);
    };
    const idOf = (it) => it[type.idKey];

    // Switching the type loads that type's job list.
    const firstList = page.waitForResponse(isList, { timeout: 30000 }).catch(() => null);
    await typeSelect.selectOption(type.uiValue);
    const firstResp = await firstList;
    if (type.prepare) await type.prepare(page);
    const before = type.prepare ? await fetchList() : (firstResp ? await readList(firstResp) : await fetchList());
    const knownIds = new Set(before.map(idOf));

    // Rewrite the create request; abort it if the body is not what we expect.
    await page.route(`**${type.createApi}*`, async (route) => {
      const req = route.request();
      let body;
      try {
        body = JSON.parse(req.postData() || '');
      } catch {
        log(`作成の要求が JSON ではありませんでした。送らずに止めます: ${String(req.postData()).slice(0, 200)}`);
        await route.abort();
        return;
      }
      const bad = type.checkBody ? type.checkBody(body) : null;
      if (bad) {
        log(`作成の要求が想定と違うので送らずに止めます: ${bad}`);
        await route.abort();
        return;
      }
      const sent = { ...body, ...type.overrides(opts) };
      log(`作成の要求: ${JSON.stringify(sent)}`);
      await route.continue({ postData: JSON.stringify(sent) });
    });

    const created = page.waitForResponse(
      (r) => r.request().method() === 'POST' && r.url().includes(type.createApi), { timeout: 30000 },
    );
    await page.getByTestId('at0029-btn-create-csv').click();
    const createResp = await created.catch(() => null);
    if (!createResp) throw new Error('The create request was not sent (or was stopped); see the log above.');
    if (createResp.status() >= 300) throw new Error(`Create request failed: HTTP ${createResp.status()}`);
    log('作成を依頼しました。できあがるのを待ちます。');

    // Find our job (new id with matching parameters) and wait for READY.
    const deadline = Date.now() + opts.readyTimeoutSec * 1000;
    let job = null;
    let list = [];
    let lastStatus = '';
    while (Date.now() < deadline) {
      list = await fetchList();
      job = list.find((it) => !knownIds.has(idOf(it)) && type.matches(it, opts)) || null;
      const status = job ? job.status : '(not listed yet)';
      if (status !== lastStatus) { log(`状態: ${status}`); lastStatus = status; }
      if (job && job.status === 'READY') break;
      if (job && /ERR|FAIL/i.test(job.status || '')) throw new Error(`Dr.JOY reported status ${job.status}`);
      await sleep(POLL_INTERVAL_MS);
    }
    if (!job || job.status !== 'READY') throw new Error(`The export was not ready within ${opts.readyTimeoutSec} seconds.`);
    const jobId = idOf(job);

    // Click that row's download button and take the file from the page's
    // fetch of Firebase Storage.
    let fileName = null;
    let body = null;
    for (let attempt = 1; attempt <= MAX_DOWNLOAD_ATTEMPTS && !body; attempt++) {
      const idx = list.findIndex((it) => idOf(it) === jobId);
      if (idx < 0) throw new Error('The created job disappeared from the list.');
      log(`ダウンロードを押します（${attempt}/${MAX_DOWNLOAD_ATTEMPTS} 回目、一覧の ${idx + 1} 行目）。`);
      const dlInfo = page.waitForResponse((r) => r.url().includes(type.downloadApi), { timeout: 30000 }).catch(() => null);
      const media = page.waitForResponse(
        (r) => STORAGE_RE.test(r.url()) && r.status() === 200 && !/json/.test(r.headers()['content-type'] || ''),
        { timeout: 60000 },
      ).catch(() => null);
      await page.locator('table tbody tr').nth(idx).locator('button', { hasText: 'ダウンロード' }).click();

      const info = await dlInfo;
      if (!info) { log('ダウンロード先の問い合わせが来ませんでした。'); list = await fetchList(); continue; }
      if (!info.url().includes(jobId)) throw new Error(`Clicked a different job (${info.url()}); nothing was saved.`);
      const { downloadUrl } = await info.json();
      if (!downloadUrl) throw new Error('Dr.JOY returned no downloadUrl.');
      fileName = path.basename(downloadUrl);

      const resp = await media;
      if (!resp) { log('ファイルが届きませんでした。'); await sleep(3000); list = await fetchList(); continue; }
      if (!decodeURIComponent(resp.url()).includes(downloadUrl)) throw new Error('Received a different file from Storage; nothing was saved.');
      body = await resp.body();
    }
    if (!body) throw new Error(`Download failed after ${MAX_DOWNLOAD_ATTEMPTS} attempts.`);

    let bytes = body;
    let result;
    if (opts.type === 'department-history' && !opts.year) opts.year = String(new Date().getFullYear());
  if (opts.type === 'beacon') {
      bytes = toShiftJisBytes(body, '職員コード');
      result = validateBeacon(bytes);
    } else if (opts.type === 'work-record') {
      bytes = toShiftJisBytes(body, '基本項目');
      result = validateWorkRecord(bytes, fileName, opts);
    } else if (opts.type === 'department-history') {
      result = validateDepartmentHistory(bytes, fileName, opts);
    } else if (opts.type === 'payroll') {
      result = validatePayroll(bytes, fileName, opts);
    } else {
      result = validateWorkingPlan(bytes, fileName, opts);
    }

    if (result.problems.length > 0) {
      const rejected = path.join(opts.outDir, `rejected-at0029-${Date.now()}${path.extname(fileName)}`);
      fs.writeFileSync(rejected, bytes);
      log(`NG: ${result.problems.join(' / ')}`);
      log(`下流に拾われないよう ${path.basename(rejected)} として保存しました。`);
      process.exitCode = 3;
      return;
    }

    const tmp = path.join(opts.outDir, `.at0029-downloading-${Date.now()}`);
    fs.writeFileSync(tmp, bytes);
    const target = uniquePath(opts.outDir, fileName);
    fs.renameSync(tmp, target);
    log(`OK: ${target} (${result.summary})`);
  } finally {
    await context.close();
  }
}

main().catch((err) => {
  log(`ERROR: ${err.message}`);
  process.exitCode = 1;
});
