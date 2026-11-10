/**
 * Flow 平台後端（Google Apps Script）
 *
 * 部署：
 *   1. 建一個 Google Sheet → 擴充功能 → Apps Script，貼上本檔
 *   2. 執行一次 setup()，會建立資料表並產生 FLOW_TOKEN（記錄在執行紀錄）
 *   3. 部署 → 新增部署作業 → 網頁應用程式
 *      執行身分：我；存取權：任何人（靠 token 保護）
 *   4. 把網址和 token 交給 flow init
 *
 * Script Properties：
 *   FLOW_TOKEN      API token（setup 自動產生）
 *   FLOW_FOLDER_ID  上傳檔案的 Drive 資料夾（setup 自動建立）
 */

const SHEETS = {
  specs:     ['project', 'version', 'title', 'status', 'specJson', 'specMd', 'createdAt', 'createdBy'],
  events:    ['ts', 'project', 'version', 'node', 'status', 'summary', 'artifacts', 'commit', 'actor'],
  questions: ['id', 'ts', 'project', 'version', 'node', 'question', 'status', 'answer', 'answeredBy', 'answeredAt', 'actor'],
};
const STATUSES = ['todo', 'working', 'done', 'blocked', 'needs_review'];
const MAX_UPLOAD_BYTES = 10 * 1024 * 1024;

/* ---------- 設定 ---------- */
function setup() {
  const ss = SpreadsheetApp.getActive();
  Object.entries(SHEETS).forEach(([name, header]) => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) {
      sh.appendRow(header);
      sh.setFrozenRows(1);
    }
  });
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('FLOW_TOKEN')) {
    props.setProperty('FLOW_TOKEN', Utilities.getUuid().replace(/-/g, ''));
  }
  if (!props.getProperty('FLOW_FOLDER_ID')) {
    const folder = DriveApp.createFolder('Flow 產出物');
    props.setProperty('FLOW_FOLDER_ID', folder.getId());
  }
  Logger.log('FLOW_TOKEN = ' + props.getProperty('FLOW_TOKEN'));
  Logger.log('產出物資料夾 = https://drive.google.com/drive/folders/' + props.getProperty('FLOW_FOLDER_ID'));
}

/* ---------- 進入點 ---------- */
function doGet(e) {
  return handle_(e.parameter || {});
}

function doPost(e) {
  let body = {};
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'bad_json' }); }
  return handle_(body);
}

function handle_(req) {
  if (!req.token || req.token !== PropertiesService.getScriptProperties().getProperty('FLOW_TOKEN')) {
    return json_({ ok: false, error: 'unauthorized' });
  }
  const actions = {
    spec: getSpec_, status: getStatus_, answers: getAnswers_, questions: getQuestions_,
    push_spec: pushSpec_, report: report_, ask: ask_, answer: answer_, upload: upload_,
  };
  const fn = actions[req.action];
  if (!fn) return json_({ ok: false, error: 'unknown_action', actions: Object.keys(actions) });
  if (!req.project && req.action !== 'answer') return json_({ ok: false, error: 'missing_project' });
  try {
    return json_(Object.assign({ ok: true }, fn(req)));
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  }
}

/* ---------- 讀取 ---------- */
function getSpec_(req) {
  const rows = rows_('specs').filter(r => r.project === req.project);
  if (!rows.length) throw new Error('no_spec');
  const want = req.version && req.version !== 'latest' ? Number(req.version) : null;
  const row = want ? rows.find(r => Number(r.version) === want) : rows.reduce((a, b) => (Number(b.version) > Number(a.version) ? b : a));
  if (!row) throw new Error('no_such_version');
  return { version: Number(row.version), title: row.title, spec: JSON.parse(row.specJson), specMd: row.specMd };
}

function getStatus_(req) {
  const latest = {};
  rows_('events').filter(r => r.project === req.project).forEach(r => { latest[r.node] = r; });
  const open = rows_('questions').filter(r => r.project === req.project && r.status === 'open').length;
  return { nodes: latest, openQuestions: open };
}

function getAnswers_(req) {
  const since = req.since ? new Date(req.since).getTime() : 0;
  return {
    answers: rows_('questions').filter(r =>
      r.project === req.project && r.status === 'answered' && (new Date(r.answeredAt).getTime() || 1) > since),
  };
}

function getQuestions_(req) {
  return { questions: rows_('questions').filter(r => r.project === req.project && (!req.status || r.status === req.status)) };
}

/* ---------- 寫入 ---------- */
function pushSpec_(req) {
  const spec = typeof req.spec === 'string' ? JSON.parse(req.spec) : req.spec;
  if (!spec || !Array.isArray(spec.nodes)) throw new Error('spec_needs_nodes');
  return withLock_(() => {
    const versions = rows_('specs').filter(r => r.project === req.project).map(r => Number(r.version));
    const version = versions.length ? Math.max.apply(null, versions) + 1 : 1;
    spec.version = version;
    sheet_('specs').appendRow([req.project, version, spec.title || req.project, 'frozen',
      JSON.stringify(spec), req.specMd || '', now_(), req.actor || '']);
    return { version };
  });
}

function report_(req) {
  if (!req.node) throw new Error('missing_node');
  if (STATUSES.indexOf(req.status) < 0) throw new Error('bad_status: ' + STATUSES.join('|'));
  return withLock_(() => {
    sheet_('events').appendRow([now_(), req.project, req.version || '', req.node, req.status,
      req.summary || '', JSON.stringify(req.artifacts || []), req.commit || '', req.actor || '']);
    return {};
  });
}

function ask_(req) {
  if (!req.question) throw new Error('missing_question');
  return withLock_(() => {
    const id = 'q' + Date.now().toString(36);
    sheet_('questions').appendRow([id, now_(), req.project, req.version || '', req.node || '',
      req.question, 'open', '', '', '', req.actor || '']);
    sheet_('events').appendRow([now_(), req.project, req.version || '', req.node || '', 'needs_review',
      '提問：' + req.question, '[]', '', req.actor || '']);
    return { id };
  });
}

/** 平台頁面或你自己回答問題時用 */
function answer_(req) {
  return withLock_(() => {
    const sh = sheet_('questions');
    const data = sh.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === req.id) {
        sh.getRange(i + 1, 7, 1, 4).setValues([['answered', req.answer || '', req.actor || '', now_()]]);
        return {};
      }
    }
    throw new Error('no_such_question');
  });
}

function upload_(req) {
  const bytes = Utilities.base64Decode(req.data || '');
  if (!bytes.length) throw new Error('empty_file');
  if (bytes.length > MAX_UPLOAD_BYTES) throw new Error('file_too_large');
  const root = DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty('FLOW_FOLDER_ID'));
  const sub = root.getFoldersByName(req.project);
  const folder = sub.hasNext() ? sub.next() : root.createFolder(req.project);
  const file = folder.createFile(Utilities.newBlob(bytes, req.mimeType || 'application/octet-stream', req.name || 'file'));
  return { url: file.getUrl(), id: file.getId() };
}

/* ---------- 工具 ---------- */
function sheet_(name) { return SpreadsheetApp.getActive().getSheetByName(name); }

function rows_(name) {
  const data = sheet_(name).getDataRange().getValues();
  const head = data.shift() || [];
  return data.map(r => Object.fromEntries(head.map((h, i) => [h, r[i] instanceof Date ? r[i].toISOString() : r[i]])));
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function now_() { return new Date().toISOString(); }

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
