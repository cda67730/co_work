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
  projects:  ['project', 'title', 'purpose', 'due', 'status', 'themesJson', 'draftJson', 'createdAt', 'createdBy', 'updatedAt', 'round'],
  needs:     ['id', 'ts', 'project', 'author', 'topic', 'pain', 'want', 'plus', 'commentsJson', 'round'],
};
const PHASES = ['collecting', 'organizing', 'drafting', 'developing'];
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
  const q = e.parameter || {};
  if (q.action) return handle_(q);
  return HtmlService.createTemplateFromFile('Page').evaluate()
    .setTitle('使用者共創')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* ---------- 平台頁面（同事用） ----------
 * 同事的連結要用另一個部署：
 *   執行身分：存取網頁應用程式的使用者
 *   誰可以存取：任何擁有 Google 帳戶的使用者（或你的組織）
 * 並把這份 Sheet 共用給同事（編輯者），才能回答問題。
 * 在給 Agent 的「所有人」部署裡，下面這些函式會拒絕匿名使用者。
 */
function pageUser_() {
  const active = Session.getActiveUser().getEmail();
  const effective = Session.getEffectiveUser().getEmail();
  if (!active || active !== effective) throw new Error('請用同事專用的連結開啟，並登入 Google 帳號。');
  return active;
}

function isOwner_(user) {
  try { return user === SpreadsheetApp.getActive().getOwner().getEmail(); } catch (err) { return false; }
}
function ownerOnly_() {
  const user = pageUser_();
  if (!isOwner_(user)) throw new Error('只有專案管理者可以做這件事。');
  return user;
}
const parse_ = (s, d) => { try { return s ? JSON.parse(s) : d; } catch (err) { return d; } };

function pageProjects() {
  const user = pageUser_();
  const latest = {};
  rows_('specs').forEach(r => {
    if (!latest[r.project] || Number(r.version) > Number(latest[r.project].version)) latest[r.project] = r;
  });
  const needCount = {};
  rows_('needs').forEach(r => { needCount[r.project] = (needCount[r.project] || 0) + 1; });
  const list = rows_('projects').map(r => ({
    project: r.project, title: r.title, purpose: r.purpose, due: r.due, status: r.status, round: Number(r.round) || 1,
    version: latest[r.project] ? Number(latest[r.project].version) : 0, needs: needCount[r.project] || 0,
  }));
  // 只有規格、沒有開專案紀錄的（例如用 flow push-spec 發布的）視為開發中
  Object.values(latest).forEach(r => {
    if (!list.some(p => p.project === r.project)) {
      list.push({ project: r.project, title: r.title, purpose: '', due: '', status: 'developing', round: 1, version: Number(r.version), needs: 0, external: true });
    }
  });
  return { user: user, isOwner: isOwner_(user), projects: list };
}

function pageCreateProject(p) {
  const user = ownerOnly_();
  const id = String(p.project || '').trim().toLowerCase();
  if (!/^[a-z0-9][a-z0-9-]{1,40}$/.test(id)) throw new Error('專案代號只能用英文小寫、數字和 -，例如 leave-form。');
  if (!String(p.title || '').trim()) throw new Error('請填專案名稱。');
  return withLock_(() => {
    if (rows_('projects').some(r => r.project === id) || rows_('specs').some(r => r.project === id)) throw new Error('這個專案代號已經用過了。');
    const now = now_();
    appendObj_('projects', { project: id, title: String(p.title).trim(), purpose: String(p.purpose || '').trim(), due: String(p.due || ''),
      status: 'collecting', themesJson: '', draftJson: '', createdAt: now, createdBy: user, updatedAt: now, round: 1 });
    return { project: id };
  });
}

function pageSetPhase(project, phase) {
  ownerOnly_();
  if (PHASES.indexOf(phase) < 0 || phase === 'developing') throw new Error('不能切換到這個階段。');
  withLock_(() => updateRow_('projects', 'project', project, { status: phase, updatedAt: now_() }));
  return pagePlanning(project);
}

/** 已發布的專案開始下一輪：回到收集需求，規格草稿以目前版本為底 */
function pageNextRound(project) {
  const user = ownerOnly_();
  withLock_(() => {
    let p = rows_('projects').find(r => r.project === project);
    const cur = getSpec_({ project: project });
    if (!p) {   // 用 flow push-spec 發布、沒有開專案紀錄的，補一筆
      appendObj_('projects', { project: project, title: cur.title || project, purpose: cur.spec.purpose || '', due: '', status: 'developing',
        themesJson: '', draftJson: '', createdAt: now_(), createdBy: user, updatedAt: now_(), round: 1 });
      p = rows_('projects').find(r => r.project === project);
    }
    if (p.status !== 'developing') throw new Error('目前已經在進行新一輪了。');
    const base = Object.assign({}, cur.spec); delete base.version;
    updateRow_('projects', 'project', project, {
      status: 'collecting', round: (Number(p.round) || 1) + 1, themesJson: '[]',
      draftJson: JSON.stringify(base, null, 2), updatedAt: now_(),
    });
  });
  return pagePlanning(project);
}

/** 取消這一輪，回到開發中（這一輪沒有需求時，輪次也一併退回） */
function pageCancelRound(project) {
  ownerOnly_();
  withLock_(() => {
    const p = rows_('projects').find(r => r.project === project);
    const round = Number(p && p.round) || 1;
    if (!p || round < 2 || p.status === 'developing') throw new Error('沒有進行中的新一輪。');
    const has = rows_('needs').some(r => r.project === project && (Number(r.round) || 1) === round);
    updateRow_('projects', 'project', project, { status: 'developing', round: has ? round : round - 1, updatedAt: now_() });
  });
  return pagePlanning(project);
}

/** 需求、主題、草稿（前半段） */
function pagePlanning(project) {
  const user = pageUser_();
  const p = rows_('projects').find(r => r.project === project);
  const needs = rows_('needs').filter(r => r.project === project).map(r => ({
    id: r.id, ts: r.ts, author: r.author, topic: r.topic || r.now || '', pain: r.pain, want: r.want, round: Number(r.round) || 1,
    plus: String(r.plus || '').split(',').filter(Boolean), comments: parse_(r.commentsJson, []),
  }));
  return {
    user: user, isOwner: isOwner_(user), project: project, round: p ? (Number(p.round) || 1) : 1,
    info: p ? { title: p.title, purpose: p.purpose, due: p.due, status: p.status } : null,
    needs: needs, themes: p ? parse_(p.themesJson, []) : [], draft: p ? (p.draftJson || '') : '',
  };
}

function pageAddNeed(project, n) {
  const user = pageUser_();
  const p = rows_('projects').find(r => r.project === project);
  if (!p) throw new Error('找不到這個專案。');
  if (p.status !== 'collecting' && !isOwner_(user)) throw new Error('這個專案已經停止收集需求。');
  const topic = String(n.topic || '').trim(), pain = String(n.pain || '').trim(), want = String(n.want || '').trim();
  if (!topic) throw new Error('請填「主題」。');
  if (!pain && !want) throw new Error('「遇到的問題」和「希望變成什麼」至少填一個。');
  withLock_(() => appendObj_('needs', { id: 'n' + Date.now().toString(36).slice(-5) + Math.random().toString(36).slice(2, 5), ts: now_(),
    project: project, author: user, topic: topic, pain: pain, want: want, plus: '', commentsJson: '[]', round: Number(p.round) || 1 }));
  return pagePlanning(project);
}

function pagePlusOne(project, id) {
  const user = pageUser_();
  withLock_(() => {
    const n = rows_('needs').find(r => r.id === id && r.project === project);
    if (!n) throw new Error('找不到這則需求。');
    const plus = String(n.plus || '').split(',').filter(Boolean);
    const i = plus.indexOf(user);
    if (i >= 0) plus.splice(i, 1); else plus.push(user);
    updateRow_('needs', 'id', id, { plus: plus.join(',') });
  });
  return pagePlanning(project);
}

function pageComment(project, id, text) {
  const user = pageUser_();
  const t = String(text || '').trim();
  if (!t) throw new Error('留言不能是空的。');
  withLock_(() => {
    const n = rows_('needs').find(r => r.id === id && r.project === project);
    if (!n) throw new Error('找不到這則需求。');
    const list = parse_(n.commentsJson, []);
    list.push({ by: user, ts: now_(), text: t });
    updateRow_('needs', 'id', id, { commentsJson: JSON.stringify(list) });
  });
  return pagePlanning(project);
}

function pageSaveThemes(project, themes) {
  ownerOnly_();
  if (!Array.isArray(themes)) throw new Error('主題格式不正確。');
  const clean = themes.map((t, i) => ({
    id: String(t.id || 't' + (i + 1)), title: String(t.title || '').trim(), summary: String(t.summary || '').trim(),
    needIds: Array.isArray(t.needIds) ? t.needIds.map(String) : [], conflict: String(t.conflict || '').trim(), decision: String(t.decision || '').trim(),
  })).filter(t => t.title);
  withLock_(() => updateRow_('projects', 'project', project, { themesJson: JSON.stringify(clean), updatedAt: now_() }));
  return pagePlanning(project);
}

function pageSaveDraft(project, text) {
  ownerOnly_();
  const spec = parse_(text, null);
  if (!spec || !Array.isArray(spec.nodes) || !spec.nodes.length) throw new Error('規格要是 JSON，而且至少要有一個節點（nodes）。');
  withLock_(() => updateRow_('projects', 'project', project, { draftJson: JSON.stringify(spec, null, 2), updatedAt: now_() }));
  return pagePlanning(project);
}

function pagePublish(project) {
  const user = ownerOnly_();
  const p = rows_('projects').find(r => r.project === project);
  const spec = p && parse_(p.draftJson, null);
  if (!spec || !Array.isArray(spec.nodes) || !spec.nodes.length) throw new Error('還沒有可以發布的規格草稿。');
  spec.project = project;
  spec.title = spec.title || p.title;
  spec.purpose = spec.purpose || p.purpose;
  const r = pushSpec_({ project: project, spec: spec, specMd: '', actor: user });
  withLock_(() => updateRow_('projects', 'project', project, { status: 'developing', updatedAt: now_() }));
  return { version: r.version };
}

function pageData(project) {
  const user = pageUser_();
  const spec = getSpec_({ project: project });
  const events = rows_('events').filter(r => r.project === project).map(r => {
    let artifacts = [];
    try { artifacts = JSON.parse(r.artifacts || '[]'); } catch (err) {}
    return Object.assign({}, r, { artifacts: artifacts });
  });
  const questions = rows_('questions').filter(r => r.project === project);
  return { user: user, isOwner: isOwner_(user), project: project, version: spec.version, title: spec.title, spec: spec.spec, events: events, questions: questions };
}

function pageAnswer(project, id, answer) {
  const user = pageUser_();
  if (!String(answer || '').trim()) throw new Error('回覆不能是空的。');
  answer_({ id: id, answer: String(answer).trim(), actor: user });
  return pageData(project);
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
function sheet_(name) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(name);
  if (!sh && SHEETS[name]) {           // 新增的分頁不必重跑 setup
    sh = ss.insertSheet(name);
    sh.appendRow(SHEETS[name]);
    sh.setFrozenRows(1);
  } else if (sh && SHEETS[name]) {     // 舊版分頁缺欄位就補在最後
    const head = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0];
    const legacy = { topic: 'now' };
    const missing = SHEETS[name].filter(h => head.indexOf(h) < 0 && head.indexOf(legacy[h]) < 0);
    if (missing.length) sh.getRange(1, head.filter(String).length + 1, 1, missing.length).setValues([missing]);
  }
  return sh;
}

/** 依標題列欄位順序寫入一列（新舊欄位順序不同也不會錯位） */
function appendObj_(name, obj) {
  const sh = sheet_(name);
  const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const legacy = { now: 'topic' };
  sh.appendRow(head.map(h => (obj[h] !== undefined ? obj[h] : obj[legacy[h]] !== undefined ? obj[legacy[h]] : '')));
}

/** 依 key 欄位找到列並更新指定欄位（rowIndex 從 1 起算，含標題列） */
function updateRow_(name, key, value, patch) {
  const sh = sheet_(name);
  const data = sh.getDataRange().getValues();
  const head = data[0];
  const k = head.indexOf(key);
  for (let i = 1; i < data.length; i++) {
    if (data[i][k] === value) {
      Object.entries(patch).forEach(([col, v]) => {
        const c = head.indexOf(col);
        if (c >= 0) sh.getRange(i + 1, c + 1).setValue(v);
      });
      return true;
    }
  }
  return false;
}

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
