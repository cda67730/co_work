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
 *   FLOW_TOKEN        API token（setup 自動產生）
 *   FLOW_FOLDER_ID    上傳檔案的 Drive 資料夾（setup 自動建立）
 *   GEMINI_API_KEY    整理需求、畫圖及規格用的 Gemini API key（自己新增）
 *   GEMINI_MODEL_FAST 選填，預設 gemini-2.5-flash（整理需求）
 *   GEMINI_MODEL_SPEC 選填，預設 gemini-2.5-pro（畫圖及規格）
 */

const SHEETS = {
  specs:     ['project', 'version', 'title', 'status', 'specJson', 'specMd', 'createdAt', 'createdBy'],
  events:    ['ts', 'project', 'version', 'node', 'status', 'summary', 'artifacts', 'commit', 'actor'],
  questions: ['id', 'ts', 'project', 'version', 'node', 'question', 'status', 'answer', 'answeredBy', 'answeredAt', 'actor'],
  projects:  ['project', 'title', 'purpose', 'due', 'status', 'themesJson', 'draftJson', 'createdAt', 'createdBy', 'updatedAt', 'round'],
  needs:     ['id', 'ts', 'project', 'author', 'topic', 'pain', 'want', 'plus', 'commentsJson', 'round'],
  drafts:    ['project', 'round', 'v', 'ts', 'by', 'note', 'specJson'],
  chats:     ['ts', 'project', 'round', 'stage', 'role', 'author', 'text', 'log'],
  comments:  ['id', 'ts', 'project', 'round', 'node', 'author', 'text', 'status', 'note'],
  activity:  ['ts', 'project', 'actor', 'kind', 'text'],
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
    log_(id, user, 'project', '開啟專案「' + String(p.title).trim() + '」，開始收集需求');
    return { project: id };
  });
}

const PHASE_ZH = { collecting: '收集', organizing: '整理', drafting: '畫圖及規格', developing: '開發' };
function pageSetPhase(project, phase) {
  const user = ownerOnly_();
  if (PHASES.indexOf(phase) < 0 || phase === 'developing') throw new Error('不能切換到這個階段。');
  withLock_(() => updateRow_('projects', 'project', project, { status: phase, updatedAt: now_() }));
  log_(project, user, 'phase', '進入「' + PHASE_ZH[phase] + '」階段');
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
    const round = (Number(p.round) || 1) + 1;
    updateRow_('projects', 'project', project, {
      status: 'collecting', round: round, themesJson: '[]', draftJson: JSON.stringify(base, null, 2), updatedAt: now_(),
    });
    appendObj_('drafts', { project: project, round: round, v: 1, ts: now_(), by: user, note: '以上線的 v' + cur.version + ' 為底稿', specJson: JSON.stringify(base) });
    log_(project, user, 'round', '開始第 ' + round + ' 輪需求（目前上線 v' + cur.version + '）');
  });
  return pagePlanning(project);
}

/** 取消這一輪，回到開發中（這一輪沒有需求時，輪次也一併退回） */
function pageCancelRound(project) {
  const user = ownerOnly_();
  log_(project, user, 'round', '取消這一輪，回到開發');
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
  const round = p ? (Number(p.round) || 1) : 1;
  const inRound = r => r.project === project && (Number(r.round) || 1) === round;
  return {
    user: user, isOwner: isOwner_(user), project: project, round: round,
    info: p ? { title: p.title, purpose: p.purpose, due: p.due, status: p.status } : null,
    needs: needs, themes: p ? parse_(p.themesJson, []) : [], draft: p ? (p.draftJson || '') : '',
    drafts: rows_('drafts').filter(inRound).map(r => ({ v: Number(r.v), ts: r.ts, by: r.by, note: r.note })),
    chats: rows_('chats').filter(inRound).map(r => ({ ts: r.ts, stage: r.stage, role: r.role, author: r.author, text: r.text, log: parse_(r.log, []) })),
    comments: rows_('comments').filter(inRound).map(r => ({ id: r.id, ts: r.ts, node: r.node, author: r.author, text: r.text, status: r.status, note: r.note })),
    aiReady: !!PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY'),
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
  log_(project, user, 'need', '提出需求「' + topic + '」');
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
  delete spec.openQuestions;
  spec.nodes.forEach(n => delete n.locked);
  const r = pushSpec_({ project: project, spec: spec, specMd: '', actor: user });
  withLock_(() => updateRow_('projects', 'project', project, { status: 'developing', updatedAt: now_() }));
  log_(project, user, 'publish', '規格定稿，發布 v' + r.version + '（' + spec.nodes.length + ' 個節點），交給 AI Agent 開發');
  return { version: r.version };
}

/* ================= AI（Gemini） ================= */
function gemini_(kind, system, user, schema) {
  const props = PropertiesService.getScriptProperties();
  const key = props.getProperty('GEMINI_API_KEY');
  if (!key) throw new Error('還沒設定 Gemini API key：到 Apps Script「專案設定」→「指令碼屬性」新增 GEMINI_API_KEY。');
  const model = kind === 'spec' ? (props.getProperty('GEMINI_MODEL_SPEC') || 'gemini-2.5-pro') : (props.getProperty('GEMINI_MODEL_FAST') || 'gemini-2.5-flash');
  const res = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/' + model + ':generateContent', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { 'x-goog-api-key': key },
    payload: JSON.stringify({
      systemInstruction: { parts: [{ text: system }] },
      contents: [{ role: 'user', parts: [{ text: user }] }],
      generationConfig: Object.assign({ responseMimeType: 'application/json', temperature: 0.4 }, schema ? { responseSchema: schema } : {}),
    }),
  });
  const code = res.getResponseCode(), body = parse_(res.getContentText(), {});
  if (code !== 200) throw new Error('AI 呼叫失敗（' + code + '）：' + ((body.error && body.error.message) || '請稍後再試'));
  const c = body.candidates && body.candidates[0];
  const text = c && c.content && c.content.parts ? c.content.parts.filter(x => !x.thought).map(x => x.text || '').join('') : '';
  const out = parse_(text.replace(/^```(?:json)?\s*|\s*```$/g, ''), null);
  if (!out) throw new Error('AI 回覆的格式不正確，請再試一次。');
  return out;
}

const NODE_TYPES = ['input', 'llm', 'retrieve', 'condition', 'tool', 'output'];

/* Gemini 回覆結構（responseSchema），避免 AI 自由發揮格式或抄回範本 */
const STR = { type: 'STRING' }, STRS = { type: 'ARRAY', items: { type: 'STRING' } };
const SC_ACC = { type: 'ARRAY', items: { type: 'OBJECT', properties: { id: STR, text: STR }, required: ['id', 'text'] } };
const SC_NODE = { type: 'OBJECT', properties: { id: STR, type: { type: 'STRING', enum: NODE_TYPES }, label: STR, description: STR, inputs: STRS, outputs: STRS, acceptance: SC_ACC },
  required: ['id', 'type', 'label', 'description', 'inputs', 'outputs', 'acceptance'] };
const SC_EDGE = { type: 'OBJECT', properties: { from: STR, to: STR, label: STR }, required: ['from', 'to'] };
const SC_SPEC = { type: 'OBJECT', properties: { title: STR, purpose: STR, nodes: { type: 'ARRAY', items: SC_NODE }, edges: { type: 'ARRAY', items: SC_EDGE }, decisions: STRS, openQuestions: STRS },
  required: ['title', 'purpose', 'nodes', 'edges', 'decisions', 'openQuestions'] };
const SC_FRESH = { type: 'OBJECT', properties: { reply: STR, spec: SC_SPEC }, required: ['reply', 'spec'] };
const SC_OPS = { type: 'OBJECT', properties: {
  reply: STR,
  ops: { type: 'ARRAY', items: { type: 'OBJECT', properties: {
    op: { type: 'STRING', enum: ['add_node', 'update_node', 'remove_node', 'connect', 'disconnect', 'set_decisions', 'set_open_questions', 'set_title', 'set_purpose'] },
    id: STR, node: SC_NODE,
    set: { type: 'OBJECT', properties: { type: { type: 'STRING', enum: NODE_TYPES }, label: STR, description: STR, inputs: STRS, outputs: STRS, acceptance: SC_ACC } },
    from: STR, to: STR, label: STR, list: STRS, text: STR }, required: ['op'] } },
  resolved: STRS,
  unresolved: { type: 'ARRAY', items: { type: 'OBJECT', properties: { id: STR, reason: STR }, required: ['id', 'reason'] } } },
  required: ['reply', 'ops', 'resolved', 'unresolved'] };
const SC_THEMES = { type: 'OBJECT', properties: { reply: STR, themes: { type: 'ARRAY', items: { type: 'OBJECT',
  properties: { id: STR, title: STR, summary: STR, needIds: STRS, conflict: STR, decision: STR }, required: ['id', 'title', 'summary', 'needIds', 'conflict', 'decision'] } } },
  required: ['reply', 'themes'] };
const isPlaceholder_ = v => /^\s*(…|\.{3}|-|—|xxx|todo|tbd)?\s*$/i.test(String(v == null ? '' : v));

function chatAppend_(project, round, stage, role, author, text, log) {
  appendObj_('chats', { ts: now_(), project: project, round: round, stage: stage, role: role, author: author, text: text, log: JSON.stringify(log || []) });
}
function recentChat_(project, round, stage, n) {
  return rows_('chats').filter(r => r.project === project && (Number(r.round) || 1) === round && r.stage === stage)
    .slice(-n).map(r => (r.role === 'user' ? '使用者' : 'AI') + '：' + r.text).join('\n');
}
function needsText_(needs) {
  return needs.map(n => '[' + n.id + '] 主題：' + (n.topic || '—') + '（' + n.plus.length + ' 人附議）\n問題：' + (n.pain || '—') + '\n希望：' + (n.want || '—') +
    (n.comments.length ? '\n留言：' + n.comments.map(c => c.text).join('／') : '')).join('\n\n');
}

/** ② 整理：AI 依需求產生或修改主題（message 為空 = 重新整理） */
function pageThemesChat(project, message) {
  const user = ownerOnly_();
  const plan = pagePlanning(project);
  const p = plan.info, round = plan.round;
  const needs = plan.needs.filter(n => n.round === round);
  if (!needs.length) throw new Error('這一輪還沒有需求可以整理。');
  const msg = String(message || '').trim() || '請依目前的需求整理主題。';
  const system = [
    '你是需求分析師，幫團隊把同事的需求歸納成主題。用繁體中文、白話。',
    '規則：',
    '1. 主題 3 到 7 個，每個主題列出來源需求編號 needIds，一則需求可屬於多個主題',
    '2. 同事之間意見衝突的地方不要合併成模糊說法，寫在 conflict，交給人決定',
    '3. 不要加入需求裡沒有的內容；附議多的需求要反映在 summary',
    '4. 已經有主題時，依使用者的指示修改：沒提到的主題保持原樣，主題的 id 和 decision 一律保留',
    '回覆欄位：reply 用一兩句話說明你做了什麼；themes 每個主題有 id（沿用原本的 id，新主題給空字串）、title（10 字內）、summary（1 到 2 句）、needIds（需求編號，例如 nsckk70e4）、conflict（沒有衝突就給空字串）、decision（保留原值，沒有就空字串）。',
    '所有內容都要根據上面的需求實際撰寫，不可以出現「…」或範例文字。',
  ].join('\n');
  const prompt = '專案：' + p.title + (p.purpose ? '（目的：' + p.purpose + '）' : '') + (round > 1 ? '\n這是第 ' + round + ' 輪，系統已上線，以下是新增需求。' : '') +
    '\n\n需求：\n' + needsText_(needs) +
    '\n\n目前主題：\n' + (plan.themes.length ? JSON.stringify(plan.themes) : '（還沒有）') +
    '\n\n最近對話：\n' + (recentChat_(project, round, 'themes', 6) || '（無）') +
    '\n\n這次的指示：' + msg;
  const out = gemini_('fast', system, prompt, SC_THEMES);
  if (!Array.isArray(out.themes) || !out.themes.length || out.themes.some(t => isPlaceholder_(t.title))) throw new Error('AI 沒有回傳有效的主題，請再試一次。');
  const old = Object.fromEntries(plan.themes.map(t => [t.id, t]));
  let n = plan.themes.length;
  const themes = out.themes.map(t => {
    const o = old[t.id];
    return { id: o ? o.id : 't' + (++n) + Math.random().toString(36).slice(2, 4), title: String(t.title || '').trim(), summary: String(t.summary || '').trim(),
      needIds: Array.isArray(t.needIds) ? t.needIds.map(String) : [], conflict: String(t.conflict || '').trim(), decision: o ? (t.decision || o.decision || '') : String(t.decision || '') };
  }).filter(t => t.title);
  withLock_(() => {
    updateRow_('projects', 'project', project, { themesJson: JSON.stringify(themes), updatedAt: now_() });
    chatAppend_(project, round, 'themes', 'user', user, msg);
    chatAppend_(project, round, 'themes', 'ai', 'AI', String(out.reply || '已更新主題。'));
  });
  log_(project, user, 'themes', 'AI 整理主題：' + themes.map(t => t.title).join('、'));
  return pagePlanning(project);
}

/** ③ 畫圖及規格：AI 產生初版或依意見修改（只輸出差異） */
function pageSpecChat(project, message, includeComments, regenerate) {
  const user = ownerOnly_();
  const plan = pagePlanning(project);
  const p = plan.info, round = plan.round;
  const draft = parse_(plan.draft, null);
  const fresh = !!regenerate || !draft || !Array.isArray(draft.nodes) || !draft.nodes.length;
  const open = includeComments ? plan.comments.filter(c => c.status === 'open') : [];
  const msg = String(message || '').trim() || (fresh ? (regenerate ? '請依需求與主題重新產生整份流程與規格。' : '請依需求與主題產生初版流程與規格。') : open.length ? '請把同事的意見整理進規格。' : '');
  if (!msg) throw new Error('請輸入要修改的內容。');
  const rules = [
    '你是系統分析師，和團隊一起設計一個工作流程。之後 AI Agent 會照規格開發，不寫程式的同事會依驗收標準驗收。用繁體中文。',
    '節點 type 只能是：input（輸入）、llm（AI 判斷或生成）、retrieve（查資料）、condition（分流）、tool（程式處理）、output（輸出）。',
    '每個節點：id（英文小寫加底線）、type、label（白話）、description（白話）、inputs、outputs（英文小寫欄位名，下游要的欄位上游一定要有）、acceptance（除了 input 都要有，具體可測試，例如「問『學費怎麼繳』要分到行政類」，格式 [{"id":"a1","text":"…"}]）。',
    'condition 的分支名稱寫在連線 label。已決定的事放 decisions；還沒決定的放 openQuestions，不要自己做決定。不要加入需求沒有的功能。',
    '所有內容都要根據下面的需求實際撰寫，不可以出現「…」、範例文字或空白的名稱。',
  ];
  let system, prompt;
  const ctx = '專案：' + p.title + (p.purpose ? '（目的：' + p.purpose + '）' : '') +
    '\n\n主題：\n' + (plan.themes.map(t => '■ ' + t.title + '：' + t.summary + (t.decision ? '\n  已決定：' + t.decision : t.conflict ? '\n  尚未決定：' + t.conflict : '')).join('\n') || '（無）') +
    '\n\n這一輪的需求：\n' + needsText_(plan.needs.filter(n => n.round === round));
  if (fresh) {
    system = rules.concat(['流程拆成 4 到 10 個節點，從 input 節點開始，到 output 節點結束，每個節點都要有連線。',
      '回覆欄位：reply 用一兩句話說明這份流程；spec 是完整規格（title、purpose、nodes、edges、decisions、openQuestions）。edges 的 from、to 要用節點 id。']).join('\n');
    prompt = ctx + '\n\n指示：' + msg;
  } else {
    const locked = draft.nodes.filter(n => n.locked).map(n => n.id);
    system = rules.concat([
      '現在要修改既有的規格。只輸出「修改動作」，不要重寫整份。沒提到的節點不要動；locked 為 true 的節點完全不能修改或刪除（可以連線到它）。',
      '修改動作（ops 陣列，每個動作用 op 指定種類）：',
      'add_node：新增步驟，完整節點放在 node；update_node：修改步驟，id 指定節點、要改的欄位放在 set；remove_node：刪除步驟，id 指定節點；',
      'connect：新增連線，用 from、to（節點 id），分支名稱放 label；disconnect：移除連線，用 from、to；',
      'set_decisions、set_open_questions：用 list 給完整的新清單；set_title、set_purpose：用 text 給新內容。',
      '有同事意見時，逐則判斷：採納的意見編號放 resolved；不採納或需要人決定的放 unresolved，並寫 reason。沒有就給空陣列。',
      'reply 用一兩句話說明改了什麼。',
    ]).join('\n');
    prompt = ctx + '\n\n目前規格：\n' + JSON.stringify(draft) +
      (locked.length ? '\n\n鎖定節點（不能改）：' + locked.join(', ') : '') +
      (open.length ? '\n\n同事意見：\n' + open.map(c => '[' + c.id + '] 節點 ' + (c.node || '整體') + '｜' + c.author.replace(/@.*/, '') + '：' + c.text).join('\n') : '') +
      '\n\n最近對話：\n' + (recentChat_(project, round, 'spec', 6) || '（無）') +
      '\n\n這次的指示：' + msg;
  }
  const out = gemini_('spec', system, prompt, fresh ? SC_FRESH : SC_OPS);
  let spec, log = [];
  if (fresh) {
    if (!out.spec || !Array.isArray(out.spec.nodes)) throw new Error('AI 沒有回傳規格，請再試一次。');
    spec = normalizeSpec_(out.spec);
    if (spec.nodes.length < 2 || spec.nodes.some(n => isPlaceholder_(n.label) || isPlaceholder_(n.id))) throw new Error('AI 回傳的流程圖不完整（步驟名稱空白或只有範本），沒有存檔，請再試一次。');
    log = [(regenerate ? '重新產生 ' : '產生 ') + spec.nodes.length + ' 個步驟'];
  } else {
    const r = applySpecOps_(draft, out.ops || []); spec = r.spec; log = r.log;
  }
  withLock_(() => {
    const v = saveDraft_(project, round, spec, 'AI', String(out.reply || '').slice(0, 80) || (fresh ? (regenerate ? 'AI 重新產生' : 'AI 產生初版') : 'AI 修改'));
    chatAppend_(project, round, 'spec', 'user', user, msg + (open.length ? '（附 ' + open.length + ' 則同事意見）' : ''));
    chatAppend_(project, round, 'spec', 'ai', 'AI', String(out.reply || '已更新。') + '（草稿 v' + v + '）', log);
    const unres = Object.fromEntries((out.unresolved || []).map(u => [u.id, u.reason]));
    (out.resolved || []).forEach(id => updateRow_('comments', 'id', id, { status: 'resolved', note: '已由 AI 納入草稿 v' + v }));
    Object.entries(unres).forEach(([id, reason]) => updateRow_('comments', 'id', id, { note: 'AI：' + reason }));
  });
  log_(project, user, 'draft', (fresh ? 'AI 產生初版規格' : 'AI 修改規格') + '：' + (String(out.reply || '').slice(0, 60)));
  return pagePlanning(project);
}

function normalizeSpec_(spec) {
  spec.nodes = (spec.nodes || []).filter(n => n && n.id).map(normNode_);
  spec.edges = (spec.edges || []).filter(e => e && e.from && e.to);
  spec.decisions = (Array.isArray(spec.decisions) ? spec.decisions : []).filter(x => !isPlaceholder_(x));
  spec.openQuestions = (Array.isArray(spec.openQuestions) ? spec.openQuestions : []).filter(x => !isPlaceholder_(x));
  const ids = new Set(spec.nodes.map(n => n.id));
  spec.edges = spec.edges.filter(e => ids.has(e.from) && ids.has(e.to));
  return spec;
}
function normNode_(n) {
  const o = Object.assign({}, n);
  o.id = String(o.id).trim().toLowerCase().replace(/[^a-z0-9_]/g, '_');
  if (NODE_TYPES.indexOf(o.type) < 0) o.type = 'tool';
  ['inputs', 'outputs'].forEach(k => { o[k] = Array.isArray(o[k]) ? o[k].map(String) : []; });
  o.acceptance = (Array.isArray(o.acceptance) ? o.acceptance : []).map((a, i) =>
    typeof a === 'string' ? { id: 'a' + (i + 1), text: a } : { id: String(a.id || 'a' + (i + 1)), text: String(a.text || '') }).filter(a => a.text);
  return o;
}

function applySpecOps_(draft, ops) {
  const spec = JSON.parse(JSON.stringify(draft)), log = [];
  const find = id => spec.nodes.find(n => n.id === id);
  const label = id => (find(id) || {}).label || id;
  spec.edges = spec.edges || [];
  (ops || []).forEach(o => {
    if (!o || !o.op) return;
    if (o.op === 'add_node' && o.node && o.node.id) {
      const n = normNode_(o.node);
      if (find(n.id) || isPlaceholder_(n.label)) return;
      spec.nodes.push(n); log.push('＋ 新增「' + n.label + '」');
    } else if (o.op === 'update_node') {
      const n = find(o.id); if (!n || n.locked) return;
      const set = o.set || {};
      Object.keys(set).forEach(k => { if (['type', 'label', 'description', 'inputs', 'outputs', 'acceptance', 'config'].indexOf(k) >= 0) n[k] = set[k]; });
      Object.assign(n, normNode_(n)); log.push('～ 修改「' + n.label + '」');
    } else if (o.op === 'remove_node') {
      const n = find(o.id); if (!n || n.locked) return;
      spec.nodes = spec.nodes.filter(x => x.id !== o.id);
      spec.edges = spec.edges.filter(e => e.from !== o.id && e.to !== o.id); log.push('－ 刪除「' + n.label + '」');
    } else if (o.op === 'connect' && find(o.from) && find(o.to) && !spec.edges.some(e => e.from === o.from && e.to === o.to)) {
      spec.edges.push({ from: o.from, to: o.to, label: o.label || '' }); log.push('連線 ' + label(o.from) + ' → ' + label(o.to));
    } else if (o.op === 'disconnect') {
      const b = spec.edges.length; spec.edges = spec.edges.filter(e => !(e.from === o.from && e.to === o.to));
      if (spec.edges.length < b) log.push('移除連線 ' + label(o.from) + ' → ' + label(o.to));
    } else if (o.op === 'set_decisions' || o.op === 'set_open_questions') {
      const f = o.op === 'set_decisions' ? 'decisions' : 'openQuestions';
      spec[f] = (o.list || []).filter(x => !isPlaceholder_(x)); log.push('更新' + (f === 'decisions' ? '已決定事項' : '待決定事項'));
    } else if ((o.op === 'set_title' || o.op === 'set_purpose') && !isPlaceholder_(o.text)) {
      spec[o.op === 'set_title' ? 'title' : 'purpose'] = o.text; log.push('更新' + (o.op === 'set_title' ? '標題' : '目的'));
    } else if (o.op === 'set' && ['decisions', 'openQuestions', 'title', 'purpose'].indexOf(o.field) >= 0) {
      spec[o.field] = o.value; log.push('更新 ' + ({ decisions: '已決定事項', openQuestions: '待決定事項', title: '標題', purpose: '目的' })[o.field]);
    }
  });
  return { spec: spec, log: log };
}

function saveDraft_(project, round, spec, by, note) {
  const vs = rows_('drafts').filter(r => r.project === project && (Number(r.round) || 1) === round).map(r => Number(r.v));
  const v = vs.length ? Math.max.apply(null, vs) + 1 : 1;
  appendObj_('drafts', { project: project, round: round, v: v, ts: now_(), by: by, note: note, specJson: JSON.stringify(spec) });
  updateRow_('projects', 'project', project, { draftJson: JSON.stringify(spec, null, 2), updatedAt: now_() });
  return v;
}

/** 手動修改節點（patch 為 null 代表刪除）；手動改過的節點自動鎖定，AI 不會再動 */
function pageEditNode(project, nodeId, patch) {
  const user = ownerOnly_();
  const plan = pagePlanning(project), draft = parse_(plan.draft, null);
  if (!draft || !Array.isArray(draft.nodes)) throw new Error('還沒有規格草稿。');
  let note;
  if (patch === null) {
    const n = draft.nodes.find(x => x.id === nodeId); if (!n) throw new Error('找不到這個節點。');
    draft.nodes = draft.nodes.filter(x => x.id !== nodeId);
    draft.edges = (draft.edges || []).filter(e => e.from !== nodeId && e.to !== nodeId);
    note = '手動刪除「' + n.label + '」';
  } else if (patch.newNode) {
    const n = normNode_(Object.assign({ id: 'step_' + Date.now().toString(36).slice(-4), type: 'tool', label: '新步驟' }, patch.newNode));
    n.locked = true; draft.nodes.push(n);
    if (patch.after && draft.nodes.some(x => x.id === patch.after)) draft.edges = (draft.edges || []).concat([{ from: patch.after, to: n.id, label: '' }]);
    note = '手動新增「' + n.label + '」';
  } else {
    const n = draft.nodes.find(x => x.id === nodeId); if (!n) throw new Error('找不到這個節點。');
    ['type', 'label', 'description', 'inputs', 'outputs', 'acceptance'].forEach(k => { if (patch[k] !== undefined) n[k] = patch[k]; });
    Object.assign(n, normNode_(n));
    n.locked = patch.locked !== undefined ? !!patch.locked : true;
    if (patch.edges) {   // [{to,label}] 取代這個節點連出去的線
      draft.edges = (draft.edges || []).filter(e => e.from !== n.id).concat(patch.edges.filter(e => e.to && draft.nodes.some(x => x.id === e.to)).map(e => ({ from: n.id, to: e.to, label: e.label || '' })));
    }
    note = (Object.keys(patch).length === 1 && patch.locked !== undefined) ? (patch.locked ? '鎖定' : '解除鎖定') + '「' + n.label + '」' : '手動修改「' + n.label + '」';
  }
  withLock_(() => saveDraft_(project, plan.round, draft, user, note));
  return pagePlanning(project);
}

function pageRestoreDraft(project, v) {
  const user = ownerOnly_();
  const plan = pagePlanning(project);
  const row = rows_('drafts').find(r => r.project === project && (Number(r.round) || 1) === plan.round && Number(r.v) === Number(v));
  if (!row) throw new Error('找不到這個版本。');
  withLock_(() => saveDraft_(project, plan.round, parse_(row.specJson, {}), user, '退回草稿 v' + v));
  return pagePlanning(project);
}

/** 節點留言（所有人都可以） */
function pageAddComment(project, node, text) {
  const user = pageUser_();
  const t = String(text || '').trim();
  if (!t) throw new Error('意見不能是空的。');
  const p = rows_('projects').find(r => r.project === project);
  if (!p) throw new Error('找不到這個專案。');
  withLock_(() => appendObj_('comments', { id: 'c' + Date.now().toString(36).slice(-5) + Math.random().toString(36).slice(2, 4), ts: now_(), project: project,
    round: Number(p.round) || 1, node: node || '', author: user, text: t, status: 'open', note: '' }));
  const dn = node ? ((parse_(p.draftJson, {}).nodes || []).find(x => x.id === node) || {}).label || node : '';
  log_(project, user, 'comment', '對' + (dn ? '「' + dn + '」' : '整體規格') + '提出意見：' + t.slice(0, 40));
  return pagePlanning(project);
}

function pageSetCommentStatus(project, id, status) {
  ownerOnly_();
  if (['open', 'resolved', 'ignored'].indexOf(status) < 0) throw new Error('狀態不正確。');
  withLock_(() => updateRow_('comments', 'id', id, { status: status, note: status === 'open' ? '' : status === 'ignored' ? '管理者決定不採納' : '管理者標為已處理' }));
  return pagePlanning(project);
}

/** ⑤ 共用看過程：整個專案的時間軸 */
function pageTimeline(project) {
  pageUser_();
  const items = rows_('activity').filter(r => r.project === project).map(r => ({ ts: r.ts, kind: r.kind, actor: r.actor, text: r.text }));
  rows_('events').filter(r => r.project === project).forEach(r => items.push({ ts: r.ts, kind: 'dev', actor: r.actor, text: r.node + '：' + (r.summary || r.status), status: r.status }));
  rows_('questions').filter(r => r.project === project && r.status === 'answered').forEach(r => items.push({ ts: r.answeredAt, kind: 'answer', actor: r.answeredBy, text: '回覆 AI 的提問「' + r.question + '」：' + r.answer }));
  rows_('specs').filter(r => r.project === project && !items.some(i => i.kind === 'publish' && i.text.indexOf('v' + r.version + '（') >= 0))
    .forEach(r => items.push({ ts: r.createdAt, kind: 'publish', actor: r.createdBy, text: '發布規格 v' + r.version }));
  items.sort((a, b) => String(a.ts).localeCompare(String(b.ts)));
  return { items: items };
}

function log_(project, actor, kind, text) {
  try { appendObj_('activity', { ts: now_(), project: project, actor: actor, kind: kind, text: text }); } catch (err) {}
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
