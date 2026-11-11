#!/usr/bin/env node
/**
 * flow — 平台與 AI Agent 之間的橋樑（Node 18+，無外部套件）
 *
 *   flow init --endpoint <網址> --project <名稱> [--token T] [--repo https://github.com/你/專案]
 *   flow pull [--version N]                      拉取規格到 .flow/
 *   flow push-spec [spec.json]                   發布新版規格（平台介面完成前用）
 *   flow report --node ID --status S --summary "…" [--file 路徑]… [--upload 路徑]… [--link "名稱|網址"]…
 *   flow ask --node ID "問題"
 *   flow answers [--all]                         讀取同事的回覆
 *   flow status                                  各節點目前狀態
 *   flow install-hook                            安裝 git commit 自動回報
 *
 * token 讀取順序：環境變數 FLOW_TOKEN → .flow/token（已加入 .gitignore）
 */
import fs from 'node:fs';
import path from 'node:path';
import { execSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';

const SELF = fileURLToPath(import.meta.url);

const DIR = '.flow';
const STATUSES = ['todo', 'working', 'done', 'blocked', 'needs_review'];
const STATUS_ZH = { todo: '未開始', working: '進行中', done: '完成', blocked: '卡住', needs_review: '待確認' };

/* ---------- 參數 ---------- */
function parseArgs(argv) {
  const out = { _: [] };
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if (a.startsWith('--')) {
      const k = a.slice(2), v = argv[i + 1] && !argv[i + 1].startsWith('--') ? argv[++i] : true;
      if (out[k] === undefined) out[k] = v;
      else out[k] = [].concat(out[k], v);
    } else out._.push(a);
  }
  return out;
}
const list = v => (v === undefined ? [] : [].concat(v));
const die = msg => { console.error('✗ ' + msg); process.exit(1); };

/* ---------- 設定 ---------- */
const p = f => path.join(DIR, f);
const readJson = (f, d) => { try { return JSON.parse(fs.readFileSync(f, 'utf8')); } catch { return d; } };
const writeJson = (f, o) => { fs.mkdirSync(path.dirname(f), { recursive: true }); fs.writeFileSync(f, JSON.stringify(o, null, 2) + '\n'); };

function config() {
  const c = readJson(p('config.json'));
  if (!c) die('找不到 .flow/config.json，請先執行 flow init');
  const token = process.env.FLOW_TOKEN || (fs.existsSync(p('token')) ? fs.readFileSync(p('token'), 'utf8').trim() : '');
  if (!token) die('缺少 token：設定環境變數 FLOW_TOKEN，或寫入 .flow/token');
  return { ...c, token };
}

function actor() {
  if (process.env.FLOW_ACTOR) return process.env.FLOW_ACTOR;
  try { return 'agent:' + execSync('git config user.name', { stdio: ['ignore', 'pipe', 'ignore'] }).toString().trim(); }
  catch { return 'agent'; }
}

function gitHead() {
  try { return execSync('git rev-parse --short HEAD', { stdio: ['ignore', 'pipe', 'ignore'] }).toString().trim(); }
  catch { return ''; }
}

/* ---------- 呼叫平台 ---------- */
async function call(action, body = {}) {
  const c = config();
  const payload = { action, token: c.token, project: c.project, actor: actor(), ...body };
  let res;
  try {
    res = await fetch(c.endpoint, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' }, // GAS 用 text/plain 可避免 CORS 預檢
      body: JSON.stringify(payload),
      redirect: 'follow',
    });
  } catch (e) { die('連不到平台：' + e.message); }
  const text = await res.text();
  let data;
  try { data = JSON.parse(text); } catch { die('平台回應不是 JSON（檢查網址或部署權限）：\n' + text.slice(0, 300)); }
  if (!data.ok) die(`平台回應錯誤：${data.error}`);
  return data;
}

/* ---------- 規格書 ---------- */
const TYPE_ZH = { input: '輸入', llm: 'LLM', retrieve: '知識檢索', condition: '條件分支', tool: '工具／程式', output: '輸出' };

function renderSpecMd(spec) {
  const L = [];
  L.push(`# ${spec.title || spec.project || '流程規格'}`, '');
  L.push(`版本 v${spec.version ?? '?'}${spec.target ? ` · 目標：${spec.target}` : ''}`, '');
  if (spec.purpose) L.push('## 目的', '', spec.purpose, '');
  L.push('## 流程', '', '```mermaid', 'flowchart LR');
  for (const n of spec.nodes) {
    const lb = (n.label || n.id).replace(/"/g, "'");
    const shape = n.type === 'condition' ? `{"${lb}"}` : n.type === 'retrieve' ? `[("${lb}")]` : `["${lb}"]`;
    L.push(`  ${n.id}${shape}`);
  }
  for (const e of spec.edges || []) L.push(`  ${e.from} -->${e.label ? `|${e.label.replace(/\|/g, '/')}|` : ''} ${e.to}`);
  L.push('```', '');
  L.push('## 節點', '');
  for (const n of spec.nodes) {
    L.push(`### ${n.label || n.id}  \`${n.id}\``, '');
    L.push(`- 類型：${TYPE_ZH[n.type] || n.type}`);
    if (n.description) L.push(`- 說明：${n.description}`);
    if (n.inputs?.length) L.push(`- 輸入：${n.inputs.map(x => '`' + x + '`').join('、')}`);
    if (n.outputs?.length) L.push(`- 輸出：${n.outputs.map(x => '`' + x + '`').join('、')}`);
    for (const [k, v] of Object.entries(n.config || {})) if (v !== '' && v != null) L.push(`- ${k}：${v}`);
    const outs = (spec.edges || []).filter(e => e.from === n.id);
    if (outs.length) L.push(`- 下一步：${outs.map(e => (e.label ? `${e.label} → ` : '') + e.to).join('；')}`);
    if (n.acceptance?.length) {
      L.push('', '**驗收標準**', '');
      for (const a of n.acceptance) L.push(`- [ ] \`${a.id}\` ${a.text}`);
    }
    L.push('');
  }
  if (spec.decisions?.length) {
    L.push('## 已決定事項', '');
    for (const d of spec.decisions) L.push(`- ${d}`);
    L.push('');
  }
  return L.join('\n');
}

/* 輸入輸出接不上就警告（不擋） */
function lintSpec(spec) {
  const warn = [];
  const ids = new Set(spec.nodes.map(n => n.id));
  for (const e of spec.edges || []) {
    if (!ids.has(e.from) || !ids.has(e.to)) warn.push(`連線 ${e.from} → ${e.to} 指到不存在的節點`);
  }
  for (const n of spec.nodes) {
    if (!n.inputs?.length) continue;
    const anc = new Set(), stack = [n.id];
    while (stack.length) {
      const cur = stack.pop();
      for (const e of spec.edges || []) if (e.to === cur && !anc.has(e.from)) { anc.add(e.from); stack.push(e.from); }
    }
    const available = new Set(spec.nodes.filter(x => anc.has(x.id)).flatMap(u => u.outputs || []));
    for (const i of n.inputs) if (anc.size && !available.has(i)) warn.push(`${n.id} 需要 \`${i}\`，但前面的節點都沒有輸出它`);
    if (n.type !== 'input' && !n.acceptance?.length) warn.push(`${n.id} 沒有驗收標準`);
  }
  return warn;
}

/* ---------- 指令 ---------- */
const commands = {
  async init(a) {
    if (!a.endpoint || !a.project) die('用法：flow init --endpoint <網址> --project <名稱> [--token <token>]');
    writeJson(p('config.json'), { endpoint: a.endpoint, project: a.project, ...(a.repo ? { repoUrl: a.repo } : {}) });
    if (a.token) fs.writeFileSync(p('token'), String(a.token) + '\n', { mode: 0o600 });
    const gi = fs.existsSync('.gitignore') ? fs.readFileSync('.gitignore', 'utf8') : '';
    const add = ['.flow/token', '.flow/seen.json'].filter(x => !gi.split('\n').includes(x));
    if (add.length) fs.appendFileSync('.gitignore', (gi && !gi.endsWith('\n') ? '\n' : '') + add.join('\n') + '\n');
    const here = path.dirname(SELF);
    const tpl = path.join(here, 'AGENTS.template.md');
    if (!fs.existsSync('AGENTS.md') && fs.existsSync(tpl)) {
      fs.copyFileSync(tpl, 'AGENTS.md');
      console.log('✓ 已建立 AGENTS.md');
    } else if (fs.existsSync('AGENTS.md')) console.log('· AGENTS.md 已存在，未覆蓋（可參考 AGENTS.template.md 手動合併）');
    console.log(`✓ 已設定專案 ${a.project}`);
    if (!a.token && !process.env.FLOW_TOKEN) console.log('· 記得設定 FLOW_TOKEN，或執行 flow init 時加上 --token');
  },

  async pull(a) {
    const d = await call('spec', { version: a.version || 'latest' });
    const spec = d.spec;
    writeJson(p('spec.json'), spec);
    fs.writeFileSync(p('SPEC.md'), d.specMd || renderSpecMd(spec));
    const prev = readJson(p('lock.json'), {});
    writeJson(p('lock.json'), { version: d.version, pulledAt: new Date().toISOString() });
    console.log(`✓ 已拉取 ${d.title} v${d.version}（${spec.nodes.length} 個節點）`);
    if (prev.version && prev.version !== d.version) {
      const old = readJson(p(`history/v${prev.version}.json`));
      if (old) {
        const changed = diffNodes(old, spec);
        if (changed.length) console.log('  與 v' + prev.version + ' 相比有變動：' + changed.join('、'));
      }
    }
    writeJson(p(`history/v${d.version}.json`), spec);
    for (const w of lintSpec(spec)) console.log('  ⚠ ' + w);
  },

  async 'push-spec'(a) {
    const file = a._[0] || p('spec.json');
    const spec = readJson(file);
    if (!spec || !Array.isArray(spec.nodes)) die(`${file} 不是有效的規格（需要 nodes 陣列）`);
    const warns = lintSpec(spec);
    for (const w of warns) console.log('  ⚠ ' + w);
    const mdFile = a.md || path.join(path.dirname(file), 'SPEC.md');
    // 有 --md 就用你寫好的規格書；否則依規格自動產生（版本號由平台給，先以「下一版」顯示）
    const specMd = a.md ? fs.readFileSync(mdFile, 'utf8') : '';
    const d = await call('push_spec', { spec, specMd });
    console.log(`✓ 已發布 v${d.version}`);
    console.log('  請執行 flow pull 取得定稿版本');
  },

  async report(a) {
    if (!a.node || !a.status) die('用法：flow report --node ID --status ' + STATUSES.join('|') + ' --summary "…"');
    if (!STATUSES.includes(a.status)) die('status 只能是：' + STATUSES.join(' | '));
    const lock = readJson(p('lock.json'), {});
    const artifacts = [];
    const c = config(), head = gitHead();
    for (const f of list(a.file)) {
      if (!fs.existsSync(f)) die('找不到檔案 ' + f);
      const rel = path.relative(process.cwd(), path.resolve(f)).split(path.sep).join('/');
      artifacts.push({ type: 'file', path: rel, ...(c.repoUrl && head ? { url: `${c.repoUrl.replace(/\/$/, '')}/blob/${head}/${rel}` } : {}) });
      rememberFile(f, a.node);
    }
    for (const u of list(a.link)) {   // --link "名稱|網址" 或只給網址
      const i = String(u).lastIndexOf('|http');
      artifacts.push(i > 0 ? { type: 'link', name: String(u).slice(0, i).trim(), url: String(u).slice(i + 1).trim() } : { type: 'link', url: u });
    }
    for (const f of list(a.upload)) {
      if (!fs.existsSync(f)) die('找不到檔案 ' + f);
      const buf = fs.readFileSync(f);
      if (buf.length > 10 * 1024 * 1024) die(f + ' 超過 10MB');
      const up = await call('upload', { name: path.basename(f), mimeType: mime(f), data: buf.toString('base64') });
      artifacts.push({ type: 'upload', name: path.basename(f), url: up.url });
    }
    await call('report', {
      version: lock.version || '', node: a.node, status: a.status,
      summary: a.summary === true ? '' : a.summary || '', artifacts, commit: gitHead(),
    });
    console.log(`✓ ${a.node} → ${STATUS_ZH[a.status]}${artifacts.length ? `（${artifacts.length} 個產出物）` : ''}`);
  },

  async ask(a) {
    const q = a._.join(' ').trim();
    if (!q) die('用法：flow ask --node ID "問題"');
    const lock = readJson(p('lock.json'), {});
    const d = await call('ask', { version: lock.version || '', node: a.node === true ? '' : a.node || '', question: q });
    console.log(`✓ 已提問（${d.id}）。同事在平台回覆後，用 flow answers 讀取`);
  },

  async answers(a) {
    const seen = readJson(p('seen.json'), { ids: [] });
    const d = await call('answers', {});
    const fresh = d.answers.filter(x => a.all || !seen.ids.includes(x.id));
    if (!fresh.length) { console.log('· 沒有新的回覆'); return; }
    for (const x of fresh) {
      console.log(`\n[${x.node || '一般'}] ${x.question}`);
      console.log(`  → ${x.answer}${x.answeredBy ? `（${x.answeredBy}）` : ''}`);
    }
    writeJson(p('seen.json'), { ids: [...new Set([...seen.ids, ...fresh.map(x => x.id)])] });
  },

  async status() {
    const d = await call('status', {});
    const spec = readJson(p('spec.json'));
    const ids = spec ? spec.nodes.map(n => n.id) : Object.keys(d.nodes);
    for (const id of ids) {
      const e = d.nodes[id];
      const label = spec?.nodes.find(n => n.id === id)?.label || id;
      console.log(`${(STATUS_ZH[e?.status] || '未開始').padEnd(4, '　')}  ${label}（${id}）${e?.summary ? ' — ' + e.summary : ''}`);
    }
    if (d.openQuestions) console.log(`\n· 尚有 ${d.openQuestions} 個問題等待回覆`);
  },

  /* git post-commit 呼叫：依改動的檔案對應節點，自動回報 working */
  async 'hook-commit'() {
    const map = readJson(p('files.json'), {});
    let files = [];
    try { files = execSync('git diff-tree --no-commit-id --name-only -r HEAD').toString().trim().split('\n').filter(Boolean); }
    catch { return; }
    const msg = execSync('git log -1 --pretty=%s').toString().trim();
    const repo = readJson(p('config.json'), {}).repoUrl, head = gitHead();
    const art = f => ({ type: 'file', path: f, ...(repo && head ? { url: `${repo.replace(/\/$/, '')}/blob/${head}/${f}` } : {}) });
    const nodes = [...new Set(files.map(f => map[f]).filter(Boolean))];
    for (const node of nodes) {
      await call('report', {
        version: readJson(p('lock.json'), {}).version || '', node, status: 'working',
        summary: 'commit：' + msg, artifacts: files.filter(f => map[f] === node).map(art), commit: head,
      }).catch(() => {});
    }
  },

  async 'install-hook'() {
    if (!fs.existsSync('.git')) die('這裡不是 git 專案根目錄');
    const self = SELF.split(path.sep).join('/'); // Git for Windows 的 sh 也吃正斜線
    const hook = path.join('.git', 'hooks', 'post-commit');
    const line = `node "${self}" hook-commit >/dev/null 2>&1 &`;
    const cur = fs.existsSync(hook) ? fs.readFileSync(hook, 'utf8') : '#!/bin/sh\n';
    if (!cur.includes('hook-commit')) fs.writeFileSync(hook, cur.trimEnd() + '\n' + line + '\n', { mode: 0o755 });
    console.log('✓ 已安裝 post-commit 自動回報');
  },
};

function rememberFile(f, node) {
  const map = readJson(p('files.json'), {});
  map[path.normalize(f)] = node;
  writeJson(p('files.json'), map);
}

function diffNodes(a, b) {
  const key = n => JSON.stringify({ ...n, x: undefined, y: undefined });
  const am = Object.fromEntries(a.nodes.map(n => [n.id, key(n)]));
  const bm = Object.fromEntries(b.nodes.map(n => [n.id, key(n)]));
  const out = [];
  for (const id of Object.keys(bm)) if (!am[id]) out.push('+' + id); else if (am[id] !== bm[id]) out.push('~' + id);
  for (const id of Object.keys(am)) if (!bm[id]) out.push('-' + id);
  return out;
}

function mime(f) {
  return ({ '.png': 'image/png', '.jpg': 'image/jpeg', '.jpeg': 'image/jpeg', '.pdf': 'application/pdf', '.md': 'text/markdown',
    '.txt': 'text/plain', '.json': 'application/json', '.html': 'text/html', '.csv': 'text/csv' })[path.extname(f).toLowerCase()] || 'application/octet-stream';
}

/* ---------- main ---------- */
const [cmd, ...rest] = process.argv.slice(2);
if (!cmd || !commands[cmd]) {
  console.log(fs.readFileSync(new URL(import.meta.url), 'utf8').split('*/')[0].replace(/^#!.*\n\/\*\*?/, '').replace(/^ \* ?/gm, ''));
  process.exit(cmd ? 1 : 0);
}
await commands[cmd](parseArgs(rest));
