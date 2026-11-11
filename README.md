# co_work

團隊共創平台的試跑專案：同事參與定義需求，AI Agent 依規格開發，進度回報到平台。

| 位置 | 內容 |
|---|---|
| `flow.mjs` | 平台與 Agent 之間的指令工具 |
| `AGENTS.md` | 給 AI Agent 的開發守則 |
| `specs/` | 規格原稿（發布到平台前的版本） |
| `gas/Code.gs` | 平台後端（Google Apps Script） |
| `gas/Page.html` | 使用者共創頁面：開專案、收集需求、整理、規格、開發追蹤，可多輪迭代 |
| `.flow/` | 從平台拉下來的定稿規格（`pull` 產生） |

## 第一次設定（Windows PowerShell）

```powershell
git clone https://github.com/cda67730/co_work
cd co_work
node flow.mjs init --endpoint <GAS 網址> --project parent-qa --token <FLOW_TOKEN> --repo https://github.com/cda67730/co_work
node flow.mjs push-spec specs/parent-qa.json
node flow.mjs pull
node flow.mjs status
```

token 存在 `.flow/token`，已列入 `.gitignore`，不會被 commit。

## 使用者共創頁面（給同事）

流程：① 收集 → ② 整理 → ③ 畫圖及規格 → ④ 開發 → ⑤ 共用看過程，畫面上只有這一條流程。

- ① 收集：同事填「主題」＋「遇到的問題／希望變成什麼」（擇一），可以＋1、留言
- ② 整理：管理者按「AI 整理」產生主題，再用對話框跟 AI 來回修改；也可以手動編輯、填寫「決定」
- ③ 畫圖及規格：AI 畫出初版流程圖與驗收標準；大家點步驟提意見；管理者用對話框請 AI 把意見整理進去（只改有變動的步驟，鎖定的步驟不動，每次都存成新草稿版本）
- ④ 開發：定稿發布後，AI Agent 用 `flow` 拉規格、回報進度
- ⑤ 共用看過程：整個專案的時間軸
- 開發中可以「開始下一輪需求」，草稿以上線版本為底，發布時只重做有變動的步驟

管理者是這份 Sheet 的擁有者；AI 功能只有管理者能使用。

### AI 設定

Apps Script「專案設定」→「指令碼屬性」新增：

| 屬性 | 說明 |
|---|---|
| `GEMINI_API_KEY` | 必填，到 Google AI Studio 申請 |
| `GEMINI_MODEL_FAST` | 選填，整理需求用，預設 `gemini-2.5-flash` |
| `GEMINI_MODEL_SPEC` | 選填，畫圖及規格用，預設 `gemini-2.5-pro` |

### 部署

1. Apps Script：`Code.gs`、HTML 檔 `Page` 用 repo 裡的版本覆蓋
2. 原本給 Agent 的部署：編輯 → 新版本
3. 給同事的部署（執行身分：存取網頁應用程式的使用者；存取：任何擁有 Google 帳戶的使用者）：編輯 → 新版本
4. Sheet 共用給同事（編輯者）

新的分頁（drafts、chats、comments、activity）第一次使用時自動建立。
