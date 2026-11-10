# co_work

團隊共創平台的試跑專案：同事參與定義需求，AI Agent 依規格開發，進度回報到平台。

| 位置 | 內容 |
|---|---|
| `flow.mjs` | 平台與 Agent 之間的指令工具 |
| `AGENTS.md` | 給 AI Agent 的開發守則 |
| `specs/` | 規格原稿（發布到平台前的版本） |
| `gas/Code.gs` | 平台後端（Google Apps Script） |
| `gas/Page.html` | 同事看的開發追蹤頁面 |
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

## 開發追蹤頁面（給同事）

1. Apps Script 編輯器：用 `gas/Code.gs` 取代原本的程式碼；新增 HTML 檔案，命名為 `Page`，貼上 `gas/Page.html`
2. 「管理部署作業」→ 編輯原本的部署 → 版本選「新版本」→ 部署（Agent 用的網址不變）
3. 「新增部署作業」→ 網頁應用程式：
   - 執行身分：**存取網頁應用程式的使用者**
   - 誰可以存取：**任何擁有 Google 帳戶的使用者**（或你的組織）
   這個新網址就是給同事的連結
4. 把 Sheet 共用給同事（編輯者），他們才能讀取資料、回答提問

同事第一次開啟時會看到授權畫面，按「允許」即可（若顯示未經驗證，選「進階」→ 前往）。
在 Agent 用的「所有人」部署開啟頁面，會顯示請改用同事連結，不會洩漏資料。
