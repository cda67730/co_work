# co_work

團隊共創平台的試跑專案：同事參與定義需求，AI Agent 依規格開發，進度回報到平台。

| 位置 | 內容 |
|---|---|
| `flow.mjs` | 平台與 Agent 之間的指令工具 |
| `AGENTS.md` | 給 AI Agent 的開發守則 |
| `specs/` | 規格原稿（發布到平台前的版本） |
| `gas/Code.gs` | 平台後端（Google Apps Script） |
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
