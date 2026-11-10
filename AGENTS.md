# 開發守則（給 AI Agent）

這個專案的需求來自團隊共創平台。規格已定稿，你負責實作，並把進度回報到平台，讓不寫程式的同事也看得到。

## 開工前

1. 執行 `node flow.mjs pull`，取得最新規格
2. 執行 `node flow.mjs answers`，讀取同事對先前問題的回覆
3. 閱讀 `.flow/SPEC.md`（給人看的規格書）與 `.flow/spec.json`（節點、輸入輸出、驗收標準）

## 開發規則

- **規格是合約。** 流程結構、節點的輸入與輸出以 `.flow/spec.json` 為準。覺得規格有問題時，用 `flow ask` 提出，不要自行改變結構
- **一次做一個節點。** 開始時回報 `working`，完成並自測後回報 `done`
- **用驗收標準自我測試。** 每個節點的 `acceptance` 就是測試案例。回報 `done` 時，summary 要寫明通過幾條、哪條沒過
- **實作細節以程式為準。** 若發現某檔案被人手動改過，先保留現有寫法，有衝突就提問，不要直接覆蓋
- 共用的程式（設定、API 呼叫、錯誤處理）放在共用檔案，不屬於任何節點，不用回報
- 程式語言依 `spec.json` 的 `target` 決定；沒有填寫時，依專案現有的檔案判斷

## 回報

```bash
# 開始做
node flow.mjs report --node classify --status working --summary "開始實作分類"

# 完成（--file 是程式檔，--upload 是給同事看的截圖或報告）
node flow.mjs report --node classify --status done \
  --summary "驗收 3 條全數通過；測試 20 題，正確 18 題" \
  --file src/classify.js --upload reports/classify-test.md

# 卡住
node flow.mjs report --node rules_kb --status blocked --summary "缺少園務規章的 Drive 資料夾權限"
```

summary 會顯示給不寫程式的同事看，請用白話寫結果，不要寫技術細節。

## 提問

只有**需要業務判斷**的問題才提問，技術問題請自己決定，並寫在 commit 訊息裡。

```bash
node flow.mjs ask --node route "家長同時問學費和孩子情緒，要分到哪一類？"
```

提問後先做其他不受影響的節點，下次開工時再用 `flow answers` 讀取回覆。
