# 家庭記帳 LINE 機器人

把機器人加進家人的 LINE 群組，直接傳「午餐 120」或拍收據，AI 會自動記到 Google 試算表。

- **記帳**：`午餐 120`、`昨天全聯 560 衛生紙`、`加油 1200、停車 60`，或傳收據／發票照片
- **查詢**：`這個月花多少？`、`上個月餐飲多少`
- **刪除**：傳 `刪除`，會撤銷自己最近一次記的帳
- **說明**：傳 `說明`
- **分類**：餐飲、交通、日用品、醫療，以及其他（在 `gas/Config.gs` 修改）

架構：LINE → Google Apps Script（免費）→ Claude Haiku 5.5 解析 → Google 試算表。
每月費用約 NT$0～30，只有 Claude API 的用量要付費（一般家庭大約每月 NT$3）。

---

## 設定步驟（約 30 分鐘）

### 1. 取得 Claude API 金鑰

1. 到 [Claude Console](https://platform.claude.com/) 註冊並登入。
2. 在 **Billing** 儲值，最低金額就夠用很久。
3. 到 **API Keys** 建立金鑰，複製 `sk-ant-...` 開頭的字串。

### 2. 建立 LINE 機器人

1. 到 [LINE Developers](https://developers.line.biz/) 用 LINE 帳號登入。
2. 建立 Provider，再建立 **Messaging API** channel（會同時建立一個 LINE 官方帳號）。
3. 在 **Messaging API** 分頁：
   - 最下方 **Channel access token** 按 Issue，複製權杖。
   - 掃 QR code 把機器人加為好友。
4. 到 [LINE Official Account Manager](https://manager.line.biz/) 的「設定 → 回應設定」：
   - **聊天**：關閉
   - **自動回應訊息**：關閉
   - **Webhook**：開啟
   - 「帳號設定」裡把 **允許加入群組**打開
5. （選填）在 **Basic settings** 分頁最下方複製 **Your user ID**（U 開頭），第 3 步會用到。

### 3. 建立試算表和 Apps Script

1. 開一個新的 [Google 試算表](https://sheets.new)，命名為「家庭帳本」。
2. 選單 **擴充功能 → Apps Script**。
3. 在 Apps Script 編輯器裡，依照本專案 `gas/` 資料夾建立檔案，並貼上內容：
   - `Code.gs`（取代預設的內容）
   - 按「＋ → 指令碼」新增 `Config`、`Claude`、`Line`、`Sheet`，各自貼上對應的 `.gs` 內容
   - 齒輪「專案設定」勾選 **在編輯器中顯示 appsscript.json**，再把 `gas/appsscript.json` 的內容貼進去
4. 「專案設定 → 指令碼屬性」新增：

   | 屬性 | 值 |
   |---|---|
   | `LINE_CHANNEL_ACCESS_TOKEN` | 第 2 步的 Channel access token |
   | `ANTHROPIC_API_KEY` | 第 1 步的 Claude API 金鑰 |
   | `LINE_BOT_USER_ID` | （選填）第 2 步的 Your user ID |

5. 回到編輯器，上方選 `setup` 函式並按 **執行**，依指示授權。試算表會多出「帳本」工作表。

### 4. 部署並連上 LINE

1. Apps Script 右上角 **部署 → 新增部署作業**，類型選 **網頁應用程式**：
   - 執行身分：**我**
   - 誰可以存取：**所有人**
2. 部署後複製 **網頁應用程式網址**（`https://script.google.com/macros/s/.../exec`）。
3. 回到 LINE Developers 的 **Messaging API** 分頁：
   - **Webhook URL** 貼上網址，按 Update
   - 打開 **Use webhook**
   - 按 Verify 有時會顯示 302 錯誤，這是 Apps Script 的轉址造成的，**不影響實際使用**
4. 私訊機器人 `午餐 120` 測試。成功後把機器人邀進家庭群組。

> ⚠️ **之後修改程式碼**，要到「部署 → 管理部署作業 → 編輯 → 版本選『新版本』」重新部署，網址不會變。

---

## 注意事項

- **Webhook 網址請勿公開**：Apps Script 讀不到 LINE 的簽章標頭，無法驗證請求來源。有設定 `LINE_BOT_USER_ID` 可以擋掉一部分錯誤的請求。
- **在群組中**，機器人只回應記帳和查詢，不會回應閒聊。
- **記錄人**會自動抓 LINE 顯示名稱。
- 試算表可以直接修改。改分類、金額都沒問題，查詢會以試算表的內容為準。

## 開發

`test/run.js` 用模擬的 Apps Script 環境測試記帳、查詢、刪除流程，不需要任何金鑰：

```bash
node test/run.js
```

## 之後可以加的功能

- 電子發票載具自動匯入
- 每月自動推送月報
- 誰付錢、月底分帳結算
- 預算提醒
