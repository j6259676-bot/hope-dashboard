# Dashboard API — 設定說明

## 1. 把程式碼貼到 GAS

開啟你的線上分部 GAS 專案（就是現有的那個），新增一個檔案叫 `dashboard-api.gs`，把 `dashboard-api.gs` 的內容貼進去。

> 如果你現有的 `doGet` 裡已有其他路由，把這段合併進去：
> ```javascript
> if (action === 'geo')     return dashRespond_(getGeoData_());
> if (action === 'growth')  return dashRespond_(getGrowthData_());
> if (action === 'youtube') return dashRespond_(getYoutubeData_());
> ```

## 2. 設定 Script Properties

GAS → 專案設定 → 指令碼屬性，新增以下四組：

| 屬性名稱 | 值 |
|---|---|
| `YT_API_KEY` | `AIzaSyDXKof8bYjsVMYcD_e-gIAD6fSccIQMLYE` |
| `YT_CHANNEL_ID` | The Hope YouTube 頻道 ID（UC 開頭，從頻道網址取得） |
| `DASHBOARD_TOKEN` | 自訂一組管理密碼（例如 `hope2026admin`） |
| `MEMBER_SHEET_ID` | 會眾 Google Sheet 的 ID（從網址取得） |

### 取得頻道 ID 的方法
1. 進入 The Hope YouTube 頻道頁面
2. 網址格式：`https://www.youtube.com/@TheHope/about`
3. 在頁面按右鍵 → 查看原始碼 → 搜尋 `channelId`
4. 或直接用 API：`https://www.googleapis.com/youtube/v3/channels?part=id&forHandle=TheHope&key=你的API_KEY`

## 3. 調整欄位索引

在 `dashboard-api.gs` 最上面的 `DASH_CFG` 區塊，確認以下索引符合你的 Sheet 結構：

**會眾紀錄 tab（colXxx）**
- `colCity`：城市欄（0-based，A=0, B=1, ...）
- `colCountry`：國家欄
- `colStatus`：狀態欄（值為 `active` 或 `inactive`）
- `colJoinDate`：加入日期欄
- `colBaptism`：受洗日期欄（空白=未受洗）

**聚會紀錄 tab（colMeetXxx）**
- `colMeetDate`：日期欄
- `colMeetSession`：場次欄（值必須是 `9:30AM` / `11:30AM` / `2PM`）
- `colMeetCount`：出席人數欄
- `colMeetNew`：新朋友人數欄

## 4. 設定自動觸發器

GAS → 觸發器 → 新增觸發器：

- 函式：`syncYoutubeData`
- 觸發類型：時間驅動
- 類型：週計時器
- 星期：週日
- 時間：晚上 10:00–11:00

## 5. 測試

在 GAS 編輯器直接跑 `syncYoutubeData()`，看 Logs 有沒有正確抓到本週直播。

API 端點測試（替換你的 GAS 部署網址）：
```
https://script.google.com/macros/s/你的ID/exec?action=youtube&token=你的DASHBOARD_TOKEN
https://script.google.com/macros/s/你的ID/exec?action=geo&token=你的DASHBOARD_TOKEN
https://script.google.com/macros/s/你的ID/exec?action=growth&token=你的DASHBOARD_TOKEN
```

## ⚠️ 注意事項

**YouTube 同時在線人數**：Data API v3 在直播結束後無法取得歷史峰值同時在線數。
目前方案：記錄 `viewCount`（累積觀看）、`likeCount`、`commentCount`。
若未來需要真正的峰值同時在線，需要在直播**進行中**每 5 分鐘 call API 記錄，或申請 YouTube Analytics API。

**Sheet tab 名稱**：`tabMembers`、`tabMeetings`、`tabYoutube` 三個名稱要和你的 Sheet 完全對應（包含空格）。
