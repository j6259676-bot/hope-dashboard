// ============================================================
// The Hope 線上分部 — Dashboard API
// 貼到現有 GAS 專案，在 Script Properties 設定以下四個值：
//   YT_API_KEY      = YouTube Data API v3 金鑰
//   YT_CHANNEL_ID   = The Hope YouTube 頻道 ID（UC 開頭）
//   DASHBOARD_TOKEN = 自訂管理密碼
//   MEMBER_SHEET_ID = 主日 Hope Nation 資料庫 Sheet ID
// ============================================================

var DASH_CFG = (function () {
  var p = PropertiesService.getScriptProperties();
  return {
    ytApiKey:    p.getProperty('YT_API_KEY'),
    channelId:   p.getProperty('YT_CHANNEL_ID'),
    token:       p.getProperty('DASHBOARD_TOKEN'),
    sheetId:     p.getProperty('MEMBER_SHEET_ID'),

    // Sheet tab 名稱
    tabPeople:   '個人紀錄',
    tabMeetings: '聚會紀錄',
    tabYoutube:  'YouTube數據',

    // 個人紀錄 欄位索引（0-based）
    colDate:     0,  // A：日期
    colSession:  1,  // B：場次（9:30 / 11:30 / 2:00）
    colName:     3,  // D：姓名
    colCountry:  6,  // G：國家
    colCity:     7,  // H：城市

    // 聚會紀錄 欄位索引（0-based）
    colMeetDate:  0,  // A：日期
    colMeetSess:  1,  // B：場次
    colMeetCount: 2,  // C：出席人數
  };
})();

// ── Token 驗證 ─────────────────────────────────────────────
function dashVerifyToken_(token) {
  return token && token === DASH_CFG.token;
}

// ── 統一回應格式 ───────────────────────────────────────────
function dashRespond_(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// doGet 路由
// 若你原本已有 doGet，把 if 區塊合併進去即可：
//
//   if (action === 'geo')     return dashRespond_(getGeoData_());
//   if (action === 'growth')  return dashRespond_(getGrowthData_());
//   if (action === 'youtube') return dashRespond_(getYoutubeData_());
// ============================================================
function doGet(e) {
  var action = e.parameter.action;
  var token  = e.parameter.token;

  if (!dashVerifyToken_(token)) {
    return dashRespond_({ error: 'Unauthorized', code: 401 });
  }

  if (action === 'geo')     return dashRespond_(getGeoData_());
  if (action === 'growth')  return dashRespond_(getGrowthData_());
  if (action === 'youtube') return dashRespond_(getYoutubeData_());

  return dashRespond_({ error: 'Unknown action', code: 400 });
}

// ============================================================
// /geo — 地理分佈
// 來源：個人紀錄 tab
// 回傳：{ countries, cities }，各自依人數由多到少排序
// ============================================================
function getGeoData_() {
  var ss   = SpreadsheetApp.openById(DASH_CFG.sheetId);
  var tab  = ss.getSheetByName(DASH_CFG.tabPeople);
  var rows = tab.getDataRange().getValues();

  var countryMap = {};
  var cityMap    = {};
  // 用姓名+國家去重，避免同一人多次出現被重複計算
  var seen = {};

  for (var i = 1; i < rows.length; i++) {
    var row     = rows[i];
    var name    = (row[DASH_CFG.colName]    || '').toString().trim();
    var country = (row[DASH_CFG.colCountry] || '').toString().trim();
    var city    = (row[DASH_CFG.colCity]    || '').toString().trim();
    if (!country) continue;

    var dedup = name + '||' + country;
    if (seen[dedup]) continue;
    seen[dedup] = true;

    countryMap[country] = (countryMap[country] || 0) + 1;
    if (city) {
      var cityKey = city + '||' + country;
      if (!cityMap[cityKey]) cityMap[cityKey] = { city: city, country: country, count: 0 };
      cityMap[cityKey].count++;
    }
  }

  var countries = Object.keys(countryMap)
    .map(function(c) { return { country: c, count: countryMap[c] }; })
    .sort(function(a, b) { return b.count - a.count; });

  var cities = Object.keys(cityMap)
    .map(function(k) { return cityMap[k]; })
    .sort(function(a, b) { return b.count - a.count; });

  return { countries: countries, cities: cities };
}

// ============================================================
// /growth — 成長指標
// 出席趨勢：來自 聚會紀錄 tab
// 新朋友趨勢：來自 個人紀錄 tab（每週新出現的人數）
// 回傳：{ attendanceTrend, newPeopleTrend, alert }
// ============================================================
function getGrowthData_() {
  var ss = SpreadsheetApp.openById(DASH_CFG.sheetId);
  var twelveAgo = new Date(Date.now() - 84 * 24 * 60 * 60 * 1000);

  // ── 出席趨勢（聚會紀錄）────────────────────────────────
  var meetTab  = ss.getSheetByName(DASH_CFG.tabMeetings);
  var meetRows = meetTab.getDataRange().getValues();
  var attendMap = {};

  for (var i = 1; i < meetRows.length; i++) {
    var r    = meetRows[i];
    var d    = new Date(r[DASH_CFG.colMeetDate]);
    if (isNaN(d) || d < twelveAgo) continue;

    var wk      = isoWeek_(d);
    var session = (r[DASH_CFG.colMeetSess]  || '').toString().trim();
    var count   = Number(r[DASH_CFG.colMeetCount]) || 0;

    if (!attendMap[wk]) attendMap[wk] = { week: wk, total: 0, s930: 0, s1130: 0, s200: 0 };
    attendMap[wk].total += count;
    if (session === '9:30')  attendMap[wk].s930  += count;
    if (session === '11:30') attendMap[wk].s1130 += count;
    if (session === '2:00')  attendMap[wk].s200  += count;
  }

  var attendanceTrend = Object.keys(attendMap).sort().map(function(w) { return attendMap[w]; });

  // ── 連續 3 週下滑預警 ────────────────────────────────────
  var alert = false;
  if (attendanceTrend.length >= 3) {
    var last3 = attendanceTrend.slice(-3);
    alert = last3[0].total > last3[1].total && last3[1].total > last3[2].total;
  }

  // ── 新朋友趨勢（個人紀錄）────────────────────────────────
  var peopleTab  = ss.getSheetByName(DASH_CFG.tabPeople);
  var peopleRows = peopleTab.getDataRange().getValues();
  var newMap = {};

  for (var j = 1; j < peopleRows.length; j++) {
    var pr = peopleRows[j];
    var pd = new Date(pr[DASH_CFG.colDate]);
    if (isNaN(pd) || pd < twelveAgo) continue;
    var wk2 = isoWeek_(pd);
    newMap[wk2] = (newMap[wk2] || 0) + 1;
  }

  var newPeopleTrend = Object.keys(newMap).sort().map(function(w) {
    return { week: w, count: newMap[w] };
  });

  return {
    attendanceTrend: attendanceTrend,
    newPeopleTrend:  newPeopleTrend,
    alert:           alert,
  };
}

// ============================================================
// /youtube — 回傳 YouTube 數據 tab 的已同步資料
// ============================================================
function getYoutubeData_() {
  var ss  = SpreadsheetApp.openById(DASH_CFG.sheetId);
  var tab = ss.getSheetByName(DASH_CFG.tabYoutube);
  if (!tab) return { rows: [], note: '尚無資料，請先手動執行一次 syncYoutubeData()' };

  var rows   = tab.getDataRange().getValues();
  var header = rows[0];
  var data   = [];

  for (var i = 1; i < rows.length; i++) {
    var obj = {};
    header.forEach(function(h, idx) { obj[h] = rows[i][idx]; });
    data.push(obj);
  }

  data.sort(function(a, b) { return new Date(b['日期']) - new Date(a['日期']); });
  return { rows: data.slice(0, 36) };
}

// ============================================================
// syncYoutubeData() — 每週日 22:00 自動觸發
// ⚠️  需在 GAS 編輯器啟用「YouTube Data API v3」進階服務
//     （左側選單 Services → YouTube Data API v3 → Add）
// ============================================================
function syncYoutubeData() {
  var apiKey = DASH_CFG.ytApiKey;
  if (!apiKey) {
    Logger.log('缺少 YT_API_KEY');
    return;
  }

  var now     = new Date();
  var sunday  = getSundayTW_(now);
  var dateStr = Utilities.formatDate(sunday, 'Asia/Taipei', 'yyyy-MM-dd');
  Logger.log('同步週日：' + dateStr);

  var dayStartUTC = new Date(dateStr + 'T00:00:00+08:00');
  var dayEndUTC   = new Date(dateStr + 'T23:59:59+08:00');

  // ── Step 1：用 YouTube 進階服務（OAuth）取得不公開直播 ──
  var broadcasts;
  try {
    broadcasts = YouTube.LiveBroadcasts.list('id,snippet,status', {
      broadcastStatus: 'completed',
      broadcastType:   'all',
      maxResults:      50
    });
  } catch (e) {
    Logger.log('YouTube 進階服務錯誤：' + e.message);
    Logger.log('→ 請在 GAS 編輯器左側 Services 加入 YouTube Data API v3');
    return;
  }

  Logger.log('取得已完成直播：' + (broadcasts.items ? broadcasts.items.length : 0) + ' 筆');
  if (!broadcasts.items || broadcasts.items.length === 0) {
    Logger.log('無已完成直播');
    return;
  }

  // ── Step 2：過濾當週日的主日直播 ──
  var validItems = broadcasts.items.filter(function(item) {
    var t         = item.snippet.title;
    var startTime = item.snippet.actualStartTime;
    if (!startTime) return false;

    var startDate  = new Date(startTime);
    var isInRange  = startDate >= dayStartUTC && startDate <= dayEndUTC;
    var isSunday   = (t.indexOf('主日') > -1 || t.indexOf('HOPE') > -1);
    var isNotQ2Q   = t.indexOf('Q2Q') === -1 && t.indexOf('Q to Q') === -1;

    Logger.log((isInRange && isSunday && isNotQ2Q ? '✓' : '✗') + ' ' + t + ' | ' + startTime);
    return isInRange && isSunday && isNotQ2Q;
  });

  Logger.log('有效主日直播：' + validItems.length + ' 筆');
  if (validItems.length === 0) {
    Logger.log('該週日無符合的主日直播（確認標題含「主日」或「HOPE」）');
    return;
  }

  // ── Step 3：用 API Key 取得統計數據（videos.list 可抓不公開影片 ID）──
  var videoIds = validItems.map(function(i) { return i.id; }).join(',');
  var videoUrl = 'https://www.googleapis.com/youtube/v3/videos'
    + '?part=snippet,statistics,liveStreamingDetails'
    + '&id=' + videoIds
    + '&key=' + apiKey;

  var videoData = JSON.parse(UrlFetchApp.fetch(videoUrl, { muteHttpExceptions: true }).getContentText());
  if (!videoData.items || videoData.items.length === 0) {
    Logger.log('videos.list 無回傳（影片 ID：' + videoIds + '）');
    return;
  }

  // ── Step 4：寫入 Sheet ──
  var ss  = SpreadsheetApp.openById(DASH_CFG.sheetId);
  var tab = ss.getSheetByName(DASH_CFG.tabYoutube);
  if (!tab) {
    tab = ss.insertSheet(DASH_CFG.tabYoutube);
    var header = ['日期','場次','影片ID','標題','觀看次數','按讚數','留言數','最高同時在線','同步時間'];
    tab.appendRow(header);
    tab.getRange(1, 1, 1, header.length).setFontWeight('bold');
  }

  var syncTime = new Date();
  var written  = 0;

  videoData.items.forEach(function(video) {
    if (tab.createTextFinder(video.id).findAll().length > 0) {
      Logger.log('已存在，跳過：' + video.snippet.title);
      return;
    }
    var stats   = video.statistics || {};
    var live    = video.liveStreamingDetails || {};
    var session = detectYtSession_(video.snippet.title);

    tab.appendRow([
      dateStr,
      session,
      video.id,
      video.snippet.title,
      parseInt(stats.viewCount)    || 0,
      parseInt(stats.likeCount)    || 0,
      parseInt(stats.commentCount) || 0,
      live.concurrentViewers || '',
      syncTime
    ]);
    Logger.log('寫入：' + video.snippet.title);
    written++;
  });

  Logger.log('syncYoutubeData 完成，新增 ' + written + ' 筆');
}

// ── 工具函式 ───────────────────────────────────────────────

function detectYtSession_(title) {
  if (title.indexOf('9:30') > -1)  return '9:30AM';
  if (title.indexOf('11:30') > -1) return '11:30AM';
  if (title.indexOf('2PM') > -1 || title.indexOf('2:00PM') > -1) return '2PM';
  if (title.indexOf('4PM') > -1)   return '4PM';
  return '其他';
}

function getSundayTW_(date) {
  var tw  = new Date(date.toLocaleString('en-US', { timeZone: 'Asia/Taipei' }));
  var day = tw.getDay();
  return new Date(tw.getFullYear(), tw.getMonth(), tw.getDate() - day);
}

function isoWeek_(date) {
  var d    = new Date(date);
  var day  = d.getDay() || 7;
  d.setDate(d.getDate() + 4 - day);
  var year = d.getFullYear();
  var week = Math.ceil((((d - new Date(year, 0, 1)) / 86400000) + 1) / 7);
  return year + '-W' + (week < 10 ? '0' + week : week);
}
