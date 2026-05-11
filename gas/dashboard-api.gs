// ============================================================
// The Hope 線上分部 — Dashboard API  v2
//
// Script Properties（GAS 專案設定 > 指令碼屬性）：
//   YT_API_KEY       = YouTube Data API v3 金鑰
//   YT_CHANNEL_ID    = The Hope YouTube 頻道 ID（UC 開頭）
//   DASHBOARD_TOKEN  = 自訂管理密碼
//   MEMBER_SHEET_ID  = 主日 Hope Nation 資料庫 Sheet ID
//   NOTIFY_EMAIL     = 錯誤通知 Email（選填）
//
// 進階服務（服務 > YouTube Analytics API）：
//   需啟用 YouTube Analytics API（updateYoutubeAnalytics 使用）
//
// 觸發器建議：
//   - syncYoutubeData       → 每週日 18:30–19:00
//   - updateYoutubeAnalytics → 每週二 20:00–21:00
// ============================================================

var DASH_CFG = (function () {
  var p = PropertiesService.getScriptProperties();
  return {
    ytApiKey:      p.getProperty('YT_API_KEY'),
    channelId:     p.getProperty('YT_CHANNEL_ID'),
    token:         p.getProperty('DASHBOARD_TOKEN'),
    sheetId:       p.getProperty('MEMBER_SHEET_ID'),
    notifyEmail:   p.getProperty('NOTIFY_EMAIL') || '',

    tabPeople:     '個人紀錄',
    tabMeetings:   '聚會紀錄',
    tabYoutube:    'YouTube數據',
    tabYoutubeGeo: 'YouTube地區',

    // 個人紀錄欄位（0-based）
    colDate:    0,
    colSession: 1,
    colName:    3,
    colCountry: 6,
    colCity:    7,

    // 聚會紀錄欄位（0-based）
    colMeetDate:  0,
    colMeetSess:  1,
    colMeetCount: 2,

    // YouTube數據欄位（0-based，v2 格式）
    ytDate:        0,   // 日期
    ytSession:     1,   // 場次
    ytVideoId:     2,   // 影片ID
    ytTitle:       3,   // 標題
    ytViews:       4,   // 觀看次數
    ytLikes:       5,   // 按讚數
    ytComments:    6,   // 留言數
    ytAvgConcur:   7,   // 平均同時在線  ← Analytics
    ytPeakConcur:  8,   // 峰值同時在線  ← Analytics
    ytAvgDuration: 9,   // 平均觀看時長(秒) ← Analytics
    ytAvgPercent:  10,  // 平均觀看%     ← Analytics
    ytTotalMins:   11,  // 總觀看分鐘    ← Analytics
    ytChat:        12,  // 聊天訊息      ← Analytics
    ytSyncTime:    13,  // 同步時間
  };
})();

// ── YouTube數據 v2 欄位標頭
var YT_HEADER_V2 = [
  '日期','場次','影片ID','標題','觀看次數','按讚數','留言數',
  '平均同時在線','峰值同時在線','平均觀看時長(秒)','平均觀看%','總觀看分鐘','聊天訊息',
  '同步時間'
];

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

// ── 錯誤通知 ───────────────────────────────────────────────
function notifyError_(msg) {
  Logger.log('[ERROR] ' + msg);
  if (!DASH_CFG.notifyEmail) return;
  try {
    MailApp.sendEmail(DASH_CFG.notifyEmail, '[The Hope Dashboard] 同步錯誤', msg);
  } catch (e) {
    Logger.log('Email notify failed: ' + e.message);
  }
}

// ============================================================
// doGet 路由
// 若原本已有 doGet，把 if 區塊合併進去：
//   if (action === 'geo')          return dashRespond_(getGeoData_());
//   if (action === 'growth')       return dashRespond_(getGrowthData_());
//   if (action === 'youtube')      return dashRespond_(getYoutubeData_());
//   if (action === 'youtube-geo')  return dashRespond_(getYoutubeGeoData_());
// ============================================================
function doGet(e) {
  var action = e.parameter.action;
  var token  = e.parameter.token;

  if (!dashVerifyToken_(token)) {
    return dashRespond_({ error: 'Unauthorized', code: 401 });
  }

  if (action === 'geo')         return dashRespond_(getGeoData_());
  if (action === 'growth')      return dashRespond_(getGrowthData_());
  if (action === 'youtube')     return dashRespond_(getYoutubeData_());
  if (action === 'youtube-geo') return dashRespond_(getYoutubeGeoData_());

  return dashRespond_({ error: 'Unknown action', code: 400 });
}

// ============================================================
// /geo — 地理分佈（會眾資料）
// ============================================================
function getGeoData_() {
  var ss   = SpreadsheetApp.openById(DASH_CFG.sheetId);
  var tab  = ss.getSheetByName(DASH_CFG.tabPeople);
  var rows = tab.getDataRange().getValues();

  var countryMap = {};
  var cityMap    = {};
  var seen       = {};

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

  return {
    countries: Object.keys(countryMap)
      .map(function (c) { return { country: c, count: countryMap[c] }; })
      .sort(function (a, b) { return b.count - a.count; }),
    cities: Object.keys(cityMap)
      .map(function (k) { return cityMap[k]; })
      .sort(function (a, b) { return b.count - a.count; }),
  };
}

// ============================================================
// /growth — 成長指標
// ============================================================
function getGrowthData_() {
  var ss = SpreadsheetApp.openById(DASH_CFG.sheetId);
  var twelveAgo = new Date(Date.now() - 84 * 24 * 60 * 60 * 1000);

  var meetTab  = ss.getSheetByName(DASH_CFG.tabMeetings);
  var meetRows = meetTab.getDataRange().getValues();
  var attendMap = {};

  for (var i = 1; i < meetRows.length; i++) {
    var r = meetRows[i];
    var d = new Date(r[DASH_CFG.colMeetDate]);
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

  var attendanceTrend = Object.keys(attendMap).sort().map(function (w) { return attendMap[w]; });

  var alert = false;
  if (attendanceTrend.length >= 3) {
    var last3 = attendanceTrend.slice(-3);
    alert = last3[0].total > last3[1].total && last3[1].total > last3[2].total;
  }

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

  return {
    attendanceTrend: attendanceTrend,
    newPeopleTrend:  Object.keys(newMap).sort().map(function (w) { return { week: w, count: newMap[w] }; }),
    alert:           alert,
  };
}

// ============================================================
// /youtube — 回傳 YouTube數據 tab，含 Analytics 聚合 stats
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
    header.forEach(function (h, idx) { obj[h] = rows[i][idx]; });
    data.push(obj);
  }

  data.sort(function (a, b) { return new Date(b['日期']) - new Date(a['日期']); });
  var recent = data.slice(0, 36);

  // 計算有 Analytics 資料的平均值
  var statsRows = recent.filter(function (r) {
    return r['平均同時在線'] !== '' && r['平均同時在線'] != null && !isNaN(Number(r['平均同時在線']));
  });

  var stats = null;
  if (statsRows.length > 0) {
    var avg = function (arr, key) {
      var vals = arr.filter(function (r) {
        return r[key] !== '' && r[key] != null && !isNaN(Number(r[key]));
      });
      if (!vals.length) return null;
      return Math.round(vals.reduce(function (s, r) { return s + Number(r[key]); }, 0) / vals.length);
    };
    stats = {
      avgConcurrent:  avg(statsRows, '平均同時在線'),
      peakConcurrent: avg(statsRows, '峰值同時在線'),
      avgDuration:    avg(statsRows, '平均觀看時長(秒)'),
      avgPercent:     avg(statsRows, '平均觀看%'),
      totalMinutes:   Math.round(statsRows.reduce(function (s, r) {
        var v = Number(r['總觀看分鐘']); return s + (isNaN(v) ? 0 : v);
      }, 0)),
      videoCount: statsRows.length,
    };
  }

  return { rows: recent, stats: stats };
}

// ============================================================
// /youtube-geo — 回傳 YouTube 觀眾地區分佈（ISO 國家碼）
// ============================================================
function getYoutubeGeoData_() {
  var ss  = SpreadsheetApp.openById(DASH_CFG.sheetId);
  var tab = ss.getSheetByName(DASH_CFG.tabYoutubeGeo);
  if (!tab) return { countries: [], note: '尚無觀眾地區數據，請先執行 updateYoutubeAnalytics()' };

  var rows = tab.getDataRange().getValues();
  if (rows.length <= 1) return { countries: [] };

  var countryMap = {};
  for (var i = 1; i < rows.length; i++) {
    var country  = (rows[i][3] || '').toString().trim(); // 國家碼
    var views    = Number(rows[i][4]) || 0;
    var duration = Number(rows[i][5]) || 0;
    if (!country) continue;

    if (!countryMap[country]) {
      countryMap[country] = { country: country, views: 0, totalDuration: 0, videoCount: 0 };
    }
    countryMap[country].views         += views;
    countryMap[country].totalDuration += duration;
    countryMap[country].videoCount++;
  }

  var countries = Object.keys(countryMap).map(function (c) {
    var d = countryMap[c];
    return {
      country:     c,
      views:       d.views,
      avgDuration: d.videoCount ? Math.round(d.totalDuration / d.videoCount) : 0,
    };
  }).sort(function (a, b) { return b.views - a.views; });

  return { countries: countries.slice(0, 20) };
}

// ============================================================
// syncYoutubeData() — 每週日 18:30–19:00 觸發
// 存基本數據（觀看/按讚/留言），Analytics 欄位留空給週二補
// ============================================================
function syncYoutubeData() {
  var apiKey    = DASH_CFG.ytApiKey;
  var channelId = DASH_CFG.channelId;
  if (!apiKey || !channelId) {
    notifyError_('缺少 YT_API_KEY 或 YT_CHANNEL_ID');
    return;
  }

  var now     = new Date();
  var sunday  = getSundayTW_(now);
  var dateStr = Utilities.formatDate(sunday, 'Asia/Taipei', 'yyyy-MM-dd');
  Logger.log('同步週日：' + dateStr);

  // 台北時間週日 06:00–19:30（影片公開時段）
  var searchStart = new Date(dateStr + 'T06:00:00+08:00');
  var searchEnd   = new Date(dateStr + 'T19:30:00+08:00');

  var searchUrl = 'https://www.googleapis.com/youtube/v3/search'
    + '?part=id,snippet'
    + '&channelId=' + encodeURIComponent(channelId)
    + '&type=video'
    + '&eventType=completed'
    + '&publishedAfter='  + searchStart.toISOString()
    + '&publishedBefore=' + searchEnd.toISOString()
    + '&maxResults=20'
    + '&order=date'
    + '&key=' + apiKey;

  var searchData = JSON.parse(
    UrlFetchApp.fetch(searchUrl, { muteHttpExceptions: true }).getContentText()
  );
  Logger.log('搜尋結果：' + (searchData.items ? searchData.items.length + ' 筆' : 'error'));

  if (!searchData.items || searchData.items.length === 0) {
    notifyError_('syncYoutubeData: 無結果（確認觸發器在週日 19:00 前執行，影片公開中）');
    return;
  }

  // 過濾主日影片
  var validItems = searchData.items.filter(function (item) {
    var t = item.snippet.title;
    return (t.indexOf('主日') > -1 || t.indexOf('HOPE') > -1)
           && t.indexOf('Q2Q') === -1
           && t.indexOf('Q to Q') === -1;
  });

  Logger.log('有效主日影片：' + validItems.length + ' 筆');
  if (validItems.length === 0) {
    Logger.log('標題過濾後無影片');
    return;
  }

  // 取得統計數據
  var videoIds = validItems.map(function (i) { return i.id.videoId; }).join(',');
  var videoUrl = 'https://www.googleapis.com/youtube/v3/videos'
    + '?part=snippet,statistics'
    + '&id=' + videoIds
    + '&key=' + apiKey;

  var videoData = JSON.parse(
    UrlFetchApp.fetch(videoUrl, { muteHttpExceptions: true }).getContentText()
  );
  if (!videoData.items || videoData.items.length === 0) {
    notifyError_('syncYoutubeData: videos.list 無回傳（ID: ' + videoIds + '）');
    return;
  }

  // 開啟 Sheet，migrate header，建 dedup Set
  var ss  = SpreadsheetApp.openById(DASH_CFG.sheetId);
  var tab = ss.getSheetByName(DASH_CFG.tabYoutube);
  if (!tab) {
    tab = ss.insertSheet(DASH_CFG.tabYoutube);
    tab.appendRow(YT_HEADER_V2);
    tab.getRange(1, 1, 1, YT_HEADER_V2.length).setFontWeight('bold');
  } else {
    migrateYtHeader_(tab);
  }

  // 讀現有 videoId 建 Set，避免重複
  var existingRows = tab.getDataRange().getValues();
  var existingIds  = {};
  for (var i = 1; i < existingRows.length; i++) {
    var vid = (existingRows[i][DASH_CFG.ytVideoId] || '').toString().trim();
    if (vid) existingIds[vid] = true;
  }

  var syncTime = new Date();
  var written  = 0;

  videoData.items.forEach(function (video) {
    if (existingIds[video.id]) {
      Logger.log('已存在，跳過：' + video.snippet.title);
      return;
    }
    var stats   = video.statistics || {};
    var session = detectYtSession_(video.snippet.title);

    // Analytics 欄位留空，由 updateYoutubeAnalytics() 週二補填
    tab.appendRow([
      dateStr,
      session,
      video.id,
      video.snippet.title,
      parseInt(stats.viewCount)    || 0,
      parseInt(stats.likeCount)    || 0,
      parseInt(stats.commentCount) || 0,
      '', '', '', '', '', '',    // 7 Analytics 欄位（留空）
      syncTime,
    ]);
    Logger.log('寫入：' + video.snippet.title);
    written++;
  });

  Logger.log('syncYoutubeData 完成，新增 ' + written + ' 筆');
}

// ============================================================
// updateYoutubeAnalytics() — 每週二 20:00 觸發
// 回補 Analytics 欄位（averageConcurrentViewers, peakConcurrentViewers,
// averageViewDuration, averageViewPercentage, estimatedMinutesWatched,
// chatMessages）並同步 YouTube地區 tab
//
// 需啟用進階服務：YouTube Analytics API
// ============================================================
function updateYoutubeAnalytics() {
  var ss  = SpreadsheetApp.openById(DASH_CFG.sheetId);
  var tab = ss.getSheetByName(DASH_CFG.tabYoutube);
  if (!tab || tab.getLastColumn() < 14) {
    Logger.log('YouTube數據 tab 不存在或尚未 migrate，請先執行 syncYoutubeData()');
    return;
  }

  var data    = tab.getDataRange().getValues();
  var updated = 0;

  for (var i = 1; i < data.length; i++) {
    var row     = data[i];
    var videoId = (row[DASH_CFG.ytVideoId] || '').toString().trim();
    if (!videoId) continue;

    // 已有資料就跳過
    var hasAnalytics = row[DASH_CFG.ytAvgConcur] !== '' && row[DASH_CFG.ytAvgConcur] != null;
    if (hasAnalytics) continue;

    var dateVal = row[DASH_CFG.ytDate];
    var dateStr = (dateVal instanceof Date)
      ? Utilities.formatDate(dateVal, 'Asia/Taipei', 'yyyy-MM-dd')
      : String(dateVal).substring(0, 10);

    var session = (row[DASH_CFG.ytSession] || '').toString().trim();
    var analytics = getVideoAnalytics_(videoId, dateStr);

    if (!analytics) {
      Logger.log('Analytics 尚未就緒，跳過：' + videoId);
      continue;
    }

    // 更新 cols 8–13（1-indexed = ytAvgConcur+1 … ytChat+1）
    tab.getRange(i + 1, DASH_CFG.ytAvgConcur + 1, 1, 6).setValues([[
      analytics.avgConcurrent,
      analytics.peakConcurrent,
      analytics.avgDuration,
      analytics.avgPercent,
      analytics.totalMinutes,
      analytics.chatMessages,
    ]]);

    syncYoutubeGeoData_(videoId, dateStr, session);
    updated++;
    Logger.log('Analytics 更新：' + videoId + ' (' + dateStr + ')');
    Utilities.sleep(300); // rate limit
  }

  Logger.log('updateYoutubeAnalytics 完成，更新 ' + updated + ' 筆');
}

// ============================================================
// getVideoAnalytics_() — 呼叫 YouTube Analytics API
// 回傳單支影片的參與指標，Analytics 未就緒時回傳 null
// ============================================================
function getVideoAnalytics_(videoId, dateStr) {
  try {
    // 7 天窗口：捕捉首週主要參與數據
    var endDate = new Date(dateStr);
    endDate.setDate(endDate.getDate() + 6);
    var endStr = Utilities.formatDate(endDate, 'Asia/Taipei', 'yyyy-MM-dd');

    var report = YouTubeAnalytics.Reports.query({
      ids:       'channel==MINE',
      startDate: dateStr,
      endDate:   endStr,
      metrics:   'averageConcurrentViewers,peakConcurrentViewers,averageViewDuration,averageViewPercentage,estimatedMinutesWatched,chatMessages',
      filters:   'video==' + videoId,
    });

    if (!report || !report.rows || !report.rows[0]) {
      Logger.log('Analytics: no rows for ' + videoId + ' (may not be ready)');
      return null;
    }

    var r = report.rows[0];
    var safe = function (val, round) {
      if (val == null || val === '') return '';
      var n = Number(val);
      return isNaN(n) ? '' : (round ? Math.round(n) : Math.round(n * 10) / 10);
    };

    return {
      avgConcurrent:  safe(r[0], true),
      peakConcurrent: safe(r[1], true),
      avgDuration:    safe(r[2], true),
      avgPercent:     safe(r[3], false),
      totalMinutes:   safe(r[4], true),
      chatMessages:   safe(r[5], true),
    };
  } catch (e) {
    Logger.log('getVideoAnalytics_ error (' + videoId + '): ' + e.message);
    return null;
  }
}

// ============================================================
// syncYoutubeGeoData_() — 寫入 YouTube地區 tab
// 依 country dimension 分解觀看數據
// ============================================================
function syncYoutubeGeoData_(videoId, dateStr, session) {
  try {
    var endDate = new Date(dateStr);
    endDate.setDate(endDate.getDate() + 6);
    var endStr = Utilities.formatDate(endDate, 'Asia/Taipei', 'yyyy-MM-dd');

    var report = YouTubeAnalytics.Reports.query({
      ids:        'channel==MINE',
      startDate:  dateStr,
      endDate:    endStr,
      dimensions: 'country',
      metrics:    'views,averageViewDuration',
      filters:    'video==' + videoId,
      sort:       '-views',
      maxResults: 30,
    });

    if (!report || !report.rows || report.rows.length === 0) {
      Logger.log('Geo: no rows for ' + videoId);
      return;
    }

    var ss  = SpreadsheetApp.openById(DASH_CFG.sheetId);
    var tab = ss.getSheetByName(DASH_CFG.tabYoutubeGeo);
    if (!tab) {
      tab = ss.insertSheet(DASH_CFG.tabYoutubeGeo);
      tab.appendRow(['同步日期', '影片ID', '場次', '國家碼', '觀看次數', '平均觀看時長(秒)']);
      tab.getRange(1, 1, 1, 6).setFontWeight('bold');
    }

    // 刪除此 videoId 的舊資料（先從後往前刪，避免 index 偏移）
    var existing = tab.getDataRange().getValues();
    for (var i = existing.length - 1; i >= 1; i--) {
      if ((existing[i][1] || '').toString() === videoId) {
        tab.deleteRow(i + 1);
      }
    }

    // 寫入新地區資料
    report.rows.forEach(function (r) {
      tab.appendRow([
        dateStr,
        videoId,
        session,
        r[0],                            // 國家碼（ISO 3166-1 alpha-2）
        Math.round(Number(r[1])) || 0,   // 觀看次數
        Math.round(Number(r[2])) || 0,   // 平均觀看時長(秒)
      ]);
    });

    Logger.log('Geo synced: ' + videoId + ', ' + report.rows.length + ' 個國家');
  } catch (e) {
    Logger.log('syncYoutubeGeoData_ error (' + videoId + '): ' + e.message);
  }
}

// ============================================================
// migrateYtHeader_() — v1（9 欄）→ v2（14 欄）Sheet 遷移
// ============================================================
function migrateYtHeader_(tab) {
  var lastCol = tab.getLastColumn();
  if (lastCol >= 14) return; // 已是 v2
  if (lastCol < 9)  return; // 不預期的格式

  // 在 col 8（最高同時在線）之後插入 5 欄，同步時間移到 col 14
  tab.insertColumnsAfter(8, 5);

  // 重命名 col 8：最高同時在線 → 平均同時在線
  tab.getRange(1, 8).setValue('平均同時在線');

  // 設定新欄標頭 cols 9–13
  tab.getRange(1, 9, 1, 5)
    .setValues([['峰值同時在線', '平均觀看時長(秒)', '平均觀看%', '總觀看分鐘', '聊天訊息']])
    .setFontWeight('bold');

  Logger.log('YouTube數據 sheet migrated: v1(9) → v2(14)');
}

// ── 工具函式 ───────────────────────────────────────────────

function detectYtSession_(title) {
  if (title.indexOf('9:30') > -1)                              return '9:30AM';
  if (title.indexOf('11:30') > -1)                             return '11:30AM';
  if (title.indexOf('2PM') > -1 || title.indexOf('2:00') > -1) return '2PM';
  if (title.indexOf('4PM') > -1)                               return '4PM';
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
