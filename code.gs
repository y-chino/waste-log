/**
 * 廃棄物報告アプリ - Backend Logic (v1.0.7 - Optimized)
 */

var APP_TITLE = "廃棄物報告";

var SHEETS = {
  CATEGORY: "Master_Category",
  LOCATION: "Master_Location",
  USER_SETTING: "UserSetting",
  DATA: "WastesData",
  CONFIG: "Config",
  USER_MASTER: "Master_User"
};

/**
 * Webアプリケーションのエントリポイント
 */
function doGet() {
  var template = HtmlService.createTemplateFromFile("index");
  var userEmail = Session.getActiveUser().getEmail() || "anonymous@mipox.co.jp";
  
  // payloadは空で返し、クライアントサイドの initApp から非同期で getInitialDataForApp を呼ぶことで
  // 画面の枠組みを先に表示させる（体感速度向上）
  template.initialPayload = "null";
  template.userEmail = userEmail;
  
  return template.evaluate()
    .setTitle(APP_TITLE)
    .addMetaTag("viewport", "width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * 組織の全ユーザーを同期する
 */
function syncOrganizationUsers() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.USER_MASTER);
  
  if (typeof AdminDirectory === 'undefined') {
    var errorMsg = "エラー: Admin SDK API が有効化されていません。";
    Logger.log(errorMsg);
    return errorMsg;
  }

  if (!sheet) {
    sheet = ss.insertSheet(SHEETS.USER_MASTER);
    sheet.appendRow(["メールアドレス", "名前"]);
    sheet.getRange("1:1").setFontWeight("bold").setBackground("#f3f3f3");
    sheet.setFrozenRows(1);
  }

  var users = [];
  var pageToken;
  try {
    do {
      var response = AdminDirectory.Users.list({
        customer: 'my_customer',
        maxResults: 500,
        pageToken: pageToken,
        orderBy: 'email',
        viewType: 'domain_public' 
      });
      if (response.users && response.users.length > 0) {
        response.users.forEach(function(user) {
          users.push([user.primaryEmail.toLowerCase(), user.name.fullName]);
        });
      }
      pageToken = response.nextPageToken;
    } while (pageToken);

    if (users.length > 0) {
      var lastRow = sheet.getLastRow();
      if (lastRow > 1) { sheet.getRange(2, 1, lastRow - 1, 2).clearContent(); }
      sheet.getRange(2, 1, users.length, 2).setValues(users);
      SpreadsheetApp.flush(); 
      return "同期完了: " + users.length + " 名を登録しました。";
    }
    return "ユーザーが見つかりませんでした。";
  } catch (e) {
    return "エラー: " + e.toString();
  }
}

/**
 * 名簿から名前を取得
 */
function getOrRegisterUserName_(email) {
  if (!email) return null;
  var lowEmail = email.toLowerCase();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.USER_MASTER);
  if (!sheet) { syncOrganizationUsers(); sheet = ss.getSheetByName(SHEETS.USER_MASTER); }
  
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase() === lowEmail) { return data[i][1]; }
  }
  return null;
}

/**
 * 初期データ一括取得（パラメータなし・Session からメール取得）
 */
function getInitialAppData() {
  var userEmail = Session.getActiveUser().getEmail()
               || Session.getEffectiveUser().getEmail()
               || "anonymous@mipox.co.jp";

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var userName = getOrRegisterUserName_(userEmail);
  if (!userName) { userName = userEmail.split('@')[0]; }

  var catData = ss.getSheetByName(SHEETS.CATEGORY).getDataRange().getValues();
  var categories = catData.slice(1).map(function(r) {
    return { id: String(r[0]).trim(), name: String(r[1]).trim(), detail: String(r[2] || "").trim(), targetBase: String(r[3] || "").trim() };
  });

  var locData = ss.getSheetByName(SHEETS.LOCATION).getDataRange().getValues();
  var locations = locData.slice(1).map(function(r) {
    return { id: String(r[0]).trim(), name: String(r[1]).trim(), dept: String(r[2]).trim(), base: String(r[3]).trim() };
  });

  var userData = ss.getSheetByName(SHEETS.USER_SETTING).getDataRange().getValues();
  var userSetting = null;
  var userExists = false;
  for (var k = 1; k < userData.length; k++) {
    if (userData[k][0] === userEmail) {
      userSetting = { base: userData[k][1], roomId: String(userData[k][2]) };
      userExists = true;
      break;
    }
  }

  var configSheet = ss.getSheetByName(SHEETS.CONFIG);
  var feedbackUrl = "", summaryUrl = "";
  if (configSheet) {
    configSheet.getDataRange().getValues().forEach(function(row) {
      var key = String(row[0]).trim().toUpperCase();
      if (key === "FEEDBACK") feedbackUrl = String(row[1]).trim();
      if (key === "SUMMARY") summaryUrl = String(row[1]).trim();
    });
  }

  var registeredData = {};
  if (userExists) {
    var myLoc = locations.find(function(l) { return l.id === userSetting.roomId; });
    if (myLoc) {
      var targetRoomIds = locations.filter(function(l) {
        return l.dept === myLoc.dept && l.base === userSetting.base;
      }).map(function(l) { return l.id; });

      var now = new Date();
      for (var i = 0; i < 3; i++) {
        var d = new Date(now.getFullYear(), now.getMonth() - i, 1);
        var monthStr = d.getFullYear() + "年" + ("0" + (d.getMonth() + 1)).slice(-2) + "月";
        Object.assign(registeredData, readMonthData_(ss, monthStr, targetRoomIds));
      }
    }
  }

  return {
    locations: locations, categories: categories, userSetting: userSetting,
    userExists: userExists, registeredData: registeredData,
    feedbackUrl: feedbackUrl, summaryUrl: summaryUrl,
    userName: userName, serverTime: Date.now()
  };
}

/**
 * 過去データ取得（textFinder で対象月の行だけ読む）
 */
function getPastWasteData(year, month) {
  var userEmail = Session.getActiveUser().getEmail()
               || Session.getEffectiveUser().getEmail()
               || "anonymous@mipox.co.jp";

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var userData = ss.getSheetByName(SHEETS.USER_SETTING).getDataRange().getValues();
  var myRoomId = "";
  for (var k = 1; k < userData.length; k++) {
    if (userData[k][0] === userEmail) { myRoomId = String(userData[k][2]); break; }
  }
  if (!myRoomId) return {};

  var locData = ss.getSheetByName(SHEETS.LOCATION).getDataRange().getValues();
  var myLoc = locData.find(function(r) { return String(r[0]) === myRoomId; });
  if (!myLoc) return {};

  var targetRoomIds = locData.slice(1).filter(function(r) {
    return String(r[2]) === myLoc[2] && String(r[3]) === myLoc[3];
  }).map(function(r) { return String(r[0]); });

  var monthStr = year + "年" + ("0" + month).slice(-2) + "月";
  return readMonthData_(ss, monthStr, targetRoomIds);
}

/**
 * textFinder で指定月の行だけ読んでデータを集計する共通処理
 */
function readMonthData_(ss, monthStr, targetRoomIds) {
  var sheet = ss.getSheetByName(SHEETS.DATA);
  var finder = sheet.getRange("K:K").createTextFinder(monthStr);
  var matches = finder.findAll();
  if (matches.length === 0) return {};

  var startRow = matches[0].getRow();
  var lastRow  = matches[matches.length - 1].getRow();
  var data = sheet.getRange(startRow, 1, lastRow - startRow + 1, 15).getValues();

  var results = {};
  data.forEach(function(row) {
    var roomId = String(row[6]);
    if (targetRoomIds.indexOf(roomId) === -1) return;
    var rawDate = new Date(row[4]);
    var dateStr = Utilities.formatDate(rawDate, "JST", "yyyy-MM-dd");
    var catId   = String(row[1]);
    var catName = String(row[2]);
    var val     = parseFloat(row[3]);
    var user    = String(row[12]);
    var time    = (row[13] instanceof Date) ? Utilities.formatDate(row[13], "JST", "yyyy/MM/dd HH:mm:ss") : String(row[13]);
    if (!results[dateStr]) results[dateStr] = {};
    if (!results[dateStr][roomId]) results[dateStr][roomId] = { total: 0, cats: {}, logs: [] };
    if (results[dateStr][roomId].cats[catId] === undefined) results[dateStr][roomId].cats[catId] = val;
    results[dateStr][roomId].logs.push({ catId: catId, catName: catName, val: val, user: user, time: time });
  });

  for (var d in results) {
    for (var r in results[d]) {
      var total = 0;
      for (var c in results[d][r].cats) total += results[d][r].cats[c];
      results[d][r].total = total;
    }
  }
  return results;
}

function registerUserSetting(base, roomId) {
  var userEmail = Session.getActiveUser().getEmail() || "anonymous@mipox.co.jp";
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.USER_SETTING);
  var data = sheet.getDataRange().getValues();
  var foundRow = -1;
  for (var i = 1; i < data.length; i++) { if (data[i][0] === userEmail) { foundRow = i + 1; break; } }
  if (foundRow > 0) { sheet.getRange(foundRow, 2, 1, 2).setValues([[base, roomId]]); } else { sheet.appendRow([userEmail, base, roomId]); }
  return { success: true };
}

function executeSaveReport(formData) {
  var userEmail = Session.getActiveUser().getEmail() || "anonymous@mipox.co.jp";
  var userName = getOrRegisterUserName_(userEmail);
  var userDisplayName = userName || (userEmail ? userEmail.split('@')[0] : "不明ユーザー");
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.DATA);
  var locSheet = ss.getSheetByName(SHEETS.LOCATION);
  var locData = locSheet.getDataRange().getValues();
  var roomName = "", deptName = "";
  for (var j = 1; j < locData.length; j++) { if (String(locData[j][0]) === String(formData.roomId)) { roomName = locData[j][1]; deptName = locData[j][2]; break; } }
  var now = new Date();
  var formattedNow = Utilities.formatDate(now, "JST", "yyyy/MM/dd HH:mm:ss");
  var dateNumStr = formData.date.replace(/-/g, "");
  var lastRow = sheet.getLastRow();
  var dataValues = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, 15).getValues() : [];
  var startRow = 2;
  for (var i = 0; i < formData.items.length; i++) {
    var item = formData.items[i];
    var inputVal = parseFloat(item.value);
    if (isNaN(inputVal)) continue;
    var matchedIdx = -1;
    for (var k = dataValues.length - 1; k >= 0; k--) {
      if (String(dataValues[k][11]).replace(/,/g, '') === dateNumStr && 
          String(dataValues[k][1]).trim() === String(item.categoryId).trim() && 
          String(dataValues[k][6]).trim() === String(formData.roomId).trim()) {
        matchedIdx = startRow + k; break;
      }
    }
    if (matchedIdx > 0) {
      sheet.getRange(matchedIdx, 4).setValue(inputVal); 
      sheet.getRange(matchedIdx, 13).setValue(userDisplayName); 
      sheet.getRange(matchedIdx, 14).setValue(formattedNow); 
    } else {
      var yearStr = new Date(formData.date).getFullYear() + "年";
      var monthStr = yearStr + ("0" + (new Date(formData.date).getMonth() + 1)).slice(-2) + "月";
      sheet.appendRow([lastRow + 1 + i, item.categoryId, item.categoryName, inputVal, formData.date, formData.baseName, formData.roomId, roomName, deptName, yearStr, monthStr, dateNumStr, userDisplayName, formattedNow, formData.baseName + "_" + roomName + "_" + item.categoryId + "_" + dateNumStr]);
    }
  }
  return { success: true, user: userDisplayName, time: formattedNow };
}