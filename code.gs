/**
 * 廃棄物報告アプリ - Backend Logic (v1.0.5 - Auto-Registration Version)
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
  
  var initialData = getInitialDataForApp_(userEmail);
  
  template.initialPayload = JSON.stringify(initialData);
  template.userEmail = userEmail;
  
  return template.evaluate()
    .setTitle(APP_TITLE)
    .addMetaTag("viewport", "width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * 組織の全ユーザーを同期する（初期構築や一括更新用）
 */
function syncOrganizationUsers() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.USER_MASTER);
  
  // Admin SDK API が有効化されているかチェック
  if (typeof AdminDirectory === 'undefined') {
    var errorMsg = "エラー: Admin SDK API が有効化されていません。GASエディタの左メニュー「サービス」から「Admin SDK API」を追加してください。";
    Logger.log(errorMsg);
    return errorMsg;
  }

  // シートがなければ作成
  if (!sheet) {
    sheet = ss.insertSheet(SHEETS.USER_MASTER);
    sheet.appendRow(["メールアドレス", "名前"]);
    sheet.getRange("1:1").setFontWeight("bold").setBackground("#f3f3f3");
    sheet.setFrozenRows(1);
  }

  var users = [];
  var pageToken;
  
  try {
    Logger.log("組織ユーザーの取得を開始します...");
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

    Logger.log("取得件数: " + users.length + " 件");

    if (users.length > 0) {
      var lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        sheet.getRange(2, 1, lastRow - 1, 2).clearContent();
      }
      sheet.getRange(2, 1, users.length, 2).setValues(users);
      SpreadsheetApp.flush(); 
      Logger.log("シートへの書き込みが完了しました。");
      return "同期完了: " + users.length + " 名を登録しました。";
    } else {
      return "警告: 組織内にユーザーが見つかりませんでした。";
    }
  } catch (e) {
    Logger.log("エラー発生: " + e.toString());
    return "エラー: " + e.toString();
  }
}

/**
 * 名簿から名前を取得。見つからない場合は組織から取得してシートに自動追加する（新ユーザー対応）
 */
function getOrRegisterUserName_(email) {
  if (!email) return null;
  var lowEmail = email.toLowerCase();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.USER_MASTER);
  
  if (!sheet) {
    var syncResult = syncOrganizationUsers();
    if (syncResult.indexOf("エラー") === 0) return null;
    sheet = ss.getSheetByName(SHEETS.USER_MASTER);
  }
  
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase() === lowEmail) {
      return data[i][1];
    }
  }

  // Admin SDK API が有効な場合のみ個別取得を試みる
  if (typeof AdminDirectory !== 'undefined') {
    try {
      var user = AdminDirectory.Users.get(lowEmail);
      if (user && user.name) {
        var newName = user.name.fullName;
        sheet.appendRow([lowEmail, newName]);
        return newName;
      }
    } catch (e) {
      console.warn("New user auto-registration failed: " + e.toString());
    }
  }
  return null;
}

/**
 * アプリ起動時の初期データを取得
 */
function getInitialDataForApp_(userEmail) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var userName = getOrRegisterUserName_(userEmail);
  if (!userName && userEmail) {
    userName = userEmail.split('@')[0];
  }

  var catSheet = ss.getSheetByName(SHEETS.CATEGORY);
  var catData = catSheet.getDataRange().getValues();
  var categories = [];
  for (var i = 1; i < catData.length; i++) {
    categories.push({
      id: String(catData[i][0]),
      name: String(catData[i][1]),
      detail: String(catData[i][2] || ""),
      targetBase: String(catData[i][3] || "")
    });
  }

  var locSheet = ss.getSheetByName(SHEETS.LOCATION);
  var locData = locSheet.getDataRange().getValues();
  var locations = [];
  for (var j = 1; j < locData.length; j++) {
    locations.push({
      id: String(locData[j][0]),
      name: String(locData[j][1]),
      dept: String(locData[j][2]),
      base: String(locData[j][3])
    });
  }

  var userSettingSheet = ss.getSheetByName(SHEETS.USER_SETTING);
  var userData = userSettingSheet.getDataRange().getValues();
  var userSetting = null;
  var userExists = false;
  for (var k = 1; k < userData.length; k++) {
    if (userData[k][0] === userEmail) {
      userSetting = {
        base: userData[k][1],
        roomId: String(userData[k][2])
      };
      userExists = true;
      break;
    }
  }

  var configSheet = ss.getSheetByName(SHEETS.CONFIG);
  var configData = configSheet.getDataRange().getValues();
  var feedbackUrl = "";
  for (var m = 1; m < configData.length; m++) {
    if (configData[m][0] === "FEEDBACK_FORM_URL") {
      feedbackUrl = configData[m][1];
      break;
    }
  }

  var wasteSheet = ss.getSheetByName(SHEETS.DATA);
  var wasteData = wasteSheet.getDataRange().getValues();
  var registeredData = {};
  var thresholdDate = new Date();
  thresholdDate.setMonth(thresholdDate.getMonth() - 3); 
  
  for (var n = wasteData.length - 1; n >= 1; n--) { 
    var row = wasteData[n];
    var rawDate = new Date(row[4]);
    if (rawDate < thresholdDate) continue;
    var catId = String(row[1]);
    var catName = String(row[2]);
    var val = parseFloat(row[3]);
    var dateStr = Utilities.formatDate(rawDate, "JST", "yyyy-MM-dd");
    var roomId = String(row[6]);
    var user = String(row[12]);
    var time = String(row[13]); 
    if (!registeredData[dateStr]) registeredData[dateStr] = {};
    if (!registeredData[dateStr][roomId]) {
      registeredData[dateStr][roomId] = { total: 0, cats: {}, logs: [] };
    }
    if (registeredData[dateStr][roomId].cats[catId] === undefined) {
      registeredData[dateStr][roomId].cats[catId] = val;
    }
    registeredData[dateStr][roomId].logs.push({
      catId: catId,
      catName: catName,
      val: val,
      user: user,
      time: time
    });
  }

  for (var d in registeredData) {
    for (var r in registeredData[d]) {
      var total = 0;
      for (var c in registeredData[d][r].cats) {
        total += registeredData[d][r].cats[c];
      }
      registeredData[d][r].total = total;
    }
  }

  return {
    locations: locations,
    categories: categories,
    userSetting: userSetting,
    userExists: userExists,
    registeredData: registeredData,
    feedbackUrl: feedbackUrl,
    userName: userName,
    serverTime: Date.now()
  };
}

/**
 * ユーザーの初期設定を保存
 */
function registerUserSetting(base, roomId) {
  var userEmail = Session.getActiveUser().getEmail() || "anonymous@mipox.co.jp";
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.USER_SETTING);
  var data = sheet.getDataRange().getValues();
  var foundRow = -1;
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === userEmail) {
      foundRow = i + 1;
      break;
    }
  }
  if (foundRow > 0) {
    sheet.getRange(foundRow, 2, 1, 2).setValues([[base, roomId]]);
  } else {
    sheet.appendRow([userEmail, base, roomId]);
  }
  return { success: true };
}

/**
 * 報告データを保存
 */
function executeSaveReport(formData) {
  var userEmail = Session.getActiveUser().getEmail() || "anonymous@mipox.co.jp";
  var userName = getOrRegisterUserName_(userEmail);
  var userDisplayName = userName || (userEmail ? userEmail.split('@')[0] : "不明ユーザー");

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.DATA);
  var locSheet = ss.getSheetByName(SHEETS.LOCATION);
  
  var locData = locSheet.getDataRange().getValues();
  var roomName = "";
  var deptName = "";
  for (var j = 1; j < locData.length; j++) {
    if (String(locData[j][0]) === String(formData.roomId)) {
      roomName = locData[j][1];
      deptName = locData[j][2];
      break;
    }
  }

  var now = new Date();
  var formattedNow = Utilities.formatDate(now, "JST", "yyyy/MM/dd HH:mm:ss");
  var dateNumStr = formData.date.replace(/-/g, "");
  var lastRow = sheet.getLastRow();
  var dataValues = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, 15).getValues() : [];
  var startRow = 2;

  for (var i = 0; i < formData.items.length; i++) {
    var item = formData.items[i];
    var inputVal = parseFloat(item.value);
    if (isNaN(inputVal) || inputVal <= 0) continue;
    var matchedIdx = -1;
    for (var k = dataValues.length - 1; k >= 0; k--) {
      if (String(dataValues[k][11]).replace(/,/g, '') === dateNumStr && 
          String(dataValues[k][1]).trim() === String(item.categoryId).trim() && 
          String(dataValues[k][6]).trim() === String(formData.roomId).trim()) {
        matchedIdx = startRow + k; 
        break;
      }
    }
    if (matchedIdx > 0) {
      sheet.getRange(matchedIdx, 4).setValue(inputVal); 
      sheet.getRange(matchedIdx, 13).setValue(userDisplayName); 
      sheet.getRange(matchedIdx, 14).setValue(formattedNow); 
    } else {
      var yearStr = new Date(formData.date).getFullYear() + "年";
      var monthStr = yearStr + ("0" + (new Date(formData.date).getMonth() + 1)).slice(-2) + "月";
      sheet.appendRow([
        lastRow + 1 + i, 
        item.categoryId, 
        item.categoryName, 
        inputVal, 
        formData.date, 
        formData.baseName, 
        formData.roomId, 
        roomName, 
        deptName, 
        yearStr, 
        monthStr,
        dateNumStr,
        userDisplayName, 
        formattedNow,
        formData.baseName + "_" + roomName + "_" + item.categoryId + "_" + dateNumStr 
      ]);
    }
  }
  return { success: true, user: userDisplayName, time: formattedNow };
}