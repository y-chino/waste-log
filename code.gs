/**
 * 廃棄物報告アプリ - Backend Logic (v1.0.5)
 */

var APP_TITLE = "廃棄物報告";

var SHEETS = {
  CATEGORY: "Master_Category",
  LOCATION: "Master_Location",
  USER_SETTING: "UserSetting",
  DATA: "WastesData",
  CONFIG: "Config" 
};

/**
 * Webアプリケーションのエントリポイント
 */
function doGet() {
  var template = HtmlService.createTemplateFromFile("index");
  var userEmail = Session.getActiveUser().getEmail() || "anonymous@mipox.co.jp";
  
  var initialData = getInitialDataForApp_(userEmail);
  
  var payload = "";
  try {
    var jsonStr = JSON.stringify(initialData);
    var blob = Utilities.newBlob(jsonStr, "application/json");
    payload = Utilities.base64Encode(blob.getBytes());
  } catch(e) {
    payload = "ERROR_ENCODE:" + e.message;
  }
  
  template.initialPayload = payload;
  template.userEmail = userEmail;
  
  return template.evaluate()
    .setTitle(APP_TITLE)
    .addMetaTag("viewport", "width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * アプリ起動時の初期データを取得（高速化のため一括取得）
 */
function getInitialDataForApp_(userEmail) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. カテゴリマスター取得
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

  // 2. 場所マスター取得
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

  // 3. ユーザー設定取得
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

  // 4. 設定情報取得 (Feedback URLなど)
  var configSheet = ss.getSheetByName(SHEETS.CONFIG);
  var configData = configSheet.getDataRange().getValues();
  var feedbackUrl = "";
  for (var m = 1; m < configData.length; m++) {
    if (configData[m][0] === "FEEDBACK_FORM_URL") {
      feedbackUrl = configData[m][1];
      break;
    }
  }

  // 5. 登録済みデータ取得（直近30日分程度に制限してパフォーマンス向上させることも可能だが、現在は全量）
  // 構造: { "yyyy-MM-dd": { "roomId": { total: X, cats: { "catId": val }, logs: [...] } } }
  var wasteSheet = ss.getSheetByName(SHEETS.DATA);
  var wasteData = wasteSheet.getDataRange().getValues();
  var registeredData = {};
  
  for (var n = 1; n < wasteData.length; n++) {
    var row = wasteData[n];
    var catId = String(row[1]);
    var catName = String(row[2]);
    var val = parseFloat(row[3]);
    var dateStr = Utilities.formatDate(new Date(row[4]), "JST", "yyyy-MM-dd");
    var roomId = String(row[6]);
    var user = String(row[12]);
    var time = String(row[13]); // フォーマット済み文字列を想定

    if (!registeredData[dateStr]) registeredData[dateStr] = {};
    if (!registeredData[dateStr][roomId]) {
      registeredData[dateStr][roomId] = { total: 0, cats: {}, logs: [] };
    }
    
    // カテゴリごとの最新値（または合計）を保持
    // 現在の仕様は「最新の上書き」を想定しているが、必要に応じて加算に変更可能
    registeredData[dateStr][roomId].cats[catId] = val;
    registeredData[dateStr][roomId].logs.push({
      catId: catId,
      catName: catName,
      val: val,
      user: user,
      time: time
    });
  }

  // 各日付・場所ごとの合計を算出
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
    serverTime: Date.now()
  };
}

/**
 * ユーザーの初期設定（拠点・場所）を保存
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
 * 報告データを保存（WastesDataシート）
 */
function executeSaveReport(formData) {
  var userEmail = Session.getActiveUser().getEmail() || "anonymous@mipox.co.jp";
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

  // 既存データの検索と更新、または新規追加
  var lastRow = sheet.getLastRow();
  var dataValues = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, 15).getValues() : [];
  var startRow = 2;

  for (var i = 0; i < formData.items.length; i++) {
    var item = formData.items[i];
    var inputVal = parseFloat(item.value);
    if (isNaN(inputVal) || inputVal <= 0) continue;
    
    var matchedIdx = -1;
    // 同じ日付・同じ場所・同じカテゴリの既存レコードを後ろから検索
    for (var k = dataValues.length - 1; k >= 0; k--) {
      if (String(dataValues[k][11]).replace(/,/g, '') === dateNumStr && 
          String(dataValues[k][1]).trim() === String(item.categoryId).trim() && 
          String(dataValues[k][6]).trim() === String(formData.roomId).trim()) {
        matchedIdx = startRow + k; 
        break;
      }
    }

    if (matchedIdx > 0) {
      // 既存レコードの上書き
      sheet.getRange(matchedIdx, 4).setValue(inputVal); // 計測値
      sheet.getRange(matchedIdx, 13).setValue(userEmail); // 更新者
      sheet.getRange(matchedIdx, 14).setValue(formattedNow); // 更新日時
    } else {
      // 新規行の追加
      var yearStr = new Date(formData.date).getFullYear() + "年";
      var monthStr = yearStr + ("0" + (new Date(formData.date).getMonth() + 1)).slice(-2) + "月";
      sheet.appendRow([
        lastRow + 1 + i, // ID (仮)
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
        userEmail,
        formattedNow,
        formData.baseName + "_" + roomName + "_" + item.categoryId + "_" + dateNumStr // タイトル
      ]);
    }
  }

  return { 
    success: true, 
    user: userEmail, 
    time: formattedNow 
  };
}