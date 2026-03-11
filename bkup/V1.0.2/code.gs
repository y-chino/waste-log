/**
 * 廃棄物報告アプリ - Backend Logic (v1.0.2)
 */

var APP_TITLE = "廃棄物報告";

var SHEETS = {
  CATEGORY: "Master_Category",
  LOCATION: "Master_Location",
  USER_SETTING: "UserSetting",
  DATA: "WastesData"
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
 * 初期データ取得
 * v1.0.2 修正: 登録者を問わず最新5,000件を取得するように変更
 */
function getInitialDataForApp_(email) {
  var result = {
    categories: [],
    locations: [],
    userSetting: null,
    registeredData: {}, 
    serverTime: Date.now(),
    userExists: false,
    error: null
  };

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sCategory = ss.getSheetByName(SHEETS.CATEGORY);
    var sLocation = ss.getSheetByName(SHEETS.LOCATION);
    var sUserSetting = ss.getSheetByName(SHEETS.USER_SETTING);
    var sData = ss.getSheetByName(SHEETS.DATA);
    
    if (!sCategory || !sLocation || !sUserSetting || !sData) throw new Error("シート欠落");

    var catValues = sCategory.getDataRange().getValues();
    var locValues = sLocation.getDataRange().getValues();
    var setValues = sUserSetting.getDataRange().getValues();
    
    result.categories = catValues.slice(1).map(function(r) {
      return { id: String(r[0]), name: String(r[1]), detail: String(r[2]), targetBase: String(r[3]), unit: "Kg" };
    });
    
    result.locations = locValues.slice(1).map(function(r) {
      return { id: String(r[0]), name: String(r[1]), dept: String(r[2]), base: String(r[3]) };
    });
    
    for (var i = 1; i < setValues.length; i++) {
      if (setValues[i][0] && String(setValues[i][0]).toLowerCase() === email.toLowerCase()) {
        result.userSetting = { base: String(setValues[i][1]), roomId: String(setValues[i][2] || "") };
        result.userExists = true;
        break;
      }
    }

    var lastRow = sData.getLastRow();
    if (lastRow > 1) {
      var scanLimit = 5000;
      var startRow = Math.max(2, lastRow - scanLimit + 1);
      var rowsToFetch = lastRow - startRow + 1;
      // N列(更新日時)まで含めて取得するため、列数を13から14へ拡張
      var dataValues = sData.getRange(startRow, 1, rowsToFetch, 14).getValues();
      
      var regData = {};
      for (var k = 0; k < dataValues.length; k++) {
        // v1.0.2 変更: 更新者(M列/index 12)によるメールアドレスのフィルタリングを削除
        var d = dataValues[k][4]; // E列: 日付
        if (d) {
          var dateObj = (d instanceof Date) ? d : new Date(d);
          if (!isNaN(dateObj.getTime())) {
            var ds = dateObj.getFullYear() + "-" + ("0" + (dateObj.getMonth() + 1)).slice(-2) + "-" + ("0" + dateObj.getDate()).slice(-2);
            var catId = String(dataValues[k][1]); // B列: カテゴリID
            var val = parseFloat(dataValues[k][3]) || 0; // D列: 計測値
            var rId = String(dataValues[k][6]); // G列: 排出元ID
            
            if (!regData[ds]) regData[ds] = {};
            if (!regData[ds][rId]) regData[ds][rId] = { total: 0, cats: {} };
            
            regData[ds][rId].total += val;
            regData[ds][rId].cats[catId] = (regData[ds][rId].cats[catId] || 0) + val;
          }
        }
      }
      result.registeredData = regData;
    }
    return result;
  } catch (e) {
    result.error = e.message;
    return result;
  }
}

/**
 * ユーザー設定保存
 */
function registerUserSetting(base, roomId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEETS.USER_SETTING);
    var email = Session.getActiveUser().getEmail();
    var lastRow = sheet.getLastRow();
    var values = (lastRow > 0) ? sheet.getRange(1, 1, lastRow, 1).getValues() : [];
    var foundIdx = -1;
    for (var i = 0; i < values.length; i++) {
      if (values[i][0] && String(values[i][0]).toLowerCase() === email.toLowerCase()) { foundIdx = i + 1; break; }
    }
    if (foundIdx > 0) { sheet.getRange(foundIdx, 2, 1, 2).setValues([[base, String(roomId)]]); } 
    else { sheet.appendRow([email, base, String(roomId)]); }
    return { success: true };
  } catch (e) { return { success: false, error: e.message }; }
}

/**
 * 報告データの保存
 * v1.0.2 修正: 既存レコードの検索時に登録者(M列)を問わないように変更
 */
function executeSaveReport(formData) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEETS.DATA);
    var userEmail = Session.getActiveUser().getEmail();
    var now = new Date();
    var dateNumStr = formData.date.replace(/-/g, "");
    var sLocation = ss.getSheetByName(SHEETS.LOCATION);
    var locs = sLocation.getDataRange().getValues();
    var deptName = "", roomName = "";
    for (var j = 1; j < locs.length; j++) {
      if (String(locs[j][0]) === String(formData.roomId)) { roomName = String(locs[j][1]); deptName = String(locs[j][2]); break; }
    }
    var lastRow = sheet.getLastRow();
    var startRow = Math.max(2, lastRow - 5000);
    var rowsToFetch = lastRow - startRow + 1;
    var dataValues = (lastRow > 1) ? sheet.getRange(startRow, 1, rowsToFetch, 14).getValues() : [];

    for (var i = 0; i < formData.items.length; i++) {
      var item = formData.items[i];
      var inputVal = parseFloat(item.value);
      if (isNaN(inputVal) || inputVal <= 0) continue;
      var matchedIdx = -1;
      for (var k = dataValues.length - 1; k >= 0; k--) {
        // v1.0.2 変更: userEmail による一致確認を削除。社内共有データとして日付・カテゴリ・場所が一致すれば上書き対象とする
        if (String(dataValues[k][11]).replace(/,/g, '') === dateNumStr && 
            String(dataValues[k][1]) === String(item.categoryId) && 
            String(dataValues[k][6]) === String(formData.roomId)) {
          matchedIdx = startRow + k; 
          break;
        }
      }
      if (matchedIdx > 0) {
        sheet.getRange(matchedIdx, 4).setValue(inputVal); // D列: 計測値
        sheet.getRange(matchedIdx, 13).setValue(userEmail); // M列: 更新者を最新の上書き者に更新
        sheet.getRange(matchedIdx, 14).setValue(Utilities.formatDate(now, "JST", "yyyy/MM/dd HH:mm")); // N列: 更新日時
      } else {
        var yearStr = new Date(formData.date).getFullYear() + "年";
        var monthStr = yearStr + ("0" + (new Date(formData.date).getMonth() + 1)).slice(-2) + "月";
        // 15列目のタイトルも拠点・場所・カテゴリ・日付の組み合わせで生成
        sheet.appendRow([lastRow + 1 + i, item.categoryId, item.categoryName, inputVal, formData.date, formData.baseName, formData.roomId, roomName, deptName, yearStr, monthStr, dateNumStr, userEmail, Utilities.formatDate(now, "JST", "yyyy/MM/dd HH:mm"), formData.baseName + "_" + roomName + "_" + item.categoryId + "_" + dateNumStr]);
      }
    }
    return { success: true };
  } catch (e) { return { success: false, error: e.message }; }
}