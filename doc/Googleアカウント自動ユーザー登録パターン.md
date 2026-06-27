# Google アカウント自動ユーザー登録パターン（GAS）

他のアプリへ転用するための設計パターンとして、廃棄物報告システムの初回登録フローを抽象化・整理したドキュメントです。

---

## 概要

このパターンは、**ユーザーが何もしなくても初回アクセス時に自動で登録が完了する**仕組みです。  
手動のアカウント作成・招待フローを排除し、Google Workspace 組織内メンバーであれば即座に使い始められます。

- **起点**: Google アカウントのセッション（ログイン済みであること）
- **名前解決**: Google OAuth2 userinfo API でアクセス中のユーザー自身のプロフィールを取得（管理者権限不要）
- **フォールバック**: API 取得失敗時はメールのローカルパートを仮名として使用
- **キャッシュ**: スプレッドシート上の `Master_User` シートをユーザー名簿として使用し、API コールを初回のみに抑制
- **重複防止**: `LockService` による同時書き込み制御

---

## データフロー

```
[アクセス発生]
    │
    ▼
[1. セッションからメール取得]
    Session.getActiveUser().getEmail()
    │
    ▼
[2. Master_User シートを検索]
    ├─ 一致あり ──→ 登録済み名前を取得 ──→ [4. 処理続行]
    └─ 一致なし ──→ [3. userinfo API で自己取得]
    │
    ▼
[3. Google OAuth2 userinfo API で自分の名前を取得]
    GET https://www.googleapis.com/oauth2/v3/userinfo
    Authorization: Bearer {ScriptApp.getOAuthToken()}
    ├─ 取得成功 ──→ Master_User に追記 ──→ [4. 処理続行]
    └─ 取得失敗 ──→ email.split('@')[0] をフォールバック名として使用
    │
    ▼
[4. 処理続行（名前を保持してデータ書き込み・画面返却 等）]
```

---

## コアロジック（転用可能な実装）

### ① セッションからメール取得

```javascript
var userEmail = Session.getActiveUser().getEmail() || "anonymous@example.com";
```

### ② 名前解決（キャッシュ照合 → userinfo API → フォールバック）

```javascript
function getOrRegisterUserName_(email) {
  if (!email) return null;
  var lowEmail = email.toLowerCase();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Master_User");

  // キャッシュシートを検索
  if (sheet) {
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).toLowerCase() === lowEmail) {
        return data[i][1]; // B列: フルネーム
      }
    }
  }

  // 見つからなければ userinfo API で自分の名前を取得（管理者権限不要）
  try {
    var token = ScriptApp.getOAuthToken();
    var res = UrlFetchApp.fetch(
      'https://www.googleapis.com/oauth2/v3/userinfo',
      { headers: { 'Authorization': 'Bearer ' + token }, muteHttpExceptions: true }
    );
    var person = JSON.parse(res.getContentText());
    var name = person.name || null;
    if (name) {
      var lock = LockService.getScriptLock();
      lock.waitLock(5000);
      try {
        if (!sheet) {
          sheet = ss.insertSheet("Master_User");
          sheet.appendRow(['メールアドレス', '名前']);
          sheet.getRange('1:1').setFontWeight('bold').setBackground('#f3f3f3');
          sheet.setFrozenRows(1);
        }
        // 重複防止：書き込み直前に再確認
        var latest = sheet.getDataRange().getValues();
        var alreadyAdded = latest.some(function(r) {
          return String(r[0]).toLowerCase() === lowEmail;
        });
        if (!alreadyAdded) sheet.appendRow([email, name]);
      } finally {
        lock.releaseLock();
      }
      return name;
    }
  } catch(e) { }
  return null; // 取得失敗
}
```

### ③ フォールバックを含む呼び出し例

```javascript
var userName = getOrRegisterUserName_(userEmail);
var userDisplayName = userName || (userEmail ? userEmail.split('@')[0] : "不明ユーザー");
```

---

## Master_User シートの構造

| 列 | 内容 | 備考 |
|----|------|------|
| A | メールアドレス（元の大文字小文字のまま） | 照合時は toLowerCase() で比較 |
| B | フルネーム | Google アカウントの表示名（`person.name`） |
| C 以降 | **拡張用に自由に追加可能** | 後述 |

---

## 拡張ポイント（他アプリへの応用時）

基本の A・B 列（メール・名前）は変えず、**C 列以降に機能を追加する**のが最もシンプルな拡張方法です。

| 拡張機能 | 追加列の例 | 実装メモ |
|----------|-----------|----------|
| 名前の手動編集 | C列: 表示名（上書き用） | 自動取得名を B、手動上書き名を C に分けて管理。C が空なら B を使う。 |
| 部署情報の登録 | D列: 部署名 | 初回登録時に手動入力、または別途管理 |
| 権限の設定 | E列: role（例: admin / user / readonly） | 初期値は `user` にしておくと安全 |

### 名前の手動編集パターン（例）

```javascript
function getDisplayName_(email) {
  // C列（手動上書き名）があればそれを優先、なければ B列（自動取得名）を使う
  var row = findUserRow_(email);
  if (!row) return email.split('@')[0];
  return row[2] || row[1]; // C列 || B列
}
```

---

## 必要な GCP / GAS 設定

| 項目 | 内容 |
|------|------|
| OAuth スコープ | `appsscript.json` に `userinfo.profile`・`userinfo.email`・`script.external_request`・`script.scriptapp`・`spreadsheets` を明示 |
| 実行権限 | `executeAs: USER_ACCESSING`（アクセスしているユーザーとして実行）で動作。管理者権限不要。 |
| スプレッドシート権限 | アプリが読み書きするスプレッドシートの編集者であること |
| 初回認証 | スコープ変更後はユーザーが再認証（myaccount.google.com/permissions でリセット → 再アクセス）が必要 |

---

## このパターンの特徴と注意点

- **管理者権限不要**: Admin SDK を使わず `oauth2/v3/userinfo` でアクセス中のユーザー自身の名前を取得するため、一般ユーザーでも動作します。
- **初回のみ API 呼び出し**: 2回目以降は `Master_User` シートのキャッシュを参照するため高速です。
- **組織外ユーザー**: Google アカウントを持つユーザーであれば、組織外でも名前が取得できます。
- **同時アクセス対策**: `LockService.getScriptLock()` で競合書き込みを防止しています。
- **人員変動への追従**: `Master_User` は自動追記のみで削除は行いません。退職者の行は手動削除が必要です。
