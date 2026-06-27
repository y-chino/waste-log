# Google アカウント自動ユーザー登録パターン（GAS）

他のアプリへ転用するための設計パターンとして、廃棄物報告システムの初回登録フローを抽象化・整理したドキュメントです。

---

## 概要

このパターンは、**ユーザーが何もしなくても初回アクセス時に自動で登録が完了する**仕組みです。  
手動のアカウント作成・招待フローを排除し、Google Workspace 組織内メンバーであれば即座に使い始められます。

- **起点**: Google アカウントのセッション（ログイン済みであること）
- **名前解決**: Admin SDK で組織情報からフルネームを自動取得
- **フォールバック**: 組織情報未登録でも動作継続（メールのローカルパートを仮名として使用）
- **キャッシュ**: スプレッドシート上のシートをユーザー名簿として使用し、API コールを最小化

---

## データフロー

```
[アクセス / アクション発生]
    │
    ▼
[1. セッションからメール取得]
    Session.getActiveUser().getEmail()
    │
    ▼
[2. Master_User シートを検索]
    ├─ 一致あり ──→ 登録済み名前を取得 ──→ [5. 処理続行]
    └─ 一致なし ──→ [3. 組織同期]
    │
    ▼
[3. Admin SDK で組織ユーザーを一括取得・シートに書き込み]
    AdminDirectory.Users.list({ customer: 'my_customer', ... })
    │
    ▼
[4. 再検索]
    ├─ 一致あり ──→ フルネームを取得 ──→ [5. 処理続行]
    └─ 一致なし ──→ email.split('@')[0] をフォールバック名として使用
    │
    ▼
[5. 処理続行（名前を保持してデータ書き込み・画面返却 等）]
```

---

## コアロジック（転用可能な実装）

### ① セッションからメール取得

```javascript
var userEmail = Session.getActiveUser().getEmail() || "anonymous@example.com";
```

### ② 名前解決（キャッシュ照合 → 組織同期 → フォールバック）

```javascript
function getOrRegisterUserName_(email) {
  if (!email) return null;
  var lowEmail = email.toLowerCase();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Master_User");

  // シート自体がなければ組織同期して作成
  if (!sheet) {
    syncOrganizationUsers();
    sheet = ss.getSheetByName("Master_User");
  }

  // キャッシュシートを検索
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase() === lowEmail) {
      return data[i][1]; // B列: フルネーム
    }
  }

  // キャッシュにいなければ組織同期してリトライ
  syncOrganizationUsers();
  data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase() === lowEmail) {
      return data[i][1];
    }
  }

  return null; // 組織外ユーザーなど
}
```

### ③ 組織同期（Admin SDK → Master_User シートへ書き込み）

```javascript
function syncOrganizationUsers() {
  if (typeof AdminDirectory === 'undefined') {
    throw new Error("Admin SDK API が有効化されていません");
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Master_User");
  if (!sheet) {
    sheet = ss.insertSheet("Master_User");
    sheet.appendRow(["email", "name"]); // ヘッダー行
  }

  var users = [];
  var pageToken;
  do {
    var response = AdminDirectory.Users.list({
      customer: 'my_customer',
      maxResults: 500,
      pageToken: pageToken,
      orderBy: 'email',
      viewType: 'domain_public'
    });
    if (response.users) {
      response.users.forEach(function(user) {
        users.push([user.primaryEmail.toLowerCase(), user.name.fullName]);
      });
    }
    pageToken = response.nextPageToken;
  } while (pageToken);

  if (users.length > 0) {
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, 2).clearContent();
    sheet.getRange(2, 1, users.length, 2).setValues(users);
    SpreadsheetApp.flush();
  }
}
```

### ④ フォールバックを含む呼び出し例

```javascript
var userName = getOrRegisterUserName_(userEmail);
var userDisplayName = userName || (userEmail ? userEmail.split('@')[0] : "不明ユーザー");
```

---

## Master_User シートの構造

| 列 | 内容 | 備考 |
|----|------|------|
| A | メールアドレス（小文字） | 照合キー |
| B | フルネーム | Admin SDK の `user.name.fullName` |
| C 以降 | **拡張用に自由に追加可能** | 後述 |

---

## 拡張ポイント（他アプリへの応用時）

基本の A・B 列（メール・名前）は変えず、**C 列以降に機能を追加する**のが最もシンプルな拡張方法です。

| 拡張機能 | 追加列の例 | 実装メモ |
|----------|-----------|----------|
| 名前の手動編集 | C列: 表示名（上書き用） | 自動取得名を B、手動上書き名を C に分けて管理。C が空なら B を使う。 |
| 部署情報の登録 | D列: 部署名 | 初回同期時は Admin SDK の `orgUnitPath` から取得することも可能 |
| 権限の設定 | E列: role（例: admin / user / readonly） | 初期値は `user` にしておくと安全 |
| グループ登録 | F列: グループID（カンマ区切り等） | Admin SDK の `Groups.list` で同様に自動同期可能 |

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
| Admin SDK API の有効化 | GAS エディタ →「サービス」→ `Admin SDK API` を追加 |
| 実行権限 | スーパー管理者または「ユーザー管理の閲覧」権限を持つアカウントで doGet/doPost を実行する必要あり |
| スプレッドシート権限 | アプリが読み書きするスプレッドシートのオーナーまたは編集者であること |
| ウェブアプリの公開設定 | 「次のユーザーとして実行: 自分（管理者アカウント）」にすることで Admin SDK が呼び出せる |

> **注意**: 「ユーザーとして実行」を「アクセスしているユーザー」にすると、一般ユーザーは Admin SDK を呼び出せないためエラーになります。

---

## このパターンの限界と注意点

- **組織外ユーザー（外部ドメイン）** は Admin SDK で取得できないため、フォールバック（メールのローカルパート）のみになります。
- **Admin SDK の同期は重い処理**（ユーザー数が多い場合）なので、同期のタイミングはアクセス時ではなく定期実行（時間トリガー）に切り出すことを検討してください。
- **Master_User シートは「キャッシュ」** なので、組織の人員変動に追従するには定期的な `syncOrganizationUsers()` の実行が必要です（例: 1日1回のトリガー設定）。
