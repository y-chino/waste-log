# clasp を使ったローカル開発・テストフロー

VSCode で編集したコードを Google Apps Script にアップロードし、実スプレッドシートのデータで動作確認する手順。

---

## 前提環境

- Node.js がインストールされていること
- GAS プロジェクトが既に存在していること（スクリプト ID: `1qRlJJFAtj__4ju5x2nkDRwB0ypz9Omx1xBCu8lmOXPBZzvwajg0kYIgd`）

---

## 初回セットアップ

### 1. clasp をインストール

```bash
npm install -g @google/clasp
```

### 2. Google アカウントでログイン

```bash
clasp login
```

ブラウザが開くのでアカウントを選択して許可する。

### 3. GAS API を有効化

[https://script.google.com/home/usersettings](https://script.google.com/home/usersettings) を開き、「Google Apps Script API」をオンにする。

### 4. 設定ファイルの確認

`.clasp.json` と `.claspignore` はすでに配置済み。

> ⚠️ `clasp clone` や `clasp pull` は既存の `index.html` / `code.gs` を上書きするため使わない。

---

## 日常の開発フロー

```
1. VSCode で index.html / code.gs を編集
         ↓
2. clasp push でGASにアップロード
         ↓
3. テストデプロイURLをブラウザで開いて動作確認
```

### コードをアップロード（appsscript.json / code.gs / index.html の3ファイル）

```bash
clasp push
```

### GAS エディタを開く（確認用）

```bash
clasp open
```

---

## デプロイ URL の確認

GAS エディタ → 右上「デプロイ」→「**テストデプロイ**」

テストデプロイは `clasp push` のたびに自動で最新コードが反映され、URLも変わらないので便利。

---

## UIとデータの確認を分ける

| 確認したいこと | 使うツール |
|--------------|-----------|
| UI の見た目・レイアウト | Live Server（`http://127.0.0.1:5500`）|
| スプシとのデータ読み書き | clasp push → テストデプロイ URL |

Live Server では `google.script.run` が使えないため、`MOCK_DATA`（`index.html` 内に定義）が表示される。

---

## 注意事項

- **`clasp pull` は使わない** — `index.html` / `code.gs` がGAS側のバージョンで上書きされる
- **GASのブラウザエディタは触らない** — ローカルとの差分が生じる原因になる
- 開発は VSCode で行い、`clasp push` で一方向に反映する運用を徹底する

---

## トラブルシューティング

### `clasp push` でエラーになる

`.clasp.json` の内容を確認する：

```json
{
  "scriptId": "1qRlJJFAtj__4ju5x2nkDRwB0ypz9Omx1xBCu8lmOXPBZzvwajg0kYIgd",
  "rootDir": "./"
}
```

### `Project contents must include a manifest file named appsscript.json` と出る

`appsscript.json` が `.claspignore` に含まれている。clasp はマニフェストファイルの push を必須とするため、`.claspignore` から除外すること。

### push 後に反映されない

テストデプロイURLを使っているか確認する。通常のデプロイURLはキャッシュが残る場合がある。
