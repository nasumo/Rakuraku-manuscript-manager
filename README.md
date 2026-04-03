# らくらく原稿管理 — Google Apps Script Backend
 
> このリポジトリは、Figmaプラグイン「**らくらく原稿管理**」が使用する Google Apps Script のバックエンドスクリプトを公開しています。

---

## これは何？

「らくらく原稿管理」は、Figmaのテキストをスプレッドシートに書き出し、修正原稿を一括反映するFigmaプラグインです。

このスクリプトは、プラグインとあなたのGoogleスプレッドシートをつなぐバックエンドです。ご自身のGoogleアカウントにデプロイして使用します。開発者のサーバーにデータは送信されません。

---

## Setup / セットアップ

プラグイン内の **歯車アイコン（初期設定）** を開くと、詳細な手順が確認できます。

以下はその概要です。

### 1. Google Apps Script を開く

[script.google.com](https://script.google.com) を開き、「新しいプロジェクト」を作成します。

### 2. スクリプトを貼り付ける

[`code.gs`](./code.gs) の全文をコピーし、エディタ内の既存コードをすべて削除してから貼り付け、保存します（Cmd+S / Ctrl+S）。

### 3. ウェブアプリとしてデプロイする

1. 「デプロイ」→「新しいデプロイ」
2. 種類：**ウェブアプリ**
3. 次のユーザーとして実行：**自分**
4. アクセスできるユーザー：**全員**
5. 「デプロイ」をクリックし、権限を承認する
6. 表示された **ウェブアプリ URL** をコピーする

### 4. プラグインに URL を貼り付ける

Figmaプラグインの「初期設定」タブで、鉛筆アイコンを押して URL を貼り付け、保存します。

---

## 動作要件

- Google アカウント
- Figma デスクトップアプリ

---

## Plugin / プラグイン

Figma Community からインストール(現在申請中):  
👉 **[Rakuraku Manuscript Manager on Figma Community]([YOUR_FIGMA_COMMUNITY_URL])**

---

## Privacy / プライバシー

Privacy Policy: https://www.notion.so/Privacy-Policy-3378e13407d08023ac3adb9bd8a6e1bb

## Support / サポート

nasu@dono.co.jp

---

## License

The GAS script in this repository is provided for use with the Rakuraku Manuscript Manager Figma plugin. Redistribution or resale of this script as a standalone product is not permitted.
