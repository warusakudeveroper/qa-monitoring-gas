# 📋 QA Monitoring System (Google Apps Script)

質疑応答連絡書を自動監視し、Discord/メール通知を行うGoogle Apps Scriptシステム

## ✨ 主な機能

- 🔍 **新規質疑の自動検知** - A〜H列が埋まったエントリを自動検出
- ✅ **回答監視** - I列(回答日)とJ列(回答内容)の追加を検知
- 🔄 **変更・削除監視** - データの変更や削除を追跡
- 💾 **自動バックアップ** - 別スプレッドシートへの定期バックアップ
- 📢 **マルチプラットフォーム通知** - Discord WebhookとEmail通知に対応

## 🚀 クイックスタート

### 1. スプレッドシートでスクリプトエディタを開く
```
拡張機能 → Apps Script
```

### 2. コードを貼り付け
`qa_monitoring.gs`の内容を全てコピーして貼り付け

### 3. 初期設定を実行
```javascript
setupScriptProperties()  // スクリプトプロパティ設定
setupTimeTrigger()       // 1時間ごとの自動実行設定
```

### 4. プロパティを設定
- `DISCORD_WEBHOOKS`: Discord Webhook URL（カンマ区切りで複数可）
- `EMAIL_ADDRESSES`: 通知先メールアドレス（カンマ区切りで複数可）
- `BACKUP_SHEET_ID`: バックアップ用スプレッドシートのID

### 5. テスト実行
```javascript
testRun()  // 手動でテスト実行
```

## 📸 通知イメージ

### Discord通知
```
📝 **新規質疑が登録されました**

**【質疑 #0918】**
👤 **発信者:** 運営者
📮 **宛先:** 設計者
```[質疑内容]```
━━━━━━━━━━━━━━━
📊 **件数:** 1件
```

### メール通知
美しいHTMLテンプレートで、質疑・回答を見やすく表示

## 📊 スプレッドシート構造

| 列 | 項目 | 説明 |
|---|---|---|
| A | 番号 | 質疑番号（一意） |
| B | 発信者 | 質問者名 |
| C | 発信者部署 | 質問者の所属 |
| D | フェーズ | プロジェクトフェーズ |
| E | カテゴリ | 質疑のカテゴリ |
| F | 参照資料 | 関連資料 |
| G | 質疑内容 | 質問内容 |
| H | 宛先 | 回答者 |
| I | 回答日 | 回答記入日 |
| J | 回答 | 回答内容 |

## ⚙️ カスタマイズ

### 監視間隔の変更
```javascript
// 30分ごとに変更
ScriptApp.newTrigger('monitorQASheet')
  .timeBased()
  .everyMinutes(30)
  .create();
```

### 通知フォーマット変更
- `formatQuestionsForDiscord()` - Discord質疑通知
- `formatAnswersForDiscord()` - Discord回答通知
- `generateQuestionEmailContent()` - メール質疑通知
- `generateAnswerEmailContent()` - メール回答通知

## 📁 ファイル構成

```
qa-monitoring-gas/
├── qa_monitoring.gs     # メインスクリプト
├── 導入手順書.md        # 詳細な導入ガイド
├── README.md           # このファイル
└── .gitignore         # Git除外設定
```

## 🔧 トラブルシューティング

### 通知が届かない
1. スクリプトプロパティの設定値を確認
2. トリガーが正しく設定されているか確認
3. 実行ログでエラーを確認

### 重複通知が発生
```javascript
resetProcessedLists()  // 処理済みリストをリセット
```

### 詳細は[導入手順書](./docs/導入手順書.md)を参照

## 🛠️ 主要な関数

| 関数名 | 説明 |
|---|---|
| `monitorQASheet()` | メイン監視処理（トリガー実行） |
| `testRun()` | 手動テスト実行 |
| `setupScriptProperties()` | 初期プロパティ設定 |
| `setupTimeTrigger()` | 時間トリガー設定 |
| `resetProcessedLists()` | 処理済みリストのリセット |

## 📋 必要な権限

- Googleスプレッドシートの読み書き
- 外部URLへのアクセス（Discord Webhook）
- メール送信（Gmail）
- スクリプトプロパティの読み書き

## 🔒 セキュリティ

- Discord Webhook URLは外部に漏らさない
- スクリプトプロパティへのアクセスを制限
- 定期的にWebhook URLを更新

## 📝 ライセンス

MIT License

## 👥 コントリビューション

Issue、Pull Requestは歓迎です。

## 📚 関連ドキュメント

- [Google Apps Script公式ドキュメント](https://developers.google.com/apps-script)
- [Discord Webhook Guide](https://discord.com/developers/docs/resources/webhook)

## 🆘 サポート

問題が発生した場合は、[Issues](https://github.com/warusaku/qa-monitoring-gas/issues)でお知らせください。

---

Made with ❤️ for efficient QA management