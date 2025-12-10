# toGoogleDrive.py 実行ガイド

このスクリプトはダウンロードフォルダからCSVファイルを自動的にGoogle Driveの指定フォルダにアップロードします。

## 実行方法

以下のコマンドを実行してください：

```bash
export PYTHONIOENCODING=utf-8 && cd "C:\Users\m-oka\car-produce-python" && python toGoogleDrive.py
```

## 重要な注意事項

**`PYTHONIOENCODING=utf-8` 環境変数は必須です。** このスクリプトは日本語テキストを含んでいるため、UTF-8エンコーディングが必要です。この環境変数を設定しないと、文字エンコーディングエラーが発生します。

## セットアップ

実行前に以下が必要です：

1. **Google API認証情報**
   - `credentials.json` ファイルをスクリプトと同じディレクトリに配置してください
   - Google Cloud Consoleから「OAuth 2.0 クライアント ID」をダウンロードしてください

2. **必要なPythonパッケージ**
   ```bash
   pip install google-api-python-client google-auth-httplib2 google-auth-oauthlib
   ```

3. **.envファイル（Notion連携時）**
   - Notion API連携機能を使用する場合は `.env` ファイルに環境変数を設定してください

## アップロード対象ファイル

スクリプトは以下のファイルを自動的に検出してアップロードします：

- **hankyobukken** を含むCSV → `マイドライブ/アクセス数/カーセンサー_アクセス数`
- **効果分析（在庫）** を含むCSV → `マイドライブ/アクセス数/グーネット_アクセス数`
- **torokubukken** を含むCSV → `マイドライブ/登録物件数/カーセンサー_登録物件数`
- **在庫検索一覧** を含むCSV → `マイドライブ/登録物件数/グーネット_登録物件数`

## トラブルシューティング

### 文字エンコーディングエラーが発生する場合
- コマンドの最初に `export PYTHONIOENCODING=utf-8` を付けてください
- このコマンド全体をコピー＆ペーストして実行してください：
  ```bash
  export PYTHONIOENCODING=utf-8 && cd "C:\Users\m-oka\car-produce-python" && python toGoogleDrive.py
  ```

### Google認証エラーが発生する場合
- `credentials.json` がスクリプトと同じディレクトリにあることを確認してください
- ブラウザで認証画面が表示される場合は、指示に従って認可してください

### フォルダが見つからないエラーが表示される場合
- Google Drive上に対応するフォルダ構造が存在することを確認してください
- スクリプトは既存フォルダにのみアップロードします（新規作成はしません）
