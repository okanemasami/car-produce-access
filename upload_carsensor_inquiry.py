# -*- coding: utf-8 -*-
"""
カーセンサー問い合わせ漏れ防止用CSVをGoogle Driveにアップロードするスクリプト

ローカルPC: C:\Users\m-oka\Downloads\カーセンサー_問い合わせ漏れ防止用.csv
Google Drive: マイドライブ > 問い合わせ数 > カーセンサー

アップロード成功後、ローカルファイルを削除します。
"""

import os
import sys
import platform
from pathlib import Path

# Windows環境でのUTF-8出力を強制設定
if platform.system() == "Windows":
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

# パッケージのインポート
try:
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from googleapiclient.discovery import build
    from googleapiclient.http import MediaIoBaseUpload
    print("[OK] すべてのGoogleライブラリのインポートに成功しました")
except ImportError as e:
    print(f"[ERROR] Googleライブラリのインポートエラー: {e}")
    print("\n解決方法:")
    print("1. 仮想環境を作成してください:")
    print("   python -m venv google_drive_env")
    print("   google_drive_env\\Scripts\\activate")
    print("   pip install google-api-python-client google-auth-httplib2 google-auth-oauthlib")
    print("\n2. または以下のコマンドで再インストール:")
    print("   pip install --force-reinstall google-api-python-client google-auth-httplib2 google-auth-oauthlib")
    sys.exit(1)

# 必要な権限スコープ
SCOPES = [
    'https://www.googleapis.com/auth/drive.metadata.readonly',
    'https://www.googleapis.com/auth/drive.file',
]

# アップロード対象ファイル
TARGET_FILE_PATH = Path(r"C:\Users\m-oka\Downloads\カーセンサー_問い合わせ漏れ防止用.csv")

# Google Drive アップロード先
INQUIRY_PARENT_FOLDER_NAME = '問い合わせ数'
INQUIRY_CARSENSOR_FOLDER_NAME = 'カーセンサー'


def authenticate_google_drive():
    """Google Drive APIの認証を行う"""
    creds = None

    # token.jsonが存在する場合は既存の認証情報を使用
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)

    # スコープ不足も再認証対象にする
    def scopes_missing(c):
        try:
            return not set(SCOPES).issubset(set(c.scopes or []))
        except Exception:
            return True

    # 認証情報が無効または存在しない、またはスコープ不足なら再認証
    if not creds or not creds.valid or scopes_missing(creds):
        if creds and creds.expired and creds.refresh_token and not scopes_missing(creds):
            print("認証情報を更新中...")
            creds.refresh(Request())
        else:
            print("新規認証を開始します（必要な権限を付与）...")
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
            print("認証が完了しました。")

        # 認証情報を保存
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    return build('drive', 'v3', credentials=creds)


def find_existing_nested_folder(service, parent_name: str, child_name: str):
    """親フォルダ名が parent_name の直下にある child_name フォルダのIDを返す（作成しない）。"""
    query = (
        f"name='{child_name}' and mimeType='application/vnd.google-apps.folder' and trashed=false"
    )
    results = service.files().list(q=query, fields="files(id,name,parents)").execute()
    items = results.get('files', [])
    for item in items:
        for parent_id in item.get('parents', []) or []:
            try:
                parent = service.files().get(fileId=parent_id, fields='id,name').execute()
                if parent.get('name') == parent_name:
                    print(f"[FOUND] フォルダ '{parent_name}/{child_name}' が見つかりました (ID: {item['id']})")
                    return item['id']
            except Exception:
                continue
    print(f"[ERROR] フォルダ '{parent_name}/{child_name}' は見つかりません。")
    return None


def file_exists_in_folder(service, filename: str, parent_folder_id: str) -> bool:
    """指定フォルダ内に同名ファイルが既に存在するかを確認（ゴミ箱除外）。"""
    query = (
        f"name='{filename}' and '{parent_folder_id}' in parents and trashed=false"
    )
    results = service.files().list(q=query, fields="files(id,name)").execute()
    return len(results.get('files', [])) > 0


def upload_file_to_drive(service, file_path: Path, parent_folder_name: str, child_folder_name: str):
    """ファイルをGoogle Driveにアップロードし、成功したらローカルファイルを削除"""
    if not file_path.exists():
        print(f"[ERROR] ファイルが存在しません: {file_path}")
        return None

    # アップロード先フォルダのIDを取得
    target_folder_id = find_existing_nested_folder(service, parent_folder_name, child_folder_name)
    if not target_folder_id:
        print(f"[ERROR] 'マイドライブ/{parent_folder_name}/{child_folder_name}' が見つからないためスキップします。")
        print("Google Driveでフォルダを作成してから再実行してください。")
        return None

    print(f"アップロード先: マイドライブ/{parent_folder_name}/{child_folder_name} (ID: {target_folder_id})")

    # 既存重複チェック
    if file_exists_in_folder(service, file_path.name, target_folder_id):
        print("[SKIP] 既に同名ファイルが存在するためアップロードをスキップします。")
        try:
            file_path.unlink()
            print(f"[DELETE] ローカルファイルを削除しました: {file_path}")
        except Exception as e:
            print(f"[WARNING] ローカルファイルの削除に失敗しました: {file_path} | {e}")
        return None

    # ファイルメタデータ
    file_metadata = {
        'name': file_path.name,
        'parents': [target_folder_id]
    }

    # アップロード実行
    print(f"アップロードを開始します: {file_path.name}")
    try:
        with open(file_path, 'rb') as fh:
            media = MediaIoBaseUpload(fh, mimetype='text/csv', resumable=False)
            file = service.files().create(
                body=file_metadata,
                media_body=media,
                fields='id,name,size,createdTime'
            ).execute()

        print("[SUCCESS] アップロード完了!")
        print(f"ファイル名: {file.get('name')}")
        print(f"ファイルID: {file.get('id')}")
        print(f"作成日時: {file.get('createdTime')}")
        print(f"アップロード先: マイドライブ/{parent_folder_name}/{child_folder_name}")
        print(f"Google DriveでのURL: https://drive.google.com/file/d/{file.get('id')}/view")

        # アップロード成功後にローカルファイルを削除
        try:
            file_path.unlink()
            print(f"[DELETE] ローカルファイルを削除しました: {file_path}")
        except Exception as e:
            print(f"[WARNING] ローカルファイルの削除に失敗しました: {file_path} | {e}")

        return file.get('id')

    except Exception as e:
        print(f"[ERROR] アップロードに失敗しました: {e}")
        return None


def main():
    """メイン実行関数"""
    print("=== カーセンサー問い合わせ漏れ防止用CSV アップローダー ===")
    print(f"対象ファイル: {TARGET_FILE_PATH}")
    print(f"アップロード先: マイドライブ/{INQUIRY_PARENT_FOLDER_NAME}/{INQUIRY_CARSENSOR_FOLDER_NAME}")
    print()

    # credentials.jsonの存在確認
    if not os.path.exists('credentials.json'):
        print("[ERROR] 'credentials.json' ファイルが見つかりません。")
        print("Google Cloud Consoleから認証情報をダウンロードして、")
        print("Pythonスクリプトと同じディレクトリに配置してください。")
        return

    # ファイル存在確認
    if not TARGET_FILE_PATH.exists():
        print(f"[WARNING] ファイルが見つかりません: {TARGET_FILE_PATH}")
        print("ファイルがダウンロードされていない可能性があります。")
        return

    # Google Drive認証
    try:
        service = authenticate_google_drive()
    except Exception as e:
        print(f"[ERROR] Google Drive認証に失敗しました: {e}")
        return

    # アップロード実行
    file_id = upload_file_to_drive(
        service,
        TARGET_FILE_PATH,
        INQUIRY_PARENT_FOLDER_NAME,
        INQUIRY_CARSENSOR_FOLDER_NAME
    )

    if file_id:
        print("\n[SUCCESS] 処理が完了しました!")
    else:
        print("\n[INFO] アップロードは実行されませんでした。")


if __name__ == '__main__':
    main()
