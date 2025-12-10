# -*- coding: utf-8 -*-
import os
import platform
import sys
from pathlib import Path

# Windows環境でのUTF-8出力を強制設定
if platform.system() == "Windows":
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

# パッケージのインポートを試行し、エラーハンドリングを追加
try:
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from googleapiclient.discovery import build
    from googleapiclient.http import MediaFileUpload, MediaIoBaseUpload
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

# 必要な権限スコープ（既存フォルダ検索 + ファイル作成）
SCOPES = [
    'https://www.googleapis.com/auth/drive.metadata.readonly',
    'https://www.googleapis.com/auth/drive.file',
]

# アップロード先のフォルダ名（マイドライブ/アクセス数/カーセンサー_アクセス数 / グーネット_アクセス数）
PARENT_FOLDER_NAME = 'アクセス数'
TARGET_FOLDER_NAME = 'カーセンサー_アクセス数'
GOONET_FOLDER_NAME = 'グーネット_アクセス数'

# 登録物件数配下のフォルダ
REG_PARENT_FOLDER_NAME = '登録物件数'
REG_CAR_SENSOR_FOLDER_NAME = 'カーセンサー_登録物件数'
REG_GOONET_FOLDER_NAME = 'グーネット_登録物件数'

def get_downloads_folder():
    """OSに応じてダウンロードフォルダのパスを取得"""
    system = platform.system()
    
    if system == "Windows":
        downloads_path = Path.home() / "Downloads"
    elif system == "Darwin":  # macOS
        downloads_path = Path.home() / "Downloads"
    else:  # Linux
        downloads_path = Path.home() / "Downloads"
    
    return downloads_path

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
    print(f"[WARNING] フォルダ '{parent_name}/{child_name}' は見つかりません（作成しません）。")
    return None

def get_target_folder_id(service):
    """マイドライブ/アクセス数/カーセンサー_アクセス数 の既存フォルダIDを取得（作成しない）"""
    return find_existing_nested_folder(service, PARENT_FOLDER_NAME, TARGET_FOLDER_NAME)

def get_child_folder_id(service, child_folder_name: str):
    """マイドライブ/アクセス数/<子> の既存フォルダIDを返す（作成しない）。"""
    return find_existing_nested_folder(service, PARENT_FOLDER_NAME, child_folder_name)

def get_nested_child_folder_id(service, parent_folder_name: str, child_folder_name: str):
    """マイドライブ/<親>/<子> の既存フォルダIDを返す（作成しない）。"""
    return find_existing_nested_folder(service, parent_folder_name, child_folder_name)

def file_exists_in_folder(service, filename: str, parent_folder_id: str) -> bool:
    """指定フォルダ内に同名ファイルが既に存在するかを確認（ゴミ箱除外）。"""
    query = (
        f"name='{filename}' and '{parent_folder_id}' in parents and trashed=false"
    )
    results = service.files().list(q=query, fields="files(id,name)").execute()
    return len(results.get('files', [])) > 0

def upload_single_file(service, file_path: Path, child_folder_name: str):
    if not file_path.exists():
        print(f"エラー: ファイルが存在しません: {file_path}")
        return None
    target_folder_id = get_child_folder_id(service, child_folder_name)
    if not target_folder_id:
        print(f"エラー: 'マイドライブ/{PARENT_FOLDER_NAME}/{child_folder_name}' が見つからないためスキップします。")
        return None
    print(f"アップロード先: マイドライブ/{PARENT_FOLDER_NAME}/{child_folder_name} (ID: {target_folder_id})")
    # 既存重複チェック
    if file_exists_in_folder(service, file_path.name, target_folder_id):
        print("[SKIP] 既に同名ファイルが存在するためアップロードをスキップします。")
        try:
            file_path.unlink()
            print(f"[DELETE] ローカルファイルを削除しました: {file_path}")
        except Exception as e:
            print(f"⚠ ローカルファイルの削除に失敗しました: {file_path} | {e}")
        return None
    file_metadata = {
        'name': file_path.name,
        'parents': [target_folder_id]
    }
    # ファイルハンドルを with で管理し、アップロード後に確実にクローズする（Windows のロック対策）
    print("アップロードを開始します...")
    with open(file_path, 'rb') as fh:
        media = MediaIoBaseUpload(fh, mimetype='text/csv', resumable=False)
        file = service.files().create(body=file_metadata, media_body=media, fields='id,name,size,createdTime').execute()
    print("[OK] アップロード完了!")
    print(f"ファイル名: {file.get('name')}")
    print(f"ファイルID: {file.get('id')}")
    print(f"作成日時: {file.get('createdTime')}")
    print(f"アップロード先: マイドライブ/{PARENT_FOLDER_NAME}/{child_folder_name}")
    print(f"Google DriveでのURL: https://drive.google.com/file/d/{file.get('id')}/view")
    
    # アップロード成功後にローカルファイルを削除（単発削除）
    try:
        file_path.unlink()
        print(f"[DELETE] ローカルファイルを削除しました: {file_path}")
    except Exception as e:
        print(f"[WARNING] ローカルファイルの削除に失敗しました: {file_path} | {e}")
    return file.get('id')

def upload_single_file_to(service, file_path: Path, parent_folder_name: str, child_folder_name: str):
    if not file_path.exists():
        print(f"エラー: ファイルが存在しません: {file_path}")
        return None
    target_folder_id = get_nested_child_folder_id(service, parent_folder_name, child_folder_name)
    if not target_folder_id:
        print(f"エラー: 'マイドライブ/{parent_folder_name}/{child_folder_name}' が見つからないためスキップします。")
        return None
    print(f"アップロード先: マイドライブ/{parent_folder_name}/{child_folder_name} (ID: {target_folder_id})")
    # 既存重複チェック
    if file_exists_in_folder(service, file_path.name, target_folder_id):
        print("[SKIP] 既に同名ファイルが存在するためアップロードをスキップします。")
        try:
            file_path.unlink()
            print(f"[DELETE] ローカルファイルを削除しました: {file_path}")
        except Exception as e:
            print(f"⚠ ローカルファイルの削除に失敗しました: {file_path} | {e}")
        return None
    file_metadata = {
        'name': file_path.name,
        'parents': [target_folder_id]
    }
    print("アップロードを開始します...")
    with open(file_path, 'rb') as fh:
        media = MediaIoBaseUpload(fh, mimetype='text/csv', resumable=False)
        file = service.files().create(body=file_metadata, media_body=media, fields='id,name,size,createdTime').execute()
    print("[OK] アップロード完了!")
    print(f"ファイル名: {file.get('name')}")
    print(f"ファイルID: {file.get('id')}")
    print(f"作成日時: {file.get('createdTime')}")
    print(f"アップロード先: マイドライブ/{parent_folder_name}/{child_folder_name}")
    print(f"Google DriveでのURL: https://drive.google.com/file/d/{file.get('id')}/view")
    try:
        file_path.unlink()
        print(f"[DELETE] ローカルファイルを削除しました: {file_path}")
    except Exception as e:
        print(f"[WARNING] ローカルファイルの削除に失敗しました: {file_path} | {e}")
    return file.get('id')

def upload_matching_downloads():
    """ダウンロードフォルダからパターンに一致するCSVをそれぞれのフォルダへアップロードする。
    - 'hankyobukken' を含むCSV → カーセンサー_アクセス数
    - '効果分析（在庫）' を含むCSV → グーネット_アクセス数
    既存フォルダのみを使用し、新規作成はしない。
    """
    try:
        downloads_folder = get_downloads_folder()
        service = authenticate_google_drive()

        # 収集
        all_csvs = [p for p in downloads_folder.glob('*.csv')]
        hankyo = [p for p in all_csvs if 'hankyobukken' in p.name.lower()]
        kouka = [p for p in all_csvs if '効果分析（在庫）' in p.name]
        toroku = [p for p in all_csvs if 'torokubukken' in p.name.lower()]
        zaikoken = [p for p in all_csvs if '在庫検索一覧' in p.name]

        print("アップロード対象(カーセンサー: hankyobukken):")
        for p in hankyo:
            print(f" - {p}")
        print("アップロード対象(グーネット: 効果分析（在庫）):")
        for p in kouka:
            print(f" - {p}")
        print("アップロード対象(カーセンサー: torokubukken/登録物件数):")
        for p in toroku:
            print(f" - {p}")
        print("アップロード対象(グーネット: 在庫検索一覧/登録物件数):")
        for p in zaikoken:
            print(f" - {p}")

        uploaded = []
        for p in hankyo:
            print(f"\n[カーセンサー] アップロード: {p}")
            fid = upload_single_file(service, p, TARGET_FOLDER_NAME)
            if fid:
                uploaded.append(fid)

        for p in kouka:
            print(f"\n[グーネット] アップロード: {p}")
            fid = upload_single_file(service, p, GOONET_FOLDER_NAME)
            if fid:
                uploaded.append(fid)

        # 登録物件数系
        for p in toroku:
            print(f"\n[カーセンサー/登録物件数] アップロード: {p}")
            fid = upload_single_file_to(service, p, REG_PARENT_FOLDER_NAME, REG_CAR_SENSOR_FOLDER_NAME)
            if fid:
                uploaded.append(fid)

        for p in zaikoken:
            print(f"\n[グーネット/登録物件数] アップロード: {p}")
            fid = upload_single_file_to(service, p, REG_PARENT_FOLDER_NAME, REG_GOONET_FOLDER_NAME)
            if fid:
                uploaded.append(fid)

        return uploaded

    except Exception as e:
        print(f"エラーが発生しました: {e}")
        return []

def main():
    """メイン実行関数"""
    print("=== Google Drive CSV アップローダー ===")
    
    # credentials.jsonの存在確認
    if not os.path.exists('credentials.json'):
        print("エラー: 'credentials.json' ファイルが見つかりません。")
        print("Google Cloud Consoleから認証情報をダウンロードして、")
        print("Pythonスクリプトと同じディレクトリに配置してください。")
        return
    
    uploaded = upload_matching_downloads()
    if uploaded:
        print(f"\n[SUCCESS] {len(uploaded)} 件のアップロードが完了しました!")
        print("Google Driveで確認してください: https://drive.google.com/")
    else:
        print("\n[ERROR] アップロード対象が見つからないか、失敗しました。")

if __name__ == '__main__':
    main()