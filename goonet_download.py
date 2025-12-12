import os
import json
import time
import glob
import shutil
import traceback
import selenium
import datetime
from pathlib import Path

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC

from webdriver_manager.chrome import ChromeDriverManager

# ============================================================
# 設定読み込み（.env → settings.json → 環境変数）
#  - HEADLESS / DOWNLOAD_DIR はカーセンサーと共通利用
#  - GOONET_USERNAME / GOONET_PASSWORD を settings.json に追記して使う
# ============================================================

def load_settings():
    settings = {}

    # .env（任意）
    try:
        from dotenv import load_dotenv
        env_path = Path(__file__).with_name(".env")
        if env_path.exists():
            load_dotenv(env_path)
    except Exception:
        pass

    # settings.json / setting.json / setteing.json（タイポも拾う）
    settings_json = None
    for name in ("settings.json", "setting.json", "setteing.json"):
        p = Path(__file__).with_name(name)
        if p.exists():
            settings_json = p
            break
    if settings_json:
        try:
            settings.update(json.loads(settings_json.read_text(encoding="utf-8")))
        except Exception as e:
            print(f"settings.json の読み込みに失敗しました: {e}")

    # 環境変数で上書き
    env_map = {
        "HEADLESS": os.getenv("HEADLESS"),
        "DOWNLOAD_DIR": os.getenv("DOWNLOAD_DIR"),
        "GOONET_USERNAME": os.getenv("GOONET_USERNAME"),
        "GOONET_PASSWORD": os.getenv("GOONET_PASSWORD"),
    }
    for k, v in env_map.items():
        if v is not None:
            settings[k] = v

    # 型整備
    headless = settings.get("HEADLESS", "false")
    if isinstance(headless, str):
        headless = headless.strip().lower() in ("1", "true", "yes", "on")
    settings["HEADLESS"] = bool(headless)

    # ダウンロード先（未指定なら OS 既定の Downloads）
    dl = settings.get("DOWNLOAD_DIR")
    if dl:
        download_dir = Path(dl).expanduser()
    else:
        candidates = [Path.home() / "Downloads", Path.home() / "ダウンロード"]
        download_dir = next((p for p in candidates if p.exists()), Path.home() / "Downloads")
    download_dir.mkdir(parents=True, exist_ok=True)
    settings["DOWNLOAD_DIR"] = str(download_dir)

    # 資格情報チェック
    if not settings.get("GOONET_USERNAME") or not settings.get("GOONET_PASSWORD"):
        raise RuntimeError(
            "GOONET の ID/PW が設定されていません。settings.json に "
            "GOONET_USERNAME / GOONET_PASSWORD を追記してください。"
        )

    return settings


# ============================================================
# WebDriver 準備
# ============================================================

def build_driver(download_dir: Path, headless: bool):
    options = webdriver.ChromeOptions()
    if headless:
        options.add_argument("--headless=new")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--start-maximized")

    # ダウンロード設定（ヘッドレスでも保存可能）
    prefs = {
        "download.default_directory": str(download_dir.resolve()),
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
    }
    options.add_experimental_option("prefs", prefs)
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_argument("--disable-blink-features=AutomationControlled")

    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    driver.set_page_load_timeout(60)
    return driver


# ============================================================
# ユーティリティ
# ============================================================

DATA_EXTS = (".csv", ".xlsx", ".xls")

def snapshot_files(directory: Path):
    return {p for p in directory.glob("*") if p.suffix.lower() in DATA_EXTS}

def has_inprogress_downloads(directory: Path):
    return any(Path(p).suffix == ".crdownload" for p in glob.glob(str(directory / "*.crdownload")))

def is_file_stable(path: Path, checks: int = 4, interval: float = 0.5) -> bool:
    """ファイルサイズが連続 checks 回変化しない＝安定"""
    try:
        last = -1
        stable = 0
        for i in range(checks):
            if not path.exists():
                print(f"[DEBUG] ファイル不存在: {path.name}")
                return False
            size = path.stat().st_size
            if size == last and size > 0:
                stable += 1
            else:
                stable = 0
            last = size
            time.sleep(interval)
        result = stable >= (checks - 1)
        print(f"[DEBUG] 安定性チェック結果: {path.name} = {result} (size={size})")
        return result
    except Exception as e:
        print(f"[DEBUG] 安定性チェック例外: {path.name} - {e}")
        return False

def wait_for_new_downloads(before: set, directory: Path, timeout: int = 180):
    """新規ダウンロード完了を待つ（.crdownloadが消えたら即返す）"""
    deadline = time.time() + timeout
    found_new_file = False

    while time.time() < deadline:
        time.sleep(1)

        # .crdownloadファイルがあれば待機
        if has_inprogress_downloads(directory):
            if not found_new_file:
                print("ダウンロード中...")
                found_new_file = True
            continue

        # 新規ファイルを確認
        after = snapshot_files(directory)
        candidates = [p for p in after - before if p.exists()]

        if candidates:
            print(f"ダウンロード完了: {len(candidates)}件")
            return candidates

    print(f"タイムアウト: ダウンロードが完了しませんでした")
    return []

def safe_rename(src: Path, dst: Path, retries: int = 20, delay: float = 0.5) -> bool:
    """Windows ロック対策付きリネーム（上書き）"""
    for _ in range(retries):
        try:
            os.replace(str(src), str(dst))
            return True
        except PermissionError:
            time.sleep(delay)
        except FileNotFoundError:
            return False
        except Exception:
            time.sleep(delay)
    return False


# ============================================================
# メイン処理（グーネット）
# ============================================================

LOGIN_URL = "https://motorgate.jp/"
TARGET_URL = "https://motorgate.jp/ana/stockeffect"

TARGET_SHOPS = [
    {"value": "1000491", "name": "ハイエース専門店　ＣＡＲ　ＰＲＯＤＵＣＥ　｜　カープロデュース", "filename_prefix": "ハイエース専門店_", "wait_seconds": 5},
    {"value": "1002529", "name": "輸入車専門店　ＣＡＲＡＤ", "filename_prefix": "CARAD_", "wait_seconds": 10},
]

def login_goonet(driver, username: str, password: str):
    driver.get(LOGIN_URL)
    print(f"ログインページにアクセス: {driver.current_url}")

    wait = WebDriverWait(driver, 30)
    client_id_field = wait.until(EC.presence_of_element_located((By.ID, "client_id")))
    # パスワードは name="client_pw" のため name 指定
    password_field = driver.find_element(By.NAME, "client_pw")

    client_id_field.clear(); client_id_field.send_keys(username)
    password_field.clear();  password_field.send_keys(password)

    login_button = driver.find_element(By.ID, "button01")
    login_button.click()

    # ログイン後の URL 変化を待つ
    wait.until(EC.url_contains("/top"))
    print(f"ログイン成功: {driver.current_url}")

def trigger_download_for_shop(driver, shop_info: dict) -> bool:
    """指定店舗で検索→エクスポートボタンをクリック（ダウンロード待機なし）"""
    try:
        print(f"\n=== {shop_info['name']} のダウンロード開始 ===")
        driver.get(TARGET_URL)

        # 店舗選択
        try:
            shop_select = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.ID, "SelectGroupShop"))
            )
            Select(shop_select).select_by_value(shop_info["value"])
            print(f"店舗選択: {shop_info['name']}")
            time.sleep(0.8)
        except Exception as e:
            print(f"店舗選択に失敗: {e}")
            return False

        # 検索ボタン
        try:
            did_click = False
            for xp in [
                "//a[contains(@href, 'click_stock_search_btn')]",
                "//*[contains(@onclick, 'click_stock_search_btn')]",
                "//a[@href='javascript:click_stock_search_btn();']",
                "//*[contains(text(), '検索') and (self::a or self::button)]",
            ]:
                try:
                    btn = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, xp)))
                    btn.click()
                    print("検索ボタンをクリック")
                    did_click = True
                    break
                except Exception:
                    continue
            if not did_click:
                print("検索ボタンが見つかりませんでした")
                return False
            time.sleep(2)
        except Exception as e:
            print(f"検索ボタン操作でエラー: {e}")
            return False

        # エクスポートボタンをクリック
        try:
            export_button = None
            for xp in [
                "//*[contains(text(), '検索結果をエクスポート')]",
                "//*[contains(text(), 'エクスポート')]",
                "//a[contains(@class, 'export') or contains(@onclick, 'export')]",
                "//*[@id='export']",
            ]:
                try:
                    export_button = WebDriverWait(driver, 8).until(EC.element_to_be_clickable((By.XPATH, xp)))
                    break
                except Exception:
                    continue

            if not export_button:
                print("エクスポートボタンが見つかりませんでした")
                return False

            export_button.click()
            print("エクスポートボタンをクリック")

        except Exception as e:
            print(f"エクスポートボタン操作でエラー: {e}")
            return False

        return True

    except Exception as e:
        print(f"{shop_info['name']} の処理でエラー: {e}")
        print(traceback.format_exc())
        return False


def main():
    settings = load_settings()
    DOWNLOAD_DIR = Path(settings["DOWNLOAD_DIR"])
    HEADLESS = settings["HEADLESS"]
    USERNAME = settings["GOONET_USERNAME"]
    PASSWORD = settings["GOONET_PASSWORD"]

    driver = None
    try:
        print(f"DOWNLOAD_DIR: {DOWNLOAD_DIR}")
        driver = build_driver(DOWNLOAD_DIR, HEADLESS)

        # ログイン
        login_goonet(driver, USERNAME, PASSWORD)

        # ダウンロード前のファイル一覧を取得
        print("\n=== ダウンロード前のファイル確認 ===")
        before_files = snapshot_files(DOWNLOAD_DIR)
        print(f"既存ファイル数: {len(before_files)}")

        # 各店舗のダウンロードボタンを順番にクリック
        print("\n=== 各店舗のダウンロード処理開始 ===")
        for shop in TARGET_SHOPS:
            ok = trigger_download_for_shop(driver, shop)
            if ok:
                wait_time = shop.get("wait_seconds", 5)
                print(f"{shop['name']}: ダウンロードボタンをクリック完了")
                print(f"{wait_time}秒待機...")
                time.sleep(wait_time)
            else:
                print(f"{shop['name']}: ダウンロードボタンクリックに失敗")

        # 全ダウンロード完了後、新規ファイルを検出
        print("\n=== 新規ファイルの検出とリネーム処理 ===")
        after_files = snapshot_files(DOWNLOAD_DIR)
        new_files = [p for p in after_files - before_files if p.exists()]

        if not new_files:
            print("新規ファイルが見つかりませんでした")
        else:
            print(f"新規ファイル検出: {len(new_files)}件")
            # ダウンロード順（古い順）にソート：最初にダウンロードしたファイル = 最初の店舗
            new_files_sorted = sorted(new_files, key=lambda p: p.stat().st_mtime)

            # 各店舗に対応するファイルをリネーム
            # TARGET_SHOPS[0] = ハイエース専門店 → 最初のファイル（古い方）
            # TARGET_SHOPS[1] = CARAD → 2番目のファイル（新しい方）
            for i, shop in enumerate(TARGET_SHOPS):
                if i < len(new_files_sorted):
                    file_to_rename = new_files_sorted[i]
                    current_time = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                    new_name = f"{shop['filename_prefix']}{current_time}_{file_to_rename.name}"
                    dst = file_to_rename.with_name(new_name)

                    # リネーム実行
                    if dst.exists():
                        try:
                            dst.unlink()
                        except Exception:
                            pass
                    ok = safe_rename(file_to_rename, dst)
                    if ok:
                        print(f"リネーム完了: {shop['name']} ({file_to_rename.name}) -> {dst.name}")
                    else:
                        print(f"リネーム失敗: {file_to_rename.name}")

        print("\n=== 処理完了 ===")

    except Exception as e:
        print(f"メイン処理でエラー: {e}")
        print(traceback.format_exc())
    finally:
        if driver:
            driver.quit()
        print("ブラウザを閉じました")


if __name__ == "__main__":
    main()
