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
# メイン処理（グーネット問い合わせ）
# ============================================================

LOGIN_URL = "https://motorgate.jp/"
INQUIRY_URL = "https://motorgate.jp/inquiry/est/search"

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

def get_current_month_range():
    """現在月の1日と末日を返す"""
    now = datetime.datetime.now()
    first_day = now.replace(day=1)
    # 翌月の1日から1日引いて当月末日を取得
    next_month = first_day.replace(day=28) + datetime.timedelta(days=4)
    last_day = next_month - datetime.timedelta(days=next_month.day)
    return first_day, last_day

def configure_search_form(driver):
    """検索フォームを設定（ステータス「すべて」、問い合わせ期間）"""
    wait = WebDriverWait(driver, 15)

    # 1. ステータス「すべて」を選択
    # name="s_status", value="99", id="check01_05", onclick="status_change(this.value);"
    print("ステータス「すべて」を選択...")
    try:
        status_radio = driver.find_element(By.ID, "check01_05")
        # ラジオボタンをチェック状態にする
        driver.execute_script("arguments[0].checked = true;", status_radio)
        # onclick イベントハンドラを直接呼び出す
        driver.execute_script("status_change('99');")
        print("ステータス「すべて」を選択しました (ID: check01_05, status_change呼び出し)")

        # ステータス変更後、URL遷移が完了するまで待機
        print("URL遷移を待機中...")
        WebDriverWait(driver, 10).until(
            lambda d: "s_status_value=99" in d.current_url or "s_status_value=20" in d.current_url
        )
        print(f"URL遷移完了: {driver.current_url}")

        # ページの再レンダリング完了を待つ
        print("ページ安定化待機...")
        time.sleep(2)
    except Exception as e:
        print(f"警告: ステータス「すべて」の選択に失敗: {e}")

    # 2. 問い合わせ期間のラジオボタンをクリック（チェックボックスではない！）
    # name="s_est_date", value="1", id="radio02_01", onclick="on_est_date_radio_change(this.value);"
    print("問い合わせ期間のラジオボタンをクリック...")
    try:
        date_radio = driver.find_element(By.ID, "radio02_01")
        # ラジオボタンをチェック状態にする
        driver.execute_script("arguments[0].checked = true;", date_radio)
        # onclick イベントハンドラを直接呼び出す
        driver.execute_script("on_est_date_radio_change('1');")
        print("問い合わせ期間ラジオボタンをクリックしました (ID: radio02_01, on_est_date_radio_change呼び出し)")
    except Exception as e:
        print(f"エラー: 問い合わせ期間ラジオボタンが見つかりません: {e}")
        raise

    # ラジオボタンクリック後、JavaScriptが日付フィールドを有効化するまで待機
    print("日付フィールドが有効化されるまで待機...")
    time.sleep(2)

    # 3. 開始日・終了日を設定（jQuery UI Datepicker対策）
    first_day, last_day = get_current_month_range()
    print(f"期間設定: {first_day.strftime('%Y/%m/%d')} 〜 {last_day.strftime('%Y/%m/%d')}")

    # 開始日フィールド: name="s_est_date_from"
    try:
        start_field = driver.find_element(By.NAME, "s_est_date_from")
        # disabled 属性が外れていることを確認
        is_disabled = start_field.get_attribute("disabled")
        if is_disabled:
            print(f"警告: 開始日フィールドがまだ無効化されています (disabled={is_disabled})")
            time.sleep(1)  # 追加待機

        # jQuery UI Datepicker対策: JavaScriptで直接値を設定
        start_date_str = first_day.strftime("%Y/%m/%d")
        driver.execute_script("arguments[0].value = arguments[1];", start_field, start_date_str)
        # datepickerのchangeイベントを発火
        driver.execute_script("$(arguments[0]).trigger('change');", start_field)
        print(f"開始日を設定しました: {start_date_str}")
    except Exception as e:
        print(f"エラー: 開始日フィールドの設定に失敗: {e}")
        raise

    # 終了日フィールド: name="s_est_date_to"
    try:
        end_field = driver.find_element(By.NAME, "s_est_date_to")
        # disabled 属性が外れていることを確認
        is_disabled = end_field.get_attribute("disabled")
        if is_disabled:
            print(f"警告: 終了日フィールドがまだ無効化されています (disabled={is_disabled})")
            time.sleep(1)  # 追加待機

        # jQuery UI Datepicker対策: JavaScriptで直接値を設定
        end_date_str = last_day.strftime("%Y/%m/%d")
        driver.execute_script("arguments[0].value = arguments[1];", end_field, end_date_str)
        # datepickerのchangeイベントを発火
        driver.execute_script("$(arguments[0]).trigger('change');", end_field)
        print(f"終了日を設定しました: {end_date_str}")
    except Exception as e:
        print(f"エラー: 終了日フィールドの設定に失敗: {e}")
        raise

def trigger_download(driver):
    """検索とダウンロードを実行"""
    # 1. 検索を実行（JavaScriptで直接関数を呼び出す）
    # <a href="JavaScript:mode_change('');search();">検索</a>
    print("検索を実行...")
    try:
        # mode_change('') と search() を直接呼び出し
        driver.execute_script("mode_change('');")
        driver.execute_script("search();")
        print("検索を実行しました (JavaScript直接呼び出し)")
    except Exception as e:
        print(f"エラー: 検索実行に失敗: {e}")
        raise

    # 2. 検索結果の読み込み待機
    print("検索結果の読み込み待機...")
    time.sleep(5)

    # 3. エクスポートボタンをクリック
    print("エクスポートボタンをクリック...")
    export_clicked = False
    for xpath in [
        "//button[contains(text(), 'エクスポート')]",
        "//a[contains(text(), 'エクスポート')]",
        "//*[@id='export']",
        "//button[contains(@class, 'export')]",
        "//a[contains(@onclick, 'export')]",
    ]:
        try:
            export_button = WebDriverWait(driver, 15).until(
                EC.element_to_be_clickable((By.XPATH, xpath))
            )
            export_button.click()
            print("エクスポートボタンをクリックしました")
            export_clicked = True
            break
        except Exception:
            continue

    if not export_clicked:
        raise RuntimeError("エクスポートボタンが見つかりませんでした")


def main():
    settings = load_settings()
    DOWNLOAD_DIR = Path(settings["DOWNLOAD_DIR"])
    HEADLESS = settings["HEADLESS"]
    USERNAME = settings["GOONET_USERNAME"]
    PASSWORD = settings["GOONET_PASSWORD"]

    driver = None
    try:
        print(f"DOWNLOAD_DIR: {DOWNLOAD_DIR}")
        print(f"HEADLESS: {HEADLESS}")
        driver = build_driver(DOWNLOAD_DIR, HEADLESS)

        # ダウンロード前のファイル確認
        print("\n=== ダウンロード前のファイル確認 ===")
        before_files = snapshot_files(DOWNLOAD_DIR)
        print(f"既存ファイル数: {len(before_files)}")

        # ログイン
        print("\n=== ログイン処理 ===")
        login_goonet(driver, USERNAME, PASSWORD)

        # 問い合わせ検索ページへ遷移
        print(f"\n=== 問い合わせ検索ページへ遷移 ===")
        driver.get(INQUIRY_URL)
        print(f"現在のURL: {driver.current_url}")
        time.sleep(3)

        # 検索フォーム設定
        print("\n=== 検索フォーム設定 ===")
        configure_search_form(driver)

        # ダウンロード実行
        print("\n=== ダウンロード実行 ===")
        trigger_download(driver)

        # ダウンロード完了待機
        print("\n=== ダウンロード完了待機 ===")
        new_files = wait_for_new_downloads(before_files, DOWNLOAD_DIR, timeout=180)

        # ファイル名変更
        if new_files:
            print(f"\n=== ファイル名変更処理 ===")
            for file_path in new_files:
                timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                new_name = f"goonet_inquiry_{timestamp}_{file_path.name}"
                dst = file_path.with_name(new_name)

                # 既存ファイルがあれば削除
                if dst.exists():
                    try:
                        dst.unlink()
                    except Exception:
                        pass

                ok = safe_rename(file_path, dst)
                if ok:
                    print(f"リネーム完了: {file_path.name} -> {dst.name}")
                else:
                    print(f"リネーム失敗: {file_path.name}")
        else:
            print("\n新規ファイルが見つかりませんでした")

        print("\n=== 処理完了 ===")

    except Exception as e:
        print(f"\nメイン処理でエラー: {e}")
        print(traceback.format_exc())
    finally:
        if driver:
            driver.quit()
        print("ブラウザを閉じました")


if __name__ == "__main__":
    main()
