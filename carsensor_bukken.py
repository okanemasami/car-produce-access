# -*- coding: utf-8 -*-
"""
ローカルPC用 修正版（多重ダウンロード防止）
- 認証/保存先は setting.json（なければ settig.json / settings.json）から読み込み
- DOWNLOAD_DIR に保存（例: C:\\Users\\m-oka\\Downloads）
- 一度だけダウンロードを発火する start_download_once() を実装
"""

import os
import json
import time
import glob
from pathlib import Path

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# ===== 設定読込 =====
def load_settings():
    candidates = ["setting.json", "settig.json", "settings.json"]
    for name in candidates:
        p = Path.cwd() / name
        if p.exists():
            with open(p, "r", encoding="utf-8") as f:
                data = json.load(f)
            print(f"設定ファイルを読み込みました: {p}")
            return data
    raise FileNotFoundError(
        "設定ファイルが見つかりませんでした。'setting.json' もしくは 'settig.json' / 'settings.json' を実行フォルダに置いてください。"
    )

settings = load_settings()

username = settings.get("CARSENSOR_USERNAME")
password = settings.get("CARSENSOR_PASSWORD")
download_dir_str = settings.get("DOWNLOAD_DIR")

if not username or not password:
    raise RuntimeError("setting.json に 'CARSENSOR_USERNAME' と 'CARSENSOR_PASSWORD' を設定してください。")
if not download_dir_str:
    raise RuntimeError("setting.json に 'DOWNLOAD_DIR' を設定してください。（例: C:\\\\Users\\\\m-oka\\\\Downloads）")

DOWNLOAD_DIR = Path(download_dir_str).expanduser().resolve()
DOWNLOAD_DIR.mkdir(parents=True, exist_ok=True)
download_path = str(DOWNLOAD_DIR)

# ===== ユーティリティ =====
def list_data_files(root_dir: Path):
    exts = {".csv", ".xlsx", ".xls"}
    files = []
    for r, d, fs in os.walk(str(root_dir)):
        for f in fs:
            p = Path(r) / f
            if p.suffix.lower() in exts:
                files.append(str(p))
    return files

def wait_for_download(before_files, dir_path: Path, timeout=90):
    """5秒待機して新規ファイルを返す"""
    time.sleep(5)
    before_set = set(before_files)
    now = list_data_files(dir_path)
    new_files = [p for p in now if p not in before_set]
    if new_files:
        print(f"ダウンロード完了: {len(new_files)}件")
        return new_files
    print("新規ファイルが見つかりませんでした")
    return []

def handle_alert_if_present(driver):
    """アラートがあれば処理（OK押下）"""
    try:
        alert = driver.switch_to.alert
        print(f"アラート: {alert.text}")
        alert.accept()
        print("アラート OK")
        time.sleep(0.5)
        return True
    except Exception:
        return False

def start_download_once(driver, locator, before_files, dir_path: Path, trigger_wait=6):
    """ダウンロードボタンをクリック（シンプル版）"""
    try:
        elem = WebDriverWait(driver, 20).until(EC.element_to_be_clickable(locator))
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", elem)
        time.sleep(0.2)
        elem.click()
        print("ダウンロードボタンをクリックしました")
        handle_alert_if_present(driver)
        return True
    except Exception as e:
        print(f"ダウンロードボタンクリック失敗: {e}")
        return False

# ===== Chrome オプション =====
options = webdriver.ChromeOptions()
# 必要なら下の1行をコメントアウトしてブラウザ表示
options.add_argument("--headless=new")
options.add_argument("--disable-gpu")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
options.add_argument("--window-size=1920,1080")

prefs = {
    "download.default_directory": download_path,   # JSONの DOWNLOAD_DIR
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True,
}
options.add_experimental_option("prefs", prefs)

# ===== WebDriver 準備（Selenium Manager → 失敗時 webdriver-manager）=====
driver = None
try:
    driver = webdriver.Chrome(options=options)
except Exception as e:
    print("Selenium Manager での起動に失敗。webdriver-manager を試します:", e)
    try:
        from webdriver_manager.chrome import ChromeDriverManager
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    except Exception as e2:
        raise RuntimeError(f"ChromeDriver の起動に失敗しました: {e2}")

# ヘッドレス時のダウンロード許可（未対応版は無視）
try:
    driver.execute_cdp_cmd("Page.setDownloadBehavior", {"behavior": "allow", "downloadPath": download_path})
except Exception:
    pass

# ===== ターゲット URL =====
login_url = "https://c-match.carsensor.net/login/"
target_url = "https://c-match.carsensor.net/vehicles/registrationList/"

try:
    # 実行前のファイル状態
    print("実行前のファイル状態を確認:")
    pre_files = list_data_files(DOWNLOAD_DIR)
    try:
        print(f"実行前に存在するCSV/Excel ファイル数: {len(pre_files)}")
    except UnicodeEncodeError:
        print(f"実行前に存在するCSV/Excel ファイル数: {len(pre_files)}")

    # --- ログイン ---
    driver.get(login_url)
    print(f"ログインページにアクセスしました: {driver.current_url}")

    username_field = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.XPATH, "//input[@name='loginId']"))
    )
    password_field = driver.find_element(By.XPATH, "//input[@name='passwordCd']")
    username_field.send_keys(username)
    password_field.send_keys(password)

    login_button = driver.find_element(By.XPATH, "//input[@id='sbtLogin']")
    login_button.click()

    wait = WebDriverWait(driver, 20)
    try:
        wait.until(EC.url_contains("login=true"))
    except Exception:
        wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
    print(f"ログイン後URL: {driver.current_url}")

    # --- 対象ページへ ---
    driver.get(target_url)
    print(f"目的のページに移動しました: {driver.current_url}")
    time.sleep(2)

    # ダウンロードボタン（最初のページ）を一度だけトリガー
    before = list_data_files(DOWNLOAD_DIR)
    print("最初のページでのダウンロードを開始します（単発トリガー制御）。")
    started = start_download_once(
        driver,
        (By.XPATH, "//*[contains(text(), 'ダウンロード')]"),
        before,
        DOWNLOAD_DIR,
        trigger_wait=6
    )
    if not started:
        raise RuntimeError("最初のページでダウンロード開始を検知できませんでした。")

    print("ダウンロードの完了を待機中...")
    new_files = wait_for_download(before, DOWNLOAD_DIR, timeout=90)
    if new_files:
        print(f"ダウンロードされたファイル数: {len(new_files)}")
    else:
        print("ダウンロードされたファイルが見つかりませんでした")
        print(f"現在のファイル数（デバッグ用）: {len(list_data_files(DOWNLOAD_DIR))}")

    # ===== 「他店舗参照」→「ハイエース専門店」での2回目DL（こちらも単発トリガー）=====
    print("\n=== ハイエース専門店のクリック処理を開始 ===")

    try:
        handle_alert_if_present(driver)

        print("「他店舗参照」ボタンを検索中...")
        tatenpobtn = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.ID, "tatenpoBtn"))
        )
        tatenpobtn.click()
        print("「他店舗参照」ボタンをクリックしました")
        time.sleep(2)

        print("「ハイエース専門店」を含む要素を検索中...")
        hiace_element = None
        # 方法1
        try:
            hiace_element = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), 'ハイエース専門店')]"))
            )
            print("方法1でハイエース専門店の要素を発見しました")
        except Exception:
            print("方法1では見つかりませんでした")
        # 方法2
        if hiace_element is None:
            try:
                hiace_element = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.XPATH, "//h1[contains(text(), 'ハイエース専門店')]"))
                )
                print("方法2でハイエース専門店の要素を発見しました")
            except Exception:
                print("方法2では見つかりませんでした")
        # 方法3
        if hiace_element is None:
            try:
                hiace_element = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), 'CAR PRODUCE')]"))
                )
                print("方法3でCAR PRODUCEの要素を発見しました")
            except Exception:
                print("方法3では見つかりませんでした")

        if hiace_element:
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", hiace_element)
            hiace_element.click()
            print("ハイエース専門店の要素をクリックしました")
            time.sleep(2)

            print("\n=== ハイエース専門店ページでのダウンロード処理（単発トリガー）開始 ===")
            pre_hiace = list_data_files(DOWNLOAD_DIR)
            started2 = start_download_once(
                driver,
                (By.XPATH, "//*[contains(text(), 'ダウンロード')]"),
                pre_hiace,
                DOWNLOAD_DIR,
                trigger_wait=6
            )
            if started2:
                print("ハイエース専門店ページのダウンロード完了待機...")
                new_hiace = wait_for_download(pre_hiace, DOWNLOAD_DIR, timeout=90)
                if new_hiace:
                    print(f"ハイエース専門店でダウンロードされたファイル数: {len(new_hiace)}")
                else:
                    print("ハイエース専門店でダウンロードされたファイルが見つかりませんでした")
            else:
                print("ハイエース専門店ページでダウンロード開始を検知できませんでした")
        else:
            print("ハイエース専門店の要素が見つかりませんでした")
            body_text = driver.find_element(By.TAG_NAME, "body").text
            if "ハイエース" in body_text:
                print("ページにはハイエースというテキストが含まれています")
            else:
                print("ページにハイエースというテキストが見つかりません")
    except Exception as e:
        print(f"ハイエース専門店のクリック処理でエラーが発生しました: {e}")
        import traceback
        print(traceback.format_exc())

except Exception as e:
    print(f"エラーが発生しました: {e}")
    import traceback
    print(traceback.format_exc())

finally:
    try:
        driver.quit()
    except Exception:
        pass
    print("処理を完了しました（ダウンロード先: " + download_path + "）")
