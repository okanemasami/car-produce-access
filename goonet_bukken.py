# -*- coding: utf-8 -*-
"""
グーネット（motorgate.jp）ローカル実行版
- 認証/設定は setting.json（なければ settig.json / settings.json）から読み込み
- DOWNLOAD_DIR にダウンロード（例: C:\\Users\\m-oka\\Downloads）
- HEADLESS 設定に従ってヘッドレス/通常表示を切替
- Selenium Manager（Selenium 4.6+）を優先、失敗時は webdriver-manager にフォールバック
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
from selenium.webdriver.common.action_chains import ActionChains

# ===================== 設定読み込み =====================
def load_settings():
    for name in ["setting.json", "settig.json", "settings.json"]:
        p = Path.cwd() / name
        if p.exists():
            with open(p, "r", encoding="utf-8") as f:
                data = json.load(f)
            print(f"設定ファイルを読み込みました: {p}")
            return data
    raise FileNotFoundError(
        "設定ファイルが見つかりません。実行フォルダに 'setting.json'（または 'settig.json' / 'settings.json'）を配置してください。"
    )

def to_bool(v, default=False):
    if isinstance(v, bool):
        return v
    if isinstance(v, str):
        return v.strip().lower() in ("1", "true", "yes", "on")
    return default

settings = load_settings()
username = os.getenv("GOONET_USERNAME") or settings.get("GOONET_USERNAME")
password = os.getenv("GOONET_PASSWORD") or settings.get("GOONET_PASSWORD")
download_dir_str = os.getenv("DOWNLOAD_DIR") or settings.get("DOWNLOAD_DIR")
# 環境変数 HEADLESS を優先（GitHub Actions用）
headless_env = os.getenv("HEADLESS")
if headless_env is not None:
    headless = to_bool(headless_env, default=True)
else:
    headless = to_bool(settings.get("HEADLESS", True), default=True)

if not username or not password:
    raise RuntimeError("setting.json に 'GOONET_USERNAME' と 'GOONET_PASSWORD' を設定してください。")
if not download_dir_str:
    raise RuntimeError("setting.json に 'DOWNLOAD_DIR' を設定してください。（例: C:\\\\Users\\\\m-oka\\\\Downloads）")

DOWNLOAD_DIR = Path(download_dir_str).expanduser().resolve()
DOWNLOAD_DIR.mkdir(parents=True, exist_ok=True)
download_path = str(DOWNLOAD_DIR)

# ===================== ユーティリティ =====================
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
    """15秒待機して新規ファイルを返す"""
    time.sleep(15)
    before_set = set(before_files)
    now = list_data_files(dir_path)
    new_files = [p for p in now if p not in before_set]
    if new_files:
        print(f"ダウンロード完了: {len(new_files)}件")
        return new_files
    print("新規ファイルが見つかりませんでした")
    return []

# ===================== Chrome 起動 =====================
options = webdriver.ChromeOptions()
if headless:
    options.add_argument("--headless=new")
options.add_argument("--disable-gpu")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
options.add_argument("--window-size=1920,1080")
# User-Agent を設定してヘッドレス検出を回避
options.add_argument("--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")

prefs = {
    "download.default_directory": download_path,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True,
    "profile.default_content_settings.popups": 0,
}
options.add_experimental_option("prefs", prefs)
options.add_experimental_option("excludeSwitches", ["enable-automation"])
options.add_experimental_option("useAutomationExtension", False)

driver = None
try:
    # Selenium Manager（推奨）
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

# ネットワークログを有効化
try:
    driver.execute_cdp_cmd('Network.enable', {})
except Exception:
    pass

# ===================== 対象URL =====================
login_url = "https://motorgate.jp/"
target_url = "https://motorgate.jp/group/stock/search"

try:
    # 実行前のファイル状態
    print("実行前のファイル状態を確認:")
    pre_files = list_data_files(DOWNLOAD_DIR)
    print(f"実行前に存在するCSV/Excel ファイル数: {len(pre_files)}")

    # --- ログイン ---
    driver.get(login_url)
    print(f"ログインページにアクセス: {driver.current_url}")

    # ログインフォーム要素
    client_id_field = WebDriverWait(driver, 30).until(
        EC.presence_of_element_located((By.ID, "client_id"))
    )
    password_field = driver.find_element(By.NAME, "client_pw")
    client_id_field.clear()
    client_id_field.send_keys(username)
    password_field.clear()
    password_field.send_keys(password)

    # ログインボタン
    login_button = driver.find_element(By.ID, "button01")
    login_button.click()

    # ログイン成功待ち（URL か body の遷移で判定）
    wait = WebDriverWait(driver, 30)
    try:
        wait.until(EC.url_contains("/top"))
    except Exception:
        wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
    print(f"ログイン後URL: {driver.current_url}")

    # --- 目的ページへ ---
    driver.get(target_url)
    print(f"検索ページへ遷移: {driver.current_url}")
    time.sleep(3)  # 画面描画待ち

    # 参考ログ
    links = driver.find_elements(By.TAG_NAME, "a")
    buttons = driver.find_elements(By.TAG_NAME, "button")
    print(f"リンク数: {len(links)} / ボタン数: {len(buttons)}")

    # --- エクスポート実行 ---
    # 1) CSS (li.export > a)
    export_link = None
    try:
        export_link = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "li.export > a"))
        )
        print(f"エクスポートリンク発見: テキスト={export_link.text} href={export_link.get_attribute('href')}")
    except Exception:
        print("CSS 'li.export > a' では見つからず → 代替手段へ")

    # 実行前ファイル一覧を保存
    before = list_data_files(DOWNLOAD_DIR)

    triggered = False

    # 優先順位1: リンク要素を直接クリック（最も確実）
    if export_link:
        try:
            print(f"エクスポートリンクを直接クリック試行...")
            # 邪魔な要素を非表示にする
            driver.execute_script("""
                var input = document.getElementById('ac1');
                if (input) input.style.display = 'none';
            """)
            time.sleep(0.5)

            # スクロールして要素を表示
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", export_link)
            time.sleep(1)

            # フォームデータを取得してHTTP POSTで直接リクエスト送信
            try:
                print("フォームデータを取得して直接POSTリクエストを送信...")

                # Cookieとフォームデータを取得
                cookies = driver.get_cookies()
                cookie_dict = {cookie['name']: cookie['value'] for cookie in cookies}

                # フォームデータを全て取得
                form_data = driver.execute_script("""
                    var formData = {};
                    var form = document.getElementById('frm');
                    if (!form) return null;

                    // すべてのinput, select, textarea要素を取得
                    var inputs = form.querySelectorAll('input, select, textarea');
                    inputs.forEach(function(input) {
                        if (input.name) {
                            if (input.type === 'checkbox' || input.type === 'radio') {
                                if (input.checked) {
                                    formData[input.name] = input.value;
                                }
                            } else {
                                formData[input.name] = input.value;
                            }
                        }
                    });

                    // export_flgを追加
                    formData['export_flg'] = '1';

                    return formData;
                """)

                if form_data is None:
                    print("エラー: フォームが見つかりません")
                    triggered = False
                else:
                    print(f"取得したフォームデータ項目数: {len(form_data)}")

                    # requestsライブラリでPOST送信
                    import requests

                    csv_url = "https://motorgate.jp/group/stock/search/csv"
                    headers = {
                        'User-Agent': driver.execute_script("return navigator.userAgent;"),
                        'Referer': 'https://motorgate.jp/group/stock/search',
                        'Origin': 'https://motorgate.jp'
                    }

                    print(f"POSTリクエスト送信先: {csv_url}")
                    response = requests.post(csv_url, data=form_data, cookies=cookie_dict, headers=headers)

                    print(f"レスポンスステータス: {response.status_code}")
                    print(f"Content-Type: {response.headers.get('Content-Type', 'N/A')}")

                    if response.status_code == 200:
                        # CSVファイルとして保存
                        from datetime import datetime
                        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                        filename = os.path.join(DOWNLOAD_DIR, f"goonet_bukken_{timestamp}.csv")

                        with open(filename, 'wb') as f:
                            f.write(response.content)

                        print(f"CSVファイルを保存しました: {filename}")
                        print(f"ファイルサイズ: {len(response.content)} bytes")
                        triggered = True
                    else:
                        print(f"エラー: ステータスコード {response.status_code}")
                        print(f"レスポンス本文（最初の500文字）: {response.text[:500]}")
                        triggered = False

            except Exception as e_post:
                print(f"POSTリクエストエラー: {e_post}")
                import traceback
                traceback.print_exc()
                triggered = False
        except Exception as e1:
            print(f"通常のクリックでエラー: {e1}")
            # JavaScriptでクリック
            try:
                driver.execute_script("arguments[0].click();", export_link)
                print("エクスポートリンクをクリックしました（JS経由）")
                time.sleep(20)
                triggered = True
            except Exception as e2:
                print(f"JS経由のクリックでもエラー: {e2}")

    # 優先順位2: JS 関数 excel() の直接実行（フォールバック）
    if not triggered:
        try:
            excel_exists = driver.execute_script("return typeof excel === 'function';")
            print(f"excel() 関数の存在確認: {excel_exists}")

            driver.execute_script("excel();")
            print("JavaScript 関数 excel() を実行しました")
            time.sleep(15)
            triggered = True
        except Exception as e:
            print(f"excel() 実行でエラー: {e}")

    # 4) さらに失敗時は “エクスポート” テキスト検索
    if not triggered:
        try:
            export_links = driver.find_elements(By.XPATH, "//a[contains(text(), 'エクスポート')]")
            if export_links:
                el = export_links[0]
                href = el.get_attribute("href") or ""
                if "javascript:excel()" in href.replace(" ", ""):
                    driver.execute_script("excel();")
                    print("excel() を再実行しました")
                else:
                    driver.execute_script("arguments[0].click();", el)
                    print("“エクスポート” リンクをクリック（テキスト一致）")
                triggered = True
        except Exception as e:
            print(f"代替テキスト検索でもエラー: {e}")

    if not triggered:
        raise RuntimeError("エクスポート操作を開始できませんでした。画面構造の変更が疑われます。")

    # 直接POSTでダウンロードした場合はここでの待機は不要
    if not triggered:
        # アラートが出る場合に備えてハンドリング
        try:
            alert = driver.switch_to.alert
            print(f"ダウンロード時のアラート: {alert.text}")
            alert.accept()
            print("アラート OK")
        except Exception:
            pass

        # --- ダウンロード完了待機 ---
        print("ダウンロード完了待機中...")
        print(f"ダウンロードディレクトリ: {DOWNLOAD_DIR}")
        print(f"実行前ファイル数: {len(before)}")

        new_files = wait_for_download(before, DOWNLOAD_DIR, timeout=120)
        if new_files:
            print(f"ダウンロードされたファイル数: {len(new_files)}")
            for nf in new_files:
                print(f"  - {nf}")
        else:
            print("ダウンロードされたファイルが見つかりませんでした。")
            current_files = list_data_files(DOWNLOAD_DIR)
            print(f"現在のファイル数（デバッグ用）: {len(current_files)}")
            print("現在のファイル一覧:")
            for cf in current_files:
                print(f"  - {cf}")

except Exception as e:
    print(f"エラーが発生しました: {e}")
    import traceback
    print(traceback.format_exc())

finally:
    try:
        driver.quit()
    except Exception:
        pass
    print(f"処理を完了しました（ダウンロード先: {download_path} / ヘッドレス: {headless}）")
