import os
import json
import time
import glob
import shutil
import traceback
import selenium
from pathlib import Path

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from webdriver_manager.chrome import ChromeDriverManager

# ---- 設定の読み込み ---------------------------------------------------------

def load_settings():
    """ .env → settings.json → OS環境変数 の優先順で読み込み """
    settings = {}

    # .env 読み込み（存在すれば）
    try:
        from dotenv import load_dotenv
        env_path = Path(__file__).with_name(".env")
        if env_path.exists():
            load_dotenv(env_path)
    except Exception:
        pass  # dotenv 未インストールでも続行

    # settings.json 読み込み（存在すれば）
    settings_json = None
    for name in ("settings.json", "setting.json", "setteing.json"):  # タイプミスも一応拾う
        p = Path(__file__).with_name(name)
        if p.exists():
            settings_json = p
            break

    if settings_json:
        try:
            settings.update(json.loads(settings_json.read_text(encoding="utf-8")))
        except Exception as e:
            print(f"settings.json の読み込みに失敗しました: {e}")

    # 環境変数の値で上書き
    env_map = {
        "CARSENSOR_USERNAME": os.getenv("CARSENSOR_USERNAME"),
        "CARSENSOR_PASSWORD": os.getenv("CARSENSOR_PASSWORD"),
        "HEADLESS": os.getenv("HEADLESS"),
        "DOWNLOAD_DIR": os.getenv("DOWNLOAD_DIR"),
    }
    for k, v in env_map.items():
        if v is not None:
            settings[k] = v

    # 型整備
    headless = settings.get("HEADLESS", "false")
    if isinstance(headless, str):
        headless = headless.strip().lower() in ("1", "true", "yes", "on")
    settings["HEADLESS"] = bool(headless)

    # ダウンロード先（未指定なら OS のダウンロードフォルダ推定）
    dl = settings.get("DOWNLOAD_DIR")
    if dl:
        download_dir = Path(dl).expanduser()
    else:
        # Windows は内部的には "Downloads" フォルダ名。日本語環境の見た目は「ダウンロード」
        candidates = [
            Path.home() / "Downloads",
            Path.home() / "ダウンロード",  # 念のため
        ]
        download_dir = next((p for p in candidates if p.exists()), Path.home() / "Downloads")
    download_dir.mkdir(parents=True, exist_ok=True)
    settings["DOWNLOAD_DIR"] = str(download_dir)

    # 資格情報チェック
    if not settings.get("CARSENSOR_USERNAME") or not settings.get("CARSENSOR_PASSWORD"):
        raise RuntimeError("ID/PW が設定されていません。.env か settings.json に CARSENSOR_USERNAME / CARSENSOR_PASSWORD を設定してください。")

    return settings


# ---- WebDriver 構築 ---------------------------------------------------------

def build_driver(download_dir: Path, headless: bool):
    options = webdriver.ChromeOptions()
    if headless:
        # 新ヘッドレスを利用（Chrome 109+）
        options.add_argument("--headless=new")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--start-maximized")

    # ダウンロード設定
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


# ---- ユーティリティ ---------------------------------------------------------

DATA_EXTS = (".csv", ".xlsx", ".xls")

def snapshot_files(directory: Path):
    return {p for p in directory.glob("*") if p.suffix.lower() in DATA_EXTS}

def wait_for_new_downloads(before: set, directory: Path, timeout: int = 120):
    """新規ダウンロードファイル（.csv/.xlsx/.xls）を待つ。5秒待機して新規ファイルを返す"""
    time.sleep(5)
    after = snapshot_files(directory)
    new_files = [p for p in after - before if p.exists()]
    if new_files:
        print(f"ダウンロード完了: {len(new_files)}件")
        return new_files
    print("新規ファイルが見つかりませんでした")
    return []


# ---- メイン処理 -------------------------------------------------------------

def main():
    settings = load_settings()
    DOWNLOAD_DIR = Path(settings["DOWNLOAD_DIR"])
    HEADLESS = settings["HEADLESS"]
    username = settings["CARSENSOR_USERNAME"]
    password = settings["CARSENSOR_PASSWORD"]

    login_url = "https://c-match.carsensor.net/login/"
    target_url = "https://c-match.carsensor.net/counter/byVehicle/"

    driver = None
    try:
        print(f"ダウンロード先: {DOWNLOAD_DIR}")
        driver = build_driver(DOWNLOAD_DIR, HEADLESS)

        # 既存ファイルのスナップショット
        print("実行前のファイル状態を確認:")
        pre_files = snapshot_files(DOWNLOAD_DIR)
        print(f"既存ファイル数: {len(pre_files)}")

        # ログイン
        driver.get(login_url)
        print(f"ログインページにアクセス: {driver.current_url}")

        wait = WebDriverWait(driver, 30)
        username_field = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@name='loginId']")))
        password_field = driver.find_element(By.XPATH, "//input[@name='passwordCd']")
        username_field.clear(); username_field.send_keys(username)
        password_field.clear(); password_field.send_keys(password)

        login_button = driver.find_element(By.XPATH, "//input[@id='sbtLogin']")
        login_button.click()

        # ログイン完了待ち
        wait.until(EC.any_of(EC.url_contains("login=true"),
                             EC.url_contains("counter"),
                             EC.presence_of_element_located((By.XPATH, "//a|//button"))))
        print(f"ログイン成功: {driver.current_url}")

        # カウント対象ページへ
        driver.get(target_url)
        print(f"目的ページへ遷移: {driver.current_url}")
        time.sleep(3)

        # ダウンロードボタンを検出
        try:
            download_button = WebDriverWait(driver, 15).until(
                EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), 'ダウンロード')]"))
            )
            print(f"ダウンロードボタン検出: タグ={download_button.tag_name}")
        except Exception:
            # デバッグ情報出力
            all_links = driver.find_elements(By.TAG_NAME, "a")
            all_buttons = driver.find_elements(By.TAG_NAME, "button")
            print(f"リンク数: {len(all_links)}, ボタン数: {len(all_buttons)}")
            raise RuntimeError("ダウンロードボタンが見つかりませんでした。")

        # クリック前スナップショット
        before_main = snapshot_files(DOWNLOAD_DIR)

        # ★「直アクセス or クリック」どちらか1回だけ
        did_action = False
        if download_button.tag_name.lower() == "a":
            href = (download_button.get_attribute("href") or "").strip()
            if href and not href.lower().startswith("javascript"):
                print(f"href 直アクセスのみ実行: {href}")
                driver.get(href)
                did_action = True

        if not did_action:
            download_button.click()
            print("ダウンロードボタンをクリック（1回のみ）")

        # 可能なアラート処理
        time.sleep(1)
        try:
            alert = driver.switch_to.alert
            print(f"アラート検出: {alert.text}")
            alert.accept()
            print("アラート OK")
        except Exception:
            pass

        # ダウンロード完了待ち（★最新の1ファイルだけ採用）
        new_main_files = wait_for_new_downloads(before_main, DOWNLOAD_DIR, timeout=180)
        if new_main_files:
            new_main_files = sorted(new_main_files, key=lambda p: p.stat().st_mtime, reverse=True)[:1]
            print("ダウンロード完了（メイン）:")
            for p in new_main_files:
                print(f"- {p}")
        else:
            print("ダウンロードされたファイルが見つかりませんでした（メイン）。")

        # ===== ハイエース専門店のクリック処理 =====
        print("\n=== ハイエース専門店のクリック処理を開始 ===")
        try:
            # アラートが残っていれば処理
            try:
                alert = driver.switch_to.alert
                print(f"アラート検出: {alert.text}")
                alert.dismiss()
                time.sleep(1)
            except Exception:
                pass

            # 「他店舗参照」押下
            tatenpo = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.ID, "tatenpoBtn")))
            tatenpo.click()
            print("「他店舗参照」をクリック")
            time.sleep(2)

            # 「ハイエース専門店」を探してクリック（複数パターン）
            hiace_element = None
            for xp in [
                "//*[contains(text(), 'ハイエース専門店')]",
                "//h1[contains(text(), 'ハイエース専門店')]",
                "//*[contains(text(), 'CAR PRODUCE')]",
            ]:
                try:
                    hiace_element = WebDriverWait(driver, 8).until(EC.element_to_be_clickable((By.XPATH, xp)))
                    print(f"要素検出: {xp}")
                    break
                except Exception:
                    pass

            if not hiace_element:
                # ページ内テキスト確認
                body_text = driver.find_element(By.TAG_NAME, "body").text
                if "ハイエース" in body_text:
                    print("ページ内に『ハイエース』テキストは存在しますが、クリック可能要素が見つかりません。")
                else:
                    print("ページ内に『ハイエース』テキストが見つかりません。")
                raise RuntimeError("ハイエース専門店の要素が見つかりませんでした。")

            hiace_element.click()
            print("『ハイエース専門店』をクリック")
            time.sleep(2)

            # ハイエースページでのダウンロード
            try:
                download_button_hiace = WebDriverWait(driver, 15).until(
                    EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), 'ダウンロード')]"))
                )
                print(f"ハイエースページのダウンロードボタン検出: タグ={download_button_hiace.tag_name}")
            except Exception:
                all_links = driver.find_elements(By.TAG_NAME, "a")
                all_buttons = driver.find_elements(By.TAG_NAME, "button")
                print(f"(ハイエース) リンク数: {len(all_links)}, ボタン数: {len(all_buttons)}")
                raise RuntimeError("ハイエースページでダウンロードボタンが見つかりませんでした。")

            before_hiace = snapshot_files(DOWNLOAD_DIR)

            # ★「直アクセス or クリック」どちらか1回だけ
            did_action = False
            if download_button_hiace.tag_name.lower() == "a":
                href = (download_button_hiace.get_attribute("href") or "").strip()
                if href and not href.lower().startswith("javascript"):
                    print(f"(ハイエース) href 直アクセスのみ実行: {href}")
                    driver.get(href)
                    did_action = True

            if not did_action:
                download_button_hiace.click()
                print("(ハイエース) ダウンロードボタンをクリック（1回のみ）")

            time.sleep(1)
            try:
                alert = driver.switch_to.alert
                print(f"(ハイエース) アラート検出: {alert.text}")
                alert.accept()
                print("(ハイエース) アラート OK")
            except Exception:
                pass

            # ダウンロード完了待ち（★最新の1ファイルだけ採用）
            new_hiace_files = wait_for_new_downloads(before_hiace, DOWNLOAD_DIR, timeout=180)
            if new_hiace_files:
                new_hiace_files = sorted(new_hiace_files, key=lambda p: p.stat().st_mtime, reverse=True)[:1]
                print("ダウンロード完了（ハイエース）:")
                for p in new_hiace_files:
                    print(f"- {p}")
                    # ★コピーではなくリネーム（*_hiace へ）
                    name, ext = p.stem, p.suffix
                    hiace_renamed = p.with_name(f"{name}_hiace{ext}")
                    try:
                        if hiace_renamed.exists():
                            hiace_renamed.unlink()  # 同名があれば削除して置き換え
                        p.rename(hiace_renamed)
                        print(f"ハイエース用にリネーム: {hiace_renamed}")
                    except Exception as e:
                        print(f"リネームに失敗: {e}")
            else:
                print("ダウンロードされたファイルが見つかりませんでした（ハイエース）。")

        except Exception as e:
            print(f"ハイエース専門店処理でエラー: {e}")
            print(traceback.format_exc())

    except Exception as e:
        print(f"エラーが発生しました: {e}")
        print(traceback.format_exc())

    finally:
        if driver:
            driver.quit()
        print("処理を完了しました。")


if __name__ == "__main__":
    main()
