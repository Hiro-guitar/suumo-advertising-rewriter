# suumo_scraper.py

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import re

# 空室確認URLとして認識するドメイン
KNOWN_VACANCY_DOMAINS = ['itandibb.com', 'es-square.net', 'bb.ielove.jp']

def login_to_fn_forrent(login_id: str, password: str):
    options = Options()
    options.add_argument("--headless")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    driver = webdriver.Chrome(options=options)

    driver.get("https://www.fn.forrent.jp/fn/")
    time.sleep(2)

    login_input = driver.find_element(By.XPATH, '//input[@type="text"]')
    password_input = driver.find_element(By.XPATH, '//input[@type="password"]')

    login_input.send_keys(login_id)
    password_input.send_keys(password)

    driver.find_element(By.ID, "Image7").click()
    print("✅ ログイン実行")
    time.sleep(3)

    WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it("navi"))
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "menu_3"))).click()
    print("✅ 掲載指示ボタンをクリックしました")

    driver.switch_to.default_content()
    time.sleep(3)
    print("✅ 現在のURL:", driver.current_url)

    return driver

def click_keisai_bukken_only(driver):
    WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it("main"))
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.LINK_TEXT, "掲載物件のみ"))).click()
    print("✅ 「掲載物件のみ」リンクをクリックしました")
    time.sleep(3)
    driver.switch_to.default_content()

def extract_properties(driver):
    def convert_manegement_fee_to_yen(mng_fee_str: str) -> int:
        try:
            if "万円" in mng_fee_str:
                num = float(mng_fee_str.replace("万円", "").strip())
                return int(num * 10000)
            else:
                digits = re.findall(r"\d+\.?\d*", mng_fee_str)
                if digits:
                    return int(float(digits[0]))
                return 0
        except:
            return 0

    def extract_area(area_str: str) -> float:
        m = re.search(r"(\d+(?:\.\d+)?)\s*(?:m2|㎡)", area_str)
        if m:
            return float(m.group(1))
        return 0.0

    WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it("main"))
    time.sleep(2)

    rows = driver.find_elements(By.CSS_SELECTOR, "tr.cnet_base32.keisaiHeight, tr.cnet_base32.keisaiHeightBg")

    properties = []
    for row in rows:
        try:
            name_cell = row.find_element(By.CSS_SELECTOR, "td:nth-child(4)")
            name = name_cell.text.strip()

            rent_cell = row.find_element(By.CSS_SELECTOR, "td:nth-child(5)")
            rent = rent_cell.find_element(By.CSS_SELECTOR, "span.bold").text.strip()
            rent_text = rent_cell.text.strip().split('\n')
            management_fee_str = rent_text[1] if len(rent_text) > 1 else ""
            management_fee_yen = convert_manegement_fee_to_yen(management_fee_str)

            layout_cell = row.find_element(By.CSS_SELECTOR, "td:nth-child(6)")
            layout_text = layout_cell.text.strip()
            lines = layout_text.split('\n')
            area_value = 0.0
            for line in lines:
                if "m2" in line or "㎡" in line:
                    area_value = extract_area(line)
                    break

            parts = name.split('\n')
            property_name = parts[0].strip() if len(parts) > 0 else ""
            room_number = ""
            if len(parts) > 1:
                room_number = parts[1].replace("号室", "").strip()

            loc_cell = row.find_element(By.CSS_SELECTOR, "td:nth-child(2)")
            loc_text = loc_cell.text.strip().split('\n')
            nearest_station = loc_text[0].strip() if len(loc_text) > 0 else ""
            address = loc_text[1].strip() if len(loc_text) > 1 else ""

            url = ""
            try:
                suumo_img = row.find_element(By.CSS_SELECTOR, "img.mk_suumopc")
                onclick_js = suumo_img.get_attribute("onclick")
                if onclick_js:
                    m = re.search(r"openPcSite\('([^']+)'\)", onclick_js)
                    if m:
                        url = m.group(1)
            except:
                pass

            properties.append({
                "物件名": property_name,
                "部屋番号": room_number,
                "所在地": address,
                "最寄り駅": nearest_station,
                "賃料": float(rent.replace("万円","").strip()),
                "管理費": management_fee_yen,
                "間取り": lines[0] if len(lines) > 0 else "",
                "専有面積": area_value,
                "URL": url,
                "空室確認URL": "",  # extract_free_comment_urls() で後から埋める
            })

        except Exception as e:
            print("⚠️ 1物件の抽出エラー:", e)
            continue

    driver.switch_to.default_content()
    return properties


def extract_free_comment_urls(driver, properties):
    """
    各物件の詳細ページを訪問し、「フリーコメント※100文字以内」欄から
    掲載元URL（itandiBB / いい生活スクエア / いえらぶBB）を取得する。

    properties リストの各要素の "空室確認URL" キーを更新する。
    """
    if not properties:
        return

    print(f"\n📋 {len(properties)} 件の詳細ページからフリーコメントURLを取得中...")

    for i in range(len(properties)):
        vacancy_url = ""
        prop_name = properties[i].get('物件名', f'Row {i}')

        try:
            # mainフレームに切り替え（一覧ページが読み込まれている状態）
            WebDriverWait(driver, 10).until(
                EC.frame_to_be_available_and_switch_to_it("main"))
            time.sleep(1)

            # 行を再取得（前回のback()で一覧ページに戻っているはず）
            rows = driver.find_elements(
                By.CSS_SELECTOR,
                "tr.cnet_base32.keisaiHeight, tr.cnet_base32.keisaiHeightBg")

            if i >= len(rows):
                print(f"  ⚠️ [{i}] {prop_name}: 一覧の行が見つかりません（rows={len(rows)}）")
                properties[i]["空室確認URL"] = ""
                driver.switch_to.default_content()
                continue

            # 物件名リンクをクリックして詳細ページへ
            row = rows[i]
            name_cell = row.find_element(By.CSS_SELECTOR, "td:nth-child(4)")
            try:
                link = name_cell.find_element(By.TAG_NAME, "a")
                link.click()
            except Exception:
                # リンクがない場合はスキップ
                print(f"  ➖ {prop_name}: 詳細リンクなし")
                properties[i]["空室確認URL"] = ""
                driver.switch_to.default_content()
                continue

            time.sleep(3)

            # 詳細ページ内からURLを抽出（まだmainフレーム内）
            vacancy_url = _find_vacancy_url_in_page(driver)

            # 一覧ページに戻る
            driver.back()
            time.sleep(3)
            driver.switch_to.default_content()

        except Exception as e:
            print(f"  ⚠️ {prop_name}: 詳細ページエラー: {e}")
            try:
                driver.switch_to.default_content()
            except:
                pass

        properties[i]["空室確認URL"] = vacancy_url
        status = '✅' if vacancy_url else '➖'
        print(f"  {status} {prop_name}: {vacancy_url or '(URLなし)'}")


def _find_vacancy_url_in_page(driver):
    """
    詳細ページ内から掲載元URL（itandiBB / いい生活スクエア / いえらぶBB）を抽出する。

    Strategy:
    1. 「フリーコメント」ラベル周辺のテキストからURLを探す
    2. 見つからなければページ全体のテキストから既知ドメインのURLを探す
    """
    try:
        # Strategy 1: フリーコメント欄の周辺テキストからURL抽出
        try:
            elements = driver.find_elements(
                By.XPATH, "//*[contains(text(), 'フリーコメント')]")
            for elem in elements:
                # 要素自身と親要素のテキストを確認
                for target in _get_element_and_ancestors(elem, levels=3):
                    text = target.text
                    url = _extract_known_domain_url(text)
                    if url:
                        return url
        except Exception:
            pass

        # Strategy 2: ページ全体から既知ドメインURLを検索
        try:
            body_text = driver.find_element(By.TAG_NAME, "body").text
            url = _extract_known_domain_url(body_text)
            if url:
                return url
        except Exception:
            pass

        return ""

    except Exception as e:
        print(f"    ⚠️ URL抽出エラー: {e}")
        return ""


def _get_element_and_ancestors(elem, levels=3):
    """要素自身と上位N階層の親要素を返すジェネレータ"""
    yield elem
    current = elem
    for _ in range(levels):
        try:
            current = current.find_element(By.XPATH, "./..")
            yield current
        except Exception:
            break


def _extract_known_domain_url(text):
    """テキストから既知ドメインのURLを抽出（クエリパラメータ除去済み）"""
    urls = re.findall(r'https?://[^\s<>"\'。、）\)\]]+', text)
    for url in urls:
        # クエリパラメータを除去（いい生活スクエアのURL対策: 100文字制限）
        clean = url.split('?')[0].rstrip('/')
        if any(domain in clean for domain in KNOWN_VACANCY_DOMAINS):
            return clean
    return ""
