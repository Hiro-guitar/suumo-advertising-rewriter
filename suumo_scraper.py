# suumo_scraper.py

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import re

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
        import re
        # 正確に「数字＋m2」や「数字＋㎡」を含む行だけ対象にする
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
            # 面積だけ抽出（例: "1K\n22.96m2\nマンション\n相談" から22.96を抜き出す）
            lines = layout_text.split('\n')
            area_value = 0.0
            for line in lines:
                if "m2" in line or "㎡" in line:
                    area_value = extract_area(line)
                    break  # 最初にマッチしたら確定

            # 物件名＋部屋番号を分割（例: "ＨＯＰＥ　ＣＩＴＹ　秋葉原\n802号室" → 「物件名」「部屋番号」）
            parts = name.split('\n')
            property_name = parts[0].strip() if len(parts) > 0 else ""
            room_number = ""
            if len(parts) > 1:
                room_number = parts[1].replace("号室", "").strip()

            # 所在地・最寄駅はtd:nth-child(2)にある。例："ＪＲ山手線/秋葉原\n千代田区岩本町３"
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
                "使わない列": "",
                "URL": url,
            })

        except Exception as e:
            print("⚠️ 1物件の抽出エラー:", e)
            continue

    driver.switch_to.default_content()
    return properties
