from suumo_scraper import login_to_fn_forrent, click_keisai_bukken_only, extract_properties
import gspread
from oauth2client.service_account import ServiceAccountCredentials

def update_sheet(properties):
    scope = ['https://spreadsheets.google.com/feeds',
             'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name('suumo-key.json', scope)
    client = gspread.authorize(creds)

    sheet = client.open('物件空室管理').worksheet('シート1')

    header = ['物件名', '部屋番号', '所在地', '最寄り駅', '賃料', '管理費', '間取り', '専有面積', '使わない列', 'URL']
    sheet.update([header], 'A1:J1')

    existing_records = sheet.get_all_records()

    # 比較用：物件名・部屋番号をトリミング＆文字列化
    def normalize(val):
        s = str(val or "").strip().lstrip("'")
        return s.zfill(4) if s.isdigit() else s

    scraped_keys = {(normalize(prop['物件名']), normalize(prop['部屋番号'])) for prop in properties}

    # --- 削除対象行の特定 ---
    rows_to_delete = []
    for idx, row in enumerate(existing_records, start=2):
        if not row['物件名']:
            continue  # 空行スキップ
        if not row['部屋番号']:
            continue  # 部屋番号が空欄なら削除しない（補完対象）

        key = (normalize(row['物件名']), normalize(row['部屋番号']))
        if key not in scraped_keys:
            rows_to_delete.append(idx)

    for row_idx in sorted(rows_to_delete, reverse=True):
        sheet.delete_rows(row_idx)
        print(f"🗑 スプレッドシートの行 {row_idx} を削除しました")

    # 最新のデータ取得（削除後）
    all_values = sheet.get_all_values()
    latest_records = sheet.get_all_records()
    latest_keys = {(normalize(row['物件名']), normalize(row['部屋番号'])) for row in latest_records}

    # 情報補完処理（空欄セルを補う）
    for idx, row in enumerate(latest_records, start=2):
        for prop in properties:
            if normalize(row['物件名']) == normalize(prop['物件名']) and \
               (not row['部屋番号'] or normalize(row['部屋番号']) == normalize(prop['部屋番号'])):

                columns = {
                    '所在地': ("C", prop['所在地']),
                    '最寄り駅': ("D", prop['最寄り駅']),
                    '賃料': ("E", prop['賃料']),
                    '管理費': ("F", prop['管理費']),
                    '間取り': ("G", prop['間取り']),
                    '専有面積': ("H", prop['専有面積']),
                    'URL': ("J", prop['URL']),
                }

                for key, (col, value) in columns.items():
                    existing_value = all_values[idx - 1][ord(col) - ord("A")].strip()
                    if not existing_value and value != "":
                        cell = f"{col}{idx}"
                        sheet.update(cell, [[value]])
                        print(f"✏️ {row['物件名']}（{row['部屋番号']}）の {key} を補完しました → {value}")

    # 新規物件の追加（存在しないもの）
    rows_to_add = []
    for prop in properties:
        key = (normalize(prop['物件名']), normalize(prop['部屋番号']))
        if key not in latest_keys:
            row = [
                prop['物件名'],
                prop['部屋番号'],
                prop['所在地'],
                prop['最寄り駅'],
                prop['賃料'],
                prop['管理費'],
                prop['間取り'],
                prop['専有面積'],
                '',  # 使わない列
                prop['URL']
            ]
            rows_to_add.append(row)

    if rows_to_add:
        last_row = len(sheet.get_all_values()) + 1
        sheet.insert_rows(rows_to_add, row=last_row)
        print(f"➕ {len(rows_to_add)} 件の新規物件を追加しました")

    print("✅ スプレッドシートの更新が完了しました")

if __name__ == "__main__":
    login_id = "f18481900101"
    password = "ban123t1a2k3"

    driver = login_to_fn_forrent(login_id, password)
    click_keisai_bukken_only(driver)

    properties = extract_properties(driver)
    driver.quit()

    update_sheet(properties)
