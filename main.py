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

    def normalize(val):
        s = str(val or "").strip().lstrip("'")
        return s.zfill(4) if s.isdigit() else s

    scraped_keys = {(normalize(p['物件名']), normalize(p['部屋番号'])) for p in properties}

    # 1. 削除対象（シートにあるがスクレイピングにない行）
    rows_to_delete = []
    for idx, row in enumerate(existing_records, start=2):
        if not row['物件名']:
            continue
        room = str(row['部屋番号'] or "").strip()
        if room == "":
            # 部屋番号空欄の行は削除しない
            continue
        key = (normalize(row['物件名']), normalize(row['部屋番号']))
        if key not in scraped_keys:
            rows_to_delete.append(idx)

    for row_idx in sorted(rows_to_delete, reverse=True):
        sheet.delete_rows(row_idx)
        print(f"🗑 スプレッドシートの行 {row_idx} を削除しました")

    # 2. 補完処理
    all_values = sheet.get_all_values()
    latest_records = sheet.get_all_records()
    latest_keys = {(normalize(r['物件名']), normalize(r['部屋番号'])) for r in latest_records}

    used_props = set()

    for idx, row in enumerate(latest_records, start=2):
        row_name = normalize(row['物件名'])
        row_room = normalize(row['部屋番号'])

        if row_room == "":
            candidates = [
                p for p in properties
                if normalize(p['物件名']) == row_name and (normalize(p['物件名']), normalize(p['部屋番号'])) not in used_props
            ]
            if candidates:
                prop = candidates[0]
                used_props.add((normalize(prop['物件名']), normalize(prop['部屋番号'])))

                row_values = [
                    prop['物件名'],
                    prop['部屋番号'],
                    prop['所在地'],
                    prop['最寄り駅'],
                    prop['賃料'],
                    prop['管理費'],
                    prop['間取り'],
                    prop['専有面積'],
                    existing_i_value,
                    prop['URL']
                ]
                cell_range = f"A{idx}:J{idx}"
                sheet.update(cell_range, [row_values])
                print(f"✏️ 物件名 '{row['物件名']}' 行 {idx} のB〜J列を補完しました（部屋番号割当：{prop['部屋番号']}）")
        else:
            key = (row_name, row_room)
            for prop in properties:
                if (normalize(prop['物件名']), normalize(prop['部屋番号'])) == key:
                    used_props.add(key)
                    columns = {
                        '所在地': ("C", prop['所在地']),
                        '最寄り駅': ("D", prop['最寄り駅']),
                        '賃料': ("E", prop['賃料']),
                        '管理費': ("F", prop['管理費']),
                        '間取り': ("G", prop['間取り']),
                        '専有面積': ("H", prop['専有面積']),
                        'URL': ("J", prop['URL']),
                    }
                    for col, value in columns.values():
                        current_val = all_values[idx - 1][ord(col) - ord("A")].strip()
                        if current_val == "" and value != "":
                            cell = f"{col}{idx}"
                            sheet.update(cell, [[value]])
                            print(f"✏️ {row['物件名']}（{row['部屋番号']}）の {col} を補完 → {value}")
                    break

    # 3. 新規追加（used_propsにない物件のみ）
    rows_to_add = []
    for prop in properties:
        key = (normalize(prop['物件名']), normalize(prop['部屋番号']))
        if key not in used_props and key not in latest_keys:
            row = [
                prop['物件名'],
                prop['部屋番号'],
                prop['所在地'],
                prop['最寄り駅'],
                prop['賃料'],
                prop['管理費'],
                prop['間取り'],
                prop['専有面積'],
                None,  # I列：未記入（元から空欄のまま）
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
