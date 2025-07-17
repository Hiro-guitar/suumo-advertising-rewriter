# main.py

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
    sheet.update([header], 'A1:J1')  # 値を先に、範囲を後に

    existing_records = sheet.get_all_records()
    scraped_names = [prop['物件名'] for prop in properties]

    # 削除対象行の特定（1行目ヘッダーなので2からスタート）
        # 削除対象行の特定（1行目ヘッダーなので2からスタート）
    rows_to_delete = []

    # 物件名の出現回数をカウント
    from collections import Counter
    name_counter = Counter([row['物件名'] for row in existing_records])

    # 既存物件を物件名 + 部屋番号で整理
    scraped_set_full = set((prop['物件名'], prop['部屋番号']) for prop in properties)
    scraped_names_only = set(prop['物件名'] for prop in properties)

    for idx, row in enumerate(existing_records, start=2):
        name = row['物件名']
        room = row['部屋番号']
        if name_counter[name] > 1:
            # 同じ物件名が複数ある場合は、部屋番号も含めてチェック
            if (name, room) not in scraped_set_full:
                rows_to_delete.append(idx)
        else:
            # 1件しかない場合は、物件名だけで判定
            if name not in scraped_names_only:
                rows_to_delete.append(idx)

    for row_idx in sorted(rows_to_delete, reverse=True):
        sheet.delete_rows(row_idx)
        print(f"🗑 スプレッドシートの行 {row_idx} を削除しました")

    # 最新のデータを取得
    all_values = sheet.get_all_values()
    latest_records = sheet.get_all_records()  # ヘッダー除いた辞書形式
    latest_names = [row['物件名'] for row in latest_records]

    # 情報補完処理（すでに存在する物件で空欄セルがある場合）
    for idx, row in enumerate(latest_records, start=2):  # 行番号2から開始（1はヘッダー）
        for prop in properties:
            if row['物件名'] == prop['物件名']:
                updates = []
                columns = {
                    '部屋番号': ("B", prop['部屋番号']),
                    '所在地': ("C", prop['所在地']),
                    '最寄り駅': ("D", prop['最寄り駅']),
                    '賃料': ("E", prop['賃料']),
                    '管理費': ("F", prop['管理費']),
                    '間取り': ("G", prop['間取り']),
                    '専有面積': ("H", prop['専有面積']),
                    'URL': ("J", prop['URL']),
                }
                for key, (col, value) in columns.items():
                    # 対象セルが空欄 or 値がない場合にのみ更新
                    existing_value = all_values[idx - 1][ord(col) - ord("A")].strip()
                    if not existing_value and value != "":
                        cell = f"{col}{idx}"
                        sheet.update(cell, [[value]])
                        print(f"✏️ {row['物件名']} の {key} を補完しました → {value}")

    # 追加対象の物件を特定（新規）
    rows_to_add = []
    for prop in properties:
        if prop['物件名'] not in latest_names:
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
