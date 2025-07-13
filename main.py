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
    existing_names = [row['物件名'] for row in existing_records]

    scraped_names = [prop['物件名'] for prop in properties]

    # 削除対象行の特定（1行目ヘッダーなので2からスタート）
    rows_to_delete = []
    for idx, row in enumerate(existing_records, start=2):
        if row['物件名'] not in scraped_names:
            rows_to_delete.append(idx)

    # 削除は後ろから行う（行番号ズレ防止）
    for row_idx in sorted(rows_to_delete, reverse=True):
        sheet.delete_rows(row_idx)
        print(f"🗑 スプレッドシートの行 {row_idx} を削除しました")

    # URLを補完する処理（物件名が一致し、URLが空欄の行に対して）
    for idx, row in enumerate(existing_records, start=2):
        for prop in properties:
            if row['物件名'] == prop['物件名'] and not row['URL'] and prop['URL']:
                cell = f"J{idx}"
                sheet.update(cell, [[prop['URL']]])
                print(f"🔗 URLを補完しました：{row['物件名']} → {prop['URL']}")

    # 削除後の最新A列を取得（ヘッダー除く）
    latest_names = sheet.col_values(1)[1:]

    # 追加対象の物件を特定
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
                '',  # 使わない列は空白
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
