from suumo_scraper import login_to_fn_forrent, click_keisai_bukken_only, extract_properties, extract_free_comment_urls
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import datetime

def update_sheet(properties):
    scope = ['https://spreadsheets.google.com/feeds',
             'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name('suumo-key.json', scope)
    client = gspread.authorize(creds)

    sheet = client.open('物件空室管理').worksheet('シート1')

    # ヘッダー設定（A-N列）
    header = ['物件名', '部屋番号', '所在地', '最寄り駅', '賃料', '管理費',
              '間取り', '専有面積', 'ステータス', 'URL', '終了日', '元付', '空室確認', 'FN消滅日']
    sheet.update([header], 'A1:N1')

    existing_records = sheet.get_all_records()
    all_values = sheet.get_all_values()

    def normalize(val):
        s = str(val or "").strip().lstrip("'")
        return s.zfill(4) if s.isdigit() else s

    scraped_keys = {(normalize(p['物件名']), normalize(p['部屋番号'])) for p in properties}
    scraped_name_set = {normalize(p['物件名']) for p in properties}

    today = datetime.date.today().isoformat()
    rows_to_delete = []

    # ================================================================
    # Phase 1: N列管理（FN Forrent 消滅トラッキング + 2週間削除）
    # ================================================================
    for idx, row in enumerate(existing_records, start=2):
        if not row.get('物件名'):
            continue

        row_name = normalize(row['物件名'])
        row_room = normalize(str(row.get('部屋番号', '') or ''))

        # 部屋番号空欄の行はN列管理しない（補完処理で部屋番号が割り当てられる）
        if row_room == "":
            continue

        key = (row_name, row_room)
        row_data = all_values[idx - 1] if idx - 1 < len(all_values) else []
        current_n = row_data[13].strip() if len(row_data) > 13 else ""  # N列 index=13

        if key in scraped_keys:
            # FN Forrent にある → N列クリア（復帰）
            if current_n:
                sheet.update(f"N{idx}", [[""]])
                print(f"🔄 {row['物件名']}（{row_room}）FN消滅日クリア（復帰）")
        else:
            # FN Forrent にない
            if not current_n:
                # 初回検出: N列に今日の日付をセット
                sheet.update(f"N{idx}", [[today]])
                print(f"📅 {row['物件名']}（{row_room}）FN Forrent から消滅 → {today}")
            else:
                # 既に消滅日あり → 2週間チェック
                try:
                    disappeared = datetime.date.fromisoformat(current_n)
                    if (datetime.date.today() - disappeared).days >= 14:
                        rows_to_delete.append(idx)
                        print(f"🗑 {row['物件名']}（{row_room}）2週間経過 → 削除予定")
                except ValueError:
                    pass  # 不正な日付形式はスキップ

    # 2週間超の行を削除（後ろから削除してインデックスずれを防止）
    for row_idx in sorted(rows_to_delete, reverse=True):
        sheet.delete_rows(row_idx)
        print(f"🗑 行 {row_idx} を削除しました")

    # ================================================================
    # Phase 2: 補完処理（削除後にデータを再取得）
    # ================================================================
    all_values = sheet.get_all_values()
    latest_records = sheet.get_all_records()
    latest_keys = {(normalize(r['物件名']), normalize(str(r.get('部屋番号', '') or ''))) for r in latest_records}

    used_props = set()

    for idx, row in enumerate(latest_records, start=2):
        row_name = normalize(row['物件名'])
        row_room = normalize(str(row.get('部屋番号', '') or ''))

        row_data = all_values[idx - 1] if idx - 1 < len(all_values) else []
        current_i = row_data[8].strip() if len(row_data) > 8 else ""   # I列 index=8
        current_k = row_data[10].strip() if len(row_data) > 10 else "" # K列 index=10
        current_m = row_data[12].strip() if len(row_data) > 12 else "" # M列 index=12

        if row_room == "":
            # ------ 部屋番号空欄行の補完 ------
            candidates = [
                p for p in properties
                if normalize(p['物件名']) == row_name and (normalize(p['物件名']), normalize(p['部屋番号'])) not in used_props
            ]
            if candidates:
                prop = candidates[0]
                used_props.add((normalize(prop['物件名']), normalize(prop['部屋番号'])))

                vacancy_url = prop.get('空室確認URL', '')
                effective_m = current_m or vacancy_url

                # I列の決定: M列にURLあり + I列空 + K列空 → '募集中'
                if not current_i and not current_k and effective_m:
                    new_i = '募集中'
                else:
                    new_i = current_i  # 既存値を保持

                row_values = [
                    prop['物件名'],
                    prop['部屋番号'],
                    prop['所在地'],
                    prop['最寄り駅'],
                    prop['賃料'],
                    prop['管理費'],
                    prop['間取り'],
                    prop['専有面積'],
                    new_i,
                    prop['URL']
                ]
                sheet.update(f"A{idx}:J{idx}", [row_values])

                # M列補完
                if not current_m and vacancy_url:
                    sheet.update(f"M{idx}", [[vacancy_url]])
                    print(f"  📎 {row['物件名']}（{prop['部屋番号']}）M列セット: {vacancy_url}")

                print(f"✏️ '{row['物件名']}' 行 {idx} を補完（部屋番号: {prop['部屋番号']}）")
        else:
            # ------ 部屋番号あり行の補完 ------
            key = (row_name, row_room)
            for prop in properties:
                if (normalize(prop['物件名']), normalize(prop['部屋番号'])) == key:
                    used_props.add(key)

                    # B-H列 + J列: 空セルのみ補完（既存ロジック）
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
                    for col_name, (col, value) in columns.items():
                        current_val = all_values[idx - 1][ord(col) - ord("A")].strip()
                        if current_val == "" and str(value) != "":
                            sheet.update(f"{col}{idx}", [[value]])
                            print(f"✏️ {row['物件名']}（{row_room}）{col}列を補完 → {value}")

                    # M列補完: 空ならスクレイパーのURLをセット（既存URLは上書きしない）
                    vacancy_url = prop.get('空室確認URL', '')
                    if not current_m and vacancy_url:
                        sheet.update(f"M{idx}", [[vacancy_url]])
                        current_m = vacancy_url
                        print(f"  📎 {row['物件名']}（{row_room}）M列セット: {vacancy_url}")

                    # I列: M列にURLあり + I列空 + K列空 → '募集中'
                    effective_m = current_m or vacancy_url
                    if not current_i and not current_k and effective_m:
                        sheet.update(f"I{idx}", [['募集中']])
                        print(f"  ✅ {row['物件名']}（{row_room}）I列を「募集中」にセット")

                    break

    # ================================================================
    # Phase 3: 新規物件の追加
    # ================================================================
    rows_to_add = []
    for prop in properties:
        key = (normalize(prop['物件名']), normalize(prop['部屋番号']))
        if key not in used_props and key not in latest_keys:
            vacancy_url = prop.get('空室確認URL', '')
            # M列にURLがあれば I列 = '募集中'、なければ空
            i_value = '募集中' if vacancy_url else ''

            row = [
                prop['物件名'],
                prop['部屋番号'],
                prop['所在地'],
                prop['最寄り駅'],
                prop['賃料'],
                prop['管理費'],
                prop['間取り'],
                prop['専有面積'],
                i_value,          # I列: ステータス
                prop['URL'],
                '',               # K列: 終了日（vacancy-checker管理）
                '',               # L列: 元付（手動）
                vacancy_url,      # M列: 空室確認URL
                '',               # N列: FN消滅日
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
    extract_free_comment_urls(driver, properties)
    driver.quit()

    update_sheet(properties)
