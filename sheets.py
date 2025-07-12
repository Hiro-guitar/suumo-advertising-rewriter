import gspread
from oauth2client.service_account import ServiceAccountCredentials

def append_to_sheet(values: list[list[str]]):
    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive"
    ]
    creds = ServiceAccountCredentials.from_json_keyfile_name(
        "credentials/suumo-key.json", scope  # ←ファイル名は適宜変更
    )
    client = gspread.authorize(creds)

    # スプレッドシートとシート（タブ）名はあなたのものに合わせて！
    spreadsheet = client.open("物件空室管理")  # ←シート名（画面上部のタイトル）
    worksheet = spreadsheet.worksheet("シート1")  # ←タブ名

    # 下に追記（既存データを消さない！）
    worksheet.append_rows(values)
