import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
from config import GOOGLE_SHEET_ID, GOOGLE_CREDENTIALS_FILE

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

def get_sheet():
    try:
        creds = Credentials.from_service_account_file(GOOGLE_CREDENTIALS_FILE, scopes=SCOPES)
        client = gspread.authorize(creds)
        sheet = client.open_by_key(GOOGLE_SHEET_ID).sheet1
        return sheet
    except Exception as e:
        print(f"Google Sheets xato: {e}")
        return None

async def log_to_sheets(order_id, from_apteka, to_apteka, tovar, miqdor, status):
    try:
        sheet = get_sheet()
        if not sheet:
            return
        sheet.append_row([
            order_id,
            datetime.now().strftime("%Y-%m-%d %H:%M"),
            from_apteka,
            to_apteka,
            tovar,
            miqdor,
            status
        ])
    except Exception as e:
        print(f"Sheets log xato: {e}")

async def update_sheet_status(order_id: int, new_status: str):
    try:
        sheet = get_sheet()
        if not sheet:
            return
        # ID ustunidan qidirish (1-ustun)
        cell = sheet.find(str(order_id), in_column=1)
        if cell:
            sheet.update_cell(cell.row, 7, new_status)  # 7-ustun = status
    except Exception as e:
        print(f"Sheets update xato: {e}")
