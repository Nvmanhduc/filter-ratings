import time
import os
import pyperclip
from collections import defaultdict
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# ---------- Configuration ----------
BASE_URL = 'https://www.chess.com/vi/ratings'
SHEET_URL = 'https://docs.google.com/spreadsheets/d/1Drh72wt8JIqdLgPFKdXeR-wH2qFAlaqVbFwWj70HCFs/edit#gid=0'
# GOOGLE_SHEETS_KEYFILE = 'credentials.json'
SPREADSHEET_ID = '1Drh72wt8JIqdLgPFKdXeR-wH2qFAlaqVbFwWj70HCFs'
SOURCE_SHEET = "Tops"
RESULT_SHEET = "Result"
COLUMN_MAP = {"Cờ cổ điển": 1, "Cờ nhanh": 2, "Cờ chớp": 3}
CHROME_USER_DATA_DIR = os.path.expanduser('~/chrome_profile')
CHROME_PROFILE_NAME = 'Mạnh Đức'
GOOGLE_SHEETS_KEYFILE = os.getenv("GOOGLE_APPLICATION_CREDENTIALS")


# ---------- Helpers ----------
def init_driver():
    opts = Options()
    opts.add_argument(f"--user-data-dir={CHROME_USER_DATA_DIR}")
    opts.add_argument(f"--profile-directory={CHROME_PROFILE_NAME}")
    opts.add_argument('--start-maximized')
    return webdriver.Chrome(options=opts)


def init_sheet():
    scope = [
        'https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/drive'
    ]
    creds = ServiceAccountCredentials.from_json_keyfile_name(GOOGLE_SHEETS_KEYFILE, scope)
    client = gspread.authorize(creds)
    return client.open_by_key(SPREADSHEET_ID)


def scrape_players_from_web(driver):
    players = []
    driver.get(BASE_URL)
    wait = WebDriverWait(driver, 10)
    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "table tbody tr")))

    rows = driver.find_elements(By.CSS_SELECTOR, "table tbody tr")
    for row in rows:
        try:
            name = row.find_element(By.CSS_SELECTOR, "td:nth-child(2) a").text.strip()
            classical = row.find_element(By.CSS_SELECTOR, "td:nth-child(4)").text.strip()
            rapid = row.find_element(By.CSS_SELECTOR, "td:nth-child(5)").text.strip()
            bullet = row.find_element(By.CSS_SELECTOR, "td:nth-child(6)").text.strip()
            players.append([name, classical, rapid, bullet])
        except Exception as e:
            print("Lỗi đọc dòng:", e)

    return players


def write_to_source_sheet(sheet, players,driver):
    driver.execute_script("window.open('');")
    driver.switch_to.window(driver.window_handles[-1])
    driver.get(SHEET_URL)

    sheet.clear()
    sheet.insert_row(["Nhập số điểm để lọc theo cột:"], 1)
    sheet.insert_row(["Tên", "Cờ cổ điển", "Cờ nhanh", "Cờ chớp"], 2)
    sheet.insert_rows(players, row=3, value_input_option="USER_ENTERED")
    print(f"Đã ghi dữ liệu vào sheet '{SOURCE_SHEET}'.")


def get_filter_settings(sheet):
    try:
        classical = sheet.acell("B1").value
        rapid     = sheet.acell("C1").value
        bullet    = sheet.acell("D1").value

        if classical and classical.strip().isdigit():
            return "Cờ cổ điển", int(classical.strip())
        elif rapid and rapid.strip().isdigit():
            return "Cờ nhanh", int(rapid.strip())
        elif bullet and bullet.strip().isdigit():
            return "Cờ chớp", int(bullet.strip())
        else:
            return None, None
    except:
        return None, None


def filter_rows(data, col_index, threshold):
    result = [data[1]]
    for row in data[2:]:
        try:
            value = int(row[col_index])
            if value >= threshold:
                result.append(row)
        except:
            continue
    return result


def write_result(sheet, data):
    sheet.clear()
    sheet.insert_rows(data, row=1, value_input_option="USER_ENTERED")
    print(f"Đã ghi {len(data)-1} dòng vào sheet '{RESULT_SHEET}'.")


def main():
    driver = init_driver()
    spreadsheet = init_sheet()
    source = spreadsheet.worksheet(SOURCE_SHEET)
    try:
        result = spreadsheet.worksheet(RESULT_SHEET)
    except gspread.exceptions.WorksheetNotFound:
        result = spreadsheet.add_worksheet(title=RESULT_SHEET, rows="100", cols="10")

    # 1. Scrape và ghi dữ liệu gốc
    players = scrape_players_from_web(driver)
    write_to_source_sheet(source, players,driver)

    # 2. Theo dõi thay đổi B1–D1 để lọc lại
    last_filter = None
    print("Đang theo dõi B1:C1:D1 để lọc dữ liệu...")
    try:
        while True:
            criteria, threshold = get_filter_settings(source)
            current_filter = (criteria, threshold)

            if current_filter != last_filter and criteria in COLUMN_MAP and threshold is not None:
                col_index = COLUMN_MAP[criteria]
                data = source.get_all_values()
                filtered = filter_rows(data, col_index, threshold)
                write_result(result, filtered)
                last_filter = current_filter
            else:
                print("Không có thay đổi. Chờ lần sau...")

            time.sleep(5)
    finally:
        driver.quit()


if __name__ == "__main__":
    main()
