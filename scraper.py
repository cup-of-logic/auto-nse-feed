from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver import ChromeOptions
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

import gspread
from google.oauth2.service_account import Credentials
import pandas as pd

import sys
import time
from datetime import datetime
import colorama
from colorama import Fore, Back, Style
from win10toast import ToastNotifier
from pprint import pprint


PATH = "https://www.nseindia.com/market-data/oi-spurts"
CHROME_DRIVER_PATH = r"C:\Program Files (x86)\Web Driver\chromedriver.exe"
# GSHEET_ID = "17oIOPPRMKJqJAzdzFmuaZE9qUu2CobqMmpn5kkkQyU4"
GSHEET_ID = "1VUB5F70BxKfKE1tQ6vESrbpSLvS3QE5khPij9QnKFK0"

TABLE_DIV_ID = "oi_sprutz_table"
TABLE_ID = "oiSpurtsTable"

TIME_INTERVAL = 250  # in seconds
SHOW_LAST_UPDATE_NOTIF = 600  # in seconds
TRY_AGAIN_WAIT = 10  # in seconds

LOG_COLORS = {
    "INFO": Fore.WHITE + Style.BRIGHT,
    "ERROR": Fore.RED + Style.BRIGHT,
    "DEBUG": Fore.CYAN + Style.DIM,
    "CRITICAL": Back.RED + Fore.WHITE + Style.BRIGHT
}

colorama.init()
toaster = ToastNotifier()


def get_current_time():
    return datetime.now().strftime('%d-%m-%Y %H:%M:%S')


def show_log(category: str, message: str):
    category = category.upper()

    log = f"[{get_current_time()}] [{category}] {message}"
    print(LOG_COLORS[category] + log + Style.RESET_ALL)

    with open("logs.txt", mode='a') as logs_file:
        logs_file.write(log + '\n')


class Scrape:
    def __init__(self):
        self.web_driver = self.table_df = self.sheet = None
        self.last_updated = datetime.now()

        try:
            self.__setup_gspread()
        except Exception as e:
            show_log(category="ERROR", message=f"Error setting up Google Sheet: {e}")
            toaster.show_toast("ERROR NOTIFICATION", "Error setting up Google Sheet", duration=10)
            sys.exit()

        self.__scrape()

    def __scrape(self):
        while True:
            if (datetime.now() - self.last_updated).total_seconds() > SHOW_LAST_UPDATE_NOTIF:
                show_log(category="CRITICAL", message=f"Google Sheet has not been updated within the last {SHOW_LAST_UPDATE_NOTIF} seconds. It is advisable to rerun the script.")
                toaster.show_toast("Error", "Error setting up Google Sheet", duration=10)

            try:
                self.__get_web_driver()
            except Exception:
                show_log(category="CRITICAL", message=f"Could not load the website. Trying Again!")
                time.sleep(TRY_AGAIN_WAIT)
                continue

            try:
                self.__get_table()
            except Exception:
                show_log(category="CRITICAL", message=f"Could not fetch the table HTML. Trying Again!")
                time.sleep(TRY_AGAIN_WAIT)
                continue

            try:
                self.__set_values()
            except Exception:
                show_log(category="CRITICAL", message=f"Could not update Google Sheet. Trying Again!")
                time.sleep(TRY_AGAIN_WAIT)
                continue

            time.sleep(TIME_INTERVAL)

    def __get_web_driver(self):
        try:
            options = ChromeOptions()
            options.add_argument('--no-sandbox')
            options.add_argument('--disable-dev-shm-usage')
            options.add_argument("--window-position=-2000,0")

            service = Service(executable_path=CHROME_DRIVER_PATH)
            self.web_driver = webdriver.Chrome(service=service, options=options)
            show_log(category='DEBUG', message=f"Set up Chrome Driver using options: {options.to_capabilities()}")
        except Exception:
            show_log(category='ERROR', message="Could not set up Chrome Driver. Check if its version matches the current chrome version")
            toaster.show_toast("Error", "Could not set up Chrome Driver. Check if its version matches the current chrome version", duration=10)
            sys.exit()

        self.web_driver.get(PATH)
        WebDriverWait(self.web_driver, 10).until(
            EC.presence_of_element_located((By.ID, TABLE_ID))
        )

        show_log(category='DEBUG', message=f"Website loaded successfully.")

    def __setup_gspread(self):
        scope = ["https://www.googleapis.com/auth/spreadsheets"]
        creds = Credentials.from_service_account_file("creds_.json", scopes=scope)
        client = gspread.authorize(creds)
        self.sheet = client.open_by_key(GSHEET_ID).get_worksheet(0)

        show_log(category='INFO', message=f"Successfully connected to Google Sheets.")

    def __get_table(self):
        html = self.web_driver.page_source

        tables = pd.read_html(html)
        self.table_df = tables[0]

        show_log(category='DEBUG', message="HTML table fetched successfully.")

        self.web_driver.quit()

    def __set_values(self):
        time_str = f"Last Updated: {get_current_time()}"
        values = [[time_str]] + [self.table_df.columns.tolist()] + self.table_df.values.tolist()

        self.sheet.update(values)
        self.sheet.merge_cells('A1:J1')
        self.last_updated = datetime.now()

        show_log(category='INFO', message="Google Sheet Updated.")


if __name__ == '__main__':
    show_log(category='INFO', message="Starting scraper.")
    Scrape()
