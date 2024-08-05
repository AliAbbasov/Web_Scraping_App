import tkinter as tk
from tkinter import ttk, messagebox
from tkinter import filedialog
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
from datetime import datetime, timedelta

class SeleniumScraper:
    def __init__(self, chrome_driver_path, username, password, start_date, finish_date, file_path):
        self.service = Service(chrome_driver_path)
        self.chrome_options = Options()
        self.chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])
        self.driver = webdriver.Chrome(service=self.service, options=self.chrome_options)
        self.username = username
        self.password = password
        self.start_date = start_date
        self.finish_date = finish_date
        self.file_path = file_path
        self.store_mapping = {
            "Store1": "storecode1",
            "Store2": "storecode2",
            "Store3": "storecode3",
            "Store4": "storecode4",
            "Store5": "storecode5",
            "Store6": "storecode6",
            "Store7": "storecode7",
            "Store8": "storecode8",
            "Store9": "storecode9",
            "Store10": "storecode10",
            "Store11": "storecode11"
        }

    def scrape_data(self):
        url = "https://example.com/overview/generalLevel"
        self.driver.get(url)
        self.driver.maximize_window()

        current_date_dt = datetime.strptime(self.start_date, "%d.%m.%Y")
        finish_date_dt = datetime.strptime(self.finish_date, "%d.%m.%Y")

        try:
            WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.ID, "signin-email"))).send_keys(self.username)
            WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.ID, "signin-password"))).send_keys(self.password)
            WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.CLASS_NAME, "login-button"))).click()
            WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.CLASS_NAME, "navTrigger"))).click()
            WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//span[text()='General Level']"))).click()
            WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//span[text()='General Insight']"))).click()
            WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//span[@class='filter-option pull-left' and text()='Select Store From List']"))).click()

            for store_name, store_code in self.store_mapping.items():
                WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.XPATH, f"//span[text()='{store_name}']"))).click()
                time.sleep(1)

            all_data = []
            while current_date_dt <= finish_date_dt:
                current_date = current_date_dt.strftime("%d.%m.%Y")
                WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.ID, "reportrange"))).click()
                start_date_input = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.NAME, "daterangepicker_start")))
                start_date_input.clear()
                start_date_input.send_keys(current_date)
                end_date_input = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.NAME, "daterangepicker_end")))
                end_date_input.clear()
                end_date_input.send_keys(current_date)
                WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//button[@class='applyBtn btn btn-sm btn-success' and text()='Apply']"))).click()
                WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.ID, "go"))).click()
                table = WebDriverWait(self.driver, 20).until(EC.presence_of_element_located((By.ID, "generalDatatable")))

                rows = table.find_elements(By.TAG_NAME, "tr")
                for row in rows:
                    cols = row.find_elements(By.TAG_NAME, "td")
                    if len(cols) > 1 and cols[0].text.lower() != 'sum':
                        filtered_cols = [cols[0].text] + [int(cols[i].text.replace(',', '')) if cols[i].text.replace(',', '').isdigit() else cols[i].text for i in range(2, len(cols)-1)]
                        filtered_cols.append(current_date)
                        all_data.append(filtered_cols)

                if current_date_dt < finish_date_dt:
                    next_button = WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//div[@class='arrows dateArrow-right glyphicon glyphicon-triangle-right hidden-xs hidden-sm shadow']")))
                    next_button.click()
                    time.sleep(2)

                current_date_dt += timedelta(days=1)

            column_names = ['Store Name'] + [str(i + 9) for i in range(1, len(all_data[0]) - 1)] + ['Date']
            df = pd.DataFrame(all_data, columns=column_names)
            df['Store Code'] = df['Store Name'].map(self.store_mapping)
            columns_order = ['Store Name', 'Store Code', 'Date'] + [str(i + 9) for i in range(1, len(all_data[0]) - 1)]
            df = df[columns_order]

            file_name = f"{self.file_path}/scraping_data_{self.start_date.replace('.', '-')}_to_{self.finish_date.replace('.', '-')}.xlsx"
            df.to_excel(file_name, index=False)
            print(f"Data saved to {file_name}")
            time.sleep(10)

        except Exception as e:
            print(f"An error occurred: {e}")

        finally:
            self.driver.quit()


class StoreDataProcessor:
    def __init__(self, start_date, finish_date, file_path):
        self.start_date = start_date
        self.finish_date = finish_date
        self.file_path = file_path
        self.input_file_name = f"{file_path}/scraping_data_{start_date.replace('.', '-')}_to_{finish_date.replace('.', '-')}.xlsx"
        self.output_file_name = f"{file_path}/data_with_query_{start_date.replace('.', '-')}_to_{finish_date.replace('.', '-')}.xlsx"

    def process_data(self):
        df = pd.read_excel(self.input_file_name)
        output_data = []
        dates = df['Date'].unique()

        for date in dates:
            date_df = df[df['Date'] == date]
            for index, row in date_df.iterrows():
                name = row["Store Name"]
                site_value = row["Store Code"] if "Store Code" in df.columns else 0
                formatted_date = pd.to_datetime(date, dayfirst=True).strftime('%d-%m-%Y')

                for time_interval, qty in row.items():
                    if time_interval not in ["Store Name", "Store Code", "Date"]:
                        try:
                            time_number = int(time_interval)
                            start_time = f"{time_number}:00"
                            end_time = f"{time_number + 1}:00"
                            main_query = (f"INSERT INTO trStoreVisitors (CompanyCode, OfficeCode, StoreCode, CurrentDate, "
                                          f"CurrentHour, InVisitorCount, OutVisitorCount) VALUES (1, '{site_value}', "
                                          f"'{site_value}', '{formatted_date}', {time_number}, {qty}, {qty})")
                            output_data.append({
                                "Store Name": name,
                                "Store Code": site_value,
                                "Date": formatted_date,
                                "Start Time": start_time,
                                "End Time": end_time,
                                "Time #": time_number,
                                "Qty": qty,
                                "Main Query": main_query
                            })
                        except ValueError:
                            continue

        output_df = pd.DataFrame(output_data)
        output_df.to_excel(self.output_file_name, index=False)
        print(f"Data saved to {self.output_file_name}")

# GUI setup
class Application(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Data Scraper and Processor")
        self.geometry("400x200")

        self.start_date_label = ttk.Label(self, text="Start Date (DD.MM.YYYY):")
        self.start_date_label.pack(pady=5)
        self.start_date_entry = ttk.Entry(self)
        self.start_date_entry.pack(pady=5)

        self.finish_date_label = ttk.Label(self, text="Finish Date (DD.MM.YYYY):")
        self.finish_date_label.pack(pady=5)
        self.finish_date_entry = ttk.Entry(self)
        self.finish_date_entry.pack(pady=5)

        self.file_path_label = ttk.Label(self, text="File Path:")
        self.file_path_label.pack(pady=5)
        self.file_path_entry = ttk.Entry(self)
        self.file_path_entry.pack(pady=5)

        self.browse_button = ttk.Button(self, text="Browse", command=self.browse_directory)
        self.browse_button.pack(pady=5)

        self.scrape_button = ttk.Button(self, text="Scrape and Process
