import re
import pandas as pd
import time
from pathlib import Path
from openpyxl import load_workbook
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException, NoSuchFrameException
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException, UnexpectedAlertPresentException, NoAlertPresentException


def run_cleanup(file_name, output_file, status_callback=None):

    def update_status(message):
        print(message)

        if status_callback:
            status_callback(message)

        time.sleep(0.25)

    def close_alert_if_present(timeout=3):
        try:
            WebDriverWait(driver, timeout).until(EC.alert_is_present())
            alert_curr = driver.switch_to.alert
            # update_status(f"Alert appeared: {alert_curr.text[:80]}...")
            alert_curr.accept()
            # update_status("Alert closed")
        except TimeoutException:
            pass

    def wait_for_login_complete(timeout=300):
        start_time = time.time()

        while time.time() - start_time < timeout:
            close_alert_if_present()

            try:
                driver.switch_to.default_content()
                driver.switch_to.frame("menu")
                return True
            except (NoSuchFrameException, NoSuchElementException, UnexpectedAlertPresentException):
                time.sleep(1)

        raise TimeoutError("Login did not complete within allowed time.")

    def clean_part_number(part_value):
        if pd.isna(part_value):
            return ""

        part_value = str(part_value).upper()

        match = re.search(r"GH\d{2}-[A-Z0-9]+", part_value)

        if match:
            return match.group(0)

        return part_value.strip()

    update_status("Reading Excel file...")

    input_path = Path(file_name)

    if input_path.suffix.lower() == ".xls":
        update_status("XLS file detected. Converting to XLSX...")

        converted_file = input_path.with_name(f"{input_path.stem}_converted.xlsx")

        xls_df = pd.read_excel(input_path, engine="xlrd")
        xls_df.to_excel(converted_file, index=False)

        file_name = converted_file

        update_status(f"Converted file created: {converted_file.name}")

    df = pd.read_excel(file_name)

    columns_to_keep = [
        "ASC Job No",
        "Description",
        "IMEI No",
        "Parts No",
        "Part Serial",
    ]

    cleaned_df = df[columns_to_keep].copy()
    cleaned_df["Tech"] = ""
    cleaned_df["Verified Serial"] = cleaned_df["Part Serial"]

    # --------------------------------------------
    # Selenium Website Access Script
    # --------------------------------------------

    service_order_url = "https://biz3.samsungcsportal.com/svctracking/lite/ServiceOrderListLite.jsp?search_status=&searchContent=&menuBlock=&menuUrl=&naviDirValue="
    options = Options()

    profile_path = Path.home() / "GSPN_Report_Cleaner_Chrome_Profile"
    options.add_argument(f"--user-data-dir={profile_path}")
    options.add_argument("--profile-directory=Default")

    update_status("Opening Chrome...")

    driver = webdriver.Chrome(options=options)

    driver.get("https://gspn3.samsungcsportal.com/main.jsp")

    update_status("Waiting for GSPN Login...")
    # input("Log in if needed, then press Enter...")

    wait_for_login_complete()

    update_status("Opening Business page...")
    wait = WebDriverWait(driver, 300)
    update_status("Login detected. Continuing...")

    business = wait.until(
        EC.element_to_be_clickable((By.XPATH, "//span[.//nobr[normalize-space()='Business']]"))
    )
    old_windows = driver.window_handles

    business.click()

    wait.until(lambda d: len(d.window_handles) > len(old_windows))
    driver.switch_to.window(driver.window_handles[-1])
    update_status("Waiting for Business page to load...")

    wait.until(lambda d: "operate.do" in d.current_url)

    time.sleep(2)

    update_status("Business page loaded.")

    driver.switch_to.default_content()
    driver.get(service_order_url)

    update_status("Loading Service Order Management Light...")

    wait.until(
        EC.element_to_be_clickable((By.ID, "service_order_no"))
    )

    update_status("Service Order page loaded. Continuing...")


    # --------------------------------------------
    # Search Tickets Logic
    # --------------------------------------------

    total_rows = len(cleaned_df)
    unique_tickets = cleaned_df["ASC Job No"].nunique()

    update_status("Retrieving technician names from each ticket number...")

    for ticket_number, (asc_job_no, ticket_rows) in enumerate(cleaned_df.groupby("ASC Job No", sort=False), start=1):
        asc_job_no = str(asc_job_no).strip()

        update_status(f"\nSearching ticket {ticket_number} of {unique_tickets}: {asc_job_no}")

        driver.get(service_order_url)

        service_order_input = wait.until(
            EC.element_to_be_clickable((By.ID, "service_order_no"))
        )

        service_order_input.clear()
        service_order_input.send_keys(asc_job_no)
        service_order_input.send_keys(Keys.ENTER)

        try:
            WebDriverWait(driver, 3).until(EC.alert_is_present())
            alert = driver.switch_to.alert
            # update_status(f"Alert appeared: {alert.text}")
            alert.accept()
            # update_status("Alert closed")
        except TimeoutException:
            pass

        engineer_element = wait.until(
            EC.presence_of_element_located((By.ID, "ENGINEERNAME"))
        )

        engineer_text = engineer_element.text.strip()

        if engineer_text:
            parts = [part.strip() for part in engineer_text.split("|")]

            name_part = None

            for part in parts:
                if not part:
                    continue

                if part.replace("-", "").replace(" ", "").isdigit():
                    continue

                if re.search(r"\d", part) and part.isupper():
                    continue

                name_part = part
                break

            if name_part:
                first_name = name_part.split()[0]
            else:
                first_name = "NOT FOUND"

        else:
            first_name = "NOT FOUND"

        update_status(f"Ticket: {asc_job_no}\nTechnician: {first_name}")

        for index, row in ticket_rows.iterrows():
            cleaned_df.loc[index, "Tech"] = first_name

        update_status("Reading repair part fields from ticket...")

        page_parts = {}

        part_code_elements = driver.find_elements(By.NAME, "PARTS_CODE")
        old_serial_elements = driver.find_elements(By.NAME, "OLD_SERIAL_MAT")
        old_fab_elements = driver.find_elements(By.NAME, "OLD_FAB_ID")

        for i, part_code_elements in enumerate(part_code_elements):
            part_code = clean_part_number(part_code_elements.get_attribute("value"))

            old_serial = ""
            old_fab = ""

            if i < len(old_serial_elements):
                old_serial = old_serial_elements[i].get_attribute("value") or ""

            if i < len(old_fab_elements):
                old_fab = old_fab_elements[i].get_attribute("value") or ""

            old_serial = old_serial.strip()
            old_fab = old_fab.strip()

            verified_from_page = old_fab if old_fab else old_serial

            if part_code:
                page_parts.setdefault(part_code, []).append({
                    "old_serial": old_serial,
                    "old_fab": old_fab,
                    "verified": verified_from_page
                })

        for index, row in ticket_rows.iterrows():
            cleaned_df.loc[index, "Tech"] = first_name

            excel_part_no = clean_part_number(row["Parts No"])

            if pd.isna(row["Part Serial"]):
                excel_part_serial = ""
            else:
                excel_part_serial = str(row["Part Serial"]).strip()

            verified_serial = excel_part_serial

            matching_parts = page_parts.get(excel_part_no, [])

            if matching_parts:
                matched_part = None

                for part in matching_parts:
                    if excel_part_serial and part["old_serial"] == excel_part_serial:
                        matched_part = part
                        break

                if matched_part is None:
                    matched_part = matching_parts[0]

                if matched_part["verified"]:
                    verified_serial = matched_part["verified"]

            cleaned_df.loc[index, "Verified Serial"] = verified_serial

            update_status(f"{excel_part_no}\nVerified Serial: {verified_serial}")

        # if pd.isna(row["Part Serial"]):
        #     verified_serial = ""
        # else:
        #     verified_serial = str(row["Part Serial"]).strip()
        #
        # try:
        #     octa_old_element = driver.find_element(By.ID, "OLD_FAB_ID")
        #     octa_old_value = octa_old_element.get_attribute("value")
        #
        #     if octa_old_value:
        #         verified_serial = octa_old_value.strip()
        #
        # except NoSuchElementException:
        #     pass
        #
        # update_status(f"Verified Serial: {verified_serial}")
        #
        # cleaned_df.loc[index, "Verified Serial"] = verified_serial


    # ---------------------------------------------
    # Ticket Page
    # engineer_element = wait.until(
    #     EC.presence_of_element_located((By.ID, "ENGINEERNAME"))
    # )

    # engineer_text = engineer_element.text.strip()
    # first_name = engineer_text.split()[0]
    # print("Engineer full text:", engineer_text)
    # print("Engineer first name:", first_name)

    # cleaned_df.loc[index, "Tech"] = first_name

    cleaned_df = cleaned_df.sort_values(by=["Tech", "ASC Job No"])

    cleaned_df.to_excel(output_file, index=False)

    wb = load_workbook(output_file)
    ws = wb.active

    update_status("\nAdjusting Excel column widths...")

    for column in ws.columns:
        max_length = 0
        column_letter = column[0].column_letter

        for cell in column:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))

        ws.column_dimensions[column_letter].width = max_length + 5

    update_status("Saving cleaned Excel file...")

    wb.save(output_file)

    output_name = Path(output_file).name

    update_status("Done!")
    update_status(f"\nSaved file name: {output_name}\n")

    driver.quit()