from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, WebDriverException
import time
import os
import pandas as pd
# import os
import json
import re
import requests


TABLEAU_URL = "https://eu-west-1a.online.tableau.com/#/site/jlrradar/projects/320642"
TABLEAU_EMAIL = "ssing232@partner.jaguarlandrover.com"


def get_edge_driver():
    options = webdriver.EdgeOptions()
    options.add_argument("--no-sandbox")
    options.add_argument("--start-maximized")

    try:
        driver = webdriver.Edge(options=options)
        return driver
    except WebDriverException as e:
        print(f"Edge WebDriver initialize karne mein error: {e}")
        print("Please ensure Edge browser aur sahi Edge WebDriver (msedgedriver.exe) installed hain aur aapke PATH mein hain.")
        exit()


driver = get_edge_driver()
wait = WebDriverWait(driver, 60)

print("WebDriver initialized. Tableau URL par navigate kar raha hai...")

original_window_handle = driver.current_window_handle

try:
    driver.get(TABLEAU_URL)
    print(f"Navigated to: {TABLEAU_URL}")

    email_input = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@id ='email']")))
    email_input.send_keys(TABLEAU_EMAIL)
    print("Email enter kiya gaya.")

    login_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@id ='login-submit']")))
    login_button.click()
    print("Login button par click kiya gaya. Authentication complete hone ka wait kar raha hai...")

    time.sleep(5)

    workbook_link = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[contains(@href,'workbooks/2537541')]")))
    workbook_link.click()
    print("Workbook link par click kiya gaya. Workbook load hone ka wait kar raha hai...")

    time.sleep(5)

    view_link = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[contains(@href,'redirect_to_view/13504758')]")))
    view_link.click()
    print("View link par click kiya gaya. View load hone aur iframe ke appear hone ka wait kar raha hai...")

    print("Tableau iframe par switch karne ka attempt kar raha hai...")
    time.sleep(5)

except Exception as e:
    print(f"Ek Timeout error hua: {e}")
    print("Iska matlab usually yeh hai ki ek element specified time ke andar nahi mila.")
    if 'driver' in locals():
        driver.quit()

try:

    iframe_locator = (By.XPATH, "//iframe[contains(@src, 'tableau')]")
    
    wait.until(EC.frame_to_be_available_and_switch_to_it(iframe_locator))
    print("Tableau iframe par successfully switch kiya gaya.")

    print("Iframe ke andar main 'Download' button search aur click karne ka attempt kar raha hai...")
    download_button = wait.until(
        EC.element_to_be_clickable((By.XPATH, "//div[@id='toolbar-container']//button[@id='download']"))
    )
    download_button.click()
    print("Main 'Download' button par click kiya gaya.")
    time.sleep(5)

    print("Download menu mein 'Data' option search aur click karne ka attempt kar raha hai...")
    data_option_button = wait.until(
        EC.element_to_be_clickable((By.XPATH, "//div[@id='viz-viewer-toolbar-download-menu']//span[text()='Crosstab']"))
    )
    data_option_button.click()
    print("'Data' option par click kiya gaya.")

    print("Naye download window/tab ke open hone ka wait kar raha hai...")
    time.sleep(5)
    
    modal_dialog_locator = (By.XPATH, "//div[@role='dialog' and contains(., 'Select Format')]") 
    wait.until(EC.visibility_of_element_located(modal_dialog_locator))
    print("Modal download dialog box successfully appear hua.")


    try:
        
        csv_radio_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//label[@data-tb-test-id='crosstab-options-dialog-radio-csv-Label']")))
        csv_radio_button.click()
        print("'CSV' format select kiya gaya.")
        time.sleep(5) 
    except TimeoutException:
        print("ERROR: Crosstab dialog (modal) mein 'CSV' radio button specified time ke andar nahi mila ya clickable nahi.")
        raise 
    except Exception as e:
        print(f"ERROR selecting CSV radio button: {e}")
        raise

    
    final_download_button = wait.until(
        EC.element_to_be_clickable((By.XPATH, "//button[@data-tb-test-id='export-crosstab-export-Button']"))
    )
    print("Naye download window mein Dialog mein final 'Download' button mil gaya aur clickable hai.")
    final_download_button.click()
    print("Naye download window mein Dialog mein final 'Download' button par click kiya gaya. Data download initiate hona chahiye.")

    time.sleep(5)
    
    # driver.close() 
    print("Download window close kiya gaya.")

except Exception as e:
    print(f"Ek Timeout error hua: {e}")
    print("Iska matlab usually yeh hai ki ek element specified time ke andar nahi mila.")
    if 'driver' in locals():
        driver.quit()


try:
    # 📁 Step 1: Set the folder path where new CSVs will be downloaded
    folder_path = "C:/Users/ssing232/Downloads/"  # 🟡 Change this to your actual download folder

    # ✅ Step 2: Get all CSV files in the folder
    csv_files = [file for file in os.listdir(folder_path) if file.lower().endswith('.csv')]

    if not csv_files:
        raise FileNotFoundError("❌ No CSV file found in the folder!")

    # ✅ Step 3: Pick the most recently created CSV file
    latest_csv_path = max(
        [os.path.join(folder_path, f) for f in csv_files],
        key=os.path.getctime
    )

    print(f"📄 Latest CSV file fetched: {latest_csv_path}")

    # ✅ Step 4: Output file names
    base_dir = os.path.dirname(latest_csv_path)
    base_file_name_with_timestamp = "clean_output"
    output_csv = os.path.join(base_dir, f"{base_file_name_with_timestamp}.csv")
    output_json = os.path.join(base_dir, f"{base_file_name_with_timestamp}.json")

    # ✅ Step 5: Extract 'range' from the file name
    filename = os.path.basename(latest_csv_path)
    # match = re.search(r"SA\s*-\s*(.+?)\.csv", filename, re.IGNORECASE)
    # range_value = match.group(1).strip() if match else "UNKNOWN"
    match = re.search(r"SA\s*-\s*(QTD|MTD|YTD)\s*\(?\d*\)?",filename, re.IGNORECASE)
    range_value = match.group(1).strip() if match else "UNKNOWN"
    filename
    # ✅ Step 6: Read CSV file with fallback encodings
    try:
        df = pd.read_csv(latest_csv_path, encoding='utf-16', sep='\t')
        if df.shape[1] == 1:
            raise ValueError("Fallback to latin1")
    except Exception:
        df = pd.read_csv(latest_csv_path, encoding='latin1', sep='\t')

    print("🔍 Columns detected:", list(df.columns))

    # ✅ Step 7: Drop first 2 and last 5 columns
    if df.shape[1] < 7:
        raise ValueError("❌ Not enough columns to drop first 2 and last 5.")

    df = df.iloc[:, 2:]

    # ✅ Step 8: Rename columns to expected names
    expected_columns = [
        "user_email", "enquiry", "unique_td",
                "enquiry_to_UTD", "enquiry_to_TD", "new_orders",
                "td_to_retail", "cancellations", "cancellation_contro","avg_enq_to_ord_days","dig_enq_to_ord_days",
                "testDrives","net_orders","retail"
    ]

    if df.shape[1] != len(expected_columns):
        raise ValueError(f"❌ Column count mismatch. Got {df.shape[1]}, expected {len(expected_columns)}.")

    df.columns = expected_columns

    # ✅ Step 9: Clean email column
    df["user_email"] = df["user_email"].astype(str)
    df = df[df["user_email"].str.contains(
        r"^[a-zA-Z0-9._%+-]+@[A-Za-z0-9.-]+\.[a-zA-Z]{2,}$", na=False
    )]

    # ✅ Step 10: Remove duplicates and reset index
    df = df.drop_duplicates(subset="user_email").reset_index(drop=True)

    print(f"Total records in the cleaned CSV: {len(df)}")

    # ✅ Step 11: Add 'range' column
    df["range"] = range_value

    # ✅ Step 12: Save cleaned data to CSV
    df.to_csv(output_csv, index=False)
    print(f"✅ Cleaned CSV saved at:\n{output_csv}")

    # ✅ Step 13: Save cleaned data to JSON
    df.to_json(output_json, orient="records", indent=2)
    print(f"📤 JSON saved at:\n{output_json}")

    # ✅ API Call
    with open(output_json, 'r', encoding='utf-8') as f:
        json_data = f.read()

    headers = {
        "Content-Type": "application/json"
    }

    response = requests.post(
        "https://api.smartassistapp.in/api/bulk-insert/analytics-records",
        data=json_data,
        headers=headers,
        verify=False
    )

    if response.status_code in [200, 201]:
        print("✅ Data upload successful!")
    else:
        print(f"❌ Upload failed. Status: {response.status_code}")
        print(response.text)

except Exception as e:
    print('error hai kya be : - ', e)

try:

    # driver.switch_to.window(original_window_handle) 
    # print("Original Tableau page par wapas switch kiya gaya.")
    

    # iframe_locator = (By.XPATH, "//iframe[contains(@src, 'tableau')]")
    
    # wait.until(EC.frame_to_be_available_and_switch_to_it(iframe_locator))
    # print("Tableau iframe par successfully switch kiya gaya.")

    # print("Iframe ke andar main 'Download' button search aur click karne ka attempt kar raha hai...")

    new_csv_page_gone = wait.until(
        EC.element_to_be_clickable((By.XPATH, "//span[@id='tableauTabbedNavigation_tab_1']"))
    )
    print("naye csv page me aache se witch hogya hu mai ")
    time.sleep(15)
    new_csv_page_gone.click()
    print("page successfully found hogya hai. well done.")

    time.sleep(15)

    print("all is well till now ")



    download_button = wait.until(
        EC.element_to_be_clickable((By.XPATH, "//button[@id='download']"))
    )
    download_button.click()
    print("Main 'Download' button par click kiya gaya.")

    print("Download menu mein 'Data' option search aur click karne ka attempt kar raha hai...")


    time.sleep(5)
    data_option_button = wait.until(
        EC.element_to_be_clickable((By.XPATH, "//div[@id='viz-viewer-toolbar-download-menu']//span[text()='Crosstab']"))
    )
    data_option_button.click()
    print("'Data' option par click kiya gaya.")

    print("Naye download window/tab ke open hone ka wait kar raha hai...")
    time.sleep(5)
    
    modal_dialog_locator = (By.XPATH, "//div[@role='dialog' and contains(., 'Select Format')]") 
    wait.until(EC.visibility_of_element_located(modal_dialog_locator))
    print("Modal download dialog box successfully appear hua.")

    # print("Modal dialog ke andar 'CSV' format select karne ka attempt kar raha hai...")

    
    try:
        
        csv_radio_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//label[@data-tb-test-id='crosstab-options-dialog-radio-csv-Label']")))
        csv_radio_button.click()
        print("'CSV' format select kiya gaya.")
        time.sleep(5) 
    except TimeoutException:
        print("ERROR: Crosstab dialog (modal) mein 'CSV' radio button specified time ke andar nahi mila ya clickable nahi.")
        raise 
    except Exception as e:
        print(f"ERROR selecting CSV radio button: {e}")
        raise

    
    final_download_button = wait.until(
        EC.element_to_be_clickable((By.XPATH, "//button[@data-tb-test-id='export-crosstab-export-Button']"))
    )
    print("Naye download window mein Dialog mein final 'Download' button mil gaya aur clickable hai.")
    final_download_button.click()
    print("Naye download window mein Dialog mein final 'Download' button par click kiya gaya. Data download initiate hona chahiye.")

    time.sleep(5)
    
    
    # driver.close() 
    print("Download window close kiya gaya.")

except Exception as e:
    print(f"Ek Timeout error hua: {e}")
    print("Iska matlab usually yeh hai ki ek element specified time ke andar nahi mila.")
    if 'driver' in locals():
        driver.quit()


try:
    # 📁 Step 1: Set the folder path where new CSVs will be downloaded
    folder_path = "C:/Users/ssing232/Downloads/"  # 🟡 Change this to your actual download folder

    # ✅ Step 2: Get all CSV files in the folder
    csv_files = [file for file in os.listdir(folder_path) if file.lower().endswith('.csv')]

    if not csv_files:
        raise FileNotFoundError("❌ No CSV file found in the folder!")

    # ✅ Step 3: Pick the most recently created CSV file
    latest_csv_path = max(
        [os.path.join(folder_path, f) for f in csv_files],
        key=os.path.getctime
    )

    print(f"📄 Latest CSV file fetched: {latest_csv_path}")

    # ✅ Step 4: Output file names
    base_dir = os.path.dirname(latest_csv_path)
    base_file_name_with_timestamp = "clean_output"
    output_csv = os.path.join(base_dir, f"{base_file_name_with_timestamp}.csv")
    output_json = os.path.join(base_dir, f"{base_file_name_with_timestamp}.json")

    # ✅ Step 5: Extract 'range' from the file name
    filename = os.path.basename(latest_csv_path)
    # match = re.search(r"SA\s*-\s*(.+?)\.csv", filename, re.IGNORECASE)
    # range_value = match.group(1).strip() if match else "UNKNOWN"
    match = re.search(r"SA\s*-\s*(QTD|MTD|YTD)\s*\(?\d*\)?",filename, re.IGNORECASE)
    range_value = match.group(1).strip() if match else "UNKNOWN"
    filename
    # ✅ Step 6: Read CSV file with fallback encodings
    try:
        df = pd.read_csv(latest_csv_path, encoding='utf-16', sep='\t')
        if df.shape[1] == 1:
            raise ValueError("Fallback to latin1")
    except Exception:
        df = pd.read_csv(latest_csv_path, encoding='latin1', sep='\t')

    print("🔍 Columns detected:", list(df.columns))

    # ✅ Step 7: Drop first 2 and last 5 columns
    if df.shape[1] < 7:
        raise ValueError("❌ Not enough columns to drop first 2 and last 5.")

    df = df.iloc[:, 2:]

    # ✅ Step 8: Rename columns to expected names
    expected_columns = [
        "user_email", "enquiry", "unique_td",
                "enquiry_to_UTD", "enquiry_to_TD", "new_orders",
                "td_to_retail", "cancellations", "cancellation_contro","avg_enq_to_ord_days","dig_enq_to_ord_days",
                "testDrives","net_orders","retail"
    ]

    if df.shape[1] != len(expected_columns):
        raise ValueError(f"❌ Column count mismatch. Got {df.shape[1]}, expected {len(expected_columns)}.")

    df.columns = expected_columns

    # ✅ Step 9: Clean email column
    df["user_email"] = df["user_email"].astype(str)
    df = df[df["user_email"].str.contains(
        r"^[a-zA-Z0-9._%+-]+@[A-Za-z0-9.-]+\.[a-zA-Z]{2,}$", na=False
    )]

    # ✅ Step 10: Remove duplicates and reset index
    df = df.drop_duplicates(subset="user_email").reset_index(drop=True)

    print(f"Total records in the cleaned CSV: {len(df)}")

    # ✅ Step 11: Add 'range' column
    df["range"] = range_value

    # ✅ Step 12: Save cleaned data to CSV
    df.to_csv(output_csv, index=False)
    print(f"✅ Cleaned CSV saved at:\n{output_csv}")

    # ✅ Step 13: Save cleaned data to JSON
    df.to_json(output_json, orient="records", indent=2)
    print(f"📤 JSON saved at:\n{output_json}")


    # ✅ API Call
    with open(output_json, 'r', encoding='utf-8') as f:
        json_data = f.read()

    headers = {
        "Content-Type": "application/json"
    }

    response = requests.post(
        "https://api.smartassistapp.in/api/bulk-insert/analytics-records",
        data=json_data,
        headers=headers,
        verify=False
    )

    if response.status_code in [200, 201]:
        print("✅ Data upload successful!")
    else:
        print(f"❌ Upload failed. Status: {response.status_code}")
        print(response.text)

except Exception as e:
    print('error hai kya be : - ', e)


try:

    # driver.switch_to.window(original_window_handle) 
    # print("Original Tableau page par wapas switch kiya gaya.")
    
    # # --- Start: Handle the 'View Data' modal dialog on the ORIGINAL page ---
    # # print("Original page par 'View Data' modal dialog ke appear hone ka wait kar raha hai...")
    # iframe_locator = (By.XPATH, "//iframe[contains(@src, 'tableau')]")
    
    # wait.until(EC.frame_to_be_available_and_switch_to_it(iframe_locator))
    # print("Tableau iframe par successfully switch kiya gaya.")

    # print("Iframe ke andar main 'Download' button search aur click karne ka attempt kar raha hai...")

    new_csv_page_gone_2 = wait.until(
        EC.element_to_be_clickable((By.XPATH, "//span[@id='tableauTabbedNavigation_tab_2']"))
    )
    print("naye csv page me aache se witch hogya hu mai ")
    time.sleep(15)
    new_csv_page_gone_2.click()
    print("page successfully found hogya hai. well done.")

    time.sleep(15)


    download_button = wait.until(
        EC.element_to_be_clickable((By.XPATH, "//button[@id='download']"))
    )
    download_button.click()
    print("Main 'Download' button par click kiya gaya.")

    print("Download menu mein 'Data' option search aur click karne ka attempt kar raha hai...")
    time.sleep(5)
    data_option_button = wait.until(
        EC.element_to_be_clickable((By.XPATH, "//div[@id='viz-viewer-toolbar-download-menu']//span[text()='Crosstab']"))
    )
    data_option_button.click()
    print("'Data' option par click kiya gaya.")

    print("Naye download window/tab ke open hone ka wait kar raha hai...")
    time.sleep(5)
    modal_dialog_locator = (By.XPATH, "//div[@role='dialog' and contains(., 'Select Format')]") 
    wait.until(EC.visibility_of_element_located(modal_dialog_locator))
    print("Modal download dialog box successfully appear hua.")

    print("Modal dialog ke andar 'CSV' format select karne ka attempt kar raha hai...")

    try:
        
        csv_radio_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//label[@data-tb-test-id='crosstab-options-dialog-radio-csv-Label']")))
        csv_radio_button.click()
        print("'CSV' format select kiya gaya.")
        time.sleep(5) 
    except TimeoutException:
        print("ERROR: Crosstab dialog (modal) mein 'CSV' radio button specified time ke andar nahi mila ya clickable nahi.")
        raise 
    except Exception as e:
        print(f"ERROR selecting CSV radio button: {e}")
        raise

    
    final_download_button = wait.until(
        EC.element_to_be_clickable((By.XPATH, "//button[@data-tb-test-id='export-crosstab-export-Button']"))
    )
    print("Naye download window mein Dialog mein final 'Download' button mil gaya aur clickable hai.")
    final_download_button.click()
    print("Naye download window mein Dialog mein final 'Download' button par click kiya gaya. Data download initiate hona chahiye.")

    time.sleep(5)
    
    
    driver.close() 
    print("teno download hogya successfully.")

except Exception as e:
    print(f"Ek Timeout error hua: {e}")
    print("Iska matlab usually yeh hai ki ek element specified time ke andar nahi mila.")
    if 'driver' in locals():
        driver.quit()



try:
    # 📁 Step 1: Set the folder path where new CSVs will be downloaded
    folder_path = "C:/Users/ssing232/Downloads/"  # 🟡 Change this to your actual download folder

    # ✅ Step 2: Get all CSV files in the folder
    csv_files = [file for file in os.listdir(folder_path) if file.lower().endswith('.csv')]

    if not csv_files:
        raise FileNotFoundError("❌ No CSV file found in the folder!")

    # ✅ Step 3: Pick the most recently created CSV file
    latest_csv_path = max(
        [os.path.join(folder_path, f) for f in csv_files],
        key=os.path.getctime
    )

    print(f"📄 Latest CSV file fetched: {latest_csv_path}")

    # ✅ Step 4: Output file names
    base_dir = os.path.dirname(latest_csv_path)
    base_file_name_with_timestamp = "clean_output"
    output_csv = os.path.join(base_dir, f"{base_file_name_with_timestamp}.csv")
    output_json = os.path.join(base_dir, f"{base_file_name_with_timestamp}.json")

    # ✅ Step 5: Extract 'range' from the file name
    filename = os.path.basename(latest_csv_path)
    # match = re.search(r"SA\s*-\s*(.+?)\.csv", filename, re.IGNORECASE)
    # range_value = match.group(1).strip() if match else "UNKNOWN"
    match = re.search(r"SA\s*-\s*(QTD|MTD|YTD)\s*\(?\d*\)?",filename, re.IGNORECASE)
    range_value = match.group(1).strip() if match else "UNKNOWN"
    filename
    # ✅ Step 6: Read CSV file with fallback encodings
    try:
        df = pd.read_csv(latest_csv_path, encoding='utf-16', sep='\t')
        if df.shape[1] == 1:
            raise ValueError("Fallback to latin1")
    except Exception:
        df = pd.read_csv(latest_csv_path, encoding='latin1', sep='\t')

    print("🔍 Columns detected:", list(df.columns))

    # ✅ Step 7: Drop first 2 and last 5 columns
    if df.shape[1] < 7:
        raise ValueError("❌ Not enough columns to drop first 2 and last 5.")

    df = df.iloc[:, 2:]

    # ✅ Step 8: Rename columns to expected names
    expected_columns = [
        "user_email", "enquiry", "unique_td",
                "enquiry_to_UTD", "enquiry_to_TD", "new_orders",
                "td_to_retail", "cancellations", "cancellation_contro","avg_enq_to_ord_days","dig_enq_to_ord_days",
                "testDrives","net_orders","retail"
    ]

    if df.shape[1] != len(expected_columns):
        raise ValueError(f"❌ Column count mismatch. Got {df.shape[1]}, expected {len(expected_columns)}.")

    df.columns = expected_columns

    # ✅ Step 9: Clean email column
    df["user_email"] = df["user_email"].astype(str)
    df = df[df["user_email"].str.contains(
        r"^[a-zA-Z0-9._%+-]+@[A-Za-z0-9.-]+\.[a-zA-Z]{2,}$", na=False
    )]

    # ✅ Step 10: Remove duplicates and reset index
    df = df.drop_duplicates(subset="user_email").reset_index(drop=True)

    print(f"Total records in the cleaned CSV: {len(df)}")

    # ✅ Step 11: Add 'range' column
    df["range"] = range_value

    # ✅ Step 12: Save cleaned data to CSV
    df.to_csv(output_csv, index=False)
    print(f"✅ Cleaned CSV saved at:\n{output_csv}")

    # ✅ Step 13: Save cleaned data to JSON
    df.to_json(output_json, orient="records", indent=2)
    print(f"📤 JSON saved at:\n{output_json}")

    # ✅ API Call
    with open(output_json, 'r', encoding='utf-8') as f:
        json_data = f.read()

    headers = {
        "Content-Type": "application/json"
    }

    response = requests.post(
        "https://api.smartassistapp.in/api/bulk-insert/analytics-records",
        data=json_data,
        headers=headers,
        verify=False
    )

    if response.status_code in [200, 201]:
        print("✅ Data upload successful!")
    else:
        print(f"❌ Upload failed. Status: {response.status_code}")
        print(response.text)



except Exception as e:
    print('error hai kya be : - ', e)


# except TimeoutException as e:
#     print(f"Ek Timeout error hua: {e}")
#     print("Iska matlab usually yeh hai ki ek element specified time ke andar nahi mila.")
#     if 'driver' in locals():
#         driver.quit()
# except NoSuchElementException as e:
#     print(f"Ek element nahi mila: {e}")
#     print("Iska matlab hai ki ek element ke liye XPath ya locator galat hai ya element page par present nahi hai.")
#     if 'driver' in locals():
#         driver.quit()
# except WebDriverException as e:
#     print(f"Ek WebDriver error hua: {e}")
#     print("Yeh browser issues, network problems, ya incorrect driver setup ke karan ho sakta hai.")
#     if 'driver' in locals():
#         driver.quit()
# except Exception as e:
#     print(f"Ek unexpected error hua: {e}")
#     if 'driver' in locals():
#         driver.quit()

finally:
    if 'driver' in locals():
        print("Sabhi open windows close kar raha hai.")
        driver.quit()
    print("Browser close ho gaya.")



