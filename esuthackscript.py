import os
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

def login_and_get_profile_info(username, password, chromedriver_path):
    # Set the path to the chromedriver
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

            # Return the extracted data
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
                'DateOfBirth': date_of_birth
            }
        else:
            print(f"Login failed for {username}. Skipping...")
            return None
    finally:
        driver.quit()

def process_accounts(input_file, output_file, chromedriver_path):
    # Read accounts data from Excel
    accounts_df = pd.read_excel(input_file)

    # Open the CSV file for writing (overwrite mode)
    with open(output_file, 'w', newline='') as csvfile:
        # Initialize CSV writer
        header = ['Username', 'Password', 'Surname', 'Firstname', 'Middlename', 'Phone', 'MatricNumber', 
                  'Department', 'Email', 'DateOfBirth']
        
        writer = pd.DataFrame(columns=header)
        writer.to_csv(csvfile, header=True, index=False)

        # Process each account and save the profile immediately after processing
        for index, row in accounts_df.iterrows():
            username = row['username']
            password = row['password']

            profile_info = login_and_get_profile_info(username, password, chromedriver_path)

            if profile_info is not None:
                # Write the valid profile info to the CSV file
                profile_info_df = pd.DataFrame([profile_info])
                profile_info_df.to_csv(csvfile, header=False, index=False)  # Append data
                csvfile.flush()  # Force the data to be written immediately
                print(f"Saved profile info for {username}")

    print(f"Results saved to {output_file}")

if __name__ == "__main__":
    input_file = "/path/to/your/input.xlsx"  # Input Excel file with usernames and passwords
    output_file = "/path/to/output.csv"  # Output CSV file to save the profiles
    chromedriver_path = "/path/to/chromedriver"  # Path to the ChromeDriver binary

    process_accounts(input_file, output_file, chromedriver_path)
