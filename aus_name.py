import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import pyautogui
import time

# Path to the Excel file containing form data
excel_file = 'C:/office work/automation/aus_name.xlsx'

# Read form data from the Excel file
form_data_df = pd.read_excel(excel_file)

# Initialize the WebDriver (make sure the path to your WebDriver is correct)
driver = webdriver.Chrome()
driver.maximize_window()

def get_cleaned_data(data_to_fill, column_name):
    value = data_to_fill.get(column_name, '')
    if isinstance(value, float) and pd.isna(value):
        return ''
    return str(value).replace("nan", "")

try:
    # Open the website
    driver.get('https://admin.ikounselor.com/login/')

    # Find the login form fields and fill them
    username_field = driver.find_element(By.ID, ':r0:')
    password_field = driver.find_element(By.ID, 'auth-login-v2-password')

    # Fill in your credentials
    username_field.send_keys('admin@ikounselor.com')
    password_field.send_keys('ADc_*6-\mv_?MJ3')

    # Submit the login form
    password_field.send_keys(Keys.ENTER)

    # Wait for the login process to complete
    WebDriverWait(driver, 30).until(EC.url_contains('dashboard'))

    # Loop through each row in the DataFrame
    for index, row in form_data_df.iterrows():
        url = row['url']
        
        # Get the row data excluding the URL column
        data_to_fill = row.drop(labels=['url']).to_dict()

        # Open the URL
        driver.get(url)
        
        # Add a small delay to ensure the page loads completely
        time.sleep(1)
        
        # Extract form data
        country = row.get('country', 'N/A')
        # title = get_cleaned_data(data_to_fill, 'title')
        # keyword = get_cleaned_data(data_to_fill, 'keywords')
        # description = get_cleaned_data(data_to_fill, 'description')

        # Fill in the form fields
        try:
            # Get all elements with contenteditable="true"

            exam_dropdown = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, 'mui-component-select-country')))
            exam_dropdown.click()
            exam_option_xpath = f"//li[contains(text(), '{country}')]"
            exam_option = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, exam_option_xpath)))
            time.sleep(1)
            exam_option.click()
            time.sleep(1)
            # pyautogui.press('tab')

            # title_element = driver.find_element(By.NAME, 'metaTags.title')
            # title_element.clear()  # Clear any existing data
            # title_element.send_keys(str(title))

            # keyword_element = driver.find_element(By.NAME, 'metaTags.keywords')
            # keyword_element.clear()  # Clear any existing data
            # keyword_element.send_keys(str(keyword))

            # description_element = driver.find_element(By.NAME, 'metaTags.description')
            # description_element.clear()  # Clear any existing data
            # description_element.send_keys(str(description))

            # Submit the form
            submit_button = driver.find_element(By.CSS_SELECTOR, '.MuiButtonBase-root.MuiButton-root.MuiButton-contained.MuiButton-containedPrimary.MuiButton-sizeLarge.MuiButton-containedSizeLarge.css-wve6q6')
            submit_button.click()

            print(f"Form submitted successfully for URL: {url}")

            # Delete the row from the DataFrame after successful submission
            form_data_df.drop(index, inplace=True)

            # Save the updated DataFrame back to the Excel file
            form_data_df.to_excel(excel_file, index=False)

        except Exception as e:
            print(f"Error filling or submitting form for URL {url}: {str(e)}")

        # Add a small delay before proceeding to the next URL
        time.sleep(2)

finally:
    # Close the WebDriver
    driver.quit()
