import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import pyautogui
import time

# Path to the Excel file containing form data
excel_file = 'C:/office work/automation/edit_course.xlsx'

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
        about_course = get_cleaned_data(data_to_fill, 'about_course')
        course_description = get_cleaned_data(data_to_fill, 'course_description')
        how_will_learn = get_cleaned_data(data_to_fill, 'how_will_learn')
        psw_opportunity = get_cleaned_data(data_to_fill, 'psw_opportunity')
        job_opportunity = get_cleaned_data(data_to_fill, 'job_opportunity')
        admission_requirements = get_cleaned_data(data_to_fill, 'admission_requirements')
        eng_level = get_cleaned_data(data_to_fill, 'eng_level')
        eng_level_body = get_cleaned_data(data_to_fill, 'eng_level_body')
        ielts = get_cleaned_data(data_to_fill, 'ielts')
        ielts_value = get_cleaned_data(data_to_fill, 'ielts_value')
        toefl_ibt = get_cleaned_data(data_to_fill, 'toefl_ibt')
        toefl_ibt_value = get_cleaned_data(data_to_fill, 'toefl_ibt_value')
        toefl_cbt = get_cleaned_data(data_to_fill, 'toefl_cbt')
        toefl_cbt_value = get_cleaned_data(data_to_fill, 'toefl_cbt_value')
        pte = get_cleaned_data(data_to_fill, 'pte')
        pte_value = get_cleaned_data(data_to_fill, 'pte_value')
        title = get_cleaned_data(data_to_fill, 'title')
        keyword = get_cleaned_data(data_to_fill, 'keywords')
        description = get_cleaned_data(data_to_fill, 'description')

        # Fill in the form fields
        try:
            # Get all elements with contenteditable="true"
            contenteditable_elements = driver.find_elements(By.CSS_SELECTOR, 'div[contenteditable="true"]')

            short_description_field = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.NAME, 'shortDescription')))
            short_description_field.clear()
            short_description_field.send_keys(about_course)

            # Clear existing data and fill course description (first contenteditable element)
            if len(contenteditable_elements) > 0:
                pyautogui.press('tab')
                pyautogui.hotkey('ctrl', 'a')
                pyautogui.press('delete')
                driver.execute_script("arguments[0].innerHTML = '';", contenteditable_elements[0])
                contenteditable_elements[0].send_keys(course_description)

            # Clear existing data and fill PSW opportunity (second contenteditable element)
            if len(contenteditable_elements) > 1:
                pyautogui.press('tab')
                pyautogui.hotkey('ctrl', 'a')
                pyautogui.press('delete')
                driver.execute_script("arguments[0].innerHTML = '';", contenteditable_elements[1])
                contenteditable_elements[1].send_keys(how_will_learn)

            # Clear existing data and fill job opportunity (third contenteditable element)
            if len(contenteditable_elements) > 2:
                pyautogui.press('tab')
                pyautogui.hotkey('ctrl', 'a')
                pyautogui.press('delete')
                driver.execute_script("arguments[0].innerHTML = '';", contenteditable_elements[2])
                contenteditable_elements[2].send_keys(psw_opportunity)

            # Clear existing data and fill admission requirements (fourth contenteditable element)
            if len(contenteditable_elements) > 3:
                pyautogui.press('tab')
                pyautogui.hotkey('ctrl', 'a')
                pyautogui.press('delete')
                driver.execute_script("arguments[0].innerHTML = '';", contenteditable_elements[3])
                contenteditable_elements[3].send_keys(job_opportunity)

            if len(contenteditable_elements) > 4:
                pyautogui.press('tab')
                pyautogui.hotkey('ctrl', 'a')
                pyautogui.press('delete')
                driver.execute_script("arguments[0].innerHTML = '';", contenteditable_elements[4])
                contenteditable_elements[4].send_keys(admission_requirements)

            title_element = driver.find_element(By.NAME, 'metaTags.title')
            title_element.clear()  # Clear any existing data
            title_element.send_keys(str(title))

            keyword_element = driver.find_element(By.NAME, 'metaTags.keywords')
            keyword_element.clear()  # Clear any existing data
            keyword_element.send_keys(str(keyword))

            description_element = driver.find_element(By.NAME, 'metaTags.description')
            description_element.clear()  # Clear any existing data
            description_element.send_keys(str(description))

            # Submit the form
            submit_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//button[@type='submit' and contains(@class, 'MuiButton-containedPrimary')]")))
            submit_button.click()

            # After submission, confirm and submit again
            submit_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//button[@type='submit' and contains(@class, 'MuiButton-containedPrimary')]")))
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
