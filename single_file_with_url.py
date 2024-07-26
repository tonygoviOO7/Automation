from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException, TimeoutException
import pandas as pd
import time
import pyautogui
 
# Initialize Selenium WebDriver (assuming Chrome)
driver = webdriver.Chrome()

try:
    # Open the website
    driver.get('https://admin.ikounselor.com/login/')   

    # Find the login form fields and fill thempip 
    username_field = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, ':r0:')))
    password_field = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'auth-login-v2-password')))

    username_field.send_keys('admin@ikounselor.com')
    password_field.send_keys('ADc_*6-\\mv_?MJ3')

    # Submit the login form
    password_field.send_keys(Keys.ENTER)

    # Wait for the login process to complete
    WebDriverWait(driver, 30).until(EC.url_contains('dashboard'))
    driver.maximize_window()

    # Define the path to your Excel file
    excel_file = 'Book1.xlsx'
 
    # Load Excel data into a DataFrame
    dataframe = pd.read_excel(excel_file)

    # Iterate over each row in the DataFrame
    for index, row in dataframe.iterrows():
        try:
            # Extract university URL from the first column
            university_url = row[0]

            # Proceed to the university URL page
            driver.get(university_url)

            # Start adding data from the second column onwards
            course_code = str(row.get('course_code', 'N/A')).replace("nan","N/A")
            course_type = str(row.get('course_type', 'N/A'))
            course_name = row.get('course_name', 'N/A')
            course_level = row.get('course_level', 'N/A')
            duration = row.get('duration', 'N/A')
            avg_processing_time = row.get('avg_processing_time', 'N/A')
            tution_fee_currency = row.get('tution_fee_currency', 'N/A')
            tution_fee = str(row.get('tution_fee', 'N/A'))
            application_fee_currency = row.get('application_fee_currency', 'N/A')
            application_fee = str(row.get('application_fee', 'N/A'))
            annual_living_cost_currency = row.get('annual_living_cost_currency', 'N/A')
            annual_living_cost = str(row.get('annual_living_cost', 'N/A'))
            intakes = str(row.get('intakes', '')).replace("nan", "")
            exam_accepted = row.get('exam_accepted', 'N/A')
            about_course = str(row.get('about_course', '')).replace("nan", "")
            course_description = str(row.get('course_description', '')).replace("nan", "")
            how_learn = str(row.get('how_learn', '')).replace("nan", "")
            psw_opportunity = str(row.get('job_opportunity', '')).replace("nan", "")
            job_opportunity = str(row.get('job_opportunity', '')).replace("nan", "")
            admission_requirements = str(row.get('admission_requirements', '')).replace("nan", "")
            eng_level = row.get('eng_level', 'N/A')
            eng_level_body = row.get('eng_level_body', 'N/A')
            ielts = row.get('ielts', 'N/A')
            ielts_value = str(row.get('ielts_value', 'N/A'))
            toefl_ibt = row.get('toefl_ibt', 'N/A')
            toefl_ibt_value = str(row.get('toefl_ibt_value', 'N/A'))
            toefl_cbt = row.get('toefl_cbt', 'N/A')
            toefl_cbt_value = str(row.get('toefl_cbt_value', 'N/A'))
            pte = row.get('pte', 'N/A')
            pte_value = str(row.get('pte_value', 'N/A'))
            title = str(row.get('title', '')).replace("nan", "")
            keyword = str(row.get('keyword', '')).replace("nan", "")
            description = str(row.get('description', '')).replace("nan", "")

            # Fill in form fields
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.NAME, 'code'))).send_keys(course_code)
            driver.find_element(By.NAME, 'type').send_keys(course_type)
            driver.find_element(By.NAME, 'name').send_keys(course_name)

            # Select course level from dropdown
            level_dropdown = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, 'mui-component-select-level')))
            level_dropdown.click()
            level_option_xpath = f"//li[contains(text(), '{course_level}')]"
            level_option = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, level_option_xpath)))
            level_option.click()

            driver.find_element(By.NAME, 'duration').clear()
            driver.find_element(By.NAME, 'duration').send_keys(duration)

            driver.find_element(By.NAME, 'processingTime').clear()
            driver.find_element(By.NAME, 'processingTime').send_keys(avg_processing_time)

            # driver.find_element(By.NAME, 'duration').send_keys(duration)
            # driver.find_element(By.NAME, 'processingTime').send_keys(avg_processing_time)

            # Select tuition fee currency from dropdown
            if tution_fee_currency:
                tuition_fee_dropdown = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.NAME, 'tuitionFee.currency')))
                tuition_fee_dropdown.click()
                tuition_fee_option_xpath = f"//li[contains(text(), '{tution_fee_currency}')]"
                tuition_fee_option = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, tuition_fee_option_xpath)))
                tuition_fee_option.click()

            driver.find_element(By.NAME, 'tuitionFee.amount').send_keys(tution_fee)

            # Select application fee currency from dropdown
            if application_fee_currency:
                application_fee_dropdown = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.NAME, 'applicationFee.currency')))
                application_fee_dropdown.click()
                application_fee_option_xpath = f"//li[contains(text(), '{application_fee_currency}')]"
                application_fee_option = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, application_fee_option_xpath)))
                application_fee_option.click()

            driver.find_element(By.NAME, 'applicationFee.amount').send_keys(application_fee)

            # Select annual living cost currency from dropdown
            if annual_living_cost_currency:
                living_cost_dropdown = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.NAME, 'annualLivingCost.currency')))
                living_cost_dropdown.click()
                living_cost_option_xpath = f"//li[contains(text(), '{annual_living_cost_currency}')]"
                living_cost_option = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, living_cost_option_xpath)))
                living_cost_option.click()

            driver.find_element(By.NAME, 'annualLivingCost.amount').send_keys(annual_living_cost)

            driver.find_element(By.NAME, 'intakes').send_keys(intakes)

            # Select exam accepted from dropdown
            exam_dropdown = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, 'mui-component-select-exams')))
            exam_dropdown.click()
            exam_option_xpath = f"//li[contains(text(), '{exam_accepted}')]"
            exam_option = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, exam_option_xpath)))
            exam_option.click()
            time.sleep(.5)
            pyautogui.press('tab')

            # Fill in other fields (short_description, course_description, psw_opportunity, job_opportunity, admission_requirements, english_language_req)
            contenteditable_elements = driver.find_elements(By.CSS_SELECTOR, 'div[contenteditable="true"]')

            short_description_field = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.NAME, 'shortDescription')))
            short_description_field.clear()
            short_description_field.send_keys(about_course)

            # Clear existing data and fill course description (first contenteditable element)
            driver.execute_script("arguments[0].innerHTML = '';", contenteditable_elements[0])
            contenteditable_elements[0].send_keys(course_description)

            # Clear existing data and fill PSW opportunity (second contenteditable element)
            driver.execute_script("arguments[0].innerHTML = '';", contenteditable_elements[1])
            contenteditable_elements[1].send_keys(how_learn)

            # Clear existing data and fill how learn (third contenteditable element)    
            driver.execute_script("arguments[0].innerHTML = '';", contenteditable_elements[2])
            contenteditable_elements[2].send_keys(psw_opportunity)

            # Clear existing data and fill job opportunity (fourth contenteditable element)
            driver.execute_script("arguments[0].innerHTML = '';", contenteditable_elements[3])
            contenteditable_elements[3].send_keys(job_opportunity)

            # Clear existing data and fill admission requirements (fifth contenteditable element)
            driver.execute_script("arguments[0].innerHTML = '';", contenteditable_elements[4])
            contenteditable_elements[4].send_keys(admission_requirements)


            button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, '.MuiButton-root.MuiButton-tonalSuccess')))
            button.click()

            # Handling English Requirements
            driver.find_element(By.NAME, 'minEnglishRequirements.0.title').clear()
            driver.find_element(By.NAME, 'minEnglishRequirements.0.title').send_keys(eng_level)

            driver.find_element(By.NAME, 'minEnglishRequirements.0.body').clear()
            driver.find_element(By.NAME, 'minEnglishRequirements.0.body').send_keys(eng_level_body)

            button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, '.MuiButton-root.MuiButton-tonalSuccess')))
            button.click()

            driver.find_element(By.NAME, 'minEnglishRequirements.1.title').clear()
            driver.find_element(By.NAME, 'minEnglishRequirements.1.title').send_keys(ielts)
                        
            driver.find_element(By.NAME, 'minEnglishRequirements.1.body').clear()
            driver.find_element(By.NAME, 'minEnglishRequirements.1.body').send_keys(ielts_value)

            button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, '.MuiButton-root.MuiButton-tonalSuccess')))
            button.click()

            driver.find_element(By.NAME, 'minEnglishRequirements.2.title').clear()
            driver.find_element(By.NAME, 'minEnglishRequirements.2.title').send_keys(toefl_ibt)
                        
            driver.find_element(By.NAME, 'minEnglishRequirements.2.body').clear()
            driver.find_element(By.NAME, 'minEnglishRequirements.2.body').send_keys(toefl_ibt_value)

            button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, '.MuiButton-root.MuiButton-tonalSuccess')))
            button.click()

            driver.find_element(By.NAME, 'minEnglishRequirements.3.title').clear()
            driver.find_element(By.NAME, 'minEnglishRequirements.3.title').send_keys(toefl_cbt)
                        
            driver.find_element(By.NAME, 'minEnglishRequirements.3.body').clear()
            driver.find_element(By.NAME, 'minEnglishRequirements.3.body').send_keys(toefl_cbt_value)

            button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, '.MuiButton-root.MuiButton-tonalSuccess')))
            button.click()

            driver.find_element(By.NAME, 'minEnglishRequirements.4.title').clear()
            driver.find_element(By.NAME, 'minEnglishRequirements.4.title').send_keys(pte)
                        
            driver.find_element(By.NAME, 'minEnglishRequirements.4.body').clear()
            driver.find_element(By.NAME, 'minEnglishRequirements.4.body').send_keys(pte_value)

            title_field = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.NAME, 'metaTags.title')))
            title_field.clear()
            title_field.send_keys(title)

            keywords_field = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.NAME, 'metaTags.keywords')))
            keywords_field.clear()
            keywords_field.send_keys(keyword)

            description_field = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.NAME, 'metaTags.description')))
            description_field.clear()
            description_field.send_keys(description)

            # Submit the form
            submit_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, '.MuiButton-containedPrimary')))
            submit_button.click()

            # Wait for submission to be processed (adjust as needed)
            time.sleep(2)

            print(f"Successfully processed row {index} for URL {university_url}")

            # Assuming submission is successful, delete the row from Excel
            dataframe.drop(index, inplace=True)
            dataframe.to_excel(excel_file, index=False)

        except (NoSuchElementException, TimeoutException) as e:
            print(f"Error processing row {index}: {e}")

finally:
    # Close the browser session
    driver.quit()
