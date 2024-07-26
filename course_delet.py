import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Initialize Selenium WebDriver (assuming Chrome)
driver = webdriver.Chrome()

# Open the website
driver.get('https://admin.ikounselor.com/login/')

# Find the login form fields and fill them
username_field = driver.find_element(By.ID, ':r0:')
password_field = driver.find_element(By.ID, 'auth-login-v2-password')

# Fill in your credentials
username_field.send_keys('admin@ikounselor.com')
password_field.send_keys('ADc_*6-\\mv_?MJ3')

# Submit the login form
password_field.submit()

# Wait for the login process to complete
WebDriverWait(driver, 30).until(EC.url_contains('dashboard'))
driver.maximize_window()

# Navigate to the specific university URL
university_url = 'https://admin.ikounselor.com/universities/660bc1b22e71bf43040d8d6c/'
driver.get(university_url)

# Wait for the page to load completely
WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.TAG_NAME, 'body')))

# Repeat the following block of code 3 times with a time delay between each iteration
for _ in range(81):
    try:
        button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[1]/div[2]/main/div/div[2]/div[2]/div[2]/div/div/div/div[1]/div[7]/div/button")))
        button.click()
    except Exception as e:
        print("Error occurred while clicking button:", e)

    #time.sleep(2)  # Add a time delay of 2 seconds between the first and second action

    try:
        submit_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[3]/div/div[2]/button[2]")))
        submit_button.click()
    except Exception as e:
        print("Error occurred while clicking submit button:", e)

    time.sleep(1)  # Add a time delay of 3 seconds between the second and third action

# button = driver.find_element(By.XPATH, "/html/body/div[1]/div[1]/div[2]/main/div/div[1]/div[2]/div/a[2]/button")
# time.sleep(2)                           
#         button.click()








# Close the browser after completing the tasks
driver.quit()
