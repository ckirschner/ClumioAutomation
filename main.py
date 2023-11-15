from selenium import webdriver
from selenium import By
from selenium import Keys
from selenium.common.exceptions import NoSuchElementException, ElementClickInterceptedException
import time

import creds


def login_and_select_microsoft_365(url, username, password):
    # Initialize WebDriver for Chrome
    driver = webdriver.Chrome()

    try:
        # Navigate to the website
        driver.get(url)

        # Wait for the login elements to load
        time.sleep(2)

        # Enter username in the field with label 'clumio-email' to enter email
        username_field = driver.find_element(By.CSS_SELECTOR, "#clumio-email")
        username_field.send_keys(username)

        # Find and click the 'Next' button to submit username
        next_button = driver.find_element(By.CSS_SELECTOR, "button[type='submit']")
        next_button.click()

        # Wait for password field
        time.sleep(2)

        # Enter password and click 'Log In'
        password_field = driver.find_element(By.CSS_SELECTOR, "#clumio-password")
        password_field.send_keys(password)
        login_button = driver.find_element(By.CSS_SELECTOR, "button[type='submit']")
        login_button.click()

        # Wait for the page to load after login
        time.sleep(10)

        # Find and click the 'Microsoft 365' button to open M365 backups
        ms365_button = driver.find_element(By.CSS_SELECTOR, "a[title='Microsoft 365']")
        ms365_button.click()

        # Wait for the next element to load
        time.sleep(5)

        # Find and click the button for the organization backup
        onmicrosoft_button = driver.find_element(By.CSS_SELECTOR, "a[title='M365x72807442.onmicrosoft.com']")
        onmicrosoft_button.click()

        # Wait for the next element to load
        time.sleep(5)

        # Click on the Exchange Tab
        exchange_tab = driver.find_element(By.XPATH, "//div[contains(text(),'Exchange')]")
        exchange_tab.click()

        time.sleep(5)

        # Click first username
        user = driver.find_element(By.LINK_TEXT, "Adele Vance")
        user.click()

        time.sleep(8)

        # Click on the date to restore from
        restore_date = driver.find_element(By.XPATH, "//div[@id='month-calendar-14-10-2023']//button[@class='cl-month-dynamic-calendar__body__date']")
        restore_date.click()

        time.sleep(5)

        # Click on the restore exchange button (Needs to be reworked to find must recent exchange backup)
        restore_exchange = driver.find_element(By.XPATH, "(//button[@id='restore-btn'])[2]")
        restore_exchange.click()

        time.sleep(5)

        # Enter the restore mailbox location
        restore_mailbox = driver.find_element(By.XPATH , "//input[@id='mailboxName']")
        restore_mailbox.click()
        time.sleep(1)
        restore_mailbox.send_keys(Keys.CONTROL, 'a')
        time.sleep(1)
        restore_mailbox.send_keys(Keys.BACKSPACE)
        time.sleep(2)
        restore_mailbox.send_keys("MOD Administrator")
        time.sleep(3)
        restore_mailbox.send_keys(Keys.ENTER)

        time.sleep(2)

        # Enter the restore folder location
        username_field = driver.find_element(By.XPATH, "//input[@id='folderName']")
        restore_mailbox.click()
        time.sleep(1)
        username_field.send_keys("AdeleV Restore1")

        time.sleep(2)

        # Disable Contact Restore
        disable_contact = driver.find_element(By.XPATH, "//label[@for='restore-contacts']")
        disable_contact.click()

        time.sleep(2)

        # Disable Calendar Restore
        disable_calendar = driver.find_element(By.XPATH, "//label[@for='restore-calendars']")
        disable_calendar.click()

        time.sleep(5)

        # Start Restore
        restore_exchange = driver.find_element(By.XPATH, "//button[@class='clumio-button-v1 primary default size-regular']")
        restore_exchange.click()

        time.sleep(25)

    except NoSuchElementException as e:
        print("Element not found:", e)
    except ElementClickInterceptedException as e:
        print("Element not clickable:", e)
    finally:
        # Close the browser (comment out if you want to keep it open)
        # driver.quit()
        pass

# Use the function
# log into clumio using user and pass stored in a creds module that doesn't sync to git.
login_and_select_microsoft_365("https://portal.clumio.com/", creds.clumio_user, creds.clumio_pass)
