from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException, ElementClickInterceptedException
import time
import csv

import creds

driver = webdriver.Chrome()
driver.implicitly_wait(15)

def login_to_clumio(url, admin_username, admin_password):
    # Initialize WebDriver for Chrome
    driver = webdriver.Chrome()
    driver.implicitly_wait(15)
    restore_user = "Adele Vance"

    try:
        # Navigate to the website
        driver.get(url)

        driver.find_element(By.CSS_SELECTOR, "#clumio-email").send_keys(admin_username)

        driver.find_element(By.CSS_SELECTOR, "button[type='submit']").click()

        # Enter password and click 'Log In'
        password_field = driver.find_element(By.CSS_SELECTOR, "#clumio-password")
        password_field.send_keys(admin_password)
        login_button = driver.find_element(By.CSS_SELECTOR, "button[type='submit']")
        login_button.click()

        # Wait for the page to load after login
        time.sleep(10)

    except NoSuchElementException as e:
        print("Element not found:", e)
    except ElementClickInterceptedException as e:
        print("Element not clickable:", e)
    finally:
        # Close the browser (comment out if you want to keep it open)
        # driver.quit()
        pass

def restore_exchange_mailbox(current_user, restore_location, date):

    calendar_selection = f"//div[@id='month-calendar-{date}']//button[@class='cl-month-dynamic-calendar__body__date']"

    try:

        # Select User mailbox
        driver.find_element(By.LINK_TEXT, current_user).click()

        # Select date to restore 
        driver.find_element(By.XPATH, calendar_selection).click()

        driver.find_element(By.XPATH, "(//button[@id='restore-btn'])[1]").click()

        # Enter the restore mailbox location
        restore_mailbox = driver.find_element(By.XPATH, "//input[@id='mailboxName']")
        restore_mailbox.click()

        time.sleep(5)
        restore_mailbox.send_keys(Keys.CONTROL, 'a', Keys.BACKSPACE, restore_location, Keys.ENTER)

        # Enter the restore folder location name it with the current user being restored
        username_field = driver.find_element(By.XPATH, "//input[@id='folderName']")
        restore_mailbox.click()
        username_field.send_keys(current_user, " - Restored")

        # Disable Contact Restore
        disable_contact = driver.find_element(By.XPATH, "//label[@for='restore-contacts']")
        disable_contact.click()

        # Disable Calendar Restore
        disable_calendar = driver.find_element(By.XPATH, "//label[@for='restore-calendars']")
        disable_calendar.click()

        # Start Restore
        restore_exchange = driver.find_element(By.XPATH,
                                               "//button[@class='clumio-button-v1 primary default size-regular']")
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


# log into clumio using user and pass stored in a creds module that doesn't sync to git.
# using this method with non secure test environment.
login_to_clumio("https://portal.clumio.com/", creds.clumio_user, creds.clumio_pass)

user_list = input("Enter path to user list CSV: ")
restore_mailbox_location = input("Enter the name of the shared mailbox to restore items to: ")
restore_date = input("Enter the date to restore in format DD-MM-YYYY: ")
users = []

with open(user_list, 'r') as csvfile:
    csv_reader = csv.DictReader(user_list)

    for row in csv_reader:
        users.append(row)


for user_to_restore in users:
    restore_exchange_mailbox(user_to_restore, restore_mailbox_location, restore_date)
