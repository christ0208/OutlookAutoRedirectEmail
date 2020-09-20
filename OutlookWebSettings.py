from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import StaleElementReferenceException
import time


class OutlookWebSettings:
    def __init__(self, driver):
        self.driver = driver

    def set_rule_settings(self, url, data):
        self.driver.get(url)

        WebDriverWait(self.driver, 15).until(expected_conditions.visibility_of_element_located(
            (By.XPATH, "//*[contains(text(), 'Add new rule')]")
        ))
        time.sleep(3)
        try:
            delete_button_rules = self.driver.find_elements_by_xpath("//button[@title='Delete rule']")

            for rule in delete_button_rules:
                rule.click()
                WebDriverWait(self.driver, 10).until(expected_conditions.visibility_of_element_located(
                    (By.XPATH, "//button[starts-with(@id,'ok-')]")
                ))
                self.driver.find_element_by_xpath(
                    "//button[starts-with(@id,'ok-')]"
                ).send_keys(Keys.ENTER)
        except NoSuchElementException:
            pass

        self.driver.find_element_by_xpath(
            "//*[contains(text(), 'Add new rule')]"
        ).click()

        WebDriverWait(self.driver, 10).until(lambda driver: driver.execute_script('return document.readyState') == 'complete')
        WebDriverWait(self.driver, 10).until(expected_conditions.visibility_of_element_located(
            (By.XPATH, "//span[contains(text(), 'Select an action')]")
        ))

        self.driver.execute_script('document.querySelector("input[placeholder=\'Search settings\']").style.display=\'none\'')
        self.driver.implicitly_wait(5)
        while(True):
            name_input = self.driver.find_element_by_xpath("//input[@placeholder='Name your rule']")
            name_input.send_keys(data["studentEmail"])
            time.sleep(1)
            current_value = name_input.get_attribute('value')

            if current_value == data['studentEmail']:
                break

        while(True):
            try:
                self.driver.find_element_by_xpath("//button[@title='Redirect to']")

                break
            except NoSuchElementException:
                self.driver.find_element_by_xpath(
                    "//span[contains(text(), 'Select an action')]"
                ).click()
            time.sleep(1)

        self.driver.find_element_by_xpath("//button[@title='Redirect to']").click()
        WebDriverWait(self.driver, 10).until(expected_conditions.visibility_of_element_located(
            (By.XPATH, "//input[@role='combobox'][@class='ms-BasePicker-input pickerInput_9afdc73b']")
        ))

        while (True):
            name_input = lambda: self.driver.find_element_by_xpath("//input[@role='combobox'][@class='ms-BasePicker-input pickerInput_9afdc73b']")
            name_input().send_keys(Keys.NULL)
            self.driver.implicitly_wait(1)
            name_input().send_keys(data["studentEmail"])
            time.sleep(2)
            current_value = name_input().get_attribute('value')

            if current_value == data['studentEmail']:
                break

        self.driver.find_element_by_xpath(
            "//input[@role='combobox']"
        ).send_keys(Keys.TAB)

        self.driver.find_element_by_xpath(
            "//*[contains(text(), 'Select a condition')]"
        ).click()
        WebDriverWait(self.driver, 10).until(expected_conditions.visibility_of_element_located(
            (By.XPATH,
             "//button[@title='From']")
        ))
        self.driver.find_element_by_xpath("//button[@title='From']").click()
        WebDriverWait(self.driver, 10).until(expected_conditions.visibility_of_element_located(
            (By.XPATH, "//input[@role='combobox'][@class='ms-BasePicker-input pickerInput_9afdc73b']")
        ))

        while (True):
            name_input = self.driver.find_element_by_xpath("//input[@role='combobox'][@class='ms-BasePicker-input pickerInput_9afdc73b']")
            name_input.send_keys(Keys.NULL)
            self.driver.implicitly_wait(1)
            name_input.send_keys(data['from'])
            time.sleep(2)
            current_value = name_input.get_attribute('value')

            if current_value == data['from']:
                break

        self.driver.find_element_by_xpath(
            "//input[@role='combobox']"
        ).send_keys(Keys.TAB)
        self.driver.find_element_by_xpath(
            "//*[contains(text(), 'Add another condition')]"
        ).click()
        WebDriverWait(self.driver,10).until(expected_conditions.visibility_of_element_located(
            (By.XPATH, "//*[contains(text(), 'And...')]")
        ))
        self.driver.find_element_by_xpath(
            "//*[contains(text(), 'And...')]"
        ).click()
        self.driver.find_element_by_xpath("//button[@title='Subject includes']").click()
        WebDriverWait(self.driver, 10).until(expected_conditions.visibility_of_element_located(
            (By.XPATH, "//input[@placeholder='Enter words to look for']")
        ))

        while (True):
            name_input = self.driver.find_element_by_xpath("//input[@placeholder='Enter words to look for']")
            name_input.send_keys(Keys.NULL)
            self.driver.implicitly_wait(1)
            name_input.send_keys(data['subject'])
            time.sleep(2)
            current_value = name_input.get_attribute('value')

            if current_value == data['subject']:
                break

        self.driver.find_element_by_xpath("//*[contains(text(), 'Stop processing more rules')]").click()

        self.driver.find_element_by_xpath(
            "//*[contains(text(), 'Save')]"
        ).click()

    def set_junk_email_settings(self, url, data):
        self.driver.get(url)

        WebDriverWait(self.driver, 15).until(expected_conditions.visibility_of_element_located(
            (By.XPATH, "//button[contains(@class,'_3B7DMxWQgF7_ejwoA86ngO')]")
        ))

        try:
            delete_button_rules = self.driver.find_elements_by_xpath("//button[contains(@class, 'ms-Button ms-Button--icon')]")

            for rule in delete_button_rules:
                if "Delete" in rule.get_attribute('innerHTML'):
                    rule.click()

        except NoSuchElementException:
            pass

        self.driver.find_elements_by_xpath("//button[contains(@class,'_3B7DMxWQgF7_ejwoA86ngO')]")[1].click()
        WebDriverWait(self.driver, 15).until(expected_conditions.visibility_of_element_located(
            (By.XPATH, "//input[contains(@placeholder, 'Example:')]")
        ))
        self.driver.find_element_by_xpath("//input[contains(@placeholder, 'Example:')]").send_keys(data['excludeDomain'])
        self.driver.find_element_by_xpath("//input[contains(@placeholder, 'Example:')]").send_keys(Keys.ENTER)

        self.driver.implicitly_wait(2)
        try:
            self.driver.find_element_by_xpath("//span[contains(text(), 'Save')]").click()
        except StaleElementReferenceException:
            pass

        self.driver.implicitly_wait(2)


