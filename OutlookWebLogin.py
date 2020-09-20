from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.support.ui import WebDriverWait


class OutlookWebLogin:
    def __init__(self, driver):
        self.driver = driver

    def redirect_login_page(self, url):
        self.driver.get(url)

    def do_login(self, credential):
        WebDriverWait(self.driver, 10).until(expected_conditions.presence_of_element_located((By.TAG_NAME, "input")))

        self.driver.find_element_by_xpath("//input[@name='loginfmt']").send_keys(credential['email'])
        self.driver.find_element_by_xpath("//input[@name='loginfmt']").send_keys(Keys.ENTER)

        WebDriverWait(self.driver, 10).until(
            expected_conditions.presence_of_element_located((By.XPATH, "//input[@name='Password']"))
        )
        self.driver.find_element_by_xpath("//input[@name='Password']").send_keys(credential['password'])
        self.driver.find_element_by_xpath("//input[@name='Password']").send_keys(Keys.ENTER)

    def bypass_stay_signed_in(self):
        WebDriverWait(self.driver, 10).until(expected_conditions.presence_of_element_located((By.TAG_NAME, "input")))

        self.driver.find_element_by_xpath("//input[@value='No']").click()

