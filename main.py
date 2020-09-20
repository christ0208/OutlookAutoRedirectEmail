from selenium import webdriver
from ExcelReader import ExcelReader
from OutlookWebLogin import OutlookWebLogin
from OutlookWebSettings import OutlookWebSettings

CURRENT_USER_AGENT = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/80.0.3987.149 Safari/537.36'
chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument('--incognito')
chrome_options.add_argument('--user-agent=' + CURRENT_USER_AGENT)
chrome_options.add_argument('--ignore-ssl-errors=yes')
chrome_options.add_argument('--ignore-certificate-errors')
chrome_options.add_argument('--start-maximized')

chrome_driver = None


def prep_chrome_driver():
    global chrome_driver
    chrome_driver = webdriver.Chrome(chrome_options=chrome_options)

if __name__ == '__main__':
    excel_reader = ExcelReader('./data.xlsx')
    collected_data = excel_reader.read()

    for data in collected_data:
        prep_chrome_driver()
    
        outlook_web_login = OutlookWebLogin(chrome_driver)
        outlook_web_login.redirect_login_page('https://outlook.live.com/owa/?nlp=1')
        outlook_web_login.do_login({"email": data["targetEmail"], "password": data["targetPassword"]})
        outlook_web_login.bypass_stay_signed_in()
    
        outlook_web_settings = OutlookWebSettings(chrome_driver)
        outlook_web_settings.set_rule_settings('https://outlook.office365.com/mail/options/mail/rules', {"studentEmail": data["studentEmail"], "from": data["from"], "subject": data["subject"]})
        outlook_web_settings.set_junk_email_settings('https://outlook.office365.com/mail/options/mail/junkEmail', {"excludeDomain": data["excludeDomain"]})
    
        chrome_driver.close()

        print(data["targetEmail"] + ' done setting')


