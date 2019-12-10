import re
import secrets
import shutil
import string
import time
from datetime import date
from os import path

import openpyxl
from Pass import Pass
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select


class ResellerOnboarder:
    def __init__(self):
        self.logbook = []  # tuples of the form (row, username, password)
        self.logins = Pass()

        # Set up webdriver
        self.driver = webdriver.Chrome(r"Drivers/chromedriver.exe")
        self.driver.maximize_window()
        self.driver.implicitly_wait(10)

    def passwordGenerator(self):
        punctuation = "!@#$%^&*-+?"
        characters = string.ascii_letters + string.digits + punctuation
        password = ""
        while self.passwordCheck(password) is not True:
            password = ""
            for x in range(8):
                password += secrets.choice(characters)
        return password

    def passwordCheck(self, password):
        if re.search(r"[!@#$%^&*\-+?]", password) and re.search(r"[A-Z]", password) and re.search(r"[a-z]",password) and re.search(r"[0-9]", password):
            return True
        return False

    def ims(self):
        self.driver.get(r"http://ims.buroserv.com.au/login.php")
        username = self.driver.find_element_by_id("login-username")
        password = self.driver.find_element_by_id("login-password")
        username.send_keys(self.logins.ims[0])
        password.send_keys(self.logins.ims[1])
        self.driver.find_element_by_id("btn-login").click()

    def supatools(self):
        self.driver.get(r"https://support.viptelecombilling.net.au/login.php")
        username = self.driver.find_element_by_name("user_id")
        password = self.driver.find_element_by_name("password")
        username.send_keys(self.logins.supatools[0])
        password.send_keys(self.logins.supatools[1])
        password.send_keys(Keys.ENTER)


    def porta(self):
        self.driver.get(r"https://billing.isphone.com.au/index.html")
        username = self.driver.find_element_by_id("pb_auth_user")
        password = self.driver.find_element_by_id("pb_auth_password")
        username.send_keys(self.logins.porta[0])
        password.send_keys(self.logins.porta[1])
        password.send_keys(Keys.ENTER)

        self.driver.find_element_by_link_text("Resellers").click()
        self.driver.find_element_by_class_name("add").click()
        customerID = self.driver.find_element_by_xpath('//*[@id="PortaBilling_Tabs_Master_Container"]/tbody/tr[1]/td/table/tbody/tr/td[2]/table/tbody/tr[1]/td[2]/span/input')
        companyName = self.driver.find_element_by_xpath('//*[@id="addressInfoTab"]/tbody/tr/td[1]/table/tbody/tr[1]/td[2]/span/input')
        firstName = self.driver.find_element_by_id("firstname_field_id")
        lastName = self.driver.find_element_by_id("lastname_field_id")
        country = Select(self.driver.find_element_by_id("country"))
        address1 = self.driver.find_element_by_id("address_field_id_1")
        city = self.driver.find_element_by_id("city_field_id")
        state = Select(self.driver.find_element_by_id("state"))
        postcode = self.driver.find_element_by_id("zip_field_id")
        phone = self.driver.find_element_by_class_name("phone1")
        email = self.driver.find_element_by_id("email_field_id")

        customerID.send_keys()
        companyName.send_keys()
        firstName.send_keys()
        lastName.send_keys()
        country.select_by_visible_text("Australia")
        address1.send_keys()
        city.send_keys()
        state.select_by_visible_text()
        postcode.send_keys()
        phone.send_keys()
        email.send_keys()


    def teardown(self):
        self.driver.quit()


def main():



if __name__ == "__main__":
    main()
