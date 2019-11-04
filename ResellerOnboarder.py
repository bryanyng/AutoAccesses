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

        time.sleep(2)
        self.driver.get(r"http://ims.buroserv.com.au/users.php")
        self.driver.find_element_by_link_text("Add User").click()
        email = self.driver.find_element_by_id("adduser-email")
        username = self.driver.find_element_by_id("adduser-username")
        password = self.driver.find_element_by_id("adduser-password")
        confirm_password = self.driver.find_element_by_id("adduser-confirm_password")
        firstName = self.driver.find_element_by_id("adduser-first_name")
        lastName = self.driver.find_element_by_id("adduser-last_name")
        email.send_keys(self.buro_email)
        username.send_keys(self.username)
        password.send_keys("Temp123")
        confirm_password.send_keys("Temp123")
        firstName.send_keys(self.firstName)
        lastName.send_keys(self.lastName)

        self.driver.find_element_by_id("btn-add-user").click()

        print("Please choose a suitable role for the user.")
        input("Press enter when done and the script will proceed.")

        # Check if user exists after creation
        search = self.driver.find_element_by_xpath('//*[@id="users-list_filter"]/label/input')
        search.send_keys(self.username)
        users_list = self.driver.find_element_by_id("users-list").text
        if self.username in users_list:
            self.logbook.append((17, self.username, "Temp123"))
            print(self.username)
            print("Temp123")
            print("IMS user successfully created...")
        else:
            print("IMS user not created!")

    def supatools(self):
        self.driver.get(r"https://support.viptelecombilling.net.au/login.php")
        username = self.driver.find_element_by_name("user_id")
        password = self.driver.find_element_by_name("password")
        username.send_keys(self.logins.supatools[0])
        password.send_keys(self.logins.supatools[1])
        password.send_keys(Keys.ENTER)

        time.sleep(2)  # wait for page to load
        frame = self.driver.find_element_by_xpath('/html/frameset/frameset/frame[1]')
        self.driver.switch_to.frame(frame)
        self.driver.find_element_by_xpath('/html/body/table/tbody/tr[2]/td/a/img').click()
        self.driver.switch_to.default_content()
        frame = self.driver.find_element_by_xpath('/html/frameset/frameset/frame[2]')
        self.driver.switch_to.frame(frame)
        self.driver.find_element_by_xpath(
            '/html/body/table/tbody/tr/td/table/tbody/tr[1]/td/table[2]/tbody/tr/td[4]/a/img').click()
        self.driver.find_element_by_name("newci").click()
        drop_down = Select(
            self.driver.find_element_by_xpath('/html/body/form/table/tbody/tr/td/table/tbody/tr[1]/td[2]/select'))
        drop_down.select_by_visible_text("PlanetTel")
        time.sleep(2)  # wait for page to load
        fullName = self.driver.find_element_by_id("__name")
        title = self.driver.find_element_by_id("__job_title")
        email = self.driver.find_element_by_id("__email")
        self.driver.find_element_by_name("__can_switch_co").click()
        access_role = Select(self.driver.find_element_by_id("__access_role_id"))
        access_role.select_by_visible_text("Planet Tel")
        access_level = Select(self.driver.find_element_by_id("__access_level"))
        access_level.select_by_visible_text("Service Desk")
        username = self.driver.find_element_by_id("__user_id")
        password = self.driver.find_element_by_id("f_password")
        confirm_password = self.driver.find_element_by_id("confirm_password")
        status = Select(self.driver.find_element_by_name("__status"))
        status.select_by_visible_text("Reset")
        send_email = Select(self.driver.find_element_by_id("user_email_template"))
        send_email.select_by_visible_text("user_creation_password.txt")
        fullName.send_keys(self.fullname)
        title.send_keys(self.title)
        email.send_keys(self.buro_email)
        username.send_keys(self.username)
        password.send_keys("Temp123")
        confirm_password.send_keys("Temp123")

        self.driver.find_element_by_id("savebtn").click()

        # Check if user exists after creation
        time.sleep(2)
        self.driver.switch_to.default_content()
        frame = self.driver.find_element_by_xpath('/html/frameset/frame')
        self.driver.switch_to.frame(frame)
        search = self.driver.find_element_by_xpath('//*[@id="filter"]')
        search.send_keys(self.fullname)
        search.send_keys(Keys.ENTER)
        try:
            self.driver.switch_to.default_content()
            frame = self.driver.find_element_by_xpath('/html/frameset/frameset/frame[2]')
            self.driver.switch_to.frame(frame)
            users_list = self.driver.find_element_by_xpath('/html/body/table/tbody/tr/td').text
            if self.username in users_list:
                self.logbook.append((34, self.username, "Temp123"))
                print(self.username)
                print("Temp123")
                print("Supatools user successfully created...")
        except:
            print("Supatools user not created!")

    def porta(self):
        self.driver.get(r"https://billing.isphone.com.au/index.html")
        username = self.driver.find_element_by_id("pb_auth_user")
        password = self.driver.find_element_by_id("pb_auth_password")
        username.send_keys(self.logins.porta[0])
        password.send_keys(self.logins.porta[1])
        password.send_keys(Keys.ENTER)

        self.driver.get(r"https://billing.isphone.com.au/users.html")
        self.driver.find_element_by_class_name("add").click()
        time.sleep(3)  # wait for js to load
        firstName = self.driver.find_element_by_name("firstname")
        lastName = self.driver.find_element_by_id("lastname_field_id")
        email = self.driver.find_element_by_id("email_field_id")
        firstName.send_keys(self.firstName)
        lastName.send_keys(self.lastName)
        email.send_keys(self.buro_email)
        self.driver.find_element_by_link_text("Web Self-Care").click()
        username = self.driver.find_element_by_id("login_input_id")
        password = self.driver.find_element_by_id("password_input_id")
        username.clear()
        username.send_keys(self.username)
        newPass = self.passwordGenerator()
        password.send_keys(newPass)
        drop_down = Select(self.driver.find_element_by_id("ACLSelect"))
        drop_down.select_by_visible_text("ISPhone - Account Manager")
        self.driver.find_element_by_id("api-token-auth").click()

        self.driver.find_element_by_class_name("save_close").click()

        self.driver.get(r"https://billing.isphone.com.au/users.html")
        print("Please verify if user is created (Y/N).")
        answer = input("")
        if answer == 'Y' or answer == 'y':
            self.logbook.append((19, self.username, newPass))
            print(self.username)
            print(newPass)
            print("Porta user successfully created...")
        else:
            print("Porta user not created!")

    def teardown(self):
        self.driver.quit()


def main():
    print("Enter the full name of the new user:")
    name = input("")
    print("Enter the job title of the user:")
    title = input("")
    print("Is the following correct? (Y/N)")
    print(name + ",", title)
    answer = input("")
    if answer == 'N' or answer == 'n':
        print("Please run the script again.")
        return

    elif answer == "Y" or answer == 'y':
        print("Proceeding to create user accounts...")
        ac = AccessCreator(name, title)
        ac.createAll()
        ac.teardown()
        print("Complete!")


if __name__ == "__main__":
    main()
