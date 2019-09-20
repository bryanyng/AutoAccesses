import re
import secrets
import shutil
import string
import time
from datetime import date
from os import path

import openpyxl
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from Pass import Pass


class PasswordResetter:
    def __init__(self, fullname):
        self.fullname = fullname
        fullname = fullname.split()
        self.username = fullname[0] + '.' + fullname[1]
        self.firstName = fullname[0]
        self.lastName = fullname[1]
        self.buro_email = self.username + "@buroserv.com.au"
        self.planettel_email = self.username + "@planettel.com.au"
        self.newPassword = ""
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
        if re.search(r"[!@#$%^&*\-+?]", password) and re.search(r"[A-Z]", password) and re.search(r"[a-z]", password) and re.search(r"[0-9]", password):
            return True
        return False

    def ims(self):
        self.driver.get(r"http://ims.buroserv.com.au/login.php")

        username = self.driver.find_element_by_id("login-username")
        password = self.driver.find_element_by_id("login-password")
        username.send_keys(self.logins.ims()[0])
        password.send_keys(self.logins.ims()[1])
        self.driver.find_element_by_id("btn-login").click()

        time.sleep(2)
        self.driver.get(r"http://ims.buroserv.com.au/users.php")

        search = self.driver.find_element_by_xpath('//*[@id="users-list_filter"]/label/input')
        search.send_keys(self.username)
        self.driver.find_element_by_xpath('//*[@id="users-list"]/tbody/tr/td[6]/div/a[2]').click()
        self.driver.find_element_by_xpath('//*[@id="users-list"]/tbody/tr/td[6]/div/ul/li[1]/a').click()
        password = self.driver.find_element_by_id("adduser-password")
        confirm_password = self.driver.find_element_by_id("adduser-confirm_password")
        password.send_keys("Temp123")
        confirm_password.send_keys("Temp123")
        self.driver.find_element_by_id("btn-add-user").click()
        self.newPassword = "Temp123"

    def tele_iboss(self):
        self.driver.get(r"https://symbio-aspire.iboss.com.au/aspireV2/login.php")
        username = self.driver.find_element_by_name("Username")
        password = self.driver.find_element_by_name("Password")
        username.send_keys(self.logins.tele_iboss()[0])
        password.send_keys(self.logins.tele_iboss()[1])
        self.driver.find_element_by_name("Submit").click()
        self.driver.find_element_by_id("ToolsDrop").click()
        self.driver.find_element_by_link_text("Wholesaler User Logins").click()
        users_list = Select(self.driver.find_element_by_id("WholesellerUserID"))
        users_list.select_by_visible_text(self.fullname)
        password = self.driver.find_element_by_id("Password")
        special_char = secrets.choice("!#%@<>")
        newPass = self.passwordGenerator()
        newPass = re.sub(r"[$%^&*\-+?]", special_char, newPass)
        password.send_keys(newPass)

        # ensure user is active
        status = Select(self.driver.find_element_by_name("Status"))
        status.select_by_visible_text("Active")

        self.driver.find_element_by_xpath('//*[@id="wholesalerpassworddiv"]/td/table/tbody/tr[9]/td/div/input').click()
        self.newPassword = newPass

    def octane(self):
        self.driver.get(r"https://octane.telcoinabox.com/tiab/Login")
        self.driver.find_element_by_id("login_button").click()
        time.sleep(2)
        username = self.driver.find_element_by_id("j_username")
        password = self.driver.find_element_by_id("predigpass")
        username.send_keys(self.logins.octane()[0])
        password.send_keys(self.logins.octane()[1])
        password.send_keys(Keys.ENTER)
        self.driver.find_element_by_xpath('//span[text()="Login"]').click()
        self.driver.get(r"https://octane.telcoinabox.com/tiab/UserList")

        self.driver.find_element_by_css_selector('button[onclick="doEdit(\'' + self.username + '\'); return false;"]').click()
        password = self.driver.find_element_by_id("predigpass")
        confirm_password = self.driver.find_element_by_id("retype-password")
        newPass = self.passwordGenerator()
        password.send_keys(newPass)
        confirm_password.send_keys(newPass)
        self.driver.find_element_by_id("btn-reset").click()
        # MISSING MODAL CONFIRM CLICK
        self.newPassword = newPass

    def clarus_genex(self):
        self.driver.get(r"https://genex.billing.com.au/Module/Main/login.aspx")
        database = self.driver.find_element_by_id("ctl00_CPH_txtDatabase")
        username = self.driver.find_element_by_id("ctl00_CPH_txtUsername")
        password = self.driver.find_element_by_id("ctl00_CPH_txtPassword")
        database.clear()
        username.clear()
        database.send_keys("Clarus")
        username.send_keys(self.logins.clarus_genex()[0])
        password.send_keys(self.logins.clarus_genex()[1])
        self.driver.find_element_by_id("ctl00_CPH_btnLogin").click()
        self.driver.get(r"https://genex.billing.com.au/module/Roles/UserManager.aspx")
        find_username = self.driver.find_element_by_xpath('//label[contains(text(),\'' + self.fullname + '\')]').text
        match = re.match("\w+ \w+ \((.+)\)", find_username)
        genex_username = match.group(1)
        self.driver.find_element_by_css_selector('input[value=\"' + genex_username + '\"]').click()
        self.driver.find_element_by_id("ctl00_CPH_UsersAndRoles_UpdateButtons_EditCancelButton").click()
        password = self.driver.find_element_by_id("ctl00_CPH_UsersAndRoles_txtPassword")
        confirm_password = self.driver.find_element_by_id("ctl00_CPH_UsersAndRoles_txtConfirmPassword")
        password.clear()
        confirm_password.clear()
        newPass = self.passwordGenerator()
        password.send_keys(newPass)
        confirm_password.send_keys(newPass)
        self.driver.find_element_by_id("ctl00_CPH_UsersAndRoles_UpdateButtons_SaveButton").click()
        self.newPassword = newPass


    def buro_genex(self):
        self.driver.get(r"https://genex.billing.com.au/Module/Main/login.aspx")
        database = self.driver.find_element_by_id("ctl00_CPH_txtDatabase")
        username = self.driver.find_element_by_id("ctl00_CPH_txtUsername")
        password = self.driver.find_element_by_id("ctl00_CPH_txtPassword")
        database.clear()
        username.clear()
        database.send_keys("Buroserv")
        username.send_keys(self.logins.buro_genex()[0])
        password.send_keys(self.logins.buro_genex()[1])
        self.driver.find_element_by_id("ctl00_CPH_btnLogin").click()
        self.driver.get(r"https://genex.billing.com.au/module/Roles/UserManager.aspx")
        find_username = self.driver.find_element_by_xpath('//label[contains(text(),\'' + self.fullname + '\')]').text
        match = re.match("\w+ \w+ \((.+)\)", find_username)
        genex_username = match.group(1)
        self.driver.find_element_by_css_selector('input[value=\"' + genex_username + '\"]').click()
        self.driver.find_element_by_id("ctl00_CPH_UsersAndRoles_UpdateButtons_EditCancelButton").click()
        password = self.driver.find_element_by_id("ctl00_CPH_UsersAndRoles_txtPassword")
        confirm_password = self.driver.find_element_by_id("ctl00_CPH_UsersAndRoles_txtConfirmPassword")
        password.clear()
        confirm_password.clear()
        newPass = self.passwordGenerator()
        password.send_keys(newPass)
        confirm_password.send_keys(newPass)
        self.driver.find_element_by_id("ctl00_CPH_UsersAndRoles_UpdateButtons_SaveButton").click()
        self.newPassword = newPass

    def sonar(self):
        self.driver.get(r"https://mvp02.symbionetworks.com/sonar_admin/")
        username = self.driver.find_element_by_name("j_username")
        password = self.driver.find_element_by_name("j_password")
        username.send_keys(self.logins.sonar()[0])
        password.send_keys(self.logins.sonar()[1])
        password.send_keys(Keys.ENTER)

        self.driver.find_element_by_class_name("rootVoices").click()
        self.driver.find_element_by_link_text("Users").click()
        user_code = self.driver.find_element_by_xpath('//td[@title=\"' + self.buro_email + '\"]/preceding-sibling::td').text
        self.driver.find_element_by_css_selector('button[onclick=\"Users(\'' + user_code + '\', \'com.symbio.sona.actor.User\', \'User\', \'User\');').click()
        actorManager = self.driver.find_element_by_id("User" + user_code + "actorManager").is_selected()
        time.sleep(5)
        if actorManager:
            self.driver.find_element_by_xpath('//*[@id="User' + user_code + 'actorManager"]').click()
            self.driver.find_element_by_xpath('//*[@id="User' + user_code + 'Form"]/button').click()
            time.sleep(5)
        self.driver.find_element_by_xpath('//*[@id="Modify_User"]/table[2]/tbody/tr[2]/td[3]/button').click()
        password = self.driver.find_element_by_name("newPassword")
        confirm_password = self.driver.find_element_by_name("confirmNewPassword")
        newPass = self.passwordGenerator()
        password.send_keys(newPass)
        confirm_password.send_keys(newPass)
        self.driver.find_element_by_xpath('//*[@id="User' + user_code + 'changePasswordForm"]/button').click()
        time.sleep(5)
        if actorManager:
            self.driver.find_element_by_xpath('//*[@id="User' + user_code + 'actorManager"]').click()
            self.driver.find_element_by_xpath('//*[@id="User' + user_code + 'Form"]/button').click()
            time.sleep(5)
        self.newPassword = newPass

    def supatools(self):
        self.driver.get(r"https://support.viptelecombilling.net.au/login.php")
        username = self.driver.find_element_by_name("user_id")
        password = self.driver.find_element_by_name("password")
        username.send_keys(self.logins.supatools()[0])
        password.send_keys(self.logins.supatools()[1])
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
        username = self.driver.find_element_by_xpath('//*[@id="value_1"]')
        username.send_keys(self.fullname)
        self.driver.find_element_by_xpath('//*[@id="sf"]/table/tbody/tr/td[1]/table[5]/tbody/tr/td/input').click()
        drop_down = Select(self.driver.find_element_by_id('column1'))
        drop_down.select_by_visible_text("User Id")
        self.driver.find_element_by_link_text("Edit").click()

    def porta(self):
        self.driver.get(r"https://billing.isphone.com.au/index.html")
        username = self.driver.find_element_by_id("pb_auth_user")
        password = self.driver.find_element_by_id("pb_auth_password")
        username.send_keys(self.logins.porta()[0])
        password.send_keys(self.logins.porta()[1])
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
            print("Porta user successfully created...")
        else:
            print("Porta user not created!")

    def viaip_utilibill(self):
        self.driver.get(r"https://viaip.utilibill.com.au/viaip/Login")
        self.driver.find_element_by_tag_name("a").click()
        username = self.driver.find_element_by_name("j_username")
        password = self.driver.find_element_by_id("predigpass")
        username.send_keys(self.logins.viaip_utilibill()[0])
        password.send_keyss(self.logins.viaip_utilibill()[1])
        self.driver.find_element_by_name("submit").click()
        self.driver.find_element_by_id("submitrequest").click()

        self.driver.get(r"https://viaip.utilibill.com.au/viaip/UserList")
        self.driver.find_element_by_css_selector('a[href="javascript:doEdit(\'' + self.username + '\');"]').click()
        newPass = self.passwordGenerator()
        password = self.driver.find_element_by_id("predigpass")
        password.send_keys(newPass)
        self.driver.find_element_by_css_selector('a[href="javascript:resetPassword()"]').click()
        self.driver.switch_to.alert.accept()
        self.newPassword = newPass

    def cloud_ultilibill(self):
        self.driver.get(r"https://eziconnect.utilibill.com.au/eziconnect/Login")
        self.driver.find_element_by_tag_name("a").click()
        username = self.driver.find_element_by_name("j_username")
        password = self.driver.find_element_by_id("predigpass")
        username.send_keys(self.logins.cloud_utilibill()[0])
        password.send_keys(self.logins.cloud_utilibill()[1])
        self.driver.find_element_by_name("submit").click()
        self.driver.find_element_by_id("submitrequest").click()

        self.driver.get(r"https://eziconnect.utilibill.com.au/eziconnect/UserList")
        self.driver.find_element_by_css_selector('a[href="javascript:doEdit(\'' + self.username + '\');"]').click()
        newPass = self.passwordGenerator()
        password = self.driver.find_element_by_id("predigpass")
        password.send_keys(newPass)
        self.driver.find_element_by_css_selector('a[href="javascript:resetPassword()"]').click()
        self.driver.switch_to.alert.accept()
        self.newPassword = newPass

    def pickPortal(self):
        print("=====================================================")
        print("Which portal access does the user require resetting?")
        print("=====================================================")
        print("[1] Supatools")
        print("[2] IMS")
        print("[3] Utilibill")
        print("[4] Octane")
        print("[5] Genex")
        print("[6] Porta")
        print("[7] Sonar")
        print("[8] iBoss")
        pick = input("")

        if pick == "1":
            self.supatools()
        elif pick == "2":
            self.ims()
        elif pick == "3":
            print("===============================")
            print("Utilibill: ViaIP or Cloudnyne?")
            print("===============================")
            print("[1] ViaIP")
            print("[2] Cloudnyne")
            pick = input("")
            if pick == "1":
                self.viaip_utilibill()
            elif pick == "2":
                self.cloud_ultilibill()
        elif pick == "4":
            self.octane()
        elif pick == "5":
            print("===============================")
            print("Genex: Clarus or Buroserv?")
            print("===============================")
            print("[1] Clarus")
            print("[2] Buroserv")
            pick = input("")
            if pick == "1":
                self.clarus_genex()
            elif pick == "2":
                self.buro_genex()
        elif pick == "6":
            self.porta()
        elif pick == "7":
            self.sonar()
        elif pick == "8":
            self.iboss()

    def teardown(self):
        self.driver.quit()


def main():
    print("Enter the full name of the user:")
    name = input("")
    print("Is the following correct? (Y/N)")
    print(name)
    answer = input("")
    if answer == 'N' or answer == 'n':
        print("Please run the script again.")
        return

    elif answer == "Y" or answer == 'y':
        pr = PasswordResetter(name)
        pr.pickPortal()
        pr.teardown()
        print("New Password: " + pr.newPassword)
        print("Complete!")


if __name__ == "__main__":
    main()
