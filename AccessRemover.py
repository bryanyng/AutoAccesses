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
from selenium.webdriver.common.action_chains import ActionChains


class AccessRemover:
    def __init__(self, fullname):
        self.fullname = fullname
        fullname = fullname.split()
        self.username = fullname[0] + '.' + fullname[1]
        self.firstName = fullname[0]
        self.lastName = fullname[1]
        self.buro_email = self.username + "@buroserv.com.au"
        self.planettel_email = self.username + "@planettel.com.au"
        self.logins = Pass()

        # Set up webdriver
        self.driver = webdriver.Chrome(r"Drivers/chromedriver.exe")
        self.driver.maximize_window()
        self.driver.implicitly_wait(10)

    def ims(self):
        self.driver.get(r"http://ims.buroserv.com.au/login.php")
        username = self.driver.find_element_by_id("login-username")
        password = self.driver.find_element_by_id("login-password")
        username.send_keys(self.logins.ims[0])
        password.send_keys(self.logins.ims[1])
        self.driver.find_element_by_id("btn-login").click()
        time.sleep(2)
        self.driver.get(r"http://ims.buroserv.com.au/users.php")
        search = self.driver.find_element_by_xpath('//*[@id="users-list_filter"]/label/input')
        search.send_keys(self.username)
        self.driver.find_element_by_xpath('//*[@id="users-list"]/tbody/tr/td[6]/div/a[2]').click()
        self.driver.find_element_by_xpath('//*[@id="users-list"]/tbody/tr/td[6]/div/ul/li[3]/a').click()

    def tele_iboss(self):
        self.driver.get(r"https://symbio-aspire.iboss.com.au/aspireV2/login.php")
        username = self.driver.find_element_by_name("Username")
        password = self.driver.find_element_by_name("Password")
        username.send_keys(self.logins.tele_iboss[0])
        password.send_keys(self.logins.tele_iboss[1])
        self.driver.find_element_by_name("Submit").click()
        self.driver.find_element_by_id("ToolsDrop").click()
        self.driver.find_element_by_link_text("Wholesaler User Logins").click()
        drop_down = Select(self.driver.find_element_by_id("WholesellerUserID"))
        drop_down.select_by_visible_text(self.fullname)
        status = Select(self.driver.find_element_by_xpath('//*[@id="wholesalerpassworddiv"]/td/table/tbody/tr[8]/td[2]/select'))
        status.select_by_visible_text("Locked Out")
        self.driver.find_element_by_xpath('//*[@id="wholesalerpassworddiv"]/td/table/tbody/tr[9]/td/div/input').click()

    def viaip_optus(self):
        self.driver.get(r"https://www2.optus.com.au/wholesalenet/")
        username = self.driver.find_element_by_id("USER")
        password = self.driver.find_element_by_id("PASSWORD")
        username.send_keys(self.logins.viaip_optus[0])
        password.send_keys(self.logins.viaip_optus[1])
        self.driver.find_element_by_name("LOGIN").click()
        self.driver.find_element_by_xpath('//*[@id="pri_1870"]/a').click()
        action = ActionChains(self.driver)
        first_menu = self.driver.find_element_by_xpath('//*[@id="sec_1870"]/li/a/span')
        second_menu = self.driver.find_element_by_xpath('//*[@id="sec_1870"]/li/div/div/div/ul/li[2]/a')
        action.move_to_element(first_menu)
        action.perform()
        second_menu.click()
        self.driver.find_element_by_xpath('//*[@id="content"]/div/div/div[1]/div[2]/a/h2').click()
        username = self.driver.find_element_by_xpath('//*[@id="Username"]')
        username.send_keys(self.username)
        self.driver.find_element_by_xpath('//*[@id="content"]/div/div/div[1]/div[2]/form/ul/li[6]/input').click()
        self.driver.find_element_by_link_text(self.lastName.upper() + " " + self.firstName).click()
        status = Select(self.driver.find_element_by_name("user_status_flag"))
        status.select_by_visible_text("Inactive")
        self.driver.find_element_by_xpath('//*[@id="editUserProfile"]/ul[2]/li[4]/input[1]').click()

    # NOT UPDATED DO NOT USE...
    def buro_optus(self):
        self.driver.get("https://www2.optus.com.au/wholesalenet/")
        username = self.driver.find_element_by_id("USER")
        password = self.driver.find_element_by_id("PASSWORD")
        username.send_keys(self.logins.buro_optus[0])
        password.send_keys(self.logins.buro_optus[1])
        self.driver.find_element_by_name("LOGIN").click()
        self.driver.find_element_by_link_text("Administration").click()
        self.driver.find_element_by_xpath('//span[text()="User Management"]').click()

        firstName = self.driver.find_element_by_id("FirstName")
        lastName = self.driver.find_element_by_id("LastName")
        username = self.driver.find_element_by_id("Username")
        email = self.driver.find_element_by_name("Email")
        firstName.send_keys(self.firstName)
        lastName.send_keys(self.lastName)
        username.send_keys(".BURO")
        email.send_keys(self.buro_email)
        # js drop-down boxes
        self.driver.find_element_by_xpath('//*[@id="addNewUser"]/ul[1]/li[1]/span/button').click()
        self.driver.find_element_by_xpath('/html/body/div[4]/ul/li[2]/label').click()
        self.driver.find_element_by_xpath(
            '//*[@id="addNewUser"]/ul[1]/li[6]/span/table[1]/tbody/tr[3]/td[2]/button').click()
        self.driver.find_element_by_xpath('/html/body/div[6]/ul/li[2]/label').click()
        self.driver.find_element_by_xpath(
            '//*[@id="addNewUser"]/ul[1]/li[6]/span/table[1]/tbody/tr[4]/td[2]/button').click()
        self.driver.find_element_by_xpath('/html/body/div[7]/ul/li[2]/label').click()
        self.driver.find_element_by_xpath(
            '//*[@id="addNewUser"]/ul[1]/li[6]/span/table[1]/tbody/tr[5]/td[2]/button').click()
        self.driver.find_element_by_xpath('/html/body/div[8]/ul/li[2]/label').click()
        self.driver.find_element_by_xpath(
            '//*[@id="addNewUser"]/ul[1]/li[6]/span/table[1]/tbody/tr[6]/td[2]/button').click()
        self.driver.find_element_by_xpath('/html/body/div[9]/ul/li[2]/label').click()
        self.driver.find_element_by_xpath(
            '//*[@id="addNewUser"]/ul[1]/li[6]/span/table[1]/tbody/tr[7]/td[2]/button').click()
        self.driver.find_element_by_xpath('/html/body/div[10]/ul/li[2]/label').click()

        # Submit
        self.driver.find_element_by_xpath('//*[@id="addNewUser"]/ul[2]/li[4]/input[1]').click()

    def octane(self):
        self.driver.get(r"https://octane.telcoinabox.com/tiab/Login")
        self.driver.find_element_by_id("login_button").click()
        time.sleep(2)
        username = self.driver.find_element_by_id("j_username")
        password = self.driver.find_element_by_id("predigpass")
        username.send_keys(self.logins.octane[0])
        password.send_keys(self.logins.octane[1])
        password.send_keys(Keys.ENTER)
        self.driver.find_element_by_xpath('//span[text()="Login"]').click()

        self.driver.get(r"https://octane.telcoinabox.com/tiab/UserList")
        self.driver.find_element_by_css_selector(
            'button[onclick="doEdit(\'' + self.username + '\'); return false;"]').click()
        self.driver.find_element_by_xpath('//*[@id="btn-bar"]').click()
        # FIX MODAL
        self.driver.find_element_by_xpath('//*[@id="modal-btn-ok"]').click()

    def clarus_genex(self):
        self.driver.get(r"https://genex.billing.com.au/Module/Main/login.aspx")
        database = self.driver.find_element_by_id("ctl00_CPH_txtDatabase")
        username = self.driver.find_element_by_id("ctl00_CPH_txtUsername")
        password = self.driver.find_element_by_id("ctl00_CPH_txtPassword")
        database.clear()
        username.clear()
        database.send_keys("Clarus")
        username.send_keys(self.logins.clarus_genex[0])
        password.send_keys(self.logins.clarus_genex[1])
        self.driver.find_element_by_id("ctl00_CPH_btnLogin").click()
        self.driver.get(r"https://genex.billing.com.au/module/Roles/UserManager.aspx")
        find_username = self.driver.find_element_by_xpath('//label[contains(text(),\'' + self.fullname + '\')]').text
        match = re.match("\w+ \w+ \((.+)\)", find_username)
        genex_username = match.group(1)
        self.driver.find_element_by_css_selector('input[value=\"' + genex_username + '\"]').click()
        self.driver.find_element_by_xpath('//*[@id="ctl00_CPH_UsersAndRoles_DisableButton"]').click()
        self.driver.switch_to.alert.accept()

    def buro_genex(self):
        self.driver.get(r"https://genex.billing.com.au/Module/Main/login.aspx")
        database = self.driver.find_element_by_id("ctl00_CPH_txtDatabase")
        username = self.driver.find_element_by_id("ctl00_CPH_txtUsername")
        password = self.driver.find_element_by_id("ctl00_CPH_txtPassword")
        database.clear()
        username.clear()
        database.send_keys("Buroserv")
        username.send_keys(self.logins.buro_genex[0])
        password.send_keys(self.logins.buro_genex[1])
        self.driver.find_element_by_id("ctl00_CPH_btnLogin").click()
        self.driver.get(r"https://genex.billing.com.au/module/Roles/UserManager.aspx")
        find_username = self.driver.find_element_by_xpath('//label[contains(text(),\'' + self.fullname + '\')]').text
        match = re.match("\w+ \w+ \((.+)\)", find_username)
        genex_username = match.group(1)
        self.driver.find_element_by_css_selector('input[value=\"' + genex_username + '\"]').click()
        self.driver.find_element_by_xpath('//*[@id="ctl00_CPH_UsersAndRoles_DisableButton"]').click()
        self.driver.switch_to.alert.accept()

    def v4_genex(self):
        self.driver.get(r"https://genex.billing.com.au/Module/Main/login.aspx")
        database = self.driver.find_element_by_id("ctl00_CPH_txtDatabase")
        username = self.driver.find_element_by_id("ctl00_CPH_txtUsername")
        password = self.driver.find_element_by_id("ctl00_CPH_txtPassword")
        database.clear()
        username.clear()
        database.send_keys("v4")
        username.send_keys(self.logins.v4_genex[0])
        password.send_keys(self.logins.v4_genex[1])
        self.driver.find_element_by_id("ctl00_CPH_btnLogin").click()
        self.driver.get(r"https://genex.billing.com.au/module/Roles/UserManager.aspx")
        find_username = self.driver.find_element_by_xpath('//label[contains(text(),\'' + self.fullname + '\')]').text
        match = re.match("\w+ \w+ \((.+)\)", find_username)
        genex_username = match.group(1)
        self.driver.find_element_by_css_selector('input[value=\"' + genex_username + '\"]').click()
        self.driver.find_element_by_xpath('//*[@id="ctl00_CPH_UsersAndRoles_DisableButton"]').click()
        self.driver.switch_to.alert.accept()

    # NOT COMPLETE
    def sonar(self):
        self.driver.get(r"https://mvp02.symbionetworks.com/sonar_admin/")
        username = self.driver.find_element_by_name("j_username")
        password = self.driver.find_element_by_name("j_password")
        username.send_keys(self.logins.sonar[0])
        password.send_keys(self.logins.sonar[1])
        password.send_keys(Keys.ENTER)

        self.driver.find_element_by_class_name("rootVoices").click()
        self.driver.find_element_by_link_text("Users").click()

    # NOT COMPLETE
    def supatools(self):
        self.driver.get(r"https://support.viptelecombilling.net.au/login.php")
        username = self.driver.find_element_by_name("user_id")
        password = self.driver.find_element_by_name("password")
        username.send_keys(self.logins.supatools[0])
        password.send_keys(self.logins.supatools[1])
        password.send_keys(Keys.ENTER)

        time.sleep(2)  # wait for page to load

    # NOT COMPLETE
    def porta(self):
        self.driver.get(r"https://billing.isphone.com.au/index.html")
        username = self.driver.find_element_by_id("pb_auth_user")
        password = self.driver.find_element_by_id("pb_auth_password")
        username.send_keys(self.logins.porta[0])
        password.send_keys(self.logins.porta[1])
        password.send_keys(Keys.ENTER)

        self.driver.get(r"https://billing.isphone.com.au/users.html")

    def buro_frontier(self):
        self.driver.get(r"https://frontier.aapt.com.au/s/login")
        username = self.driver.find_element_by_name("j_username")
        password = self.driver.find_element_by_name("j_password")
        username.send_keys(self.logins.buro_frontier[0])
        password.send_keys(self.logins.buro_frontier[1])
        self.driver.find_element_by_id("submit").click()

        self.driver.get(r"https://frontier.aapt.com.au/s/manageusers")
        drop_down = Select(self.driver.find_element_by_id("userList"))
        drop_down.select_by_visible_text(self.lastName + ", " + self.firstName)
        self.driver.find_element_by_xpath('//*[@id="deleteUser"]/a').click()
        self.driver.find_element_by_xpath('//*[@id="popup_ok"]').click()
        # Sign out
        self.driver.find_element_by_id("nameMenu").click()
        self.driver.find_element_by_link_text("Sign Out").click()

    def cloud_frontier(self):
        self.driver.get(r"https://frontier.aapt.com.au/s/login")
        username = self.driver.find_element_by_name("j_username")
        password = self.driver.find_element_by_name("j_password")
        username.clear()
        username.send_keys(self.logins.cloud_frontier[0])
        password.send_keys(self.logins.cloud_frontier[1])
        self.driver.find_element_by_id("submit").click()

        self.driver.get(r"https://frontier.aapt.com.au/s/manageusers")
        drop_down = Select(self.driver.find_element_by_id("userList"))
        drop_down.select_by_visible_text(self.lastName + ", " + self.firstName)
        self.driver.find_element_by_xpath('//*[@id="deleteUser"]/a').click()
        self.driver.find_element_by_xpath('//*[@id="popup_ok"]').click()
        # Sign out
        self.driver.find_element_by_id("nameMenu").click()
        self.driver.find_element_by_link_text("Sign Out").click()

    def viaip_utilibill(self):
        self.driver.get(r"https://viaip.utilibill.com.au/viaip/Login")
        self.driver.find_element_by_tag_name("a").click()
        username = self.driver.find_element_by_name("j_username")
        password = self.driver.find_element_by_id("predigpass")
        username.send_keys(self.logins.viaip_utilibill[0])
        password.send_keys(self.logins.viaip_utilibill[1])
        self.driver.find_element_by_name("submit").click()
        self.driver.find_element_by_id("submitrequest").click()

        self.driver.get(r"https://viaip.utilibill.com.au/viaip/UserList")
        self.driver.find_element_by_css_selector('a[href="javascript:doEdit(\'' + self.username + '\');"]').click()
        self.driver.find_element_by_xpath(
            '//*[@id="utbFrmDiv"]/form/table/tbody/tr[5]/td/center/table/tbody/tr/td[2]/a').click()
        self.driver.switch_to.alert.accept()

    def cloud_ultilibill(self):
        self.driver.get(r"https://eziconnect.utilibill.com.au/eziconnect/Login")
        self.driver.find_element_by_tag_name("a").click()
        username = self.driver.find_element_by_name("j_username")
        password = self.driver.find_element_by_id("predigpass")
        username.send_keys(self.logins.cloud_utilibill[0])
        password.send_keys(self.logins.cloud_utilibill[1])
        self.driver.find_element_by_name("submit").click()
        self.driver.find_element_by_id("submitrequest").click()

        self.driver.get(r"https://eziconnect.utilibill.com.au/eziconnect/UserList")
        self.driver.find_element_by_css_selector('a[href="javascript:doEdit(\'' + self.username + '\');"]').click()
        self.driver.find_element_by_xpath('//*[@id="utbFrmDiv"]/form/table/tbody/tr[5]/td/center/table/tbody/tr/td[2]/a').click()
        self.driver.switch_to.alert.accept()

    def selcomm(self):
        self.driver.get("https://support.selcomm.com:8443/secure/Dashboard.jspa")
        username = self.driver.find_element_by_id("login-form-username")
        password = self.driver.find_element_by_id("login-form-password")
        username.send_keys(self.logins.selcomm[0])
        password.send_keys(self.logins.selcomm[1])
        self.driver.find_element_by_id("login").click()
        self.driver.find_element_by_id("create_link").click()

        drop_down = Select(self.driver.find_element_by_id("customfield_11502"))
        drop_down.select_by_visible_text("Bryan Yeung")
        summary = self.driver.find_element_by_id("summary")
        description = self.driver.find_element_by_id("description")
        if len(self.lastName) > 5:
            bu_id = "bu" + self.lastName[0:5].lower() + self.firstName[0].lower()
            bw_id = "bw" + self.lastName[0:5].lower() + self.firstName[0].lower()
        else:
            bu_id = "bu" + self.lastName.lower() + self.firstName[0].lower()
            bw_id = "bw" + self.lastName.lower() + self.firstName[0].lower()
        summary.send_keys("Remove User Access")
        description.send_keys(
            "Hi Team,\n\nCould you please remove the access to BU & BW of:\n" + bu_id + " / " + bw_id + "\n" + self.fullname + " (" + self.buro_email + ")\n"
            + "\n\nThanks!")
        self.driver.find_element_by_id("create-issue-submit").click()

    def removeAll(self):
        self.ims()
        self.viaip_optus()
        # self.buro_optus()
        self.supatools()
        self.buro_frontier()
        self.cloud_frontier()
        self.viaip_utilibill()
        self.cloud_ultilibill()
        # self.clarus_genex()
        self.buro_genex()
        self.v4_genex()
        self.octane()
        self.sonar()
        self.iboss()
        self.porta()
        self.selcomm()

        print("Writing user credentials to a spreadsheet...")

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
        print("Proceeding to remove user accounts...")
        ar = AccessRemover(name)
        ar.v4_genex()
        # ar.removeAll()
        ar.teardown()
        print("Complete!")


if __name__ == "__main__":
    main()
