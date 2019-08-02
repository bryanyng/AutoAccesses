import re
import secrets
import shutil
import string
import time
import openpyxl

from datetime import date
from os import path
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.action_chains import ActionChains


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
        username.send_keys("bryan.yeung")
        password.send_keys("+;q`bn3(U(m6Xe@3")
        self.driver.find_element_by_id("btn-login").click()

        time.sleep(2)
        self.driver.get(r"http://ims.buroserv.com.au/users.php")

        search = self.driver.find_element_by_xpath('//*[@id="users-list_filter"]/label/input')
        search.send_keys(self.username)

    def iboss(self):
        self.driver.get(r"https://symbio-aspire.iboss.com.au/aspireV2/login.php")
        username = self.driver.find_element_by_name("Username")
        password = self.driver.find_element_by_name("Password")
        username.send_keys("telemates.com.au")
        password.send_keys("ChangeMeMore!2m")
        self.driver.find_element_by_name("Submit").click()
        self.driver.find_element_by_id("ToolsDrop").click()
        self.driver.find_element_by_link_text("Wholesaler User Logins").click()

        username = self.driver.find_element_by_id("Username")
        password = self.driver.find_element_by_id("Password")
        name = self.driver.find_element_by_id("Name")
        email = self.driver.find_element_by_id("Email")
        iboss_username = self.firstName + self.lastName
        username.send_keys(iboss_username)
        newPass = self.passwordGenerator()
        newPass = re.sub(r"[$%^&*\-+?]", "@", newPass)
        password.send_keys(newPass)
        email.send_keys(self.buro_email)
        name.send_keys(self.fullname)

        self.driver.find_element_by_name("Submit").click()

        # Check if user exists after creation
        users_list = self.driver.find_element_by_id("WholesellerUserID").text
        if self.fullname in users_list:
            self.logbook.append((31, iboss_username, newPass))
            print("IBoss user successfully created...")
        else:
            print("IBoss user not created!")

    def octane(self):
        self.driver.get(r"https://octane.telcoinabox.com/tiab/Login")
        self.driver.find_element_by_id("login_button").click()
        time.sleep(2)
        username = self.driver.find_element_by_id("j_username")
        password = self.driver.find_element_by_id("predigpass")
        username.send_keys("bryan")
        password.send_keys("Douchebag!2")
        password.send_keys(Keys.ENTER)
        self.driver.find_element_by_xpath('//span[text()="Login"]').click()
        self.driver.get(r"https://octane.telcoinabox.com/tiab/NewUser.jsp")
        username = self.driver.find_element_by_id("uid")
        password = self.driver.find_element_by_id("predigpass")
        username.send_keys(self.username)
        newPass = self.passwordGenerator()
        password.send_keys(newPass)

        self.driver.find_element_by_id("submit-button").click()

        # Check if user exists after creation
        self.driver.get(r"https://octane.telcoinabox.com/tiab/UserList")
        users_list = self.driver.find_element_by_xpath('//*[@id="user-list"]/div[2]/table/tbody').text
        if self.username in users_list:
            self.logbook.append((7, self.username, newPass))
            print("Octane user successfully created...")
        else:
            print("Octane user not created!")
        # TODO: add partitions??

    def clarus_genex(self):
        self.driver.get(r"https://genex.billing.com.au/Module/Main/login.aspx")
        database = self.driver.find_element_by_id("ctl00_CPH_txtDatabase")
        username = self.driver.find_element_by_id("ctl00_CPH_txtUsername")
        password = self.driver.find_element_by_id("ctl00_CPH_txtPassword")
        database.clear()
        username.clear()
        database.send_keys("Clarus")
        username.send_keys("Bryan.Yeung")
        password.send_keys("1qazXSW@")
        self.driver.find_element_by_id("ctl00_CPH_btnLogin").click()
        self.driver.get(r"https://genex.billing.com.au/module/Roles/UserManager.aspx")
        self.driver.find_element_by_id("ctl00_CPH_UsersAndRoles_AddButton").click()
        username = self.driver.find_element_by_id("ctl00_CPH_UsersAndRoles_txtUserName")
        firstName = self.driver.find_element_by_id("ctl00_CPH_UsersAndRoles_txtFirstName")
        lastName = self.driver.find_element_by_id("ctl00_CPH_UsersAndRoles_txtLastName")
        password = self.driver.find_element_by_id("ctl00_CPH_UsersAndRoles_txtPassword")
        confirm_password = self.driver.find_element_by_id("ctl00_CPH_UsersAndRoles_txtConfirmPassword")
        email = self.driver.find_element_by_id("ctl00_CPH_UsersAndRoles_txtEmailAddress")
        if len(self.username) > 12:
            genex_username = self.firstName + self.lastName[0]
            if len(genex_username) > 12:
                genex_username = input("Please enter a 12 character username for Genex.")
        else:
            genex_username = self.username
        username.send_keys(genex_username)
        firstName.send_keys(self.firstName)
        lastName.send_keys(self.lastName)
        newPass = self.passwordGenerator()
        password.send_keys(newPass)
        confirm_password.send_keys(newPass)
        email.send_keys(self.buro_email)
        self.driver.find_element_by_id("ctl00_CPH_UsersAndRoles_cblGroups_0").click()

        print("Please choose appropriate roles for this user.")
        input("Press enter when done and the script will proceed.")

        self.driver.find_element_by_id("ctl00_CPH_UsersAndRoles_UpdateButtons_SaveButton").click()
        self.driver.switch_to.alert.accept()

        # Check if user exists after creation
        self.driver.get(r"https://genex.billing.com.au/module/Roles/UserManager.aspx")
        users_list = self.driver.find_element_by_id("ctl00_CPH_UsersAndRoles_rblUsers").text
        if self.fullname in users_list:
            self.logbook.append((8, genex_username, newPass))
            print("Clarus Genex user successfully created...")
        else:
            print("Clarus Genex user not created!")

    def v4_genex(self):
        self.driver.get(r"https://genex.billing.com.au/Module/Main/login.aspx")
        database = self.driver.find_element_by_id("ctl00_CPH_txtDatabase")
        username = self.driver.find_element_by_id("ctl00_CPH_txtUsername")
        password = self.driver.find_element_by_id("ctl00_CPH_txtPassword")
        database.clear()
        username.clear()
        database.send_keys("Buroserv")
        username.send_keys("Bryan.Y")
        password.send_keys("bynej2")
        self.driver.find_element_by_id("ctl00_CPH_btnLogin").click()
        self.driver.get(r"https://genex.billing.com.au/module/Roles/UserManager.aspx")
        self.driver.find_element_by_id("ctl00_CPH_UsersAndRoles_AddButton").click()
        username = self.driver.find_element_by_id("ctl00_CPH_UsersAndRoles_txtUserName")
        firstName = self.driver.find_element_by_id("ctl00_CPH_UsersAndRoles_txtFirstName")
        lastName = self.driver.find_element_by_id("ctl00_CPH_UsersAndRoles_txtLastName")
        password = self.driver.find_element_by_id("ctl00_CPH_UsersAndRoles_txtPassword")
        confirm_password = self.driver.find_element_by_id("ctl00_CPH_UsersAndRoles_txtConfirmPassword")
        email = self.driver.find_element_by_id("ctl00_CPH_UsersAndRoles_txtEmailAddress")
        if len(self.username) > 12:
            genex_username = self.firstName + self.lastName[0]
            if len(genex_username) > 12:
                genex_username = input("Please enter a 12 character username for Genex.")
        else:
            genex_username = self.username
        username.send_keys(genex_username)
        firstName.send_keys(self.firstName)
        lastName.send_keys(self.lastName)
        newPass = self.passwordGenerator()
        password.send_keys(newPass)
        confirm_password.send_keys(newPass)
        email.send_keys(self.buro_email)
        self.driver.find_element_by_id("ctl00_CPH_UsersAndRoles_cblGroups_0").click()

        print("Please choose appropriate roles for this user.")
        input("Press enter when done and the script will proceed.")

        self.driver.find_element_by_id("ctl00_CPH_UsersAndRoles_UpdateButtons_SaveButton").click()
        self.driver.switch_to.alert.accept()

        # Check if user exists after creation
        self.driver.get(r"https://genex.billing.com.au/module/Roles/UserManager.aspx")
        users_list = self.driver.find_element_by_id("ctl00_CPH_UsersAndRoles_rblUsers").text
        if self.fullname in users_list:
            self.logbook.append((9, genex_username, newPass))
            print("V4 Genex user successfully created...")
        else:
            print("V4 Genex user not created!")

    def sonar(self):
        self.driver.get(r"https://mvp02.symbionetworks.com/sonar_admin/")
        username = self.driver.find_element_by_name("j_username")
        password = self.driver.find_element_by_name("j_password")
        username.send_keys("Bryan_Yeung")
        password.send_keys("Bry@nY123!")
        password.send_keys(Keys.ENTER)

        self.driver.find_element_by_class_name("rootVoices").click()
        self.driver.find_element_by_link_text("Users").click()
        self.driver.find_element_by_xpath('//button[text()="Create User"]').click()
        time.sleep(2)  # wait for js to load
        self.driver.find_element_by_id("EndUserActor2671createUserenabled").click()
        username = self.driver.find_element_by_name("username")
        password = self.driver.find_element_by_name("password")
        surname = self.driver.find_element_by_name("surname")
        firstName = self.driver.find_element_by_name("firstName")
        email = self.driver.find_element_by_name("emailAddress")
        dob = self.driver.find_element_by_name("dateOfBirth")
        phone = self.driver.find_element_by_name("phoneNumberBH")
        sonar_username = self.firstName + self.lastName[0]
        username.send_keys(sonar_username)
        newPass = self.passwordGenerator()
        password.send_keys(newPass)
        surname.send_keys(self.lastName)
        firstName.send_keys(self.firstName)
        email.send_keys(self.buro_email)
        dob.send_keys("2019-01-01 00:00")
        phone.send_keys("1")

        self.driver.find_element_by_xpath('//button[text()="Submit"]').click()

        # Check if user exists after creation
        users_list = self.driver.find_element_by_id("EndUserActor2671UsersListingTable").text
        if sonar_username in users_list:
            self.logbook.append((29, sonar_username, newPass))
            print("Sonar user successfully created...")
        else:
            self.logbook.append((29, sonar_username, newPass)) #FIX THIS IF HAVE TIME!!!
            print("Sonar user not created!")

    def supatools(self):
        self.driver.get(r"https://support.viptelecombilling.net.au/login.php")
        username = self.driver.find_element_by_name("user_id")
        password = self.driver.find_element_by_name("password")
        username.send_keys("Bryan.Yeung")
        password.send_keys("As77917791")
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
                print("Supatools user successfully created...")
        except:
            print("Supatools user not created!")

    def porta(self):
        self.driver.get(r"https://billing.isphone.com.au/index.html")
        username = self.driver.find_element_by_id("pb_auth_user")
        password = self.driver.find_element_by_id("pb_auth_password")
        username.send_keys("Bryan.Yeung")
        password.send_keys("7NdZrXu2L6gkXcK")
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
        username.send_keys("bryan.yeung")
        password.send_keys("hR4XyL}(")
        self.driver.find_element_by_name("submit").click()
        self.driver.find_element_by_id("submitrequest").click()

        self.driver.get(r"https://viaip.utilibill.com.au/viaip/NewUser.jsp")
        username = self.driver.find_element_by_name("j_username")
        password = self.driver.find_element_by_id("predigpass")
        drop_down = Select(self.driver.find_element_by_id("group"))
        drop_down.select_by_visible_text("ViaIP Administrator")
        username.send_keys(self.username)
        newPass = self.passwordGenerator()
        password.send_keys(newPass)

        self.driver.find_element_by_class_name("buttonLrg").click()

        self.driver.get(r"https://viaip.utilibill.com.au/viaip/UserList")
        users_list = self.driver.find_element_by_id("utbListDiv").text
        match = re.search(self.username, users_list)
        if match:
            self.logbook.append((5, self.username, newPass))
            print("ViaIP Utilibill user successfully created...")
        else:
            print("ViaIP Utilibill user not created!")

    def cloud_ultilibill(self):
        self.driver.get(r"https://eziconnect.utilibill.com.au/eziconnect/Login")
        self.driver.find_element_by_tag_name("a").click()
        username = self.driver.find_element_by_name("j_username")
        password = self.driver.find_element_by_id("predigpass")
        username.send_keys("bryan.yeung")
        password.send_keys("\-9y2vUTSee9P]=(")
        self.driver.find_element_by_name("submit").click()
        self.driver.find_element_by_id("submitrequest").click()

        self.driver.get(r"https://eziconnect.utilibill.com.au/eziconnect/NewUser.jsp")
        username = self.driver.find_element_by_name("j_username")
        password = self.driver.find_element_by_id("predigpass")
        drop_down = Select(self.driver.find_element_by_id("group"))
        drop_down.select_by_visible_text("Eziconnect Administrator")
        username.send_keys(self.username)
        newPass = self.passwordGenerator()
        password.send_keys(newPass)

        self.driver.find_element_by_class_name("buttonLrg").click()

        self.driver.get(r"https://viaip.utilibill.com.au/viaip/UserList")
        users_list = self.driver.find_element_by_id("utbListDiv").text
        match = re.search(self.username, users_list)
        if match:
            self.logbook.append((6, self.username, newPass))
            print("Cloudnyne Utilibill user successfully created...")
        else:
            print("Cloudnyne Utilibill user not created!")

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
            print("=========================================")
            print("Supatool: Buroserv Employee or Reseller?")
            print("=========================================")
            print("[1] Buroserv Employee")
            print("[2] Reseller")
            pick = input("")
            if pick == "1":
                print("hi")
            elif pick == "2":
                print("bye")
        elif pick == "2":
            self.ims()
        elif pick == "3":
            print("===============================")
            print("Utilibill: ViaIP or Cloudnyne?")
            print("===============================")
            print("[1] ViaIP")
            print("[2] Cloudnyne")
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
            if pick == "1":
                self.clarus_genex()
            elif pick == "2":
                self.v4_genex()
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
        ac = PasswordResetter(name)
        ac.pickPortal()
        print("New Password: " + ac.newPassword)
        ac.teardown()
        print("Complete!")


if __name__ == "__main__":
    main()
