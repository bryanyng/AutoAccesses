import os
import re
import subprocess
import time

from selenium import webdriver


class LOLOCreator:

    def __init__(self, fullname, phone):
        self.fullname = fullname
        fullname = fullname.split()
        self.username = fullname[0] + '.' + fullname[1]
        self.firstName = fullname[0]
        self.lastName = fullname[1]
        self.buro_email = self.username + "@buroserv.com.au"
        self.phone = phone

        # Set up webdriver
        profile = webdriver.FirefoxProfile(
            profile_directory=r"Firefox Profile/ln545vx6.default")
        self.driver = webdriver.Firefox(firefox_profile=profile,
                                        executable_path=r"Drivers/geckodriver.exe")
        self.driver.maximize_window()
        self.driver.implicitly_wait(60)

        # Create new folder for uesr
        path = r"C:/Users/Bryan/OneDrive - Buroserv Australia Pty Ltd/Telstra LOLO/LOLO Secondary Certificates/" + self.fullname
        if not os.path.exists(path):
            os.mkdir(path)
            print("Directory ", path, " Created.")
        else:
            print("Directory ", path, " already exists.")

    def create(self):
        lolo_list = ['BHR', 'VIAIP', 'BVV', 'BWA', 'BAA', 'BFS']

        for lolo in lolo_list:
            lolo_name = lolo + self.lastName.upper()
            self.driver.get("https://www.telstrawholesale.com.au/")
            self.driver.get("https://portal.telstrawholesale.com.au")
            subprocess.call("Autoit Scripts/Select" + lolo + "Cert.exe")
            self.driver.find_element_by_link_text("User management").click()
            self.driver.get("https://shopfront.telstra.com.au/online/rne?&rne_eventType=registrar.event.DisplayUsersEvent&rne_sortBy=userName&rne_sortOrder=asc")
            self.driver.find_element_by_xpath('/html/body/table[4]/tbody/tr/td[2]/table[2]/tbody/tr/td[2]/table/tbody/tr[7]/td[2]/p[3]/input').click()
            firstName = self.driver.find_element_by_id("firstname")
            lastName = self.driver.find_element_by_id("lastname")
            phone = self.driver.find_element_by_id("contactNo")
            uniqueId = self.driver.find_element_by_id("uniqueId")
            firstName.send_keys(self.firstName)
            lastName.send_keys(self.lastName)
            phone.send_keys(self.phone)
            uniqueId.send_keys(lolo_name)

            self.driver.find_element_by_xpath('/html/body/table[4]/tbody/tr/td[2]/table[2]/tbody/tr/td[2]/table/tbody/tr[7]/td[2]/div/form/table/tbody/tr[15]/td/div/input[1]').click()

            certId = self.driver.find_element_by_xpath('/html/body/table[4]/tbody/tr/td[2]/table[2]/tbody/tr/td[2]/table/tbody/tr[7]/td[2]/table[4]/tbody/tr[2]/td[1]/div').text
            pin = self.driver.find_element_by_xpath('/html/body/table[4]/tbody/tr/td[2]/table[2]/tbody/tr/td[2]/table/tbody/tr[7]/td[2]/table[4]/tbody/tr[2]/td[3]/div').text

            # Assign permissions to certificate
            self.driver.get("https://portal.telstrawholesale.com.au/group/twcp/user-management")
            self.driver.get("https://shopfront.telstra.com.au/online/rne?&rne_eventType=registrar.event.DisplayUsersEvent&rne_sortBy=userName&rne_sortOrder=asc")
            self.driver.refresh()
            self.driver.find_element_by_link_text(self.lastName.upper() + ', ' + self.firstName.upper()).click()
            self.driver.find_element_by_link_text("Product Access").click()
            self.driver.find_element_by_xpath('//label[text()="LinxOnline Ordering SP Super User"]').click()
            self.driver.find_element_by_xpath('//label[text()="LinxOnline Ordering External User"]').click()
            self.driver.find_element_by_xpath('//label[text()="B2BLineTest"]').click()
            self.driver.find_element_by_xpath('//label[text()="TWCP Line Test Summary"]').click()
            if not lolo == 'BAA':
                self.driver.find_element_by_xpath('//label[text()="TWCP Online User"]').click()
            self.driver.find_element_by_xpath('//label[text()="TWCP Line Test Modem Resynch"]').click()
            self.driver.find_element_by_xpath('//label[text()="NBN LOLO SP Admin"]').click()
            self.driver.find_element_by_xpath('//label[text()="NBN LOLO SP User"]').click()
            if not lolo == 'BFS':
                self.driver.find_element_by_xpath('//label[text()="LinxOnline Service Wholesale   Reporting User"]').click()
            self.driver.find_element_by_xpath('//label[text()="LinxOnline Service Wholesale User"]').click()
            if lolo == 'BHR' or lolo =='VIAIP':
                self.driver.find_element_by_xpath('//label[text()="My Network"]').click()
            self.driver.find_element_by_xpath('/html/body/table[4]/tbody/tr/td[2]/table[2]/tbody/tr/td[2]/table/tbody/tr[8]/td[1]/table/tbody/tr/td/form[1]/table[3]/tbody/tr/td/input[1]').click()
            alert = self.driver.switch_to.alert
            match = re.search('(E.+)\'s', alert.text)
            uniqueId = match.group(1)

            alert.accept()

            # Create LOLO Ordering Profile
            self.driver.get("https://portal.telstrawholesale.com.au/group/twcp/ordering")
            self.driver.get("https://twlinxonlineordering.telstra.com.au/LOLOPRODapp/LoloSplash.jsp")
            time.sleep(2)
            self.driver.switch_to.window(self.driver.window_handles[1])
            self.driver.find_element_by_xpath('/html/body/form/table/tbody/tr[2]/td[2]/table/tbody[2]/tr[2]/td/a[1]/img').click()
            time.sleep(2)
            self.driver.switch_to.window(self.driver.window_handles[1])
            self.driver.find_element_by_link_text("Admin").click()
            self.driver.find_element_by_link_text("Maintain User Profiles").click()
            time.sleep(1)
            lastName = self.driver.find_element_by_xpath('/html/body/form/table[4]/tbody/tr/td[2]/table/tbody/tr[2]/td[2]/center/table/tbody/tr[10]/td[2]/input')
            firstName = self.driver.find_element_by_xpath('/html/body/form/table[4]/tbody/tr/td[2]/table/tbody/tr[2]/td[2]/center/table/tbody/tr[11]/td[2]/input')
            uniqueId_input = self.driver.find_element_by_xpath('/html/body/form/table[4]/tbody/tr/td[2]/table/tbody/tr[2]/td[2]/center/table/tbody/tr[12]/td[2]/input')
            lastName.send_keys(self.lastName)
            firstName.send_keys(self.firstName)
            uniqueId_input.send_keys(uniqueId)

            input("Please add permissions to this user...Press Enter to continue.")

            self.driver.find_element_by_xpath('/html/body/form/table[4]/tbody/tr/td[2]/table/tbody/tr[2]/td[2]/center/table/tbody/tr[17]/td/table/tbody/tr[3]/td[2]/a[1]/img').click()
            self.driver.switch_to.window(self.driver.window_handles[1])
            self.driver.find_element_by_xpath('/html/body/form/table[4]/tbody/tr/td[2]/table/tbody/tr[2]/td[2]/center/table/tbody/tr[18]/td/a/img').click()

            print("===" + lolo_name + "===")
            print(certId)
            print(pin)
            print(uniqueId)

            if lolo_name is not 'BFS' or len(lolo_list) is not 1:
                self.reopen()

    def reopen(self):
        self.teardown()

        # Set up webdriver
        self.driver = webdriver.Ie("Drivers/IEDriverServer.exe")
        self.driver.maximize_window()
        self.driver.implicitly_wait(10)

    def teardown(self):
        self.driver.quit()


def main():
    lolo = LOLOCreator("Sabrina Ongpin", "0284888556")
    lolo.create()
    print("Completed!")
    lolo.teardown()


if __name__ == "__main__":
    main()
