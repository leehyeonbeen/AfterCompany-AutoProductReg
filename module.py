from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    NoSuchElementException,
    NoAlertPresentException,
    TimeoutException,
)
import time
import urllib.request
import openpyxl


class AFT_Automator:
    def __init__(self) -> None:
        self.standby = 1.5
        self.standby_pageshift = 5
        self.driver = webdriver.Safari()
        self.driver.maximize_window()
        self.driver.implicitly_wait(15)
        self.wait = WebDriverWait(self.driver, 15)

    def FillTextbox(self, XPATH, text, enter=False):
        time.sleep(self.standby)
        element = self.driver.find_element(By.XPATH, XPATH)
        element.clear()
        element.send_keys(text)
        if enter:
            element.send_keys(Keys.RETURN)
        time.sleep(self.standby)

    def ClickButton(self, XPATH):
        time.sleep(self.standby)
        element = self.driver.find_element(By.XPATH, XPATH)
        element.click()
        time.sleep(self.standby)

    def SelectBox(self, XPATH, visible_txt):
        selectbox = Select(self.driver.find_element(By.XPATH, XPATH))
        selectbox.select_by_visible_text(visible_txt)

    def AFT_Login(self):
        ID = "ceo@8division.com"
        PW = "8division11@@"
        self.driver.get("https://premium.aftercompany.co.kr")
        self.FillTextbox('//*[@id="container"]/div/div[1]/form/div/div[1]/input', ID)
        self.FillTextbox('//*[@id="container"]/div/div[1]/form/div/div[2]/input', PW)
        self.ClickButton('//*[@id="container"]/div/div[1]/form/button')
        self.driver.switch_to.alert().accept()
        # time.sleep(15)
        pass

    def Web_RegProducts(self):
        self.AFT_Login()
        """
        Create a new category
        """
        self.ClickButton('//*[@id="userMenu"]/li[1]/a')
        self.ClickButton('//*[@id="userMenu"]/li[1]/ul/li[3]/a')
        urllib.request.urlretrieve(
            "https://premium.aftercompany.co.kr/storage/template/excel_template_for_prod_v3.xlsx",
            "RegProduct_template.xlsx",
        )
        xl = openpyxl.load_workbook("RegProduct_template.xlsx")
        pass
        """
        Process excel form
        """

    def Web_RegOfflineStocks(self):
        """
        Read excel file and list all available item names
        """
        self.ClickButton('//*[@id="userMenu"]/li[2]/a')
        self.ClickButton('//*[@id="userMenu"]/li[2]/ul/li[4]/a')
        """
        Register offline stocks with quantity 999
        Click register button
        """

    def POS_AutoAcceptRequests(self):
        # Login
        ID = "8division12"
        PW = "8division12"
        self.driver.get("https://pos.aftercompany.co.kr")
        self.FillTextbox('//*[@id="aftpos-app"]/div/div[1]/form/div/div[1]/input', ID)
        self.FillTextbox('//*[@id="aftpos-app"]/div/div[1]/form/div/div[2]/input', PW)
        self.ClickButton('//*[@id="aftpos-app"]/div/div[1]/form/button')

        # Get to the request menu
        time.sleep(self.standby_pageshift)
        self.ClickButton('//*[@id="aftpos-app"]/header/div[1]/ul/li[2]/a')
        self.ClickButton('//*[@id="aftpos-app"]/header/div[2]/ul[1]/li[3]/a')
        time.sleep(self.standby_pageshift)
        counter = 0
        while True:
            try:
                self.SelectBox(
                    '//*[@id="aftpos-body"]/div/form/div[2]/div[1]/select',
                    "승인(진행중)",
                )
                self.SelectBox(
                    '//*[@id="aftpos-body"]/div/form/div[2]/div[2]/select',
                    "추가요청",
                )
                self.ClickButton(
                    '//*[@id="aftpos-body"]/div/form/div[3]/div/button'
                )  # 조회
                self.ClickButton('//*[@id="dataBody"]/tr[1]/td[1]')
                self.ClickButton('//*[@id="acceptItem"]')
                alert = self.driver.switch_to.alert
                alert.accept()
                WebDriverWait(self.driver, 10).until(
                    EC.visibility_of_element_located(
                        (By.XPATH, '//*[starts-with(@id, "msgbox_")]/div/div/button')
                    )
                )
                self.driver.find_element(
                    By.XPATH, '//*[starts-with(@id, "msgbox_")]/div/div/button'
                ).click()
                time.sleep(self.standby_pageshift)
                counter += 1
                print(f"I have accepted {counter} requests.")
            except (NoSuchElementException, NoAlertPresentException) as error:
                print(f"No more requests found. Exiting...")
                break
            
    def Delete_Items(self, excel_path):
        """
        Read excel file and makeup a list 
        """
        
        """
        Read web-registered items and compare with the list
        """
        
        """
        Loop-deletion
        """
        
