import pytest
from selenium import webdriver
from time import sleep
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import Xutilities
import pytest_html

class TestDemo():

    @pytest.fixture()
    def setup(self):
        self.driver = webdriver.Chrome()
        self.driver.get('https://opensource-demo.orangehrmlive.com/')
        self.driver.maximize_window()
        sleep(2)
        yield
        sleep(5)
        self.driver.close()
        
    def test_url_title(self, setup):
        print("Website URL:", self.driver.current_url)
        print("Website Title:", self.driver.title)
        try:
            assert self.driver.title == "OrangeHrm"
        except(AssertionError):
            assert self.driver.title == "OrangeHRM"
    
    def test_case_login_page(self, setup):
        file_path = 'D:\Python\Execersice file\open pyxl excersice.xlsx'
        rows = Xutilities.get_row_count(file_path, "Sheet1")

        for r in range(2, rows+1):
            self.username = Xutilities.read_data(file_path, "Sheet1", r, 2)
            self.password = Xutilities.read_data(file_path, "Sheet1", r, 3)
            self.driver.find_element(By.ID,"txtUsername").send_keys(self.username)
            sleep(2)
            self.driver.find_element(By.ID,'txtPassword').send_keys(self.password)
            sleep(2)
            self.driver.find_element(By.NAME,'Submit').click()
            sleep(5)

            if self.driver.current_url == "https://opensource-demo.orangehrmlive.com/index.php/dashboard":
                print("Test Passed")
                Xutilities.write_data(file_path, "Sheet1", r, 4, "Test Passed")
                sleep(2)
                self.driver.find_element(By.ID,'welcome').click()
                sleep(2)
                self.driver.find_element(By.CSS_SELECTOR,"div:nth-child(1) div.panelContainer:nth-child(14) ul:nth-child(1) li:nth-child(3) > a:nth-child(1)").click()
                sleep(2)
            else:
                print("Test Failed")
                Xutilities.write_data(file_path, "Sheet1", r, 4, "Test Failed")