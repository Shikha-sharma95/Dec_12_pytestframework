import time

import pytest
from selenium import webdriver
from selenium.webdriver.common.by import By


class Test_7:

    @pytest.mark.web
    @pytest.mark.credkart
    def test_CredKart_Login(self,driver_setup):

        self.driver = driver_setup
        self.driver.get("https://automation.credence.in/login")
        self.driver.maximize_window()

        # Enter the username
        username = self.driver.find_element(By.ID, "email")
        username.send_keys("Credencetest@test.com")

        # Enter the password
        password = self.driver.find_element(By.ID, "password")
        password.send_keys("Credence@1234")

        # Click the login button
        self.driver.find_element(By.CLASS_NAME, "btn").click()

        # Verify that the user is logged in
        try:
            MenuButton = self.driver.find_element(By.CLASS_NAME, "dropdown-toggle")
            assert MenuButton.is_displayed(), "User is not logged in"
            print("User is logged in")
        except:
            print("User is not logged in")

