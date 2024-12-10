import time

import pytest
from selenium import webdriver
from selenium.webdriver.common.by import By


class Test_7:

    @pytest.mark.web
    @pytest.mark.credkart
    def test_CredKart_Login(self,driver_setup, get_data_Credkart_login):
        email = get_data_Credkart_login[0]
        print("Email",email)
        password = get_data_Credkart_login[1]
        print("Password", password)
        expected_result = get_data_Credkart_login[2]
        print("Expected",expected_result)
        # email, password, expected_result = get_data_Credkart_login
        self.driver = driver_setup
        self.driver.get("https://automation.credence.in/login")
        self.driver.maximize_window()

        # Enter the username
        username = self.driver.find_element(By.ID, "email")
        username.send_keys(email)

        # Enter the password
        time.sleep(2)
        password_input = self.driver.find_element(By.ID, "password")
        password_input.send_keys(password)

        # Click the login button
        self.driver.find_element(By.CLASS_NAME, "btn").click()

        # Verify that the user is logged in
        try:
            MenuButton = self.driver.find_element(By.CLASS_NAME, "dropdown-toggle")
            assert MenuButton.is_displayed(), "User is not logged in"
            print("User is logged in")
            actual_result = "login_pass"
        except:
            actual_result = "login_fail"
            print("User is not logged in")

        if expected_result == actual_result:
            assert  True
        else:
            assert False
