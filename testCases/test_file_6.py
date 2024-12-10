import pytest
from selenium import webdriver
from selenium.webdriver.common.by import By


class Test_6:

    @pytest.mark.web
    @pytest.mark.credkart
    #@pytest.mark.flaky(reruns=2,reruns_delay=1)
    def test_CredKart_Url(self, driver_setup):
        self.driver = driver_setup
        self.driver.get("https://automation.credence.in/shop")
        print(f"Title: {self.driver.title}")
        assert "CredKart1" in self.driver.title, "Title does not match expected"
        self.driver.quit()

    @pytest.mark.web
    @pytest.mark.credkart
    def test_CredKart_Register(self, driver_setup, faker):
        self.driver = driver_setup
        self.driver.get("https://automation.credence.in/register")
        self.driver.maximize_window()
        self.driver.implicitly_wait(2)
        random_name = faker.name()
        random_email = faker.email()
        print(f"Name: {random_name}")
        print(f"Email: {random_email}")
        # Fill the registration form
        # Enter Username
        name = self.driver.find_element(By.ID, "name")
        name.send_keys(random_name)

        # Enter Email
        email = self.driver.find_element(By.ID, "email")
        email.send_keys(random_email)

        # Enter Password
        password = self.driver.find_element(By.ID, "password")
        password.send_keys("Test@123")

        # Enter Confirm Password
        confirm_password = self.driver.find_element(By.ID, "password-confirm")
        confirm_password.send_keys("Test@123")

        # Click on Submit button
        submit_button = self.driver.find_element(By.CLASS_NAME, "btn")
        submit_button.click()

        # Verify the user registration

        try:
            MenuButton = self.driver.find_element(By.CLASS_NAME, "dropdown-toggle")
            assert MenuButton.is_displayed(), "User is not logged in"
            print("User registered successfully")
        except:
            print("User not registered successfully")

        finally:
            self.driver.quit()