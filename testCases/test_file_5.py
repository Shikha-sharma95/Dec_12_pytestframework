import pytest
from selenium import webdriver


class Test_5:
    @pytest.mark.web
    @pytest.mark.credence
    def test_credence_url(self, driver_setup):
        self.driver = driver_setup
        self.driver.get("https://credence.in")
        print(f"Title: {self.driver.title}")
        assert "Credence" in self.driver.title, "Title does not match expected"



    @pytest.mark.web
    @pytest.mark.bankapp
    #@pytest.mark.skip
    def test_bankapp_url(self, driver_setup):
        self.driver = driver_setup
        self.driver.get("https://apps.credence.in/")
        print(f"Title: {self.driver.title}")
        assert "Bank Application" in self.driver.title, "Title does not match expected"
        self.driver.quit()


