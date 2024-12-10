import pytest
from selenium import webdriver
from selenium.webdriver.chrome.options import Options

@pytest.fixture
def demo_fixture():
    print("Test run started--> This is the fixture code")
    yield # after running the testcases below code should run
    print("Test run ended--> This is the fixture code")

#
# @pytest.fixture
# def driver_setup():
#     print("Browser started")
#     driver = webdriver.Chrome()
#     driver.maximize_window()
#     yield driver # driver is return to the test cases or method
#     driver.quit()
#     print("Browser closed")



"""

pytest_addoption --> hook up function --> extended pytest functionality
pytest_addoption functions is used to pass the command line arguments
here we have added argument "--browser to pytest command line.
"""

def pytest_addoption(parser):
    parser.addoption("--browser")


@pytest.fixture
def driver_setup(request):
    browser = request.config.getoption("--browser")
    if browser == "chrome":
        print("Chrome browser started")
        driver = webdriver.Chrome()
    elif browser == "firefox":
        print("Firefox browser started")
        driver = webdriver.Firefox()
    elif browser == "edge":
        print("Edge browser started")
        driver = webdriver.Edge()
    elif browser == "headless":
        print("chrome headless browser started")
        chrome_options = Options()
        chrome_options.add_argument("--headless")  # Run Chrome in headless mode (no visible browser windows)
        driver = webdriver.Chrome(options=chrome_options)
    else:
        print("Chrome browser started")
        driver = webdriver.Chrome()

    driver.maximize_window()
    yield driver # driver is return to the test cases or method
    driver.quit()
    print("Browser closed")


@pytest.fixture(params=[
    ("Credencetest@test.com", "Credence@123", "login_pass"),
    ("Credencetest@test.com1", "Credence@123", "login_fail"),
    ("Credencetest@test.com", "Credence@1231", "login_fail"),
    ("Credencetest@test.com1", "Credence@1231", "login_fail")
])
def get_data_Credkart_login(request):
    return request.param


