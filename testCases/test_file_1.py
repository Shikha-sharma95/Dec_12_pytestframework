"""
1. Configure Interpreter
2. Check Terminal
Setting--> tools--> terminal--> check shell path --> "C:\\WINDOWS\\system32\\cmd.exe"

3. Install Libraries -->
pip install selenium pytest pytest-xdist pytest-html pytest-sugar pytest-rerunfailures pytest-faker pytest-cov openpyxl

4. Check default test runner in setting

setting --> tools --> python integration tool --> default test runner--> it should be "pytest"


5.Pytest Rules for file, class and method name
    1. testcase file should start with test_ or end with _test
    2. class should start with Test
    3. method/testcase should start with test_

6. How to run test
1. command line --> pytest
it will search all the pytest testcases file means file those are having name _test or test_
and run all these files

2. If you want to run single file then the command is
pytest filepath
e.g.
pytest testCases/test_file_1.py
pytest "D:\Batch Notes\PythonAutomation\CT20\Python_Pytest_Framework\testCases\test_file_1.py"


3. command line --> pytest -v -s (run button in pycharm by default taking pytest -v -s for script runt)
 v  verbose
 s standard output

additional detailing add in output for test run

and when you are running the test case file from pycharm run button
then default it is taking pytest -v -s


4. Run specific testcases with testcase namee
commandline --> pytest -k "keyword"

e.g.
pytest -k "CredKart"
pytest -k "url"
pytest -k "CredKart" or pytest -k "url"
pytest -k not "keyword"



5. Markers in Pytest
how to execute specific test case having specific marker.

command -->
1. pytest -m "marker_name"
2. pytest  -m "marker1 and marker2"
3. pytest  -m "not marker1"
4. pytest  -m "not marker1 and not marker2"
5. pytest -m "marker1 or marker2"




6. How to disable warnings
commandline --> pytest -W ignore
pytest --disable-warnings
pytest -p no:warnings



7. How to Create Html Report
pytest --html=report.html   ---> This will create html report in root folder

pytest --html=Html_Reports/my_report.html ---> This will create html report in


8. How to run test cases parallel

pytest -n=4

by default, it is 1
n--> no of worker processor
n--> depends on number of testcases, system configuration >processor, ram, network connectivity

mostly in real time  n = (5 - 20 )

# -n automatically : this will automatically create worker processor automatically based on number of testcases





9. configured conftest for multiple browser run by setting the user define arg ---browser
such that we can change the browser of test run just by changing the command in pytest command liner
instead of created of change the code.

pytest -v -s --html=Html_Reports/my_report.html -n auto --reruns 1 --browser chrome
pytest -v -s --html=Html_Reports/my_report.html -n auto --reruns 1 --browser firefox
pytest -v -s --html=Html_Reports/my_report.html -n auto --reruns 1 --browser edge
pytest -v --html=Html_Reports/my_report.html -n auto --reruns 1 --browser headless
pytest -v --html=Html_Reports/my_report.html -n auto --browser headless



10. how to re run failed test cases
first we need to install pytest-rerunfailure library

pytest command :
pytest -v -s --html=Html_Reports/my_report.html -n=4 --reruns 1 reruns_delay 2
this will applicable for all test cases


or you can use decorator for specific testcase
@pytest.mark.flaky(reruns=2,reruns_delay=1)


11. pip install pytest-cov
this is use to check code coverage of test run means
checking how much code is use for specific test run.
this is just for information level.

pytest --cov=project_folder_name


12. pip install pytest-faker
this is use to generate fake data for test run

Check test_CredKart_Register testcase in file test_file_6


13. pip install openpyxl
this is use to read and write excel file
create the excel sheet and data and use it in testcases

"""


import pytest


class Test1:


    @pytest.mark.group1 # user define marker--> you have to register
    @pytest.mark.print
    def test_print_1(self, demo_fixture): # testcase
        print("test 1")

    @pytest.mark.group1
    @pytest.mark.print
    def test_print_2(self):
        print("test 2")

    @pytest.mark.group1
    @pytest.mark.math
    def test_add_3(self):
        a = 10
        b = 20
        print(f" Addition of a and b is {a + b}")
        assert a + b == 20, "Test Failed"


    def Sub_4(self):
        a = 10
        b = 20
        print(f" Subtraction of a and b is {a - b}")
        assert a + b == -10, "Test Failed"

    @pytest.mark.group1
    @pytest.mark.math
    def test_Sub_4(self):
        a = 10
        b = 20
        print(f" Subtraction of a and b is {a - b}")
        assert a - b == -10, "Test Failed"


    @pytest.mark.group1
    @pytest.mark.math
    @pytest.mark.skip
    @pytest.mark.flaky(reruns=2, reruns_delay=1)
    def test_Sub_15(self):
        a = 10
        b = 20
        print(f" Subtraction of a and b is {a - b}")
        assert a - b == -10, "Test Failed"


    @pytest.mark.xfail
    def test_Sub_16(self):
        a = 10
        b = 20
        print(f" Subtraction of a and b is {a - b}")
        assert a - b == -10, "Test Failed"



"""
1. To Run All TestCases command:
--> pytest -v -s -p no:warnings

2. To run all tests with html report and disable warnings

--> pytest -v -s -p no:warnings  --html=Html_Reports/my_report.html

3. To run all tests parallel with html report and disable warnings

--> pytest -v -s -p no:warnings --html=Html_Reports/my_report.html -n=4


4. To rerun failed testcases with html report and disable warnings

--> pytest -v -s  -k "test_CredKart_Url" -p no:warnings --html=Html_Reports/my_report.html -n=4 --reruns 1 

5. command to run in multiple browser
pytest -v -s --html=Html_Reports/my_report.html -n auto --reruns 1 --browser chrome
pytest -v -s --html=Html_Reports/my_report.html -n auto --reruns 1 --browser firefox
pytest -v -s --html=Html_Reports/my_report.html -n auto --reruns 1 --browser edge
pytest -v --html=Html_Reports/my_report.html -n auto --reruns 1 --browser headless
pytest -v --html=Html_Reports/my_report.html -n auto --browser headless


6. command to check code coverage
pytest -v --html=Html_Reports/my_report.html -n auto --browser headless --cov=testCases 

"""
