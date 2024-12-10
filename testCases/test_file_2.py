import pytest


class Test2:

    @pytest.mark.math
    def test_Mul_4(self):
        a = 10
        b = 20
        print(f" Multiplication of a and b is {a * b}")
        assert a * b == 200, "Test Failed"

    @pytest.mark.math
    def test_add_5(self):
        a = 10
        b = 200
        print(f" Addition of a and b is {a + b}")
        assert a + b == 210, "Test Failed"

