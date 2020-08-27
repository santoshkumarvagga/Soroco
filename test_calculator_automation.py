'''This module illustrates sample automation system like calculator'''

import unittest

class TestCalculator(unittest.TestCase):
    def setUp(self):
        '''Provides sample input'''
        self.args = [2,3,4]

    # utility methods from here on
    def add(self):
        sum = 0
        for i in self.args:
            sum = sum + i
        return sum

    def sub(self):
        diff = self.args[0]
        for i in self.args[1:]:
            diff = diff -i
        return diff

    def mul(self):
        mul = 1
        for i in self.args:
            mul *= i
        return mul

    def div(self):
        rem = self.args[0]
        for i in self.args[1:]:
            if i==0:
                raise ZeroDivisionError
            else:
                rem = rem/i
        return rem

    # test methods from here on
    def test_addition(self):
        assert self.add() == 9, 'addition didnt work as expected..!!' 

    def test_subtraction(self):
        assert self.sub() == -5, 'subtraction didnt work as expected..!!'

    def test_multiplication(self):
        assert self.mul() == 24, 'multiplication didnt work as expected..!!'

    def test_division(self):
        assert self.div() == 0.16666666666666666, 'division didnt work as expected..!!'

    def tearDown(self):
        return "Testing ended..!!"

a = TestCalculator()

if __name__ == '__main__':
    unittest.main()
