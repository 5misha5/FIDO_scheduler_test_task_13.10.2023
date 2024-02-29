import unittest
import sys
import os

# Specify the path to the directory containing the 'scheduler' package
scheduler_directory = 'E:\Misha\GitHub\FIDO_scheduler_test_task_13.10.2023'

# Check if the directory exists before adding it
if os.path.exists(scheduler_directory):
    sys.path.append(scheduler_directory)
else:
    print(f"The directory '{scheduler_directory}' does not exist.")


from scheduler.handle import *
from scheduler.read import *
from scheduler.cols import *
import numpy as np

FILES_PATH = "/files/"

class TestWeeksToArray(unittest.TestCase, Handler):

    def test1(self):
        self.assertEqual(set(self.weeks_to_list("12,15")), {12,15})
    def test2(self):
        self.assertEqual(set(self.weeks_to_list("3 7")), {3,7})
    def test3(self):
        self.assertEqual(set(self.weeks_to_list("2-5")), {2,3,4,5})
    def test4(self):
        self.assertEqual(set(self.weeks_to_list("12-15")), {12,13,14,15})
    def test5(self):
        self.assertEqual(set(self.weeks_to_list("8-10, 13, 15")), {8,9,10,13,15})
    def test6(self):
        self.assertEqual(set(self.weeks_to_list("3, 5, 7, 10-15")), {3,5,7,10,11,12,13,14,15})
    def test7(self):
        self.assertEqual(set(self.weeks_to_list("3 6-8 f16, 18" )), {3,6,7,8,16,18})
    def test8(self):
        self.assertEqual(set(self.weeks_to_list("3 6-8 у16, 18")), {3,6,7,8,16,18})
    def test9(self):
        self.assertEqual(set(self.weeks_to_list("-12 14, 15 -")), {12, 14, 15})
    def test10(self):
        self.assertEqual(set(self.weeks_to_list("7-10a-13, 5")), {7,8,9,10, 11, 12, 13, 5})
    def test11(self):
        self.assertEqual(set(self.weeks_to_list("6-9, 7 8 11")), {6,7,8,9,11})
    def test12(self):
        self.assertEqual(set(self.weeks_to_list("-12 14-л17 -")), {12, 14,15,16,17})
    '''def test13(self):
        self.assertEqual(set(self.weeks_to_list(""), {})
    def test14(self):
        self.assertEqual(set(self.weeks_to_list(""), {})
    def test15(self):
        self.assertEqual(set(self.weeks_to_list(""), {})'''


class TestHandler(unittest.TestCase):
    def test_is_working_Handler(self):
        files = filter(lambda x: os.path.splitext(x)[1] in {".xlsx", ".doc", ".docx"}, os.listdir(FILES_PATH))
        for f in files:
            if not f.startswith("~"):
                print(f)
                handler = AbsoluteReader(FILES_PATH+f).read()
                handler.handle()
                self.assertTrue(
                    handler.data
                )

class TestToDict(unittest.TestCase):
    def test_is_working_Handler(self):
        files = filter(lambda x: os.path.splitext(x)[1] in {".xlsx", ".doc", ".docx"}, os.listdir(FILES_PATH))
        for f in files:
            if not f.startswith("~"):
                print(f)
                try:
                    data = AbsoluteReader(FILES_PATH+f).read()
                    handler = Handler(data)
                    handler.handle()
                    to_dict(handler.data, [COURSE, GROUPS, DAYS_OF_WEEKS], {"час": TIME, 
                                                                        "аудиторія": LECT_HALL, 
                                                                        "тижні": WEEKS})

                except Exception as e:
                    self.fail("to_dict raised Exception unexpectedly!\n", e)

# TODO add groups_to_set tests

#TODO spec_name tests


if __name__ == '__main__':
    unittest.main()