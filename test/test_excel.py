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

# Now, you can import modules from the 'scheduler' package


from scheduler.read import *
from scheduler.handle import *
import numpy as np

FILES_PATH = "./files/"

class TestExcel(unittest.TestCase):
    def test_not_null(self):
        excel_files = filter(lambda x: os.path.splitext(x)[1]==".xlsx", os.listdir(FILES_PATH))
        for f in excel_files:
            print(f)
            to_dict(handle(read_excel(FILES_PATH+f)),
                       [COURSE, GROUPS, DAYS_OF_WEEKS],
                       {"time": TIME,
                        "lect_hall": LECT_HALL,
                        "weeks": WEEKS})
            
if __name__ == '__main__':
    unittest.main()