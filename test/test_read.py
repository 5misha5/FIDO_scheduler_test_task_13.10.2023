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


from scheduler.read import AbsoluteReader
import numpy as np

FILES_PATH = "./files/"

class TestReadExcel(unittest.TestCase):
    def test_is_working_Ablosute(self):
        files = filter(lambda x: os.path.splitext(x)[1] in {".xlsx", ".doc", ".docx"}, os.listdir(FILES_PATH))
        for f in files:
            if not f.startswith("~"):
                print(f)
                self.assertTrue(
                    AbsoluteReader(FILES_PATH+f).read()
                )


if __name__ == '__main__':
    unittest.main()