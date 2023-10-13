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


from scheduler.read import read_excel, read_docx, read_doc
import numpy as np

FILES_PATH = "./files/"

class TestReadExcel(unittest.TestCase):
    def test_not_null(self):
        excel_files = filter(lambda x: os.path.splitext(x)[1]==".xlsx", os.listdir(FILES_PATH))
        for f in excel_files:
            if ( "БП-1" not in f) and\
                f not in [
                    "Математика__(Комп`ютерна_математика)_МП-1_Осінь_2023–2024.xlsx",
                    "Прикладна_математика_БП-4_Осінь_2023–2024.xlsx"
                ]:
                print(f)
                self.assertTrue(
                    all(np.array(read_excel(FILES_PATH+f)).flatten())
                )
    def test_cases(self):
        ...



if __name__ == '__main__':
    unittest.main()