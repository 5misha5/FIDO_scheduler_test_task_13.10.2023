import openpyxl
import re

def read_excel(path):

    def get_cell_val(cell, default = None):
        sheet = cell.parent
        rng = [s for s in sheet.merged_cells.ranges if cell.coordinate in s]
        if len(rng)!=0:
            return sheet.cell(rng[0].min_row, rng[0].min_col).value
        elif cell.value:
            return cell.value
        else: 
            return default
    
    def get_nearest_up_cell_val(cell, default = None):
        
        if get_cell_val(cell):
            return get_cell_val(cell)
        
        sheet = cell.parent
        for row in range(cell.row-1, 0, -1):
            cell_val = get_cell_val(sheet.cell(row = row, column = cell.column))
            if cell_val:
                return cell_val
        
        return default

    def get_schedule_rows_range(ws):

        days_of_week = {
            "Понеділок", 
            "Вівторок", 
            "Середа", 
            "Четвер",
            "П'ятниця",
            "Субота",
            "Неділя"
        }
        

        rows = []
        for cellA, cellB in zip(ws["A"], ws["B"]):
            cellA_val = get_cell_val(cellA, default="")
            cellB_val = get_cell_val(cellB, default="")

            if  cellB_val and \
                (
                re.match("[0-9]:[0-9][0-9]-[0-9]:[0-9][0-9]"
                        , cellB_val) or 
                re.match("[0-9]:[0-9][0-9]-[0-9][0-9]:[0-9][0-9]"
                        , cellB_val) or
                re.match("[0-9][0-9]:[0-9][0-9]-[0-9]:[0-9][0-9]"
                        , cellB_val) or 
                re.match("[0-9][0-9]:[0-9][0-9]-[0-9][0-9]:[0-9][0-9]"
                        , cellB_val) or
                cellA_val.capitalize() in days_of_week
                ):
                rows.append(cellB.row)
        return rows[0], rows[-1]
    
    
    ws = openpyxl.load_workbook(path).active

    schedule = []

    rows_range = get_schedule_rows_range(ws)

    #print(rows_range)
    for row in ws.iter_rows(rows_range[0], rows_range[1]):

        row_data = []

        for col in range(0,6):
            #print(row,col)
            if col in {0,1}:
                row_data.append(get_nearest_up_cell_val(row[col]))
            else:
                cell_val = get_cell_val(row[col])
                if col == 2 and not cell_val:
                    break
                row_data.append(get_cell_val(row[col]))
        else:
            schedule.append(row_data)

    return schedule

def read_docx(path):
    ...

def read_doc(path):
    ...




if __name__ == "__main__":
    from pprint import pprint
    import numpy as np

    #schedule = read_excel("./files/xlsx/3.xlsx")
    schedule = read_excel("./files/Прикладна_математика_БП-4_Осінь_2023–2024.xlsx")

    print(all(np.array(schedule).flatten()))

    import numpy as np
    import pandas as pd

    my_array = np.array(schedule)

    df = pd.DataFrame(my_array, columns = ['День','Час','Дисципліна, викладач', "Група", "Тижні", "Аудиторія"])

    print(df.to_string())
    #pprint(schedule)
