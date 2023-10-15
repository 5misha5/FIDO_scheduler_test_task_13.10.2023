import zipfile
import xml.etree.ElementTree

import cols

WORD_NAMESPACE = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
PARA = WORD_NAMESPACE + 'p'
TEXT = WORD_NAMESPACE + 't'
TABLE = WORD_NAMESPACE + 'tbl'
ROW = WORD_NAMESPACE + 'tr'
CELL = WORD_NAMESPACE + 'tc'



def read_docx(path):
    
    with zipfile.ZipFile(path) as docx:
        tree = xml.etree.ElementTree.XML(docx.read('word/document.xml'))
    
    schedule = []

    blank_filler = {
        cols.DAYS_OF_WEEKS: "",
        cols.TIME: ""
    }
    
    for table_node in tree.iter(TABLE):
        
        for row, row_node in enumerate(table_node.iter(ROW)):
            if row == 0:
                continue
            row_list = []
            for col, cell_node in enumerate(row_node.iter(CELL)):
                cell_value = ''.join(node.text for node in cell_node.iter(TEXT))
                if col in blank_filler.keys():
                    if (not cell_value) or cell_value.isspace():
                        cell_value = blank_filler[col]
                    else:
                        blank_filler[col] = cell_value
                      
                row_list.append(cell_value)

            schedule.append(row_list)
    
    return schedule
    

def read_doc(path):
    ...


if __name__ == "__main__":
    from pprint import pprint
    import numpy as np

    #schedule = read_excel("./files/xlsx/3.xlsx")
    schedule = read_docx("./files/doc/3.docx")

    print(all(np.array(schedule).flatten()))

    import numpy as np
    import pandas as pd

    my_array = np.array(schedule)

    df = pd.DataFrame(my_array, columns = ['День','Час','Дисципліна, викладач', "Група", "Тижні", "Аудиторія"])

    print(df.to_markdown())