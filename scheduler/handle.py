import re
from collections import defaultdict
import string
import json

DAYS_OF_WEEKS = 0
TIME = 1
COURSE = 2
GROUPS = 3
WEEKS = 4
LECT_HALL = 5

def remove_none(data):
    for i, row in enumerate(data):
        for j, col in enumerate(row):
            if col == None:
                data[i][j] = ""

    return data

def group_by(data, column):
    grouped = defaultdict(lambda: [])
    for row in data:
        grouped[row[column]].append(row)

    return grouped


def weeks_to_list(weeks: str):
    int_arr = list(map(int, filter(lambda x: x, re.split(r'\D+', weeks))))
    del_arr = re.split(r'\d+', weeks)

    weeks_set = set()

    is_range = False
    for i,week, delim in zip(range(len(int_arr)), int_arr, del_arr[1:]):
        if is_range:
            weeks_set.update(set(range(int_arr[i-1], week+1)))
        else:
            weeks_set.add(week)
        
        is_range = "-" in delim
    
    return list(weeks_set)

def groups_to_list(groups: str):
    delimeters = "".join(set(string.punctuation)-{"-"})
    return list("".join([" " if i in delimeters else i for i in groups]).split())

def split_groups(data):
    new_data = []
    for row in data:
        for group in row[GROUPS]:
            new_row = row[:]
            new_row[GROUPS] = group
            new_data.append(new_row)

    return new_data



def handle_by_column(data, column, function):
    for i, row in enumerate(data):
            data[i][column] = function(data[i][column])

    return data

def to_dict(data, nesting: list, last_data: dict):
    
    if not nesting:
        last_data_list = []
        for d in data:
            last_data_list.append({key: d[item] for key, item in last_data.items()})

        return last_data_list

    data_dict = dict()

    for key, item in group_by(data, column=nesting[0]).items():
        data_dict[key] = to_dict(item, nesting[1:], last_data=last_data)
    return data_dict

def handle(data):
    data = remove_none(data)
    data = handle_by_column(data, WEEKS, weeks_to_list)
    data = handle_by_column(data, GROUPS, groups_to_list)
    data = split_groups(data)

    return data

if __name__ == "__main__":
    from read import read_excel

    from pprint import pprint
    import numpy as np

    schedule = to_dict(handle(read_excel("./files/xlsx/3.xlsx")),
                       [COURSE, GROUPS, DAYS_OF_WEEKS],
                       {"time": TIME,
                        "lect_hall": LECT_HALL,
                        "weeks": WEEKS})
    #schedule = group_by(handle(read_excel("./files/Прикладна_математика_БП-4_Осінь_2023–2024.xlsx")),2)

    with open("data.json", 'w', encoding='utf8') as json_file:
        json.dump(schedule, json_file, ensure_ascii=False, indent=4)
'''
    import numpy as np
    import pandas as pd

    for key,item in schedule.items():
        my_array = np.array(list(map(lambda x: list(map(str, x)),item)))

        df = pd.DataFrame(my_array, columns = ['День','Час','Дисципліна, викладач', "Група", "Тижні", "Аудиторія"])

        print(key)
        print(df.to_markdown())
        print()
        print()'''

