import re
from collections import defaultdict
import string
import json
import Levenshtein
import itertools

DAYS_OF_WEEKS = 0
TIME = 1
COURSE = 2
GROUPS = 3
WEEKS = 4
LECT_HALL = 5


FEN_SPEC_CUT = {"мен","фін", "екон", "мар"}
FEN_SPEC = {
    "менеджмент": "мен",
    "фінанси": "фін",
    "економіка": "екон",
    "маркетинг": "мар"
}

def remove_without_course(data):
    data_copy = []
    for row in data:
        if not ((not row[COURSE]) or row[COURSE].isspace()):
            data_copy.append(row)

    return data_copy


def remove_none(data):
    for i, row in enumerate(data):
        for j, col in enumerate(row):
            if (not col) or col.isspace():
                data[i][j] = ""

    return data

def group_by(data, column):
    grouped = defaultdict(lambda: [])
    for row in data:
        grouped[row[column]].append(row)

    return grouped

def day_of_week_to_normal(day_of_week: str):
    days_of_week = {
        "Понеділок", 
        "Вівторок", 
        "Середа", 
        "Четвер",
        "П'ятниця",
        "Субота",
        "Неділя"
    }

    return max(days_of_week, key=lambda day: Levenshtein.ratio(day_of_week.capitalize(), day))


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



def filter_fen_spec(data, spec):
    def is_spec_appropriate(data, spec: str):

        def spec_name(word: str):

            app_word = None
            app_word_l_dist = 0

            for spec in FEN_SPEC.keys():
                if word and len(spec) >= len(word):
                    l_dist = Levenshtein.ratio(spec[:len(word)], word)
                    if l_dist > 0.6 and l_dist>app_word_l_dist:
                        app_word_l_dist = l_dist
                        app_word = spec
            return FEN_SPEC[app_word] if app_word else app_word


        course = data[COURSE]
        group = data[GROUPS]
        brackets = list(map(lambda x: x[1:-1], re.findall(r'\(.*?\)', course)))
        
        group_words = ["".join(group) for key, group in itertools.groupby(group, str.isalpha) if key]

        for group_w in group_words:
            group_spec_name = spec_name(group_w)
            if group_spec_name == spec:
                return True
            elif group_spec_name in FEN_SPEC_CUT:
                return False



        for brack_w in brackets:
            words = ["".join(w) for key, w in itertools.groupby(brack_w, str.isalpha) if key]
            brack_spec_names = set(map(spec_name, words))

            if all(brack_spec_names):
                if spec in brack_spec_names:
                    return True
        return False
    
    return list(filter(lambda x: is_spec_appropriate(x, spec), data))

def handle(data, fen_mode:bool = False, spec:str = None):
    data = remove_without_course(data)
    data = remove_none(data)
    data = handle_by_column(data, DAYS_OF_WEEKS, day_of_week_to_normal)
    data = handle_by_column(data, WEEKS, weeks_to_list)
    data = handle_by_column(data, GROUPS, groups_to_list)
    data = split_groups(data)

    if fen_mode:
        if spec not in FEN_SPEC_CUT:
            raise Exception('Parameter "spec" must be one of {fen_spec}'.format(fen_spec = list(FEN_SPEC_CUT)))
        data = filter_fen_spec(data, spec)

    return data

if __name__ == "__main__":
    from read import read_excel
    from read_docx import read_docx

    from pprint import pprint
    import numpy as np
    import pandas as pd

    '''schedule = to_dict(handle(read_docx("./files/doc/3.docx")),
                       [COURSE, GROUPS, DAYS_OF_WEEKS],
                       {"time": TIME,
                        "lect_hall": LECT_HALL,
                        "weeks": WEEKS})
    
    with open("data.json", 'w', encoding='utf8') as json_file:
        json.dump(schedule, json_file, ensure_ascii=False, indent=4)'''


    

    schedule = remove_none((remove_without_course(read_docx("./files/doc/3.docx"))))
    my_array = np.array(schedule)
    
    schedule = handle(read_docx("./files/doc/3.docx"),
                               fen_mode=True,
                               spec="екон")
    my_array = np.array(list(map(lambda x: list(map(str, x)),schedule)))
    

    df = pd.DataFrame(my_array, columns = ['День','Час','Дисципліна, викладач', "Група", "Тижні", "Аудиторія"])


    print(df.to_markdown())
    print()
    print()

    #!

    # schedule = group_by(handle(read_docx("./files/doc/3.docx"), 
    #                            fen_mode=True,
    #                            spec="екон"),COURSE)
    
    # for key,item in schedule.items():
    #     my_array = np.array(list(map(lambda x: list(map(str, x)),item)))

    #     df = pd.DataFrame(my_array, columns = ['День','Час','Дисципліна, викладач', "Група", "Тижні", "Аудиторія"])

    #     print(key)
    #     print(df.to_markdown())
    #     print()
    #     print()

