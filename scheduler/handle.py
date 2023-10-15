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

def group_by(data, column):
    grouped = defaultdict(lambda: [])
    for row in data:
        grouped[row[column]].append(row)

    return grouped

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


class Handler():

    def __init__(self, data, fen_mode = False, spec=None) -> None:
        self.data = data
        self.fen_mode = fen_mode
        self.spec = spec

    def remove_without_course(self):
        data_copy = []
        for row in self.data:
            if not ((not row[COURSE]) or row[COURSE].isspace()):
                data_copy.append(row)

        self.data = data_copy

    def remove_none(self):
        for i, row in enumerate(self.data):
            for j, col in enumerate(row):
                if (not col) or col.isspace():
                    self.data[i][j] = ""

    def day_of_week_to_normal(self, day_of_week: str):
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

    def weeks_to_list(self, weeks: str):
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

    def groups_to_list(self, groups: str):
        delimeters = "".join(set(string.punctuation)-{"-"})
        return list("".join([" " if i in delimeters else i for i in groups]).split())

    def split_groups(self):
        new_data = []
        for row in self.data:
            for group in row[GROUPS]:
                new_row = row[:]
                new_row[GROUPS] = group
                new_data.append(new_row)

        self.data = new_data

    def handle_by_column(self, column, function):
        for i, row in enumerate(self.data):
                self.data[i][column] = function(self.data[i][column])

    def handle(self, fen_mode:bool = False, spec:str = None):
        self.remove_without_course()
        self.remove_none()
        self.handle_by_column(DAYS_OF_WEEKS, self.day_of_week_to_normal)
        self.handle_by_column(WEEKS, self.weeks_to_list)
        self.handle_by_column(GROUPS, self.groups_to_list)
        self.split_groups()

        if self.fen_mode:
            fen_filter = FENFilter(self.data, self.spec)
            if self.spec not in fen_filter.FEN_SPEC_CUT:
                raise Exception('Parameter "spec" must be one of {fen_spec}'.format(fen_spec = list(fen_filter.FEN_SPEC_CUT)))
            self.data = fen_filter.filter_spec()

class FENFilter():

    def __init__(self, data, spec) -> None:
        self.data = data
        self.spec = spec

        self.FEN_SPEC_CUT = {"мен","фін", "екон", "мар", "рб"}
        self.FEN_SPEC = {
            "менеджмент": "мен",
            "фінанси": "фін",
            "економіка": "екон",
            "маркетинг": "мар",
            "розвиток": "рб",
            "рб": "рб"
        }

    def filter_spec(self):
        return list(filter(lambda x: self.is_spec_appropriate(x), self.data))
    
    def is_spec_appropriate(self, spec_data):
        
        spec = self.spec

        course = spec_data[COURSE]
        group = spec_data[GROUPS]
        brackets = list(map(lambda x: x[1:-1], re.findall(r'\(.*?\)', course)))
        
        group_words = ["".join(group) for key, group in itertools.groupby(group, str.isalpha) if key]

        for group_w in group_words:
            group_spec_name = self.spec_name(group_w)
            if group_spec_name == spec:
                return True
            elif group_spec_name in self.FEN_SPEC_CUT:
                return False



        for brack_w in brackets:
            words = ["".join(w) for key, w in itertools.groupby(brack_w, str.isalpha) if key]
            brack_spec_names = set(map(self.spec_name, words))

            if all(brack_spec_names):
                if spec in brack_spec_names:
                    return True
        return False
    

    def spec_name(self, word: str):

        app_word = None
        app_word_l_dist = 0

        for spec in self.FEN_SPEC.keys():
            if word and len(spec) >= len(word):
                l_dist = Levenshtein.ratio(spec[:len(word)].lower(), word.lower())
                if l_dist > 0.6 and l_dist>app_word_l_dist:
                    app_word_l_dist = l_dist
                    app_word = spec
        return self.FEN_SPEC[app_word] if app_word else app_word

if __name__ == "__main__":
    from read import *

    from pprint import pprint
    import numpy as np
    import pandas as pd

    handler = Handler(AbsoluteReader("./files\Економіка_БП-3_Осінь_2023–2024.doc").read(),
                               fen_mode=True,
                               spec="екон")
    handler.handle()
    schedule = handler.data
    
    with open("data.json", 'w', encoding='utf8') as json_file:
        json.dump(schedule, json_file, ensure_ascii=False, indent=4)


    

    # schedule = remove_none((remove_without_course(read_docx("./files/doc/3.docx"))))
    # my_array = np.array(schedule)
    
    # handler = Handler(AbsoluteReader("./files\Економіка_БП-3_Осінь_2023–2024.doc").read(),
    #                            fen_mode=True,
    #                            spec="екон")
    # handler.handle()
    # schedule = handler.data
    # my_array = np.array(list(map(lambda x: list(map(str, x)),schedule)))
    

    # df = pd.DataFrame(my_array, columns = ['День','Час','Дисципліна, викладач', "Група", "Тижні", "Аудиторія"])


    # print(df.to_markdown())
    # print()
    # print()

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

