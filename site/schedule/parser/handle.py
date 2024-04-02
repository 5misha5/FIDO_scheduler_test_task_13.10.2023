import re
from collections import defaultdict
import string
import json

import itertools
import pandas as pd
import Levenshtein

import cols


def group_by(data, column):
    """
    Group a list of dictionaries based on a specified column.

    Parameters:
        data (list): The 2d list to be grouped.
        column (str): The column index by which to group the data.

    Returns:
        dict: A dictionary where keys are unique values from the specified column, and values are lists of corresponding dictionaries.
    """
    grouped = defaultdict(lambda: [])
    for row in data:
        grouped[row[column]].append(row)

    return grouped


def to_dict(data, nesting: list, last_data: dict):
    """
    Convert a list of dictionaries into a nested dictionary structure based on specified nesting levels.

    Parameters:
        data (list): The 2d list to be converted.
        nesting (list): A list of column names to use for nesting levels.
        last_data (dict): A dictionary specifying the last level of data to include.

    Returns:
        dict: A nested dictionary structure representing the data.
    """
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
    """
    A class for handling and processing data.

    Parameters:
        data (list): The input data to be processed.

        fen_mode (bool): Whether to operate in FEN mode.

        spec (str): The specialization (required if in FEN mode). One of ["мен","фін", "екон", "мар", "рб"].

    Attributes:
        data (list): The data to be processed.

        fen_mode (bool): Whether the class operates in FEN mode.

        spec (str): The specialization to filter data for (only relevant in FEN mode).  One of ["мен","фін", "екон", "мар", "рб"].

    Methods:
        remove_without_course(self):
            Remove rows with missing or empty course information.

        remove_none(self):
            Remove empty or whitespace values in the data.

        day_of_week_to_normal(self, day_of_week: str):
            Convert a day of the week to its normalized form.

        weeks_to_list(self, weeks: str):
            Convert a string of weeks to a list of integers.

        groups_to_list(self, groups: str):
            Convert a string of groups to a list of group names.

        split_groups(self):
            Split rows with multiple groups into separate rows.

        handle_by_column(self, column, function):
            Apply a function to a specific column in the data.

        handle(self, fen_mode: bool = False, spec: str = None):
            Main data processing routine, including various data transformations.
            If in FEN mode, it filters the data based on the provided specialization.

    """

    def __init__(self, data: pd.DataFrame, fen_mode: bool = False, spec=None) -> None:
        """
        Initialize a Handler instance.

        Parameters:
            data (list): The input data to be processed.

            fen_mode (bool): Whether to operate in FEN mode.

            spec (str): The specialization (required if in FEN mode).  One of ["мен","фін", "екон", "мар", "рб"].

        """
        self.data = data
        self.fen_mode = fen_mode
        self.spec = spec

    def remove_without_course(self):
        """
        Remove rows with missing or empty course information.
        """

        self.data.dropna(subset=['course_lecturer'], inplace=True)
        self.data = self.data[self.data["course_lecturer"].str.len() > 4]
        self.data = self.data[~self.data["course_lecturer"].str.isspace()]

    def fill_empty(self, replace_value):
        """
        Remove empty or whitespace values in the data.
        """

        self.data.fillna(replace_value, inplace=True)
        self.data.map(lambda x: replace_value if (not x) or x.isspace() else x)

    def day_of_week_to_index(self, day_of_week: str):
        """
        Convert a day of the week to its normalized form.

        Parameters:
            day_of_week (str): The original day of the week.

        Returns:
            str: The normalized day of the week.
        """
        days_of_week = [
            "Понеділок",
            "Вівторок",
            "Середа",
            "Четвер",
            "П'ятниця",
            "Субота",
            "Неділя"
        ]

        return days_of_week.index(max(days_of_week, key=lambda day: Levenshtein.ratio(day_of_week.capitalize(), day)))

    def weeks_to_list(self, weeks: str):
        """
        Convert a string of weeks to a list of integers.

        Parameters:
            weeks (str): The string representing weeks.

        Returns:
            list of int: A list of week numbers.
        """
        int_arr = list(map(int, filter(lambda x: x, re.split(r'\D+', weeks))))
        del_arr = re.split(r'\d+', weeks)

        weeks_set = set()

        is_range = False
        for i, week, delim in zip(range(len(int_arr)), int_arr, del_arr[1:]):
            if is_range:
                weeks_set.update(set(range(int_arr[i - 1], week + 1)))
            else:
                weeks_set.add(week)

            is_range = "-" in delim

        return list(weeks_set)

    def time_to_lesson_number(self, time:str):
        time_intervals = [
            "8", "10", "11", "13", "15", "16", "18", "19"]
        for i, t in enumerate(time_intervals):
            if time.strip().startswith(t):
                return i+1
        else:
            raise Exception("time '{time}' not found in time_intervals".format(time=time))

    def remove_headers(self):
        headers = {
            "День",
            "Час",
            "Дисципліна, викладач",
            "Група",
            "Тижні",
            "Аудиторія"
        }

        for index, row in self.data.iterrows():
            for column_name in self.data.columns:
                value_to_check = str(row[column_name])
                if any(Levenshtein.ratio(value_to_check.lower(), header.lower()) > 0.8 for header in headers):
                    # If the Levenshtein distance ratio is above a certain threshold (e.g., 0.8), remove the row
                    self.data = self.data.drop(index)
                    break


    def split_course_lecturer(self, input_string):

        # Define the delimiters as a regular expression pattern
        delimiters = r'викл\.|доц\.|ст\.|проф\.|ас\.'

        delims = re.findall(delimiters + r'.*?(?=' + delimiters + '|$)', input_string)

        parts = re.split(delimiters, input_string, 1)

        return pd.Series({'course': parts[0], 'lecturer': delims[0]+parts[1] if len(parts) > 1 else None})

    def handle(self, fen_mode: bool = False, spec: str = None):
        """
        Main data processing routine, including various data transformations.
        If in FEN mode, it filters the data based on the provided specialization.

        Parameters:
            fen_mode (bool): Whether to operate in FEN mode.

            spec (str): The specialization (required if in FEN mode).  One of ["мен","фін", "екон", "мар", "рб"].

        """
        self.remove_without_course()
        self.fill_empty(replace_value="")

        # time_pattern = re.compile(r'^([01]?[0-9]|2[0-3])[:\.][0-5][0-9]-([01]?[0-9]|2[0-3])[:\.][0-5][0-9]$')
        # self.data = self.data[~self.data["time_column"].str.match(time_pattern)]

        self.remove_headers()

        self.data["day_of_week"] = self.data["day_of_week_name"].map(self.day_of_week_to_index)
        self.data["weeks"] = self.data["weeks"].map(self.weeks_to_list)

        self.data["lesson_number"] = self.data["time"].map(self.time_to_lesson_number)

        # Split course_name and lector_name
        self.data[['course_name', 'lecturer_name']] = self.data['course_lecturer'].apply(self.split_course_lecturer)

        if self.fen_mode:
            fen_filter = FENFilter(self.data, self.spec)
            if self.spec not in fen_filter.FEN_SPEC_CUT:
                raise Exception(
                    'Parameter "spec" must be one of {fen_spec}'.format(fen_spec=list(fen_filter.FEN_SPEC_CUT)))
            self.data = fen_filter.filter_spec()


class FENFilter():
    """
    A class for filtering data based on specified criteria.

    Parameters:
        data (list): A 2d list of data to be filtered.

        spec (str): The specialization to filter data for.  One of ["мен","фін", "екон", "мар", "рб"].

    Attributes:
        data (list): The input data to be filtered.

        spec (str): The specialization to filter data for.  One of ["мен","фін", "екон", "мар", "рб"].

        FEN_SPEC_CUT (set): A set of abbreviation names for specializations.

        FEN_SPEC (dict): A dictionary mapping specialization names to abbreviations.

    Methods:
        filter_spec(self):
            Filter the data based on the provided specialization.

        is_spec_appropriate(self, spec_data):
            Check if a specific data entry is appropriate for the given specialization.

        spec_name(self, word):
            Find the specialization abbreviation for a given word.

    """

    def __init__(self, data: list, spec: str) -> None:
        """
        Initialize a FENFilter instance.

        Parameters:
            data (list): A 2d list of data to be filtered.

            spec (str): The specialization to filter data for.  One of ["мен","фін", "екон", "мар", "рб"].

        """
        self.data = data
        self.spec = spec

        self.FEN_SPEC_CUT = {"мен", "фін", "екон", "мар", "рб"}
        self.FEN_SPEC = {
            "менеджмент": "мен",
            "фінанси": "фін",
            "економіка": "екон",
            "маркетинг": "мар",
            "розвиток": "рб",
            "рб": "рб"
        }

    def filter_spec(self):
        """
        Filter the data based on the provided specialization.

        Returns:
            list: A filtered 2d list of data entries.

        """
        return self.data.loc[self.data.apply(lambda row: self.is_appropriate(row), axis=1)]

    def is_appropriate(self, row):
        """
        Check if a specific data entry is appropriate for the given specialization.

        Parameters:
            spec_data (list): Data entry to be checked.

        Returns:
            bool: True if the data entry is appropriate, False otherwise.

        """

        def is_course_appropriate(spec: str, course: str):
            brackets = list(map(lambda x: x[1:-1], re.findall(r'\(.*?\)', course)))
            for brack_w in brackets:
                words = ["".join(w) for key, w in itertools.groupby(brack_w, str.isalpha) if key]
                brack_spec_names = set(map(self.spec_name, words))

                if all(brack_spec_names):
                    if spec in brack_spec_names:
                        return True
                    else:
                        return False

            return True

        def is_group_appropriate(spec: str, group: str):

            group_words = ["".join(group) for key, group in itertools.groupby(group, str.isalpha) if key]

            for group_w in group_words:
                group_spec_name = self.spec_name(group_w)
                if group_spec_name == spec:
                    return True
                elif group_spec_name in self.FEN_SPEC_CUT:
                    return False
            return True

        return is_course_appropriate(self.spec, row["course_lecturer"]) and \
            is_group_appropriate(self.spec, row["group_name"])

    def spec_name(self, word: str):
        """
        Find the specialization abbreviation for a given word

        Parameters:
            word (str): The word to be matched with specialization names.

        Returns:
            str or None: The abbreviation of the specialization if found, or None if not found.

        """
        app_word = None
        app_word_l_dist = 0

        for spec in self.FEN_SPEC.keys():
            if word and len(spec) >= len(word):
                l_dist = Levenshtein.ratio(spec[:len(word)].lower(), word.lower())
                if l_dist > 0.6 and l_dist > app_word_l_dist:
                    app_word_l_dist = l_dist
                    app_word = spec
        return self.FEN_SPEC[app_word] if app_word else None


if __name__ == "__main__":
    from read import *

    from pprint import pprint
    import numpy as np
    import pandas as pd

    # handler = Handler(AbsoluteReader("../files\Економіка_БП-3_Осінь_2023–2024.doc").read(),
    #                   fen_mode=True,
    #                   spec="фін")

    #handler = Handler(AbsoluteReader("../files\Філософія_БП-1_Осінь_2023–2024.docx").read())
    handler = Handler(AbsoluteReader("files/Право_БП-1_Осінь_2023–2024.xlsx").read())

    handler.handle()
    print(handler.data.to_markdown())
    # schedule = to_dict(handler.data, [cols.COURSE, cols.GROUPS, cols.DAYS_OF_WEEKS], {"час": cols.TIME,
    #                                                                                   "аудиторія": cols.LECT_HALL,
    #                                                                                   "тижні": cols.WEEKS})

    # with open("data.json", 'w', encoding='utf8') as json_file:
    #     json.dump(schedule, json_file, ensure_ascii=False, indent=4)

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

    # !

    # schedule = group_by(handle(read_docx("./files/doc/3.docx"), 
    #                            fen_mode=True,
    #                            spec="екон"),cols.COURSE)

    # for key,item in schedule.items():
    #     my_array = np.array(list(map(lambda x: list(map(str, x)),item)))

    #     df = pd.DataFrame(my_array, columns = ['День','Час','Дисципліна, викладач', "Група", "Тижні", "Аудиторія"])

    #     print(key)
    #     print(df.to_markdown())
    #     print()
    #     print()
