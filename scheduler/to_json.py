import json
import os

from scheduler.read import AbsoluteReader
from scheduler.handle import Handler, to_dict
import scheduler.cols as cols

def file_to_json(data_path, json_path = "data.json", fen_mode = False, fen_spec = None):
    """
    Read schedule from a file, process it, and save it as JSON.

    Parameters:
        data_path (str): The path to the input data file.

        json_path (str): The path to the output JSON file. Default is "data.json".

        fen_mode (bool): Whether to operate in FEN mode. Default is False.
        
        fen_spec (str): The specialization (required if in FEN mode). One of ["мен","фін", "екон", "мар", "рб"].

    """
    data = AbsoluteReader(os.path.abspath(data_path)).read()
    handler = Handler(data)
    if fen_mode:
        handler.handle(fen_mode=True, spec=fen_spec)
    else:
        handler.handle()
    schedule = to_dict(handler.data, [cols.COURSE, cols.GROUPS, cols.DAYS_OF_WEEKS], {"час": cols.TIME, 
                                                        "аудиторія": cols.LECT_HALL, 
                                                        "тижні": cols.WEEKS})
    
    with open(os.path.abspath(json_path), 'w', encoding='utf8') as json_file:
        json.dump(schedule, json_file, ensure_ascii=False, indent=4)
    


if __name__ == "__main__":
    #file_to_json("./files/Аналіз_вразливостей_інформаційних_систем_БП-1_Осінь_2023–2024.xlsx")
    #file_to_json("./files/Соціальна_робота_МП-1_Осінь_2023–2024.doc")
    file_to_json("./files/Економіка_БП-1_Осінь_2023–2024 (2).doc", fen_mode=True, fen_spec="екон")

