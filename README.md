# UKMA_schedule_file_scraper

## About
Цей проєкт допомагає перевести файл розкладу НАУКМА (.xlsx, .docx, .doc) в .json файл.

Інформація зберігається в форматі
```
{
    "Назва предмету": {
        "групи": {
            "день тижня": {
                "назва": "може бути номер групи та лекції",
                "час": "",
                "тижні": [список тижнів, коли йде предмет],
                "аудиторія": ""
            }
        }
    }
}
```

## Getting Started

### Prerequisites

Для роботи програми потрібні такі бібліотеки:

- openpyxl
- zipfile
- xml
- win32com
- json
- Levenshtein

### Usage

Для того щоб перетоворити дані в json потрібно

```
import scheduler

data_path = "file.docx" # Шлях до вашого файлу
json_path = "data.json" # Шлях до місця збереження файлу з назвою файлу, наприклад "C:/Розклад/ФІ/ПМ1.json"
scheduler.to_json.file_to_json(data_path, json_path)
```

Для того щоб обробити дані ФЕНу потрібно додатково вказати fen_mode = True а також вказати назву спеціальності для якої збираєтеся обробити дані, наприклад
```
spec = "екон" # Одне з ["мен","фін", "екон", "мар", "рб"]
scheduler.to_json.file_to_json(data_path, json_path, fen_mode = True, spec = spec) 
# В файлі json буде зберігатися інформація тільки про економістів
```
Якщо потрібно отримати інформацію про весь факультет - запустіть функцію змінюючи spec

Для детальнішої інформації читайте docs

## Issues
Може некоректно працювати розділення на групи для деяких значень, наприкад "1гр. (с)" не буде визначатися як одна група

