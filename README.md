# mts_web_scraping
Для запуска проекта понадобится предварительно установить Python 3.9 и Tesseract OCR версии 5 (альфа) с https://github.com/UB-Mannheim/tesseract/wiki (у меня Windows 10, поэтому все пути прописывал под нее, на линух придется переписать пути в файле с разгадыванием капчи). Замените путь к тессеракту в task_1/captcha/captcha.py глобальная переменная pytesseract.pytesseract.tesseract_cmd.

Все зависимости лежат в requirements.txt в корне проекта.

Решение задания 1 находится в папке task_1. Запуск из main файла fssprus_parser.py. (.run_excel_persons())

Занесите ФИО должников для поиска в task_1/persons.xlsx, либо измените файл с именами в task_1/excel.py класс FSSP_excel поле input_file_path.

Результат работы программы: task_1/checked_debtors.xlsx.
