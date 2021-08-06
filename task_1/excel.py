import io
import os
import sys

import pandas as pd
import win32com.client


# не юзаем эту либу
def generate_excel():
    df = pd.DataFrame({'Фамилия': ['fuck', 'skasd'],
                       'Имя': ['Вася', 'Леха'],
                       'Отчество': ['Максимович', 'Андреевич'],
                       'Дата рождения': ['2019-07-12', '2020-07-12']})

    df.to_excel('./template_1.xlsx', sheet_name='Первый', index=False)


# не юзаем эту либу
def read_excel():
    people_list = pd.read_excel('./template_1.xlsx', index_col='Фамилия')
    print(people_list.head())


def excel_pywin32():
    excel = win32com.client.Dispatch('Excel.Application')
    excel.Visible = True

    current_path = os.getcwd()

    excel_file = os.path.join(current_path, 'persons.xlsx')
    wb_data = excel.Workbooks.Open(excel_file)

    first = wb_data.Worksheets("Лист1").Range("A2")
    print(first)

    wb_data.Close(True)

    excel.Quit()


if __name__ == '__main__':
    # generate_excel()

    # read_excel()

    excel_pywin32()
