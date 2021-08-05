import pandas as pd
import win32com.client


def generate_excel():
    df = pd.DataFrame({'Фамилия': ['fuck', 'skasd'],
                       'Имя': ['Вася', 'Леха'],
                       'Отчество': ['Максимович', 'Андреевич'],
                       'Дата рождения': ['2019-07-12', '2020-07-12']})

    df.to_excel('./template_1.xlsx', sheet_name='Первый', index=False)


def read_excel():
    people_list = pd.read_excel('./template_1.xlsx', index_col='Фамилия')
    print(people_list.head())


def excel_pywin32():
    excel = win32com.client.Dispatch('Excel.Application')
    excel.Visible = True

    file = 'path_to_file'
    workbook = excel.Workbook.Open(file)

    _ = input('Press enter to close Excel')
    excel.Quit()


if __name__ == '__main__':
    # generate_excel()

    read_excel()
