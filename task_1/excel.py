import json
import os

import win32com.client

from utility.paths import get_project_root_path

class FSSP():
    # прописать свой путь с должниками либо сделать такой же: <project_path>/task_1/persons.xlsx
    input_file_path = 'persons.xlsx'
    input_sheet_name = 'Лист1'
    output_file_path = 'checked_debtors.xlsx'
    output_sheet_name = 'Должники'

    def get_debtors_to_check(self):
        excel = win32com.client.Dispatch('Excel.Application')
        excel.Visible = True

        current_path = os.getcwd()

        excel_file = os.path.join(current_path, self.input_file_path)
        wb_data = excel.Workbooks.Open(excel_file)
        ws = wb_data.Worksheets(self.input_sheet_name)

        # lastCol = ws.UsedRange.Columns.Count
        last_row = ws.UsedRange.Rows.Count

        persons = list()

        for i in range(2, last_row + 1):

            person = dict()

            person['second_name'] = str(ws.Range("A" + str(i)))
            person['first_name'] = str(ws.Range("B" + str(i)))
            person['third_name'] = str(ws.Range("C" + str(i)))
            person['birth_date'] = str(ws.Range("D" + str(i)))

            persons.append(person)

        # не помню зачем тру оставил
        wb_data.Close(True)

        excel.Quit()

        return persons

    def save_checked_debtors(self):
        excel = win32com.client.Dispatch('Excel.Application')

        current_path = os.getcwd()

        # джоиним текущий путь и путь до файла, куда будем сохранять
        excel_file = os.path.join(current_path, self.output_file_path)
        # book = excel.Workbooks.open(r'checked_debtors.xlsx')
        wb_data = excel.Workbooks.Open(excel_file)
        wb_data.Worksheets(1).Name = self.output_sheet_name

        ws = wb_data.Worksheets(self.output_sheet_name)

        debtors_json_path = f'{get_project_root_path()}/task_1/parsed_debtors.json'

        with open(debtors_json_path, encoding='utf-8') as parsed_debtors_json_file:
            parsed_debtors_json_data = parsed_debtors_json_file.read()

        if parsed_debtors_json_data:
            parsed_debtors_json = json.loads(parsed_debtors_json_data)
        else:
            parsed_debtors_json = None

        ws.Cells(1, 2).Value = 'hue'
        ws.Cells(1, 4).Value = 'ADSDDADW'

        if parsed_debtors_json:

            # TODO где то тут неправильно отсчитывает ячейки
            for debtor_id in range(0, len(parsed_debtors_json)):

                debtor = parsed_debtors_json[f'debtor_{debtor_id + 1}']

                cell_count = 0

                for debtor_part in debtor.values():

                    # TODO мб надо по кейсам идти)

                    # print(debtor_part)
                    # print(type(debtor_part))
                    # print(len(debtor_part))

                    if not isinstance(debtor_part, list):

                        debtor_part_values = list(debtor_part.values())

                        for j in range(0, len(debtor_part)):

                            ws.Cells(debtor_id + 2, cell_count + 1).Value = debtor_part_values[j]

                            cell_count += 1
                    else:
                        for j in range(0, len(debtor_part)):

                            print(debtor_part)
                            debtor_part_part = debtor_part[j]

                            if isinstance(debtor_part_part, dict):
                                debtor_part_part_values = list(debtor_part_part.values())

                                for k in range(0, len(debtor_part_part)):
                                    ws.Cells(debtor_id + 2, cell_count + 1).Value = debtor_part_part_values[k]

                                    cell_count += 1
                            else:
                                ws.Cells(debtor_id + 2, cell_count + 1).Value = debtor_part_part[j]

                                cell_count += 1



        wb_data.Save()
        wb_data.Close()


# не юзаем эту либу
# def generate_excel():
#     df = pd.DataFrame({'Фамилия': ['fuck', 'skasd'],
#                        'Имя': ['Вася', 'Леха'],
#                        'Отчество': ['Максимович', 'Андреевич'],
#                        'Дата рождения': ['2019-07-12', '2020-07-12']})
#
#     df.to_excel('./template_1.xlsx', sheet_name='Первый', index=False)


# не юзаем эту либу
# def read_excel():
#     people_list = pd.read_excel('./template_1.xlsx', index_col='Фамилия')
#     print(people_list.head())


if __name__ == '__main__':
    # generate_excel()

    # read_excel()

    fssp = FSSP()

    # fssp.ffsp_get_debtors_to_check()

    fssp.save_checked_debtors()
