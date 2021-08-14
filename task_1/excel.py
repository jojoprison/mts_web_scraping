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
        # не надо, чтоб появлялось окно с экселем и потом исчезало
        # excel.Visible = True

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

    def save_checked_debtors(self, debtors_json):
        excel = win32com.client.Dispatch('Excel.Application')

        current_path = os.getcwd()

        # джоиним текущий путь и путь до файла, куда будем сохранять
        excel_file = os.path.join(current_path, self.output_file_path)
        # book = excel.Workbooks.open(r'checked_debtors.xlsx')
        wb_data = excel.Workbooks.Open(excel_file)

        # сколько уже листов есть в файле
        sheet_count = wb_data.Worksheets.Count

        # если стоит дефолтный Лист1, то пишем прямо в него
        if sheet_count == 1 and wb_data.Worksheets(1).Name == 'Лист1':
            # имя первого листа
            new_sheet_name = f'{self.output_sheet_name}_1'
        else:
            # лист добавляется в НАЧАЛО всех листов, поэтому работаем с ним
            wb_data.Worksheets.Add()
            # имя свежесозданного листа
            new_sheet_name = f'{self.output_sheet_name}_{sheet_count + 1}'

        print(new_sheet_name)

        wb_data.Worksheets(1).Name = new_sheet_name

        ws = wb_data.Worksheets(new_sheet_name)

        # debtors_json_path = f'{get_project_root_path()}/task_1/parsed_debtors.json'
        #
        # with open(debtors_json_path, encoding='utf-8') as parsed_debtors_json_file:
        #     parsed_debtors_json_data = parsed_debtors_json_file.read()
        #
        # if parsed_debtors_json_data:
        #     debtors_json = json.loads(parsed_debtors_json_data)
        # else:
        #     debtors_json = None

        # формируем заголовки
        ws.Cells(1, 1).Value = 'Должник.ФИО'
        ws.Cells(1, 2).Value = 'Должник.Дата рождения'
        ws.Cells(1, 3).Value = 'Должник.Адрес'

        ws.Cells(1, 4).Value = 'ИП.Номер'

        ws.Cells(1, 5).Value = 'Реквизиты ИП.Первые'
        ws.Cells(1, 6).Value = 'Реквизиты ИП.Вторые (если есть)'
        ws.Cells(1, 7).Value = 'Реквизиты ИП.Орган'

        ws.Cells(1, 8).Value = 'Окончание ИП.Причина'
        ws.Cells(1, 9).Value = 'Окончание ИП.Дата'

        ws.Cells(1, 10).Value = 'Долг.Что'
        ws.Cells(1, 11).Value = 'Долг.Сколько'
        ws.Cells(1, 12).Value = 'Долг 2.Что'
        ws.Cells(1, 13).Value = 'Долг 2.Сколько'

        ws.Cells(1, 14).Value = 'Отдел судебных приставов.Название'
        ws.Cells(1, 15).Value = 'Отдел судебных приставов.Адрес'

        ws.Cells(1, 16).Value = 'Судебный пристав.ФИО'
        ws.Cells(1, 17).Value = 'Судебный пристав.Телефон'

        if debtors_json:

            for debtor_id in range(0, len(debtors_json)):

                row = debtor_id + 2

                debtor = debtors_json[f'debtor_{debtor_id + 1}']

                # инфа о должнике
                ws.Cells(row, 1).Value = debtor['debtor_info']['name']
                ws.Cells(row, 2).Value = debtor['debtor_info']['birth_date']
                ws.Cells(row, 3).Value = debtor['debtor_info']['place']

                # исполнительные производства
                ws.Cells(row, 4).Value = debtor['enforcement_proceedings']

                # детали доков
                ws.Cells(row, 5).Value = debtor['document_details']['order']
                # если есть второе постановление
                if 'order_2' in debtor['document_details']:
                    ws.Cells(row, 6).Value = debtor['document_details']['order_2']
                ws.Cells(row, 7).Value = debtor['document_details']['authority']

                # окончание ИП
                ws.Cells(row, 8).Value = debtor['ep_end']['reason']
                ws.Cells(row, 9).Value = debtor['ep_end']['date']

                # штрафы
                performance_subject = debtor['performance_subject']

                # если там лист, значит точно 2 долга (я так записывал)) - добавляем еще в две ячейки значения
                if isinstance(performance_subject, list):
                    ws.Cells(row, 10).Value = debtor['performance_subject'][0]['name']
                    ws.Cells(row, 11).Value = debtor['performance_subject'][0]['amount']
                    ws.Cells(row, 12).Value = debtor['performance_subject'][1]['name']
                    ws.Cells(row, 13).Value = debtor['performance_subject'][1]['amount']
                else:
                    ws.Cells(row, 10).Value = performance_subject['name']
                    ws.Cells(row, 11).Value = performance_subject['amount']

                # инфа о департаменте
                ws.Cells(row, 14).Value = debtor['department_of_bailiffs']['name']
                ws.Cells(row, 15).Value = debtor['department_of_bailiffs']['address']

                # инфа о приставе
                ws.Cells(row, 16).Value = debtor['bailiff']['name']
                ws.Cells(row, 17).Value = debtor['bailiff']['phone']

            # старый алгоритмический способ, не до конца робит, где-то неправильно отсчитывает ячейки
            # for debtor_id in range(0, len(parsed_debtors_json)):
            #
            #     debtor = parsed_debtors_json[f'debtor_{debtor_id + 1}']
            #
            #     row = debtor_id + 3
            #     cell_count = 1
            #
            #     for debtor_part in debtor.values():
            #
            #         # TODO мб надо по кейсам идти)
            #
            #         # print(debtor_part)
            #         # print(type(debtor_part))
            #         # print(len(debtor_part))
            #
            #         if not isinstance(debtor_part, list):
            #
            #             debtor_part_values = list(debtor_part.values())
            #
            #             for z in range(0, len(debtor_part)):
            #
            #                 # print(f'write_1: {debtor_part_values[z]}')
            #
            #                 ws.Cells(row, cell_count).Value = debtor_part_values[z]
            #
            #                 cell_count += 1
            #         else:
            #             for z in range(0, len(debtor_part)):
            #
            #                 # print(debtor_part)
            #                 debtor_part_part = debtor_part[z]
            #
            #                 if isinstance(debtor_part_part, dict):
            #                     debtor_part_part_values = list(debtor_part_part.values())
            #
            #                     for k in range(0, len(debtor_part_part)):
            #
            #                         # print(f'write_2: {debtor_part_part_values[k]}')
            #
            #                         ws.Cells(row, cell_count).Value = debtor_part_part_values[k]
            #
            #                         cell_count += 1
            #                 else:
            #
            #                     # print(f'write_3: {debtor_part_part}')
            #
            #                     ws.Cells(row, cell_count).Value = debtor_part_part
            #
            #                     cell_count += 1

        # wb_data.Save()
        wb_data.Close(SaveChanges=1)


# TODO переместить в отдельный файл
def save_to_json(debtors_json):
    debtors_json_path = f'{get_project_root_path()}/task_1/parsed_debtors.json'

    with open(debtors_json_path, encoding='utf-8') as parsed_debtors_json_file:
        parsed_debtors_json_data = parsed_debtors_json_file.read()

    if parsed_debtors_json_data:
        parsed_debtors_json = json.loads(parsed_debtors_json_data)
    else:
        parsed_debtors_json = None

    if parsed_debtors_json:
        parsed_debtors_json[f'debtor_{len(parsed_debtors_json) + 1}'] = debtors_json
    else:
        parsed_debtors_json = {'debtor_1': debtors_json}

    parsed_debtors_list_json = json.dumps(parsed_debtors_json,
                                          ensure_ascii=False, indent=4)

    with open(debtors_json_path, 'w', encoding='utf-8') as json_file:
        json_file.write(parsed_debtors_list_json)


# old
def get_debtors_json():
    debtors_json_path = f'{get_project_root_path()}/task_1/parsed_debtors.json'

    with open(debtors_json_path, encoding='utf-8') as parsed_debtors_json_file:
        parsed_debtors_json_data = parsed_debtors_json_file.read()

    if parsed_debtors_json_data:
        parsed_debtors_json = json.loads(parsed_debtors_json_data)
    else:
        parsed_debtors_json = None

    return parsed_debtors_json


def clear_json_file():
    debtors_json_path = f'{get_project_root_path()}/task_1/parsed_debtors.json'

    # вот таким нехитрым способом очищаем его
    f = open(debtors_json_path, 'w')
    f.close()

    return True


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
