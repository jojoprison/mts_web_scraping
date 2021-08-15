import json
import os
import datetime

import win32com.client

from utility.paths import get_project_root_path


class FSSP_excel():
    # прописать свой путь с должниками либо сделать такой же: <project_path>/task_1/persons.xlsx
    input_file_path = 'persons.xlsx'
    input_sheet_name = 'Лист1'
    output_file_path = 'checked_debtors.xlsx'
    output_sheet_name = 'Должники'

    # забирает ФИО и дату должников на проверку из файла excel, путь прописывается в поле класса выше
    def get_debtors_to_check(self):
        excel = win32com.client.Dispatch('Excel.Application')
        # не надо, чтоб появлялось окно с экселем и потом исчезало
        # excel.Visible = True

        current_path = os.getcwd()

        excel_file = os.path.join(current_path, self.input_file_path)
        wb_data = excel.Workbooks.Open(excel_file)
        ws = wb_data.Worksheets(self.input_sheet_name)

        last_row = ws.UsedRange.Rows.Count

        persons = list()

        # со второй, потому что в первой заголовки
        for i in range(2, last_row + 1):

            person = dict()

            index_str = str(i)

            # при таком преобразовании вместо None становится 'None', поэтому ствим условие
            if is_not_none_str(str(ws.Range("A" + index_str))):
                person['second_name'] = str(ws.Range("A" + index_str))
            else:
                person['second_name'] = None

            if is_not_none_str(str(ws.Range("B" + index_str))):
                person['first_name'] = str(ws.Range("B" + index_str))
            else:
                person['first_name'] = None

            if is_not_none_str(str(ws.Range("C" + index_str))):
                person['third_name'] = str(ws.Range("C" + index_str))
            else:
                person['third_name'] = None

            if is_not_none_str(str(ws.Range("D" + index_str))):
                birth_date = str(ws.Range("D" + index_str))

                # сохраняется вот в таком формате - конвертируем, чтоб корректно вбивать на сайте
                birth_date_converted = datetime.datetime.strptime(birth_date, '%Y-%m-%d %H:%M:%S%z')

                person['birth_date'] = birth_date_converted.strftime("%d.%m.%Y")
            else:
                person['birth_date'] = None

            persons.append(person)

        # не помню зачем тру оставил
        wb_data.Close(True)

        excel.Quit()

        return persons

    def save_checked_debtors(self, debtors_json):

        # проверяем, есть ли смысл вообще корячиться-то
        if debtors_json:

            excel = win32com.client.Dispatch('Excel.Application')

            current_path = os.getcwd()

            # джоиним текущий путь и путь до файла, куда будем сохранять
            excel_file = os.path.join(current_path, self.output_file_path)
            # book = excel.Workbooks.open(r'checked_debtors.xlsx')
            wb_data = excel.Workbooks.Open(excel_file)

            # TODO проверять скок строк - если больше 1000 - делаем новый файл

            # сколько уже листов есть в файле
            sheet_count = wb_data.Worksheets.Count

            # если листов всего 1 и стоит дефолтный Лист1, то пишем прямо в него
            if sheet_count == 1 and wb_data.Worksheets(1).Name == 'Лист1':
                # имя первого листа
                first_sheet_name = f'{self.output_sheet_name}_1'
                wb_data.Worksheets(1).Name = first_sheet_name

            # print(sheet_count)
            # print(wb_data.Worksheets(1).Name)
            # print(wb_data.Worksheets(sheet_count).Name.split(f'{self.output_sheet_name}_')[1])

            last_sheet_id = max([wb_data.Worksheets(sheet_id).Name.split(f'{self.output_sheet_name}_')[1]
                                 for sheet_id in range(1, sheet_count + 1)])
            # print(last_sheet_id)

            # for sheet_id in range(1, sheet_count):
            #     max(wb_data.Worksheets(1).Name

            last_sheet_name = f'{self.output_sheet_name}_{last_sheet_id}'

            last_sheet = wb_data.Worksheets(last_sheet_name)

            last_row = last_sheet.UsedRange.Rows.Count

            sheet_empty = False

            if last_row >= 1000:
                # лист добавляется в НАЧАЛО всех листов, поэтому работаем с ним
                wb_data.Worksheets.Add()
                # имя свежесозданного листа
                sheet_name = f'{self.output_sheet_name}_{sheet_count + 1}'
                # меняем имя новосозданного листа
                wb_data.Worksheets(1).Name = sheet_name
                # и рабочий лист у нас будет как раз он)
                ws = wb_data.Worksheets(sheet_name)
                # меняем счетчик последней строки
                last_row = 1
                # флаг что лист пустой
                sheet_empty = True
            else:
                # если нет, то рабочий лист - лист с последним айдишником, который проверяли
                ws = last_sheet
                # сохраняю, чтоб в ретерне упомянуть
                sheet_name = last_sheet_name

            # debtors_json_path = f'{get_project_root_path()}/task_1/parsed_debtors.json'
            #
            # with open(debtors_json_path, encoding='utf-8') as parsed_debtors_json_file:
            #     parsed_debtors_json_data = parsed_debtors_json_file.read()
            #
            # if parsed_debtors_json_data:
            #     debtors_json = json.loads(parsed_debtors_json_data)
            # else:
            #     debtors_json = None

            if sheet_empty:
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

            # т.к. начинаем с 1, увеличиваем индекс, до которого идем по в цикле на 1, чтоб не потерять ласт
            for debtor_id in range(1, len(debtors_json) + 1):

                # высчитывем, в какую строку будем записывать
                row = last_row + debtor_id

                # если лист пустой, первую строку будут занимать заголовки, поэтому еще +1
                if sheet_empty:
                    row += 1

                debtor = debtors_json[f'debtor_{debtor_id}']

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

            return {'saved': True,
                    'desc': f'Saved {len(debtors_json)} debtors into {self.output_file_path}, sheet {sheet_name}'}
        else:
            return {'saved': False, 'desc': 'Nothing to save'}


def is_not_none_str(cell_value_str):
    if cell_value_str == 'None':
        return None
    else:
        return cell_value_str


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

    # fssp = FSSP_excel()

    # fssp.ffsp_get_debtors_to_check()

    # fssp.save_checked_debtors({'temp': None})

    do = datetime.datetime.strptime('1991-07-14 00:00:00+00:00', '%Y-%m-%d %H:%M:%S%z')
    print(do)
    print(do.strftime("%d.%m.%Y"))
