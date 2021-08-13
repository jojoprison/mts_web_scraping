import json

from base64 import decodebytes

try:
    from PIL import Image, ImageDraw, ImageOps, ImageEnhance
except ImportError:
    import Image, ImageOps, ImageEnhance, imread
# TODO чтоб работал тессеракт, нужно установить 5 (альфа) версию с https://github.com/UB-Mannheim/tesseract/wiki
import pytesseract
from utility.paths import get_project_root_path

# где лежит исполняемый файл нейронки tesseract (у меня версия 5.0.0 альфа)
pytesseract.pytesseract.tesseract_cmd = r'D:\program_files\tesseract_ocr\tesseract.exe'

# путь к джсону с проверенными капчами
CAPTCHA_LIST_JSON_PATH = f'{get_project_root_path()}/task_1/captcha/captcha_list.json'
# путь к дире, где лежат все изображения капч
CAPTCHA_IMG_DIR = f'{get_project_root_path()}/task_1/captcha/received'
# путь к файлу, куда сохраняется свежевыжатая капча
NEW_CAPTCHA_PATH = f'{get_project_root_path()}/task_1/captcha/captcha.png'


def get_captcha_img(captcha_src_base64):

    # т.к. знаем сигнатуру src, разделим данной последовательностью на 2 части
    # после запятой будет искомая последовательность байтов изображения
    captcha_base64 = captcha_src_base64.split('base64,')[1]

    # кодируем строку в последовательность байтов
    captcha_base64 = captcha_base64.encode('utf-8')

    # запихнем в temp файл, где храним новую капчу
    # пишем байты, поэтому wb, with - чтобы питон сам закрыл поток после завершения
    with open(NEW_CAPTCHA_PATH, 'wb') as f:
        # декодируем байты и пишем в файл
        f.write(decodebytes(captcha_base64))

    return True


# разгадываем капчу
# TODO !!! сделать проверку по существующим капчам из json и подставлять другие значения символов или убирать лишние
def solve_captcha():
    # открываем картинку с капчей, конвертируем в RGB формат
    img = Image.open(NEW_CAPTCHA_PATH).convert('RGB')
    # делаем контраст
    img = ImageOps.autocontrast(img)

    # json с информацией о капче, будем хранить, чтобы поюзать в случае повтора и
    # для статистики
    captcha_json = dict()

    # получаем айдишник новой капчи
    # TODO сделать бд SQLite чтобы нормально забирать айдишник, а не хранить все в json
    captcha_last_id = get_captcha_last_id()
    if captcha_last_id:
        captcha_id = get_captcha_last_id() + 1
    else:
        captcha_id = 1

    # сразу заносим в джон, чтобы потом отыскать нужное изображение в папке
    captcha_json['id'] = captcha_id

    # сохраняем в папке с капчами
    img_file_name = f'{CAPTCHA_IMG_DIR}/{captcha_id}.png'
    img.save(img_file_name)

    # разгадываем капчу с помощью тессеракта
    text = pytesseract.image_to_string(Image.open(img_file_name), lang='rus')
    # срезаем последний символ в капче (он при каждой конвертации вылазит  - '\x0c' - FF)
    # убираем все пробелы - иногда вылазят
    # TODO сделать проверку на число символов - иногда 4, иногда 6
    # TODO проверять на символы типа ) и @, иногда вылазят
    # TODO думаю не сохранять картинки, чтобы не разрастался проект - достаточно джсонов
    res_text = text.split('\n')[0].replace(' ', '')

    # сам текст капчи
    captcha_json['text'] = res_text
    # успешно ли разгадана капча
    captcha_json['success'] = None
    # список с альтернативными решениями в будущем, tesseract разгадает неправильно
    captcha_json['alternative_text'] = list()

    return captcha_json


def save_captcha(captcha_json):
    captcha_list_json = _get_captcha_list_json()

    # TODO переписать под общий код с пометкой mode='captcha', 'user_agent', добавить дату
    if captcha_list_json:
        captcha_list_json[f'captcha_{captcha_json["id"]}'] = captcha_json
    else:
        captcha_list_json = dict()
        captcha_list_json[f'captcha_1'] = captcha_json

    captcha_list_json = json.dumps(captcha_list_json, ensure_ascii=False, indent=4)

    with open(CAPTCHA_LIST_JSON_PATH, 'w', encoding='utf-8') as json_file:
        json_file.write(captcha_list_json)

    return f'captcha_saved: {captcha_json}'


# TODO чтобы джсона с капчами доставать по тексту уже попадавшуюся - проверять наличие
def get_captcha(captcha_text):
    captcha_list_json = _get_captcha_list_json()

    if not captcha_list_json:
        return None
    else:
        # TODO непонятно вообще куда ретернит список значений, доделать
        captcha_list_json.values()


def get_captcha_last_id():
    captcha_list_json = _get_captcha_list_json()

    if not captcha_list_json:
        return None
    else:
        captcha_id_list = list(captcha_list_json.keys())

    # обрезаем первое слово, там точно будет 'captcha'

    return int(captcha_id_list[-1].split('captcha_')[-1])


def _get_captcha_list_json():
    with open(CAPTCHA_LIST_JSON_PATH, encoding='utf-8') as captcha_list_json_file:
        captcha_list_json_data = captcha_list_json_file.read()

    if captcha_list_json_data:
        return json.loads(captcha_list_json_data)
    else:
        return None


if __name__ == '__main__':
    # captcha_img_path = 'captcha_ex.jpg'
    #
    # print('-- Resolving')
    # captcha_text = solve_captcha(captcha_img_path)
    # print(f'-- Result: {captcha_text}')

    # url = 'data:image/jpeg;base64,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'
    # get_captcha_img(url)

    print(solve_captcha())

    # print(get_captcha_last_id())
