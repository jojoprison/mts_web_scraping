import datetime
import locale


def switch_month(month):

    switcher = {
        'января': 'январь',
        'февраля': 'февраль',
        'марта': 'март',
        'апреля': 'апрель',
        'мая': 'май',
        'июня': 'июнь',
        'июля': 'июль',
        'августа': 'август',
        'сентября': 'сентябрь',
        'октября': 'октябрь',
        'ноября': 'ноябрь',
        'декабря': 'декабрь'
    }

    return switcher.get(month, 'Неправильно введен месяц')


# с буквой г. в конце
def convert_reg_date(reg_date):

    reg_date_split = reg_date.split(' ')
    # удаляем слово г. в конце
    reg_date_split.pop()
    reg_month = reg_date_split[1]

    month_converted = switch_month(reg_month)

    date_converted = ''

    for idx, reg_date_part in enumerate(reg_date_split):
        # если сейчас находимся на части с месяцем
        if not idx == 1:
            date_converted = date_converted + str(reg_date_part)
        else:
            date_converted = date_converted + month_converted

        if idx != len(reg_date_split) - 1:
            date_converted += ' '

    return date_converted


def prepare_to_db(reg_date):
    # ставим русскую локаль, можно просто оставить '', считает локаль с системы
    locale.setlocale(locale.LC_ALL, 'ru_RU.UTF-8')

    temp_date = datetime.datetime.strptime(reg_date, '%d %B %Y').date()

    return temp_date
