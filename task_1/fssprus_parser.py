import json
import multiprocessing
import random
import sys
import time

from bs4 import BeautifulSoup
from fake_useragent import UserAgent
from selenium import webdriver
from selenium.common.exceptions import TimeoutException, SessionNotCreatedException, NoSuchElementException
from selenium.webdriver.chrome.options import Options as ChromeOptions
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.firefox.options import Options as FirefoxOptions
from webdriver_manager.chrome import ChromeDriverManager
from webdriver_manager.firefox import GeckoDriverManager
import requests

from utility.paths import get_project_root_path
from utility.proxy.proxy import parse, random_proxy
from utility.user_agent import save_user_agent, get_user_agent
from captcha.captcha import solve_captcha, get_captcha_img, save_captcha
from excel import save_to_json, clear_json_file, get_debtors_json
from excel import FSSP_excel


class ParserFSSP:
    site_url = 'https://fssp.gov.ru/iss/ip'
    user_agent = None
    proxy_str = None
    # класс для работы с экселем
    fssp_excel = FSSP_excel()

    def __init__(self, proxy=None):

        driver_name_list = ['chrome', 'firefox']
        # выбираем браузер из списка
        driver_name = random.choice(driver_name_list)
        # driver_name = 'chrome'

        # user_agent
        print('get user_agent...')
        try:
            self.user_agent = UserAgent(cache=False, use_cache_server=False).random
            # сейвим юзер агента чтобы в случае превышения лимита обращений
            # к API либы webdriver_manager забирать уже записанные в json
            save_user_agent(self.user_agent)
        except ValueError:
            print('value_err')
            self.user_agent = get_user_agent()
        except Exception:
            print('er_err')
            self.user_agent = get_user_agent()

        print('user_agent: ', self.user_agent)

        # proxy
        # if not proxy:
        #     print('get proxy...')
        #     self.proxy_str = random_proxy()
        #     print('proxy: ', self.proxy_str)
        #     proxy = parse(self.proxy_str).proxy_signature()
        # else:
        #     self.proxy_str = proxy.__str__()
        #     proxy = proxy.proxy_signature()

        if driver_name == 'chrome':
            chrome_options = ChromeOptions()
            # скрывает окно браузера
            # options.add_argument(f'--headless')
            # изменяет размер окна браузера
            chrome_options.add_argument(f'--window-size=800,600')
            chrome_options.add_argument("--incognito")
            # вырубаем палево с инфой что мы webdriver
            chrome_options.add_argument('--disable-blink-features=AutomationControlled')
            # user-agent
            if self.user_agent:
                chrome_options.add_argument(f'user-agent={self.user_agent}')
            # headless mode
            # chrome_options.headless = True

            # chrome_options.add_experimental_option("mobileEmulation",
            #                                        {"deviceName": "Galaxy S5"})  # or whatever

            if proxy:
                chrome_options.add_argument(f'--proxy-server={proxy}')

            try:
                driver = webdriver.Chrome(executable_path=ChromeDriverManager(cache_valid_range=14).install(),
                                          options=chrome_options)
            except (ValueError, SessionNotCreatedException):
                driver = webdriver.Chrome(executable_path=f'{get_project_root_path()}/drivers/chromedriver.exe',
                                          options=chrome_options)
        else:
            firefox_profile = webdriver.FirefoxProfile()
            if self.user_agent:
                # меняем user-agent, можно через FirefoxOptions если ЧЕ
                firefox_profile.set_preference('general.useragent.override', self.user_agent)
            firefox_profile.set_preference('dom.file.createInChild', True)

            # proxy
            # firefox_profile.set_preference("network.proxy.type", 1)
            # firefox_profile.set_preference("network.proxy.http", proxy)
            # firefox_profile.set_preference("network.proxy.http_port", port)
            # firefox_profile.set_preference("network.proxy.ssl", proxy)
            # firefox_profile.set_preference("network.proxy.ssl_port", port)

            firefox_profile.set_preference("privacy.clearOnShutdown.offlineApps", True)
            firefox_profile.set_preference("privacy.clearOnShutdown.passwords", True)
            firefox_profile.set_preference("privacy.clearOnShutdown.siteSettings", True)
            firefox_profile.set_preference("privacy.sanitize.sanitizeOnShutdown", True)
            firefox_profile.set_preference("network.cookie.lifetimePolicy", 2)
            firefox_profile.set_preference("network.dns.disablePrefetch", True)
            firefox_profile.set_preference("network.http.sendRefererHeader", 0)
            # firefox_profile.set_preference("javascript.enabled", False)

            firefox_profile.update_preferences()

            if proxy:
                firefox_capabilities = webdriver.DesiredCapabilities.FIREFOX
                firefox_capabilities['marionette'] = True

                firefox_capabilities['proxy'] = {
                    'proxyType': 'MANUAL',
                    'httpProxy': proxy,
                    'ftpProxy': proxy,
                    'sslProxy': proxy
                }
            else:
                firefox_capabilities = None

            firefox_options = FirefoxOptions()
            # размер окна браузера
            firefox_options.add_argument('--width=800')
            firefox_options.add_argument('--height=600')
            # вырубаем палево с инфой что мы webdriver
            firefox_options.set_preference('dom.webdriver.enabled', False)
            # headless mode
            # firefox_options.headless = True

            # инициализируем firefox
            try:
                driver = webdriver.Firefox(firefox_profile=firefox_profile,
                                           capabilities=firefox_capabilities,
                                           executable_path=GeckoDriverManager(cache_valid_range=14).install(),
                                           options=firefox_options,
                                           service_log_path=f'{get_project_root_path()}'
                                                            f'/logs/geckodriver.log')
            except (ValueError, SessionNotCreatedException):
                driver = webdriver.Firefox(firefox_profile=firefox_profile,
                                           capabilities=firefox_capabilities,
                                           executable_path=f'{get_project_root_path()}/drivers/geckodriver.exe',
                                           options=firefox_options,
                                           service_log_path=f'{get_project_root_path()}'
                                                            f'/logs/geckodriver.log')

        driver.delete_all_cookies()

        self.driver = driver

    # метод для закрытия браузера
    def close_driver(self):
        self.driver.close()
        self.driver.quit()

    def wait_and_close_driver(self):
        input('Press enter if you want to stop browser right now')
        self.close_driver()
        sys.exit()

    # чтобы протестить прокси
    def get_site_title(self):
        self.driver.set_page_load_timeout(10)

        try:
            self.driver.get(self.site_url)
        except TimeoutException:
            print('GET TIMEOUT')

        title = self.driver.title

        self.close_driver()

        return title

    def check_person(self, person):

        print(person)

        if self.driver.current_url != self.site_url:
            self.driver.get(self.site_url)

        time.sleep(2)

        territory_chooser = self.driver.find_element_by_id('region_id_chosen')
        territory_chooser.click()

        territory_li = territory_chooser.find_element(
            # владимирская обл
            # By.CSS_SELECTOR, '[data-option-array-index="7"]')
            # московская область
            By.CSS_SELECTOR, '[data-option-array-index="30"]')
        territory_li.click()

        time.sleep(1)

        second_name_input = self.driver.find_element_by_id('input01')
        second_name_input.clear()

        for symbol in person['second_name']:
            second_name_input.send_keys(symbol)
            time.sleep(0.5)

        time.sleep(2)

        first_name_input = self.driver.find_element_by_id('input02')
        first_name_input.clear()

        for symbol in person['first_name']:
            first_name_input.send_keys(symbol)
            time.sleep(0.5)

        time.sleep(2)

        third_name_input = self.driver.find_element_by_id('input05')
        third_name_input.clear()

        if person['third_name']:

            for symbol in person['third_name']:
                third_name_input.send_keys(symbol)
                time.sleep(0.5)

            time.sleep(2)

        birth_date_input = self.driver.find_element_by_id('input06')
        birth_date_input.clear()

        if person['birth_date']:

            birth_date_input.send_keys(person['birth_date'])

            # с датой не работает такое
            # for symbol in person['birth_date']:
            #     birth_date_input.send_keys(symbol)
            #     time.sleep(0.5)

            time.sleep(2)

        find_btn = self.driver.find_element_by_id('btn-sbm')
        # имитируем нажатие клавиши ENTER
        # find_btn.send_keys('\ue007')
        find_btn.send_keys(Keys.RETURN)

        time.sleep(4)

        self.captcha_handler()

        # очищаем json с должниками, чтобы туда потом новых засунуть
        clear_json_file()

        parse_page(self.driver.page_source)

        print('try to find pagination')

        # находим блок с пагинацией по страницам, если есть
        if self.element_exist(By.CLASS_NAME, 'pagination'):

            print('found pagination')

            count = 2

            # на след страницы надо переходить, только если кнопка с текстом 'Следующая' есть на странице,
            # на последней она пропадает
            while self.element_exist(By.XPATH, '// a[contains( text(), "Следующая")]'):
                # через файнд его надо каждый раз заново искать, DOM обновляется
                div_pagination = self.driver.find_element_by_class_name('pagination')

                # через контейнс находим нужную линку в диве
                pagination_button_next = div_pagination.find_element_by_xpath(
                    '// a[contains( text(), "Следующая")]')

                # тыкаем, переходим на след страницу
                pagination_button_next.click()

                time.sleep(3)

                # после перехода на след страницу может вылести капча - ставим ее обработчик
                self.captcha_handler()

                print('parsing page: ', count)

                # парсим страницу с ИП
                parse_page(self.driver.page_source)

                count += 1

        # обновляем excel файл со спарсенными должниками, все сохраненне данные лежат в json файле
        save_result = self.fssp_excel.save_checked_debtors(get_debtors_json())

        print(save_result)

    def element_exist(self, search_by, search_pattern):
        try:
            self.driver.find_element(search_by, search_pattern)
            return True
        except NoSuchElementException:
            return False
        except Exception as ex:
            print(ex)
            return False

    # обработчик капчи, ставим в возможных местах возникновения
    def captcha_handler(self):

        # создаем объект пройденной капчи, чтобы после выхода из цикла обозначить ее как решенную УСПЕШНО
        passed_captcha_json = dict()
        # флаг для цикла - понимаем, решаем ли капчу повторно
        captcha_again = False

        # проверяем, всплыло ли окно с капчей
        while self.captcha_exist():

            if captcha_again:
                # не решили предыдущую капчу, обозначаем это в json, сохраняем на будущее
                passed_captcha_json['success'] = False
                # сохраняем старый, чтоб не потерять, потом будет юзать)
                save_captcha(passed_captcha_json)

            # обновляем словарь с инфой о капче, пытаясь преодолеть ее
            passed_captcha_json = self.overcome_captcha()

            # сразу ставим флаг в значение true, чтоб в случае следующего захода в цикл сохранять
            # старый словарь с инфой о капче
            captcha_again = True

        # если там вообще чет есть
        if passed_captcha_json:
            # решили капчу успешно
            passed_captcha_json['success'] = True
            # сохраняем о ней инфу
            save_captcha(passed_captcha_json)

        return True

    def captcha_exist(self):
        return self.element_exist(By.ID, 'captcha-popup')

    # преодолеваем капчу
    def overcome_captcha(self):

        # TODO сделать распознавание речи капчи
        captcha_elem = self.driver.find_element_by_id('capchaVisual')

        # забираем значения аттрибута src у картинки с капчей, чтоб скачать
        captcha_src = captcha_elem.get_attribute('src')
        # скачиваем изображение капчи
        print('download captcha img')
        get_captcha_img(captcha_src)

        # разгадываем только что скачанную капчу
        captcha_json = solve_captcha()
        print(captcha_json)
        captcha_text = captcha_json['text']

        print('captcha send keys')
        captcha_text_input = self.driver.find_element_by_id('captcha-popup-code')

        # вбиваем текст капчи посимвольно, чтоб не палиться, будто мы человек
        for symb in captcha_text:
            captcha_text_input.send_keys(symb)
            time.sleep(0.5)

        # имитируем нажатие клавиши ENTER
        print('captcha send enter')
        # captcha_text_input.send_keys('\ue007')
        captcha_text_input.send_keys(Keys.RETURN)

        time.sleep(random.choice([4, 6]))

        return captcha_json

    def run_excel_persons(self):
        person_list = self.fssp_excel.get_debtors_to_check()

        for person in person_list:
            self.check_person(person)

        # закрывем браузер чтоб избавиться от процесса
        self.wait_and_close_driver()


def parse_page(page_source):
    # будем парсить html супом
    soup = BeautifulSoup(page_source, "html.parser")

    # находим блок с результатом поиска
    div_results = soup.find('div', attrs={'class': 'results'})

    # проверяем, нашлось ли вообще что-то
    # такой контейнер с классом появляется только в случае, если нет результатов
    if not div_results.find('div', attrs={'class': 'b-search-message'}):

        # сразу находим тело таблицы с результатом запроса
        table = soup.find("tbody")

        # собираем с таблицы инфу о всех должниках
        tr_list = table.find_all('tr')
        # удаляем заголовки, не понадобятся, один раз их увидеть достаточно
        del tr_list[0]
        # print(tr_list)

        parse_table(tr_list)

        return True
    else:
        # хочу подумать тута
        time.sleep(3)
        return False


def parse_table(tr_list):
    print(f'start parsing table, table len: {len(tr_list)}')
    count = 0

    for tr_elem in tr_list:
        count += 1

        print(f'parse tr: {count}')

        ep_res_dict = parse_tr(tr_elem)

        # на всякий случай сохраняем в json только что проверенных
        save_to_json(ep_res_dict)

        # print(f'debtor_saved: {ep_res_dict}')

    return True


def parse_tr(tr_elem):
    # преобразуем полученную строку в суповскую (я для дебага это делаю, мб убрать)
    # tr_elem = BeautifulSoup(tr_elem, 'html.parser')

    # находим все ячейки сразу, будем по ним ходить
    td_list = tr_elem.find_all('td')

    # сюда будет заносить результат
    ep_res_dict = dict()

    # инфа по должнику
    debtor_info = td_list[0]
    # находим все разделители <br/>, т.к. инфа из первой ячейки разделена именно ими, от них будем вести навигацию
    br_list = debtor_info.findAll('br')

    # при существовании второго разделителя br мы точно знаем, что у должника указан адрес в первой ячейке
    if len(br_list) > 1:
        place = br_list[1].next.replace('  ', ' ')
    else:
        place = None

    ep_res_dict['debtor_info'] = {'name': br_list[0].previous.strip(),
                                  'birth_date': br_list[0].next.strip(),
                                  # заменяем двойные пробелы одинарными
                                  'place': place}

    # исполнительное производство
    enforcement_proceedings = td_list[1]

    br = enforcement_proceedings.find('br')
    # проверяем, есть ли там графа СД (иногда бывает)
    if br:
        # убираем сразу лишние пробелы
        first_ep = br.previous.strip()
        # вызываем так, потому что без обертки str() будем получать класс NavigableString
        second_ep = str(br.next)

        enforcement_proceedings = [first_ep, second_ep]
    else:
        enforcement_proceedings = [enforcement_proceedings.text]

    ep_res_dict['enforcement_proceedings'] = enforcement_proceedings

    # реквизиты исполнительного документа
    executive_document_details = td_list[2]
    # может быть 2 вида исполнительных доков, соответственно - 2 разделителя
    br_list = executive_document_details.findAll('br')

    # выносим в переменную количество видов исполнительных документов по должнику, будем юзать ниже в расчете долга
    debtor_order_count = len(br_list)

    if debtor_order_count > 1:
        ep_res_dict['document_details'] = {'order': str(br_list[0].previous),
                                           'order_2': str(br_list[1].previous),
                                           'authority': str(br_list[1].next)}
    else:
        ep_res_dict['document_details'] = {'order': str(br_list[0].previous),
                                           'authority': str(br_list[0].next)}

    # инфа о дате, причине окончания и прекращения ИП
    ep_end = td_list[3]

    # иногда бывает пустым
    if ep_end.text:
        br_list = ep_end.findAll('br')
        # причина - статья, часть, пункт основания (ст. %, ч. %, п. %)

        ep_end_reason_parts = []

        # иногда бывает пишут только статью, в данном случае исключаем поле с датой
        ep_end_reason_len = len(br_list)

        # проходимся 3 раза через разделители, т.к. там три <br/>
        for i in range(0, len(br_list)):
            ep_end_reason_parts.append(str(br_list[i].next))
        # джоиним к пробелу, там нет пробелов между частями
        reason = ' '.join(ep_end_reason_parts)

        if ep_end_reason_len > 1:
            reason_date = str(br_list[0].previous)
        else:
            reason_date = None

        ep_res_dict['ep_end'] = {'reason': reason,
                                 'date': reason_date}

    else:
        ep_res_dict['ep_end'] = {'reason': None,
                                 'date': None}

    # 4 пункт пропускаем - там конпка 'оплатить' на сайте

    # инфа о предмете исполнения и сумме непогашенной задолженности
    for_what_how_many = td_list[5]

    if debtor_order_count > 1:

        performance_subject_list = []

        br = for_what_how_many.find('br')

        for_what_how_many_first = br.previous.split(':')

        # если указанна сумма задолженности
        if len(for_what_how_many_first) > 1:
            first_amount = for_what_how_many_first[1].strip()
        else:
            first_amount = None

        performance_subject_list.append({'name': for_what_how_many_first[0], 'amount': first_amount})

        for_what_how_many_second = br.next.split(':')

        # если указанна сумма задолженности
        if len(for_what_how_many_second) > 1:
            second_amount = for_what_how_many_second[1].strip()
        else:
            second_amount = None

        performance_subject_list.append({'name': for_what_how_many_second[0], 'amount': second_amount})

        ep_res_dict['performance_subject'] = performance_subject_list
    else:

        performance_subject_info = for_what_how_many.text.split(':')

        # если указанна сумма задолженности
        if len(performance_subject_info) > 1:
            amount = performance_subject_info[1].strip()
        else:
            amount = None

        ep_res_dict['performance_subject'] = {'name': performance_subject_info[0],
                                              'amount': amount}

    # инфа об отделе судебных приставов
    department_of_bailiffs = td_list[6]
    br = department_of_bailiffs.find('br')
    # мне конечно не нравится слайс в конце, но юзаю его ввиду торопливости
    # двойные запятые заменяю на одинарные, дабы привести к нормальному виду
    ep_res_dict['department_of_bailiffs'] = {'name': str(br.previous),
                                             # тут иногда криво заканчивается строка с адресом, то запятые, то -
                                             'address': br.next.replace(', ,', ',')}

    # инфа о судебном приставе
    bailiff_telephone = td_list[7]

    # в любом случае будет, там телефон в тег <b> засовывается
    br = bailiff_telephone.find('br')

    # вытаскиваем из текст из тега болд, в нем хранится наш заветный номер телефона пристава
    phone_number = str(br.next.next).replace('\n', '')

    # проверяем на пустоту
    if phone_number:
        phone = phone_number
    else:
        phone = None

    ep_res_dict['bailiff'] = {'name': str(br.previous), 'phone': phone}

    return ep_res_dict


# TODO мультипроцессы на будущее
def parse_massive():
    list_org = ParserFSSP()

    cursor = list_org.con.cursor()

    # TODO подумать, оставлять ли (нужна БД или нет)
    cursor.execute('SELECT ogrn FROM companies LIMIT 10')
    ogrn_list = cursor.fetchall()

    ogrn_list = [ogrn[0] for ogrn in ogrn_list]
    print(ogrn_list)

    pool = multiprocessing.Pool(processes=4)

    pool_res = pool.map(list_org.update_company_by_ogrn, ogrn_list)

    list_org.con.commit()

    print(pool_res)


if __name__ == '__main__':
    parser = ParserFSSP()
    parser.run_excel_persons()
