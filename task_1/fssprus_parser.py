# TODO сделать парсер через селениум, через обычные реквесты форбидден 403
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
from selenium.webdriver.firefox.options import Options as FirefoxOptions
from webdriver_manager.chrome import ChromeDriverManager
from webdriver_manager.firefox import GeckoDriverManager
import requests

from utility.paths import get_project_root_path
from utility.proxy.proxy import parse, random_proxy
from utility.user_agent import save_user_agent, get_user_agent
from excel import excel_pywin32
from captcha.captcha import solve_captcha, get_captcha_img, save_captcha


class ParserFSSP:
    site_url = 'https://fssp.gov.ru/iss/ip'
    user_agent = None
    proxy_str = None

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
                                           executable_path=GeckoDriverManager(cache_valid_range=14).install()   ,
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

    def check_person(self, second_name, first_name, third_name, birth_date):

        self.driver.get(self.site_url)

        territory_chooser = self.driver.find_element_by_id('region_id_chosen')
        territory_chooser.click()

        territory_li = territory_chooser.find_element(
            # владимирская обл
            # By.CSS_SELECTOR, '[data-option-array-index="7"]')
            # московская область
            By.CSS_SELECTOR, '[data-option-array-index="30"]')
        territory_li.click()

        time.sleep(1)

        first_name_input = self.driver.find_element_by_id('input01')
        first_name_input.send_keys(second_name)

        second_name_input = self.driver.find_element_by_id('input02')
        second_name_input.send_keys(first_name)

        # TODO вставить отчество и ДР

        find_btn = self.driver.find_element_by_id('btn-sbm')
        # имитируем нажатие клавиши ENTER
        find_btn.send_keys('\ue007')

        time.sleep(30)

        # создаем объект пройденной капчи, чтобы после выхода из цикла обозначить ее как решенную УСПЕШНО
        passed_captcha_json = None

        # проверяем, всплыло ли окно с капчей
        while self.captcha_exist():
            # не решили предыдущую капчу, ничего не меняем в json, сохраняем как есть на будущее
            save_captcha(passed_captcha_json)
            passed_captcha_json = self.overcome_captcha()

        # решили капчу успешно
        passed_captcha_json['success'] = True
        save_captcha(passed_captcha_json)

        # будем парсить html супом
        soup = BeautifulSoup(self.driver.page_source, "html.parser")
        # сразу находим тело таблицы с результатом запроса
        table = soup.find("tbody")

        # собираем с таблицы инфу о всех должниках
        tr_list = table.find_all('tr')
        # удаляем заголовки, не понадобятся, один раз их увидеть достаточно
        del tr_list[0]
        print(tr_list)



        # TODO сделать переключатель страниц внизу
        self.wait_and_close_driver()

    def captcha_exist(self):
        try:
            self.driver.find_element_by_id('captcha-popup')
            return True
        except NoSuchElementException:
            return False
        except Exception as ex:
            print(ex)
            return False

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

        print('send keys')
        captcha_text_input = self.driver.find_element_by_id('captcha-popup-code')

        # вбиваем текст капчи посимвольно, чтоб не палиться, будто мы человек
        for symb in captcha_text:
            captcha_text_input.send_keys(symb)
            time.sleep(1)

        # имитируем нажатие клавиши ENTER
        print('send enter')
        captcha_text_input.send_keys('\ue007')

        time.sleep(20)

        return captcha_json


# TODO пока сюда, потом в другой файл можно засунуть
def parse_table(tr_list):
    pass


def parse_tr(tr_elem):
    # преобразуем полученную строку в суповскую (я для дебага это делаю, мб убрать)
    tr_elem = BeautifulSoup(tr_elem, 'html.parser')

    # находим все ячейки сразу, будем по ним ходить
    td_list = tr_elem.find_all('td')

    # сюда будет заносить результат
    debtor_dict = dict()

    debtor_info = td_list[0]
    # находим все разделители <br/>, т.к. инфа из первой ячейки разделена именно ими, от них будем вести навигацию
    br_list = debtor_info.findAll('br')
    # TODO мб сюда вставить dict для удобства, но запарюсь пока что
    debtor_dict['name'] = br_list[0].previous.strip()
    debtor_dict['birth_date'] = br_list[0].next.strip()
    # заменяем двойные пробелы одинарными
    debtor_dict['place'] = br_list[1].next.replace('  ', ' ')

    # исполнительное производство
    enforcement_proceedings = td_list[1]
    # TODO тут еще в блок try except засунуть над будет
    br = enforcement_proceedings.find('br')
    # проверяем, есть ли там графа СД (иногда бывает)
    if br:
        # убираем сразу лишние пробелы
        first_ep = br.previous.strip()
        second_ep = br.next

        enforcement_proceedings = [first_ep, second_ep]
    else:
        enforcement_proceedings = [enforcement_proceedings.text]

    # TODO мб сюда вставить dict для удобства, но запарюсь пока что
    debtor_dict['enforcement_proceedings'] = enforcement_proceedings

    executive_document_details = td_list[2]
    # TODO мб сюда вставить dict для удобства, но запарюсь пока что
    print(executive_document_details)

    ep_end = td_list[3]
    # причина - статья, часть, пункт основания (ст. %, ч. %, п. %)
    ep_end_res = {'reason': None, 'date': None}

    # 4 пункт пропускаем - там конпка 'оплатить' на сайте
    for_what_how_many = td_list[5]
    fc = {'subject': None, 'amount': None}

    department_of_bailiffs = td_list[6]
    db = {'name': None, 'address': None}

    bailiff_telephone = td_list[7]
    bailiff = {'name': None, 'phone': None}

    print(debtor_dict)


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
    # parser = ParserFSSP()

    # persons = excel_pywin32()

    # parser.check_person(persons[0].get('first_name'), persons[0].get('second_name'),
    #                     persons[0].get('third_name'), persons[0].get('birth_date'))

    # parser.check_person('Иванов', 'Илья', 'Владимирович', '03.05.1981')

    tr = '''<tr class="">
<td class="first">ИВАНОВ ИЛЬЯ ВЛАДИМИРОВИЧ <br/>03.05.1981 <br/>142100,  РОССИЯ,  МОСКОВСКАЯ ОБЛ.,  Г. ПОДОЛЬСК</td>
<td class="">11057/21/50032-ИП от 01.02.2021 <br/>153243/20/50032-СД</td>
<td class="">Судебный приказ от 28.04.2018 № 2-565/2018<br/>СУДЕБНЫЙ УЧАСТОК № 190 МИРОВОГО СУДЬИ ПОДОЛЬСКОГО СУДЕБНОГО РАЙОНА МОСКОВСКОЙ ОБЛАСТИ</td>
<td class="">15.04.2021<br/>ст. 46<br/>ч. 1<br/>п. 3</td>
<td class=""><script type="text/javascript">window["_ipServices"] = {"receipt":{"title":"Квитанция","hide_title":true,"banner":"form.svg","subtitle":"<br>Квитанция","url":"https://is.fssp.gov.ru/get_receipt/?receipt="},"epgu":{"title":"Оплата через ЕПГУ","hide_title":true,"url":"https://is.fssp.gov.ru/pay/?service=epgu&pay=","banner":"pay_gos.svg","subtitle":"<br>Оплата любыми картами"}};</script></td>
<td class="">Иные взыскания имущественного характера в пользу физических и юридических лиц<br/></td>
<td class="">Подольский РОСП ГУФССП России по Московской области<br/>142100, Россия, Московская  обл., , г. Подольск, , ул. Курская, д. 6, , </td>
<td class="">ЧИСТОБАЕВА С. Х.<br/><b></b></td>
</tr>'''

    parse_tr(tr)
