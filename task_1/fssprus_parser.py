# TODO сделать парсер через селениум, через обычные реквесты форбидден 403
import multiprocessing
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
from captcha.captcha import solve_captcha, get_captcha_img


class ParserFSSP:
    site_url = 'https://fssp.gov.ru/iss/ip'
    user_agent = None
    proxy_str = None

    def __init__(self, proxy=None):

        driver_name_list = ['chrome', 'firefox']
        # выбираем браузер из списка
        # driver_name = random.choice(driver_name_list)
        driver_name = 'chrome'

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
            # это владимирская обл
            # By.CSS_SELECTOR, '[data-option-array-index="7"]')
            # московская область
            By.CSS_SELECTOR, '[data-option-array-index="30"]')
        territory_li.click()

        time.sleep(1)

        first_name_input = self.driver.find_element_by_id('input01')
        first_name_input.send_keys(second_name)

        second_name_input = self.driver.find_element_by_id('input02')
        second_name_input.send_keys(first_name)

        find_btn = self.driver.find_element_by_id('btn-sbm')
        # имитируем нажатие клавиши ENTER
        find_btn.send_keys('\ue007')

        time.sleep(30)

        # проверяем, всплыло ли окно с капчей
        try:
            self.driver.find_element_by_id('captcha-popup')
            captcha_exist = True
        except NoSuchElementException:
            captcha_exist = False
        except Exception as ex:
            print(ex)
            captcha_exist = False

        # если надо разгадать капчу
        # TODO сделать распознавание речи капчи
        if captcha_exist:
            # закрываем окно алертов
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

            # имитируем нажатие клавиши ENTER
            print('send enter')
            captcha_text_input.send_keys('\ue007')

            time.sleep(30)

        soup = BeautifulSoup(self.driver.page_source, "html.parser")
        table = soup.find("tbody")

        tr_list = table.find_all('tr')
        print(tr_list)
        del tr_list[0]
        print(tr_list)


def parse_massive():

    list_org = ParserFSSP()

    cursor = list_org.con.cursor()

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

    # persons = excel_pywin32()

    # parser.check_person(persons[0].get('first_name'), persons[0].get('second_name'),
    #                     persons[0].get('third_name'), persons[0].get('birth_date'))

    parser.check_person('Иванов', 'Илья', 'Владимирович', '03.05.1981')
    # res = rus_profile.parse_page('https://www.rusprofile.ru/codes/430000/3760')
    # print(res)