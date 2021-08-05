import multiprocessing
import random
import time
from datetime import datetime
from pathlib import Path

import requests
from bs4 import BeautifulSoup
from requests.exceptions import ProxyError, ConnectTimeout, ConnectionError

from utility.paths import get_project_root_path

FREE_FILENAME = f'{get_project_root_path()}/utility/proxy/free_proxy_list.txt'
WORKING_FILENAME = f'{get_project_root_path()}/utility/proxy/working_proxy_list.txt'


class Proxy:
    ip = None
    port = None
    protocol = None
    country = None
    date_founded = None

    def __init__(self, ip, port, protocol, country, date_founded):
        self.ip = ip
        self.port = port
        self.protocol = protocol
        self.country = country
        self.date_founded = date_founded

    def proxy_signature(self):
        return self.ip + ':' + self.port

    def for_request(self):
        return self.protocol + '://' + self.ip + ':' + self.port

    def __str__(self):
        return self.for_request() + ' - ' + self.country \
               + ', ' + str(self.date_founded)

    def __repr__(self):
        return self.for_request() + ' - ' + self.country \
               + ', ' + str(self.date_founded)


def parse(proxy_str):
    if proxy_str.find('://') != -1:
        protocol, proxy_str = proxy_str.split('://')
        # ставим доп параметр сплиту, т.к. нам нужны разделить только первым двоеточием которое отделяет порт
        ip, proxy_str = proxy_str.split(':', 1)
        port, proxy_str = proxy_str.split(' - ')
        country, date_founded = proxy_str.split(', ')

        return Proxy(ip, port, protocol, country, date_founded)
    else:
        print('wrong proxy_str: ', proxy_str)
        return None


def free_proxy_list():
    req = requests.get('https://scrapingant.com/free-proxies/')

    soup = BeautifulSoup(req.text, 'html.parser')
    proxy_list = soup.find('table', class_='proxies-table').find_all('tr')

    # удаляем заголовки самой таблицы
    del proxy_list[0]

    result_proxy_list = []

    for proxy_tr in proxy_list:
        proxy_data = proxy_tr.find_all('td')

        if proxy_data[3].text.endswith('Unknown'):
            country = proxy_data[3].text.split(' ')[-1]
        else:
            country = ' '.join(proxy_data[3].text.split(' ')[1:])

        proxy = Proxy(
            # ip
            proxy_data[0].text,
            # port
            proxy_data[1].text,
            # protocol
            proxy_data[2].text,
            country,
            # date founded
            datetime.now()
        )

        # передаю именно строковое значение чтобы потом записать в файл writelines list[str]
        result_proxy_list.append(proxy)

    return result_proxy_list


# TODO пробнуть хранить в json
# запускать периодически чтоб обновлять лист проксей
def save_to_file():
    filename = FREE_FILENAME
    proxy_file_path = Path(filename)
    # проверяет наличие файла, если его нет - создает
    proxy_file_path.touch(exist_ok=True)

    # забираем все прокси из файла в формате стринги
    with open(proxy_file_path) as file:
        file_proxy_list_str = file.readlines()

    # преобразуем в объекты кастомного класса Proxy и до вида ip:port
    file_proxy_ip_list = []
    for proxy_str in file_proxy_list_str:
        file_proxy_ip_list.append(parse(proxy_str).proxy_signature())

    # забираем свежие фришные прокси с сайта
    fresh_proxy_list = free_proxy_list()

    # преобразуем их до вида ip:port
    fresh_proxy_ip_list = []
    for proxy in fresh_proxy_list:
        fresh_proxy_ip_list.append(proxy.proxy_signature())

    # убираем прокси из списка, если они уже есть в файле
    new_proxy_ip_list = list(set(fresh_proxy_ip_list) - set(file_proxy_ip_list))

    print('new proxies: ', new_proxy_ip_list)

    # записываем новые прокси в файл
    with open(proxy_file_path, 'a') as file:
        for new_proxy_ip in new_proxy_ip_list:
            for fresh_proxy in fresh_proxy_list:
                if fresh_proxy.proxy_signature() == new_proxy_ip:
                    file.write(str(fresh_proxy) + '\n')


# возвращает список проксей из файла, если есть
def get_proxy_list(working=True):

    # читать из файла с проверенными работающими проксями, или из файла со всеми
    if working:
        proxy_file_name = WORKING_FILENAME
    else:
        proxy_file_name = FREE_FILENAME

    proxy_file_path = Path(proxy_file_name)

    # проверяем наличие файла с бесплатными проксями
    if proxy_file_path.exists():

        with open(proxy_file_name) as file:
            proxy_list_str = file.readlines()

        return proxy_list_str
    else:
        return []


def random_proxy():
    file_proxy_ip_list = get_proxy_list()

    rand_proxy = random.choice(file_proxy_ip_list)

    return rand_proxy


def write_proxy_file(proxy, file_name):

    proxy_file_path = Path(file_name)
    # проверяет наличие файла, если его нет - создает
    proxy_file_path.touch(exist_ok=True)

    # забираем все прокси из файла в формате стринги
    with open(proxy_file_path) as file:
        file_proxy_list_str = file.readlines()

    # удаляем пустые строки из списка во избежании ошибок
    empty_space = '\n'
    while empty_space in file_proxy_list_str:
        file_proxy_list_str.remove(empty_space)

    # преобразуем в объекты кастомного класса Proxy и до вида ip:port
    file_proxy_ip_list = []
    for proxy_str in file_proxy_list_str:
        file_proxy_ip_list.append(parse(proxy_str).proxy_signature())

    # если прокси еще нет в файле
    if proxy.proxy_signature() not in file_proxy_ip_list:
        # записываем рабочую проксю в файл
        with open(proxy_file_path, 'a') as file:
            file.write(str(proxy))

        return True
    else:
        return False


def get_new_proxies():
    while True:
        save_to_file()
        print('wait 600 sec...')
        time.sleep(600)


def prepare_for_request(proxy):

    if proxy:
        proxy_req = proxy.for_request()

        proxies = {
            'http': proxy_req,
            'https': proxy_req
        }

        return proxies
    else:
        return None


def test_request(proxy_str):

    # парсим проксю из строки
    proxy = parse(proxy_str)
    print('check:', proxy)

    # если получилось спарсить проксю из строки (иногда кривятся строки)
    if proxy:

        url = 'https://www.lagado.com/tools/proxy-test'
        # получаем словарь с проксями для корректного реквеста
        proxies = prepare_for_request(proxy)

        try:
            req = requests.get(url, proxies=proxies, timeout=15)

            # проверяю айпишник с сайта
            # bs = BeautifulSoup(req.text, 'html.parser')
            #
            # ip_addr = bs.find('b', string='IP Address')
            # print('ip: ', ip_addr.parent.text.split(' ')[-1])
            #
            # forwarder = bs.find('b', string='X-Forwarded-For')
            # print('forwarder: ', forwarder.parent.parent.find_all('td')[-1].text)

            write_proxy_file(proxy, WORKING_FILENAME)

            if req.status_code == 200:
                return proxies.get('http')
            else:
                print(req.status_code)
                return None

        except (ConnectTimeout, ProxyError, ConnectionError):
            return None
    else:
        return None


def test_proxies():

    # забираем прокси из общего файла
    proxy_list = get_proxy_list(working=False)
    print(len(proxy_list))

    # downloader = multiprocessing.Process(target=get_req, args=(proxies,))
    # downloader.start()
    #
    # timeout = 10
    # time.sleep(timeout)
    #
    # downloader.terminate()

    # будет открыто максимум 2 процесса, остальные будут открыты после завершения предыдущих
    pool = multiprocessing.Pool(processes=50)

    pool_res = pool.map(test_request, proxy_list)
    clear_res = [proxy for proxy in pool_res if proxy]

    return clear_res
    # result = pool.apply_async(get_req, (proxy_list[0],))

    # try:
    #     res = result.get(timeout=10)
    #     print(res)
    # except (ProxyError, multiprocessing.context.TimeoutError) as ex:
    #     print('FUICK')
    #     print(ex)


if __name__ == '__main__':
    get_new_proxies()

    # res = test_proxies()
    # print(len(res))
    # print(res)

    # save_to_file()
