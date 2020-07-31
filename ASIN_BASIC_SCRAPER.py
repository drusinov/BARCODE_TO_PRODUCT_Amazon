import time
import re
import os
import wmi
import win32serviceutil

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from bs4 import BeautifulSoup
from humanfriendly import format_timespan


def browser_refresh():
    options = Options()
    options.headless = True
    drive = webdriver.Chrome(executable_path="chromedriver.exe", options=options)
    drive.implicitly_wait(1)

    drive.set_window_size(1080, 1080)
    drive.set_window_position(1085, 0, windowHandle='current')
    drive.set_page_load_timeout(15)
    return drive


def service_shutdown():
    c = wmi.WMI()
    for process in c.Win32_Process():
        if 'chrome' in str(process.Name).lower():
            # print(process.ProcessId, process.Name)
            os.system(f"TASKKILL /F /IM {process.Name} /T")


def write_output(id_, asin_, sales_rank_):
    print('ID: ' + id_)
    print('ASIN: ' + asin_)
    print('SALES_RANK: ' + sales_rank_)
    print()

    string_out = f'{id_},{asin_},{sales_rank_}\n'

    with open('asin_output.csv', 'a') as a:
        a.write(string_out)


def get_sales_rank(link):
    rank = '0'
    try:
        page = f'https://www.amazon.co.uk/dp/{link}'
        driver.get(page)

    except Exception as ex:
        print('Page TIMED OUT!')
        raise TimeoutError

    item_source = driver.page_source
    item_soup = BeautifulSoup(item_source, 'lxml')

    soup_text = item_soup.text
    try:
        rank = re.findall(r'Amazon Bestsellers Rank\s+(\d+,*\.*\d*)\sin\s\w+ \(See Top 100 in \w+\)', soup_text)[0]
        rank = rank.replace(',', '')
    except:
        rank = '9999999999'

    return rank


header_out = 'ID,ASIN,SALES_RANK\n'

delete_output = input('Would you like to delete the output file? (Y/N)... ')
print()
if delete_output.lower() == 'y' or delete_output.lower() == 'yes':
    with open('asin_output.csv', 'w') as f:
        f.write(header_out)

driver = browser_refresh()
time_zero = time.time()

OK = False
while not OK:
    try:
        with open('id_input.csv', 'r') as f:
            data_in = f.read().split('\n')
            data_in = data_in[1:]

        with open('asin_output.csv', 'r') as f:
            data_out = f.read().split('\n')
            data_out = data_out[1:]

        id_out = []
        for row in data_out:
            rows = row.split(',')
            id_out.append(rows[0])

        count_lines = 1
        for id_ in data_in:
            item_id = id_
            item_asin = 'NONE'
            item_sales_rank = 'NONE'

            if item_id in id_out:
                print(f'{item_id} already in output!\n')
                count_lines += 1
                continue

            print('Row: ' + str(count_lines) + ' of ' + str(len(data_in)))

            try:
                href = f'https://www.amazon.co.uk/s?k={item_id}'
                driver.get(href)
            except Exception as ex:
                print('Page TIMED OUT!')
                raise TimeoutError

            try:
                pg_source = driver.page_source
                soup = BeautifulSoup(pg_source, 'lxml')
            except:
                print('>>>> NO PAGE SOURCE ?!? --row 98')
                count_lines += 1
                continue

            try:  # Check for NO results
                zero_found = soup.find('div', class_='a-section a-spacing-base a-spacing-top-medium')
                zero_found = zero_found.find('div', class_='a-row').text.strip()
                if f'No results for {item_id}.' in zero_found:
                    item_asin = '---NOT FOUND---'
                    item_sales_rank = '---NOT FOUND---'

                    write_output(item_id, item_asin, item_sales_rank)
                    count_lines += 1
                    continue

            except Exception as ex:
                zero_found = False

            try:  # Get the listing ASIN
                item_asin = soup.find('h2', class_='a-size-mini a-spacing-none a-color-base s-line-clamp-2')
                item_asin = item_asin.find('a', href=True)
                item_asin = item_asin['href']

                item_asin = re.findall(r'/dp/(\w+)/ref=', item_asin)[0]

            except Exception as ex:
                print('>>>> MORE THAN 1 ITEMS ?!? --row 125')
                item_asin = 'MORE THAN 1 ITEMS'
                item_sales_rank = 'MORE THAN 1 ITEMS'
                write_output(item_id, item_asin, item_sales_rank)
                count_lines += 1
                continue

            item_sales_rank = get_sales_rank(item_asin)
            write_output(item_id, item_asin, item_sales_rank)

            count_lines += 1
            time_passed = time.time() - time_zero
            print(f'Time running: {str(format_timespan(round(time_passed, 0)))}')
            print()

        OK = True

    except Exception as ex:  # TimeoutError:
        print(ex)  # optional
        service_shutdown()
        time.sleep(5)
        for i in range(10, 0, -1):
            print(f'Countdown from 10 to restart:... {i}')
            time.sleep(1)
        driver = browser_refresh()

driver.quit()
