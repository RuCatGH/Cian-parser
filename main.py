import time
import pickle

from openpyxl import Workbook
from selenium import webdriver
from selenium.webdriver.common.by import By
from multiprocessing import Pool
from multiprocessing import cpu_count
from fake_useragent import UserAgent

ua = UserAgent()

options = webdriver.ChromeOptions()

options.add_argument(f'user-agent={ua.random}')

options.add_experimental_option('excludeSwitches', ['enable-logging'])
# отключение вебдрайвера от логирования
options.add_argument('--disable-blink-features=AutomationControlled')
# оключение появления всплывающего окна
options.headless = True

def get_data(url):
    try:
        driver = webdriver.Chrome(options=options)
        driver.get(url=url)
        data_page = []

        # Get cookies
        # pickle.dump(driver.get_cookies(), open('cookies', 'wb'))
 
        # for cookie in pickle.load(open('cookies', 'rb')):
        #     driver.add_cookie(cookie)
        
        # driver.refresh()
        # Первая ссылка
        # for item in driver.find_elements(By.CLASS_NAME, '_93444fe79c--container--Povoi'): 
        #         try:
        #             name = item.find_element(By.CLASS_NAME, '_93444fe79c--jk--dIktL').text
        #         except: 
        #             name = ''
        #         adress = item.find_element(By.CLASS_NAME, '_93444fe79c--labels--L8WyJ').text
        #         data_page.append([name,adress])
        # Вторая ссылка
        # for item in driver.find_elements(By.CLASS_NAME, '_32bbee5fda--commercialWrapper--vwmUi'): 
        #     try:
        #         name = item.find_element(By.CLASS_NAME, 'c6e8ba5398--building-link--aX3XW').text
        #     except: 
        #         name = ''
        #     adress = item.find_element(By.CLASS_NAME, 'c6e8ba5398--address-path--b_cYO').text
        #     data_page.append([name,adress])        
        # 3 ссылка
        # for item in driver.find_elements(By.CLASS_NAME, 'e2ff468f85--item--H5oSR'): 
        #     try:
        #         name = item.find_element(By.CLASS_NAME, 'e2ff468f85--item-title--mtvit').text
        #     except: 
        #         name = ''
        #     adress = item.find_element(By.CLASS_NAME, 'e2ff468f85--address--tqmyO').text
        #     data_page.append([name,adress])   
            
        # for item in driver.find_elements(By.CLASS_NAME, '_93444fe79c--container--Povoi'): 
        #     name = ''
        #     adress = item.find_element(By.CLASS_NAME, '_93444fe79c--labels--L8WyJ').text
        #     data_page.append([name,adress])           
        # 4 ссылка
        for item in driver.find_elements(By.CLASS_NAME, '_93444fe79c--container--Povoi'): 
            try:
                name = item.find_element(By.CLASS_NAME, '_93444fe79c--kp--Yvf9W').text
            except: 
                name = ''
            adress = item.find_element(By.CLASS_NAME, '_93444fe79c--labels--L8WyJ').text
            price = item.find_element(By.CSS_SELECTOR, '[data-mark="MainPrice"]').text
            data_page.append([name,adress, price])           
        return data_page

    except Exception as ex:
        print(ex)

    finally:
        driver.close()
        driver.quit()

def main():
    # Таблица xlsx
    wb = Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet('Недвижимость')
    ws.append(['Название КП', 'Адрес', 'Цена'])
    # Мультипроцессорность
    # Вторая ссылка
    # urls1 = [f'https://www.cian.ru/cat.php?building_class_type%5B0%5D=1&building_class_type%5B1%5D=2&building_class_type%5B2%5D=4&currency=2&deal_type=sale&engine_version=2&m2=1&minarea=80&minprice=500000&offer_type=offices&office_type%5B0%5D=1&region=1&p={page}' for page in range(1,12)]
    # urls2 = [f'https://www.cian.ru/cat.php?building_class_type%5B0%5D=1&building_class_type%5B1%5D=2&building_class_type%5B2%5D=3&building_class_type%5B3%5D=4&currency=2&deal_type=rent&engine_version=2&m2=1&minarea=50&minprice=100000&offer_type=offices&office_type%5B0%5D=1&region=1&p={page}' for page in range(1,5)]
    # Третья ссылка
    # urls1 = [f'https://www.cian.ru/kottedzhnye-poselki/?commuteKmTo=50&priceFrom=70000000&road=2,5,9,12,13,14,16,17,19,21,24,25,27,30,38,44,49&p={page}' for page in range(1,11)]
    # urls2 = [f'https://www.cian.ru/cat.php?bbox=55.29279917336737%2C35.242767487499975%2C56.36101717051293%2C39.19784561249998&center=55.83059088623774%2C37.220306549999975&currency=2&deal_type=rent&engine_version=2&in_polygon[0]=37.4180604_56.1077716%2C37.4784852_56.0767716%2C37.5444032_56.0442215%2C37.5856019_56.0147715%2C37.6295472_55.9868715%2C37.637787_55.9496715%2C37.6213075_55.9093715%2C37.5965883_55.8721715%2C37.5608827_55.8365215%2C37.5334169_55.8024215%2C37.5086976_55.7667715%2C37.4922181_55.7295714%2C37.4729921_55.6861714%2C37.4345399_55.6551714%2C37.3768617_55.6319214%2C37.3136903_55.6195214%2C37.2450258_55.5993714%2C37.170868_55.5807714%2C37.0939637_55.5621714%2C37.0252992_55.5559714%2C36.9593812_55.5544214%2C36.8934633_55.5451214%2C36.8275453_55.5544214%2C36.8028261_55.5900714%2C36.8028261_55.6334714%2C36.8028261_55.6784214%2C36.8028261_55.7171714%2C36.8138124_55.7543714%2C36.835785_55.7900215%2C36.8632509_55.8318715%2C36.8852235_55.8690715%2C36.8934633_55.9062715%2C36.868744_55.9419215%2C36.9126893_55.9698215%2C36.9731141_55.9853215%2C37.0170594_56.0147715%2C37.0747377_56.0442215%2C37.1324159_56.0674716%2C37.1900941_56.0860716%2C37.2560121_56.1015716%2C37.3191835_56.1139716%2C37.3905946_56.1139716%2C37.4180604_56.1077716&minprice=400000&object_type[0]=1&offer_type=suburban&origin=map&polygon_name[0]=Область%20поиска&type=4&zoom=9&p={page}' for page in range(1,35)]
    # 4 ссылка
    urls = [f'https://krasnodar.cian.ru/cat.php?bbox=41.18590000002797%2C19.6389%2C67.78536109724004%2C131.52366562499998&center=56.7714732603733%2C75.58128281249998&currency=2&deal_type=sale&engine_version=2&land_status[0]=1&land_status[1]=2&land_status[2]=5&land_status[3]=7&maxsite=120&minprice=45000000&minsite=20&object_type[0]=3&offer_type=suburban&origin=map&zoom=4&p={page}' for page in range(1,51)]
    # urls = urls1 + urls2
    with Pool(cpu_count()) as p:
        data = p.map(get_data, urls)
        p.close()
        p.join()
    for rows in data:
        for row in rows:
            ws.append(row)
    wb.save('Data.xlsx')
if __name__ == '__main__':
    main()