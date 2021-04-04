import requests
from lxml import html
import openpyxl
from bs4 import BeautifulSoup
import pandas as pd


def read_file(filename):
    with open(filename) as input_file:
        text = input_file.read()
    return text


# url = "https://auto.ru/moskva/cars/vendor-foreign/all/do-300000/?sort=cr_date-desc&transmission=MECHANICAL&year_from=2005"
#
# r = requests.get(url)
# r.encoding = 'utf-8'
# with open('test.html', 'w', encoding='utf-8') as output_file:
#     output_file.write(r.text)


def parse_data_bs(filename):
    results = []
    text = read_file(filename)

    soup = BeautifulSoup(text, features="lxml")

    car_desc = soup.find_all('div', {'class': 'ListingItem-module__description'})
    for item in car_desc:
        car_name = \
            item.find('h3', {'class': 'ListingItemTitle-module__container ListingItem-module__title'}
                      ).find('a').text
        car_link = \
            item.find('h3', {'class': 'ListingItemTitle-module__container ListingItem-module__title'}
                      ).find('a').get('href')
        car_tech_summary = \
            item.find_all('div', {'class': 'ListingItemTechSummaryDesktop__cell'})
        # строка - 1.8 л. / 125 л.с. / Бензин
        car_engine_volume = car_tech_summary[0].text.split('/')[0]  # 1.8 л.
        car_engine_horse_volume = car_tech_summary[0].text.split('/')[1]  # 125 л.с.
        car_engine_type = car_tech_summary[0].text.split('/' )[2]  # Бензин

        car_transmition_type = car_tech_summary[1].text # Механика
        car_type = car_tech_summary[2].text  # седан, хэтчбек, универсал, купе, внедорожник,
        # минивэн, пикап, кабриолет, фургон
        car_wheel_drive = car_tech_summary[3].text  # передний, задний, полный
        car_color = car_tech_summary[4].text  # цвет

        car_price = item.find('div', {'class': 'ListingItemPrice-module__content'}).text[:-2]
        car_year = item.find('div', {'class': 'ListingItem-module__year'}).text
        car_miliage = item.find('div', {'class': 'ListingItem-module__kmAge'}).text[:-3]

        results.append({
            'car_name': car_name,
            'car_link': car_link,
            'car_engine_volume': car_engine_volume,
            'car_engine_horse_volume': car_engine_horse_volume,
            'car_engine_type': car_engine_type,
            'car_transmition_type': car_transmition_type,
            'car_type': car_type,
            'car_wheel_drive': car_wheel_drive,
            'car_color': car_color,
            'car_price, Р': car_price,
            'car_year': car_year,
            'car_miliage, км': car_miliage
    })
    return results


car_s = parse_data_bs('test.html')
car_df = pd.DataFrame(car_s)
car_df.to_excel('output.xlsx')
print(car_df)
