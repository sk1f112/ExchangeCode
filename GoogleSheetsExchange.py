from oauth2client.service_account import ServiceAccountCredentials
import gspread
import requests
from datetime import datetime

spreadsheet_key = '1-pxeNgh0UU0eAAuIW6ytWvacX8DlAJP8Q4vaOidarb4'
worksheet_name = 'ExchangeTable'
password = 1111

authorization = int(input('Введіть пароль для подальшого виконання запиту: '))
if authorization == password:
    default_date = datetime.now().strftime('%Y-%m-%d')
    start_date = input(f"Введіть початкову дату, з якої починати визначення курсу валюти (у форматі yyyy-mm-dd): ") or default_date
    end_date = input(f"Введіть кінцеву дату, для визначення курсу валюти (у форматі yyyy-mm-dd): ") or default_date
    valcode = "usd"
    sort = "exchangedate"
    order = "asc"
    format = "json"


    url = f"https://bank.gov.ua/NBU_Exchange/exchange_site?start={start_date.replace('-', '')}&end={end_date.replace('-', '')}&valcode={valcode}&sort={sort}&order={order}&{format}"

    try:
        response = requests.get(url)

        if response.status_code == 200:
            data = response.json()
            scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
            credentials = ServiceAccountCredentials.from_json_keyfile_name('drive.json', scope)
            gc = gspread.authorize(credentials)

            worksheet = gc.open_by_key(spreadsheet_key).worksheet(worksheet_name)

            headers = ["date", "type_value", "cash"]

            worksheet.clear()

            worksheet.insert_row(headers, 1)

            for index, item in enumerate(data):
                worksheet.insert_row([item['exchangedate'], item['cc'], item['rate_per_unit']], index + 2)
            print("Дані оновлені у Google Sheets.")
        else:
            print(f"Помилка запиту до API. Статус-код: {response.status_code}")
    except Exception as e:
        print(f"Помилка: {str(e)}")
else:
    print('Пароль введений не правильно!')