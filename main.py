from secret_settings import TOKEN, CREDENTIALS_FILE
from tinkoff.invest import Client, Bond
import pygsheets
import datetime
import time
import requests


SPREADSHEET_URL = 'https://docs.google.com/spreadsheets/d/1B_CVHeSpNr00YwoFMeJYyrgT5sySUX5ZfbWZr_JClx4/edit'


def authorize_google_sheets(credentials_file: str, spreadsheet_url: str) -> pygsheets.Worksheet:
    """
    Авторизует в Google Sheets и открывает указанный документ
    """
    gc = pygsheets.authorize(service_file=credentials_file)
    sh = gc.open_by_url(spreadsheet_url)
    ws = sh[0]    # Берем первый лист
    ws.clear()    # Очищаем лист
    return ws


def create_spreadsheet_header(ws: pygsheets.Worksheet) -> pygsheets.Worksheet:
    """
    Создаёт заголовок таблицы
    """
    headers = [
        'Тикер', 'Название', 'Номинал', 'Цена', 'НКД', 'Комиссия', 'Дата оферты', 'Дата погашения', 'Доход, %/год', \
        'Доход после налога, %/год', 'Только для квалов', 'Доход после налога, %/год (число)'
    ]
    ws.update_values('A1', [headers])
    
    # Список ячеек для изменения
    cells = ['A1', 'B1', 'C1', 'D1', 'E1', 'F1', 'G1', 'H1', 'I1', 'J1', 'K1']

    # Задаем свойства для каждой ячейки
    for cell in cells:
        cell_object = ws.cell(cell)
        
        cell_object.set_horizontal_alignment(pygsheets.custom_types.HorizontalAlignment.CENTER)     # Выравнивание
        cell_object.set_vertical_alignment(pygsheets.custom_types.VerticalAlignment.MIDDLE)
        
        cell_object.wrap_strategy = "WRAP"      # Перенос текста
        cell_object.set_text_format('fontSize', 12).set_text_format('bold', True)
        cell_object.color = (0.9, 0.9, 0.9)

    ws.adjust_column_width(10, 10, 150)     # Устанавливаем ширину колонки J
    ws.frozen_rows = 1                      # Закрепляем строку 1

    # Название и текущая дата
    title = "Топ облигаций по реальной доходности"
    current_date = str(datetime.datetime.now().strftime('%d.%m.%Y'))
    
    # Объединяем ячейки
    ws.merge_cells('A2', 'H2')
    ws.merge_cells('I2', 'K2')

    ws.update_values('A2:H2', [[title]])
    ws.update_values('I2:K2', [[current_date]])

    cell_title = ws.cell('A2')
    cell_date = ws.cell('I2')
    
    cell_title.set_text_format('bold', True).set_text_format('fontSize', 24)
    cell_date.set_text_format('bold', True).set_text_format('fontSize', 24)
    cell_title.color = (0.6, 0.8, 1)
    cell_date.color = (0.4, 0.6, 0.8)

    cell_title.set_horizontal_alignment(pygsheets.custom_types.HorizontalAlignment.CENTER)
    cell_title.set_vertical_alignment(pygsheets.custom_types.VerticalAlignment.MIDDLE)
    cell_date.set_horizontal_alignment(pygsheets.custom_types.HorizontalAlignment.CENTER)
    cell_date.set_vertical_alignment(pygsheets.custom_types.VerticalAlignment.MIDDLE)

    return ws


def update_spreadsheet_values(data: list, ws: pygsheets.Worksheet, row_index: int) -> dict:
    """
    Обновляет значения в Google-таблице
    """
    values = [
        [
            entry['ticker'], entry['name'], entry['nominal'], entry['price'], entry['aci'], entry['fee'], \
            entry['offerdate'], entry['maturity_date'], entry['profit_per_year'], entry['profit_per_year_after_tax'], \
            entry['qual'], entry['profit_per_year_after_tax_numeric']
        ] for entry in data]
    
    ws.update_values(f'A{row_index + 3}', values=values)


def get_bond_data(client: Client, bond: Bond) -> dict:
    """
    Получает данные об облигации
    """
    # Берём только рублёвые и не бессрочные облигации
    if bond.currency != 'rub' or bond.perpetual_flag:
        return None

    # Если срок погашения истёк, но облигация ещё есть в БД, то не берём её
    if bond.maturity_date == datetime.datetime(1970, 1, 1, 0, 0, tzinfo=datetime.timezone.utc):
        return None
    
    # Номинал
    nominal = bond.nominal.units + int(str(bond.nominal.nano)[:2]) / 100    

    # Получаем последние цены облигации
    prices = client.market_data.get_last_prices(instrument_id=[bond.figi])

    # Последняя цена
    price = (prices.last_prices[0].price.units + int(str(prices.last_prices[0].price.nano)[:2]) / 100) / 100 * nominal

    # НКД
    aci = bond.aci_value.units + int(str(bond.aci_value.nano)[:2]) / 100 

    # Комиссия 0,3%   
    fee = round(((price + aci) * 0.003), 2)                                       

    # Получаем дату оферты (если есть)
    max_retries = 3
    for attempt in range(max_retries):
        try:
            response = requests.get(f'https://iss.moex.com/iss/engines/stock/markets/bonds/securities/{bond.ticker}.json?iss.meta=off', timeout=20)
            columns = response.json()["securities"]["columns"]
            data_rows = response.json()["securities"]["data"]
            for row in data_rows:
                offerdate = row[columns.index("OFFERDATE")]
            break
        except requests.exceptions.Timeout:
            print(f"Попытка {attempt + 1} завершилась неудачей из-за тайм-аута.")
        except requests.exceptions.ConnectionError as ce:
            print(f"Попытка {attempt + 1} завершилась неудачей из-за ошибки подключения: {ce}")
        except requests.exceptions.RequestException as e:
            print(f"Попытка {attempt + 1} завершилась ошибкой: {e}")
        if attempt < max_retries - 1:
            print("Повторяем попытку...")
            time.sleep(2)

    # Получаем купоны для облигации
    if offerdate:
        offerdate = datetime.datetime.strptime(offerdate, "%Y-%m-%d").replace(tzinfo=datetime.timezone.utc)
        coupons = client.instruments.get_bond_coupons(figi=bond.figi, from_=datetime.datetime.now(), to=offerdate).events
    else:
        coupons = client.instruments.get_bond_coupons(figi=bond.figi, from_=datetime.datetime.now(), to=bond.maturity_date).events

    # Считаем сумму купонов
    sum_coupons = 0
    for coupon in coupons:
        if coupon.pay_one_bond.units == 0 and coupon.pay_one_bond.nano == 0:    # Отсекаем облигации, в которых
            return None                                                         # есть неизвестные купоны
        sum_coupons += coupon.pay_one_bond.units + int(str(coupon.pay_one_bond.nano)[:2]) / 100

    # Доходность в рублях
    profit_rub = round((nominal - price - aci - fee + sum_coupons), 2)
    if profit_rub < 0 or price == 0:  # Отсекаем облигации с неправильным отображением данных в БД Т-Инвестиций
        return None

    # Доходность в процентах
    profit_per = format((profit_rub / (price + aci + fee)), '.2%')

    # Доходность в процентах годовых
    if offerdate:
        days_left = (offerdate - datetime.datetime.now(datetime.timezone.utc)).days    # Дней до оферты
        offerdate = offerdate.strftime('%d.%m.%Y')
    else:
        days_left = (bond.maturity_date - datetime.datetime.now(datetime.timezone.utc)).days    # Дней до погашения
    if days_left == 0:  # Отсекаем облигации с погашением сегодня
        return None
    profit_per_year = format((profit_rub / (price + aci + fee) / days_left * 365), '.2%')

    # Доходность в рублях после налога
    profit_rub_after_tax = round((0.87 * profit_rub), 2)

    # Доходность в процентах после налога
    profit_per_after_tax = format((0.87 * (profit_rub / (price + aci + fee))), '.2%')

    # Доходность в процентах (и числовом значении для сортировки) годовых после налога
    profit_per_year_after_tax = format((0.87 * (profit_rub / (price + aci + fee) / days_left * 365)), '.2%')
    profit_per_year_after_tax_numeric = 0.87 * (profit_rub / (price + aci + fee) / days_left * 365)

    # Доступность только квалифицированным инвесторам
    qual = 'Да' if bond.for_qual_investor_flag else 'Нет'

    # Результирующий словарь
    bond_data = {
        'ticker': bond.ticker,                                      # Тикер
        'name': bond.name,                                          # Название
        'nominal': nominal,
        'price': price,
        'aci': aci,
        'fee': fee,
        'offerdate': offerdate,
        'maturity_date': bond.maturity_date.strftime('%d.%m.%Y'),   # Дата погашения
        'profit_per_year': profit_per_year,
        'profit_per_year_after_tax': profit_per_year_after_tax,
        'profit_per_year_after_tax_numeric': profit_per_year_after_tax_numeric,
        'qual': qual
    }

    return bond_data


def main():
    # Авторизация в Google Sheets и открытие таблицы
    ws = authorize_google_sheets(CREDENTIALS_FILE, SPREADSHEET_URL)

    # Создание заголовка таблицы
    create_spreadsheet_header(ws)

    # Инициализируем клиента Tinkoff Invest API
    with Client(TOKEN) as client:

        # Получаем список облигаций
        bonds = client.instruments.bonds()

        # Инициализируем счетчик для строки в таблице
        row_index = 0

        # Обрабатываем каждую облигацию
        for bond in bonds.instruments:
            bond_data = get_bond_data(client, bond)
            if bond_data:
                update_spreadsheet_values([bond_data], ws, row_index)
                row_index += 1

        # Сортируем таблицу после заполнения
        ws.sort_range('A2', f'L{row_index + 2}', basecolumnindex=11, sortorder='DESCENDING')

        # Удаление данных из столбца L (столбец нужен был только для сортировки)
        cell_list = ws.get_col(12, include_tailing_empty=False)  # Получаем список ячеек в столбце L
        ws.update_col(12, [''] * len(cell_list))  # Обновляем значения ячеек на пустые строки


if __name__ == "__main__":
    main()