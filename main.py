from secret_settings import TOKEN, CREDENTIALS_FILE
from tinkoff.invest import Client, Bond
import pygsheets
import datetime


SPREADSHEET_URL = 'https://docs.google.com/spreadsheets/d/1B_CVHeSpNr00YwoFMeJYyrgT5sySUX5ZfbWZr_JClx4/edit'


def authorize_google_sheets(credentials_file: str, spreadsheet_url: str) -> pygsheets.Worksheet:
    """
    Авторизует в Google Sheets и открывает указанный документ
    """
    gc = pygsheets.authorize(service_file=credentials_file)
    sh = gc.open_by_url(spreadsheet_url)
    ws = sh[0]    # Берем первый лист
    ws.clear()    # Очищаем лист

    # Обновляем заголовки таблицы
    headers = ['Тикер', 'Название', 'Номинал', 'Цена', 'НКД', 'Комиссия', 'Сумма купонов', 'Дата погашения']
    ws.update_values('A1', [headers])
    
    # Делаем заголовки жирными
    for cell in ['A1', 'B1', 'C1', 'D1', 'E1', 'F1', 'G1', 'H1']:
        ws.cell(cell).set_text_format('bold', True)

    return ws


def update_spreadsheet_values(data: list, ws: pygsheets.Worksheet, row_index: int) -> dict:
    """
    Обновляет значения в Google-таблице
    """
    values = [
        [
            entry['ticker'], 
            entry['name'], 
            entry['nominal'], 
            entry['price'], 
            entry['aci'],
            entry['fee'],
            entry['sum_coupons'],
            entry['maturity_date']
        ] for entry in data]
    
    ws.update_values(f'A{row_index+2}', values=values)


def get_bond_data(client: Client, bond: Bond) -> dict:
    """
    Получает данные об облигации
    """
    # Берём только рублёвые и не бессрочные облигации, и не флоатеры
    if bond.currency != 'rub' or bond.perpetual_flag or bond.floating_coupon_flag:
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

    # Получаем купоны для облигации
    coupons = client.instruments.get_bond_coupons(figi=bond.figi, from_=datetime.datetime.now(), to=bond.maturity_date).events

    # Считаем сумму купонов
    sum_coupons = 0
    for coupon in coupons:
        sum_coupons += coupon.pay_one_bond.units + int(str(coupon.pay_one_bond.nano)[:2]) / 100

    # Результирующий словарь
    bond_data = {
        'ticker': bond.ticker,                                      # Тикер
        'name': bond.name,                                          # Название
        'nominal': nominal,
        'price': price,
        'aci': aci,
        'fee': fee,
        'sum_coupons': sum_coupons,
        'maturity_date': bond.maturity_date.strftime('%d.%m.%Y')    # Дата погашения
    }

    return bond_data


def main():
    # Авторизация в Google Sheets и открытие таблицы
    ws = authorize_google_sheets(CREDENTIALS_FILE, SPREADSHEET_URL)

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


if __name__ == "__main__":
    main()