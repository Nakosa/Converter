"""Xlsx converter."""
import os

import pandas as pd
from StyleFrame import StyleFrame

from pprint import pprint

PERCENTAGE_GETT = 0.177
PARK_PERCENTAGE_GETT = 0.045

COLLS_NAME = {
    'name': 'Имя',
    'plan': 'Выручка всего',
    'cash': 'Выплачено клиентом',
    'driver_salary': 'Выплата водителю',
    'parking_salary': 'Комиссия парка',
    'drives': 'Выездов',
    'tips': 'Чаевые',
    'toll_roads': 'Платные дороги',
    'cancels': 'Отменённые заказы',
}

def convert_gett(filename):
    #все данные из экселя
    df = pd.read_excel(filename)
    #уникальные индексы водителей
    drivers = set(df['ID водителя'])

    out_data = []

    driver_sum = {
        'name': 'ИТОГО',
        'plan': 0,
        'cash': 0,
        'driver_salary': 0,
        'parking_salary': 0,
        'drives': 0,
        'tips': 0,
        'toll_roads': 0,
        'cancels': 0,
    }

    for driver in drivers:
        #данные по водителю
        driver_data = df[df['ID водителя'] == driver]
        #отменённые заказы
        cancels = driver_data['Тип оплаты'].str.contains('Отмененн')
        #состоявшиеся заказы
        driver_data_good = driver_data[~cancels]
        #не состоявшиеся заказы
        driver_data_cancels = driver_data[cancels]

        data_good_count = driver_data_good.shape[0]                                             #XX
        #если прошедших заказов нет (driver_data_good.shape[0] равно 0)
        if not data_good_count:
            continue

        #ФИО
        name = driver_data['Имя водителя'].values[0]
        #план по состоявшимся заказам
        to_drv_total = driver_data['Тариф для водителя всего'].values.sum()                     #I
        #план по состоявшимся заказам без чаевых
        to_drv_without_tips = driver_data_good['Тариф для водителя без чаевых'].values.sum()    #O
        #наличка от клиента
        from_psgr = driver_data_good['Получено от клиента'].values.sum()                        #H
        #чаевые от клиента
        tips = driver_data_good['Чаевые для водителя'].values.sum()                             #K
        #платные дороги
        toll_roads_array = driver_data_good['Стоимость парковки / платной дороги'].values
        toll_roads = sum([x for x in toll_roads_array if x >= 0])                               #U
        #отменённые
        cancels_money = driver_data_cancels['Тариф для водителя всего'].values.sum()

        #комиссия парка
        to_parking = (to_drv_total + toll_roads) * PARK_PERCENTAGE_GETT                         #PP
        #комиссия gett
        to_gett = (to_drv_without_tips - toll_roads) * PERCENTAGE_GETT                          #PG
        #выдать водителю
        to_drv_perc = to_drv_total - from_psgr - to_parking - to_gett                           #SS

        driver_calc = {}
        driver_calc['name'] = ' '.join(name.split()[:2])
        driver_calc['plan'] = round(to_drv_total, 2)
        driver_calc['cash'] = round(from_psgr, 2)
        driver_calc['driver_salary'] = round(to_drv_perc, 2)
        driver_calc['parking_salary'] = round(to_parking, 2)
        driver_calc['drives'] = driver_data.shape[0]
        driver_calc['tips'] = round(tips, 2)
        driver_calc['toll_roads'] = round(toll_roads, 2)
        driver_calc['cancels'] = round(cancels_money, 2)
        out_data.append(driver_calc)

        driver_sum['plan'] += driver_calc['plan']
        driver_sum['cash'] += driver_calc['cash']
        driver_sum['driver_salary'] += driver_calc['driver_salary']
        driver_sum['parking_salary'] += driver_calc['parking_salary']
        driver_sum['drives'] += driver_calc['drives']
        driver_sum['tips'] += driver_calc['tips']
        driver_sum['toll_roads'] += driver_calc['toll_roads']
        driver_sum['cancels'] += driver_calc['cancels']

    out_data.append(driver_sum)

    driver_out_data = {
        COLLS_NAME['name']: [d['name'] for d in out_data],
        COLLS_NAME['plan']: [d['plan'] for d in out_data],
        COLLS_NAME['cash']: [d['cash'] for d in out_data],
        COLLS_NAME['driver_salary']: [d['driver_salary'] for d in out_data],
        COLLS_NAME['parking_salary']: [d['parking_salary'] for d in out_data],
        COLLS_NAME['drives']: [d['drives'] for d in out_data],
        COLLS_NAME['tips']: [d['tips'] for d in out_data],
        COLLS_NAME['toll_roads']: [d['toll_roads'] for d in out_data],
        COLLS_NAME['cancels']: [d['cancels'] for d in out_data],
    }

    """
    for key in driver_out_data.keys():
        if key != 'name':
            driver_out_data[key].append(round(sum(driver_out_data[key]), 2))
        else:
            driver_out_data[key].append(None)
    """

    df_out = pd.DataFrame(driver_out_data)
    sf_out = StyleFrame(df_out)
    sf_out.set_column_width(columns=sf_out.columns, width=20)

    out_name, ext = os.path.splitext(filename)
    excel_writer = StyleFrame.ExcelWriter('%s-out%s' % (out_name, ext))
    sf_out.to_excel(excel_writer=excel_writer, header=True, columns=[
        COLLS_NAME['name'], 
        COLLS_NAME['plan'],
        COLLS_NAME['cash'], 
        COLLS_NAME['driver_salary'], 
        COLLS_NAME['parking_salary'], 
        COLLS_NAME['drives'], 
        COLLS_NAME['tips'], 
        COLLS_NAME['toll_roads'],
        COLLS_NAME['cancels']
        ])
    try:
        excel_writer.save()
    except Exception as e:
        print('Permission denied for output file')
        print('Error message:\r\n'+str(e))
    finally:
        pass




PARK_PERCENTAGE_UBER = 0.04

def convert_uber(filename):
    df = pd.read_csv(filename)
    df = df.fillna(0)
    drivers = set(df['Driver Name'])
    out_data = []

    for driver in drivers:
        driver_calc = {}
        driver_data = df[df['Driver Name'] == driver]

        name = driver_data['Driver Name'].values[0]
        total = driver_data["Fare"].values.sum()
        to_parking = total * PARK_PERCENTAGE_UBER
        total_payment = driver_data["Total Payment"].values.sum()
        to_driver = total_payment - to_parking

        driver_calc['name'] = ' '.join(name.split()[:2])
        driver_calc['total'] = round(total, 2)
        driver_calc['total_payment'] = round(total_payment, 2)
        driver_calc['to_driver'] = round(to_driver, 2)
        driver_calc['to_parking'] = round(to_parking, 2)

        out_data.append(driver_calc)

    driver_out_data = {
        'name': [d['name'] for d in out_data],
        'total': [d['total'] for d in out_data],
        'total_payment': [d['total_payment'] for d in out_data],
        'to_driver': [d['to_driver'] for d in out_data],
        'to_parking': [d['to_parking'] for d in out_data],
    }

    for key in driver_out_data.keys():
        if key != 'name':
            driver_out_data[key].append(round(sum(driver_out_data[key]), 2))
        else:
            driver_out_data[key].append(None)

    df_out = pd.DataFrame(driver_out_data)
    sf_out = StyleFrame(df_out)
    sf_out.set_column_width(columns=sf_out.columns, width=30)

    out_name, _ = os.path.splitext(filename)
    excel_writer = StyleFrame.ExcelWriter('%s-out.xlsx' % (out_name))
    sf_out.to_excel(excel_writer=excel_writer, header=True, columns=[
        'name', 'total', 'total_payment', 'to_driver', 'to_parking'])
    excel_writer.save()








def debug(str):
    print()
    print(str)
    print()
