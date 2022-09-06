#!/usr/bin/env python3
import json
import sys
import os
import logging
import logging.handlers
import xmltodict
import xlsxwriter
import requests

ZABBIX_API_URL = "http://127.0.0.1/zabbix/api_jsonrpc.php"
UNAME = "Admin"
PWORD = "admin"

spisok_1=[]

itemids_a=(input('Введите itemids: ').split())
type_a=(input('Введите тип items: '))
time_from=(input('время начала: '))
time_till=(input('время окончания: '))

max_length_A = 0
max_length_B = 0
max_length_C = 0
max_length_D = 0
max_length_E = 0
max_length_F = 0
max_length_G = 0
max_length_H = 0
max_length_I = 0
max_length_J = 0

xlsx_dir = "excel"
excelfile = '%s/items.xlsx' % xlsx_dir
workbook = xlsxwriter.Workbook(excelfile)                           # Создает новый excel файл с именем items.xlsx 
bold = workbook.add_format(
        {
            'bold': True
            }
        )
item_header = workbook.add_format(
        {
            'bg_color': '#CCCCCC',
            'bold': True,
            'border': 1
            }
        )
item_text = workbook.add_format(
        {
            'border': 1
            }
        )

index = 3

r = requests.post(ZABBIX_API_URL,
                  json={
                        "jsonrpc": "2.0",
                        "method": "user.login",
                        "params": {
                        "user": UNAME,
                        "password": PWORD},
                        "id": 1
                  })

#print(json.dumps(r.json(), indent=4, sort_keys=True))


AUTHTOKEN = r.json()["result"]

try:
    worksheet = workbook.add_worksheet("testworksheeaaat")                   # Создает файл
except Exception as ex:
    print("Error while adding worksheet")

worksheet.write('A1', 'Название метрики', item_header)
worksheet.write('B1', 'Сервер', item_header)
worksheet.write('C1', 'ID метрики', item_header)
worksheet.write('D1', 'Количество значений', item_header)
worksheet.write('E1', 'Сумма элементов в списке', item_header)
worksheet.write('F1', 'Максимальное значение', item_header)
worksheet.write('G1', 'Минимальное значение', item_header)
worksheet.write('H1', 'Среднее значение', item_header)
worksheet.write('I1', 'Время получения первого значения', item_header)
worksheet.write('J1', 'Время получения последнего значения', item_header)
worksheet.write('I2', time_till, bold)
worksheet.write('J2', time_from, bold)

max_length_A = max(max_length_A, len('Название метрики'))
max_length_B = max(max_length_B, len('Сервер'))
max_length_C = max(max_length_C, len('ID метрики'))
max_length_D = max(max_length_D, len('Количество значений'))
max_length_E = max(max_length_E, len('Сумма элементов в списке'))
max_length_F = max(max_length_F, len('Максимальное значение'))
max_length_G = max(max_length_G, len('Минимальное значение'))
max_length_H = max(max_length_H, len('Среднее значение'))
max_length_I = max(max_length_I, len('Время получения первого значения'))
max_length_J = max(max_length_H, len('Время получения последнего значения'))






for name_item_ids in itemids_a:
    r = requests.post(ZABBIX_API_URL,
                json={
                    "jsonrpc": "2.0",
                    "method": "item.get",
                    "params": {
                        "output":"extend",
                        "itemids": name_item_ids,
#                        "filter":{"name":"Outgoing network traffic on ens33"}
                    },
                    "id": 2,
                    "auth": AUTHTOKEN
                })
    data = json.dumps(r.json(), indent=4, sort_keys=True)
    data_1 = json.loads(data)
    name_item_1 = ((data_1)['result'][0]['name'])
    name_item_2 = ((data_1)['result'][0]['hostid'])
#    print(name_item_1)
#    print(name_item_2)
#    print()
#    print()

    
    r = requests.post(ZABBIX_API_URL,
                json={
                    "jsonrpc": "2.0",
                    "method":"host.get",
                    "params": {
                        "output":["name"],
                        "hostids": name_item_2,
                     },
                    "id": 2,
                    "auth": AUTHTOKEN
                })
    data_2 = json.dumps(r.json(), indent=4, sort_keys=True)
    data_3 = json.loads(data_2)
    host_name_1 = ((data_3)['result'][0]['name'])
#    print(host_name_1)
#    print()
#    print()

        
    r = requests.post(ZABBIX_API_URL,
                json={
                    "jsonrpc": "2.0",
                    "method": "trend.get",
                    "params": {
                        "output": "extend",
#                        "history": type_a,
                        "itemids": name_item_ids,
                        "time_from": time_from,
                        "time_till": time_till,
                        "sortfield": "clock",
                    },
                    "id": 2,
                    "auth": AUTHTOKEN
                })
    data_4 = json.dumps(r.json(), indent=4, sort_keys=True)
    data_5 = json.loads(data_4)
#    print(data_5)
#    print()
#    print()

    for item in data_5['result']:

                spisok_1.append(float(item['value_avg']))
    print('id метрики', name_item_ids)
    print('количество элементов в списке:', (len(spisok_1)))
    print('сумма элементов в списке:',(sum(spisok_1)))
    print('максимальное значение:',(max(spisok_1)))
    print('минимальное значение:',(min(spisok_1)))
    print('среднее значение:', ((sum(spisok_1))/(len(spisok_1))))
    print(name_item_1)
    print(host_name_1)
    print('Метрика', name_item_ids, 'Done')
    print()
    print()

    b = str(len(spisok_1))
    n = str(sum(spisok_1))
    m = str(max(spisok_1))
    h = str(min(spisok_1))
    j = str((sum(spisok_1))/(len(spisok_1)))

    worksheet.write('C%s' % index, name_item_ids , bold)
    worksheet.write('D%s' % index, b , bold)
    worksheet.write('E%s' % index, n , bold)
    worksheet.write('F%s' % index, m , bold)
    worksheet.write('G%s' % index, h , bold)
    worksheet.write('H%s' % index, j , bold)
    worksheet.write('A%s' % index, name_item_1 , bold)
    worksheet.write('B%s' % index, host_name_1 , bold)

    max_length_A = max(max_length_A, len(name_item_ids))
    max_length_B = max(max_length_B, len(b))
    max_length_C = max(max_length_C, len(n))
    max_length_D = max(max_length_D, len(m))
    max_length_E = max(max_length_E, len(h))
    max_length_F = max(max_length_F, len(j))

    index = index + 1 

    worksheet.set_column('A:A', max_length_A)
    worksheet.set_column('B:B', max_length_B)
    worksheet.set_column('C:C', max_length_C)
    worksheet.set_column('D:D', max_length_D)
    worksheet.set_column('E:E', max_length_E)
    worksheet.set_column('F:F', max_length_F)
    worksheet.set_column('G:G', max_length_G)
    worksheet.set_column('H:H', max_length_H)
    worksheet.set_column('I:I', max_length_G)
    worksheet.set_column('J:J', max_length_H)

    del spisok_1[:]

workbook.close()

print('Done')
