import xlsxwriter
from datetime import date, datetime
from math import floor
import csv
from db import delete_all_data, store_all_data, get_alerts_by_sla, get_alerts_by_state, get_alerts_by_priority, \
    get_alerts_by_type, get_alerts_by_severity, get_alerts_by_owner

date_s = date.today()
date_s = floor(date_s.day / 10) * 10
if date_s < 10:
    week = 'Week 3'
elif 10 <= date_s < 20:
    week = 'Week 1'
else:
    week = 'Week 2'
name_dest = 'Monthly KPI Data ' + week + '.xlsx'
csv_reader, file_flag = 0, 0
while file_flag == 0:
    try:
        # file = input('Enter the path or name of you file: ')
        file = 'test2.csv'
        data_file = open(file=file)
        csv_reader = csv.reader(data_file, delimiter=',')
        file_flag = 1
    except FileNotFoundError:
        print('File Not Found.')
    except OSError:
        print('Enter a Valid Path or Name.')

data = []
flag = 0
header = []
for key in csv_reader:
    temp = {}
    if flag == 0:
        for key1 in key:
            header.append(key1.lower())
        flag = 1
        continue
    for key1 in range(len(key)):
        if key[key1].__contains__('-') and key[key1].__contains__(':'):
            temp2 = key[key1].split('-')
            try:
                temp1 = datetime(day=int(temp2[0]), month=int(temp2[1]), year=int('20' + temp2[2].split()[0]),
                                 hour=int(temp2[2].split(':')[0].split()[1]), minute=int(temp2[2].split(':')[1]))
                temp.update({header[key1]: temp1})
            except:
                temp.update({header[key1]: key[key1]})
        else:
            temp.update({header[key1]: key[key1]})
    data.append(temp)

ignore = ['number', 'customerid', 'firstresponsetimeworkingtime', 'firstresponsetime', 'impact', 'review required',
          'decision result', 'decision date', 'due date']
for key in data:
    for key1 in ignore:
        key.__delitem__(key1)

service = {
    'atm': 'Gold',
    'gloghana': 'Gold',
    'glonigeria': 'Gold',
    'gosoft': 'Gold',
    'mtnaf': 'Platinum',
    'mtnbenin': 'Platinum',
    'mtnci': 'Platinum',
    'mtncongob': 'Platinum',
    'mtngc': 'Gold',
    'mtnghana': 'Platinum',
    'mtnliberia': 'Silver',
    'mtnsouthsudan': 'Gold',
    'mtnsyria': 'Platinum',
    'mtnye': 'Platinum',
    'mtnzambia': 'Platinum',
    'newcobahamas': 'Platinum',
    'other': 'Others',
    'starlink': 'Gold',
    'swazimobile': 'Platinum',
    'mtnbissau': 'Platinum',
    'mtnsudan': 'Platinum'
}
sites = {
    'atm': 'Sweden ATM',
    'gloghana': 'Glo Ghana',
    'glonigeria': 'Glo Nigeria',
    'gosoft': 'Gosoft Thailand',
    'mtnaf': 'MTN Afghanistan',
    'mtnbenin': 'MTN Benin',
    'mtnci': 'MTN Ivory Coast',
    'mtncongob': 'MTN Congo',
    'mtngc': 'MTN GC',
    'mtnghana': 'MTN Ghana',
    'mtnliberia': 'MTN Liberia',
    'mtnsouthsudan': 'MTN South Sudan',
    'mtnsyria': 'MTN Syria',
    'mtnye': 'MTN Yemen',
    'mtnzambia': 'MTN Zambia',
    'newcobahamas': 'Newco Bahamas',
    'other': 'Others',
    'starlink': 'Starlink Qatar',
    'swazimobile': 'Swazi Mobile',
    'mtnbissau': 'MTN Bissau',
    'mtnsudan': 'MTN Sudan'
}

delete = []
for key in data:
    if key['queue'].split(':')[-1].lower().__contains__('inbox'):
        key['queue'] = 'other'
        key['customer origin'] = sites['other']
    else:
        key['queue'] = key['queue'].split(':')[-1].lower().replace(' ', '')
        if key['queue'] == 'it':
            delete.append(key)
        else:
            key['customer origin'] = sites[key['queue']]
    try:
        key['service'] = service[key['queue']] + ' SLA'
        key['sla'] = service[key['queue']]
    except:
        if not delete.__contains__(key):
            print(key['ticket#'] + 'Not Added')
    key['severity'] = key['priority']

book = xlsxwriter.Workbook(name_dest)
try:
    sheet0 = book.add_worksheet(name='Data')

    formats = {
            'number': book.add_format({'num_format': '0', 'align': 'center', 'border': 1}),
            'border': book.add_format({'align': 'center', 'border': 1}),
            'heading': book.add_format({'bg_color': '#000000', 'align': 'center', 'font_color': '#FFFFFF'}),
            'date': book.add_format({'num_format': 'dd/mmm/yy hh:mm', 'border': 1})
        }
    sheet0.hide_gridlines(2)
    col = 0
    header = {}
    for key in list(data[0].keys()):
        temp = ''
        if key == 'sla':
            temp = 'SLA'
        elif key == 'firstresponse':
            temp = 'First Response'
        elif key == 'solutiontime':
            temp = 'Solution Time'
        elif key == 'agent/owner':
            temp = 'Agent/Owner'
        else:
            for key1 in key.split():
                temp += ' ' + key1[0].upper() + key1[1:]
        if key.lower().__contains__('queue'):
            sheet0.write(0, len(list(data[0].keys())) - 1, temp, formats['heading'])
            header.update({key: len(list(data[0].keys())) - 1})
        else:
            sheet0.write(0, col, temp, formats['heading'])
            header.update({key: col})
            col += 1

    row = 1
    for key in data:
        if not delete.__contains__(key):
            for key1 in header:
                try:
                    if type(key[key1]) is datetime:
                        sheet0.write(row, header[key1], key[key1], formats['date'])
                    else:
                        sheet0.write(row, header[key1], int(key[key1]), formats['number'])
                except:
                    sheet0.write(row, header[key1], key[key1], formats['border'])
            # print(key)
            row += 1


finally:
    book.close()
input('Enter to Close.')
