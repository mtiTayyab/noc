import os
import xlsxwriter

path = '.\\W2 May\\'
name_dest = 'NOC Data.xlsx'
l = os.listdir(path)
txt_name = []
for key in l:
    if (key.__contains__('CRITICAL') or key.__contains__('PROBLEM') or key.__contains__('WARNING') or key.__contains__(
            'UNKNOWN') or key.__contains__('DOWN')):
        txt_name.append(key)

final = []
for key in txt_name:
    data_f = open(path + key, encoding='UTF-8')
    data = data_f.read()
    if (key.__contains__('DOWN')):
        if (data.__contains__('Notification Type:') and data.__contains__('Host:') and data.__contains__('State:')):
            final.append(key)
    else:
        if (data.__contains__('Service:') and data.__contains__('Host:') and data.__contains__('State:')):
            final.append(key)
    if (not data_f._checkClosed()):
        data_f.close()

from_line = []
subject = []
_from = []
service = []
host = []
address = []
state = []
date = []
t = ''
for key in final:
    data_f = open(path + key, encoding='UTF-8')
    for key1 in data_f.readlines():
        if (key1.__contains__('Subject:')):
            t = key1.replace('Subject:', '')
            t = t.replace(chr(10), '')
            subject.append(t)
        if (key1.__contains__('From:')):
            t = key1.replace('From:', '')
            t = t.replace(chr(10), '')
            _from.append(t)
        if (key1.__contains__('Service:') or ((key.__contains__("DOWN")) and key1.__contains__('Notification Type:'))):
            t = key1.replace('Service:', '')
            t = key1.replace('Notification Type:', '')
            t = t.replace(chr(10), '')
            service.append(t)
        if (key1.__contains__('Host:')):
            t = key1.replace('Host:', '')
            t = t.replace(chr(10), '')
            host.append(t)
        if (key1.__contains__('Address:')):
            t = key1.replace('Address:', '')
            t = t.replace(chr(10), '')
            address.append(t)
        if (key1.__contains__('State:')):
            t = key1.replace('State:', '')
            t = t.replace(chr(10), '')
            state.append(t)
        if (key1.__contains__('Date:')):
            t = key1.replace('Date:', '')
            t = t.replace(chr(10), '')
            date.append(t)

temp = []

for key in range(len(subject)):
    if (final[key].__contains__("__")):
        k = final[key].split("__")
        final[key] = k[1]
    temp.append(final[key])
    temp.append(subject[key])
    temp.append(_from[key])
    temp.append(service[key])
    temp.append(host[key])
    temp.append(address[key])
    temp.append(state[key])
    temp.append(date[key])
    from_line.append(temp)
    temp = []

critical = []
warning = []
unknown = []

for key in range(len(state)):
    if (state[key].__contains__('CRITICAL')):
        temp.append(subject[key])
        temp.append(_from[key])
        temp.append(service[key])
        temp.append(address[key])
        temp.append(date[key])
        critical.append(temp)
        temp = []
    if (state[key].__contains__('UNKNOWN')):
        temp.append(subject[key])
        temp.append(_from[key])
        temp.append(service[key])
        temp.append(address[key])
        temp.append(date[key])
        unknown.append(temp)
        temp = []
    if (state[key].__contains__('WARNING')):
        temp.append(subject[key])
        temp.append(_from[key])
        temp.append(service[key])
        temp.append(address[key])
        temp.append(date[key])
        warning.append(temp)
        temp = []

for key in from_line:
    if (key[2].lower().__contains__('sds.noc@seamless.se')):
        from_line.remove(key)

otrs = ['ye-mtn', 'af-mtn', 'sy-mtn', 'glo-ng', 'starlink', 'newco', 'mtn-c', 'gosoft', 'DNA', 'atm', 'bjmtn', 'gc-mtn',
        'mtnliberia', 'gh-mtn', 'telecelBF', 'mtnsouthsudan', 'globenin', 'Datora', 'mtnzambia', 'mtnci', 'mtnbissau',
        'gloghana', 'swazimobile']
flag = 0
for key in from_line:
    for key1 in otrs:
        temp = key[1].lower()
        if (temp.__contains__(key1)):
            key[1] = key1
            flag = 1
            break
    if (flag == 0):
        for key1 in otrs:
            temp = key[2].lower()
            if (temp.__contains__(key1)):
                key[1] = key1
                flag = 1
                break
    else:
        flag = 0

# book = load_workbook('Structure.xlsx')
book = xlsxwriter.Workbook(name_dest)
sheet = book.add_worksheet()

# sheet = book['Sheet1']
sheet.write(0, 0, 'FileName')
sheet.write(0, 1, 'Subject')
sheet.write(0, 2, 'From')
sheet.write(0, 3, 'Service')
sheet.write(0, 4, 'Host')
sheet.write(0, 5, 'Address')
sheet.write(0, 6, 'State')
sheet.write(0, 7, 'Date')

row = 1

for key in range(len(from_line)):
    sheet.write(row, 0, from_line[key][0])
    sheet.write(row, 1, from_line[key][1])
    sheet.write(row, 2, from_line[key][2])
    sheet.write(row, 3, from_line[key][3])
    sheet.write(row, 4, from_line[key][4])
    sheet.write(row, 5, from_line[key][5])
    sheet.write(row, 6, from_line[key][6])
    sheet.write(row, 7, from_line[key][7])
    # sheet['B' + str(row)] = from_line[key][1]
    # sheet['C' + str(row)] = from_line[key][2]
    # sheet['D' + str(row)] = from_line[key][3]
    # sheet['E' + str(row)] = from_line[key][4]
    # sheet['F' + str(row)] = from_line[key][5]
    # sheet['G' + str(row)] = from_line[key][6]
    # sheet['H' + str(row)] = from_line[key][7]
    row += 1
book.close()
