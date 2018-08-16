import os
import xlsxwriter

path = '.\\source\\'
name_dest = '.\\Final_Data.xlsx'
l = os.listdir(path)
txt_name = []
for key in l:
    if (key.__contains__('CRITICAL') or key.__contains__('PROBLEM') or key.__contains__('WARNING') or key.__contains__('UNKNOWN') or key.__contains__('DOWN') or key.__contains__('Forwarded')):
        txt_name.append(key)
    else:
        print()
        # os.remove(path+key)

final = []
flag = 0
for key in txt_name:
    data_f = open(path + key, encoding='UTF-8')
    data = data_f.read()
    if (key.__contains__('DOWN')):
        if (data.__contains__('Notification Type:') and data.__contains__('Host:') and data.__contains__('State:')):
            final.append(key)
        else:
            flag=1
            # os.remove(path + key)
    elif (data.__contains__('Service:') and data.__contains__('Host:') and data.__contains__('State:')):
            final.append(key)
    else:
        flag=1
        # os.remove(path + key)
    if (not data_f._checkClosed()):
        data_f.close()
        if(flag==1):
            os.remove(path+key)
            flag=0

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
    sub_f = False
    _f_f = False
    serv_f = False
    h_f = False
    a_f = False
    stat_f = False
    d_f = False
    for key1 in data_f.readlines():
        if (key1.__contains__('Subject:') and sub_f is False):
            t = key1.replace('Subject:', '')
            t = t.replace(chr(10), '')
            subject.append(t)
            sub_f = True
        if (key1.__contains__('From:') and _f_f is False):
            t = key1.replace('From:', '')
            t = t.replace(chr(10), '')
            _from.append(t)
            _f_f = True
        if ((key1.__contains__('Service:') or ((key.__contains__("DOWN")) and key1.__contains__('Notification Type:')))and serv_f is False):
            if(key1.__contains__('Service:')):
                t = key1.replace('Service:', '')
            else:
                t = key1.replace('Notification Type:', '')
            t = t.replace(chr(10), '')
            service.append(t)
            serv_f = True
        if (key1.__contains__('Host:') and h_f is False) :
            t = key1.replace('Host:', '')
            t = t.replace(chr(10), '')
            host.append(t)
            h_f = True
        if (key1.__contains__('Address:') and a_f is False):
            t = key1.replace('Address:', '')
            t = t.replace(chr(10), '')
            address.append(t)
            a_f = True
        if (key1.__contains__('State:') and stat_f is False):
            t = key1.replace('State:', '')
            t = t.replace(chr(10), '')
            state.append(t)
            stat_f = True
        if (key1.__contains__('Date:') and d_f is False):
            t = key1.replace('Date:', '')
            t = t.replace(chr(10), '')
            date.append(t)
            d_f = True
    if(sub_f is False):
        subject.append('')
    if(_f_f is False):
        _from.append('')
    if(serv_f is False):
        service.append('')
    if(h_f is False):
        host.append('')
    if(a_f is False):
        address.append('')
    if(stat_f is False):
        state.append('')
    if(d_f is False):
        date.append('')


temp = []

for key in range(len(final)):
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

otrs = ['ye-mtn', 'af-mtn', 'sy-mtn', 'glo-ng', 'starlink', 'newco', 'mtn-c', 'gosoft', 'DNA', 'atm', 'bjmtn', 'gc-mtn','mtnliberia', 'gh-mtn', 'telecelBF', 'mtnsouthsudan', 'globenin', 'Datora', 'mtnzambia', 'mtnci', 'mtnbissau','gloghana','glo-gh', 'swazimobile']
flag = 0
for key in from_line:
    flag=0
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
