import os
import xlsxwriter
from datetime import datetime
import pymysql
from db import get_site_by_count_desc,store_all_data,delete_data
from miscellaneous import filter_characters

path = '.\\source\\'
name_dest = '.\\Final_Data.xlsx'
l = os.listdir(path)
txt_name = []
for key in l:
    if (key.__contains__('CRITICAL') or key.__contains__('PROBLEM') or key.__contains__('WARNING') or key.__contains__('UNKNOWN') or key.__contains__('DOWN') or key.__contains__('Forwarded')):
        txt_name.append(key)
    # else:
    #     print()
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
    if not data_f._checkClosed():
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
for key in range(len(final)):
    data_f = open(path + final[key], encoding='UTF-8')
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
            t = filter_characters(t)
            subject.append(t)
            sub_f = True
        if (key1.__contains__('From:') and _f_f is False):
            t = key1.replace('From:', '')
            t = filter_characters(t)
            _from.append(t)
            _f_f = True
        if ((key1.__contains__('Service:') or ((final[key].__contains__("DOWN")) and key1.__contains__('Notification Type:')))and serv_f is False):
            if(key1.__contains__('Service:')):
                if(final[key].lower().__contains__('forward')):
                    final[key] = key1
                # t = key1.replace('Service:', '')
                t = key1.split('Service:')[1].split('Host:')[0]
                t = filter_characters(t)
            else:
                # t = key1.replace('Notification Type:', '')
                t = key1.split('Notification Type:')[1].split('State:')[0]
                t = filter_characters(t)
            service.append(t)
            serv_f = True
        if (key1.__contains__('Host:') and h_f is False) :
            # t = key1.replace('Host:', '')
            t = key1.split('Host:')[1].split('State:')[0]
            t = filter_characters(t)
            host.append(t)
            h_f = True
        if (key1.__contains__('Address:') and a_f is False):
            t = key1.replace('Address:', '')
            t = filter_characters(t)
            address.append(t)
            a_f = True
        if (key1.__contains__('State:') and stat_f is False):
            t = key1.split('State:')[1].split('Date:')[0]
            t = filter_characters(t)
            state.append(t)
            stat_f = True
        if (key1.__contains__('Date:') and d_f is False):
            t = key1.replace('Date:', '')
            t = filter_characters(t)
            # t = t.split(" ")
            if t.split()[2].lower() == 'pm':
                if int(t.split()[1].split(':')[0])==12:
                    t = datetime(day=int(t.split('/')[1]), month=int(t.split('/')[0]), year=int(t.split('/')[2].split()[0]),hour=int(t.split()[1].split(':')[0])-12, minute=int(t.split()[1].split(':')[1]))

                else:
                    t = datetime(day=int(t.split('/')[1]), month=int(t.split('/')[0]),year=int(t.split('/')[2].split()[0]), hour=int(t.split()[1].split(':')[0]) + 12,minute=int(t.split()[1].split(':')[1]))
            else:
                t = datetime(day=int(t.split('/')[1]), month=int(t.split('/')[0]), year=int(t.split('/')[2].split()[0]),hour=int(t.split()[1].split(':')[0]), minute=int(t.split()[1].split(':')[1]))
            # t=t[0]
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



delete = []
for key in from_line:
    if (key[2].lower().__contains__('sds.noc@seamless.se')):
        delete.append(key)
    if (key[0].lower().__contains__('acknowledgement')):
        delete.append(key)
    if (key[3].lower().__contains__('acknowledgement')):
        delete.append(key)
    if key[6].lower()=='ok':
        delete.append(key)
    if key[1].lower().__contains__('hospitall'):
        delete.append(key)
    if key[1].lower().__contains__('localhost'):
        delete.append(key)
    if key[1].lower().__contains__('jira'):
        delete.append(key)
    if key[1].lower().__contains__('odoo'):
        delete.append(key)
    if key[1].lower().__contains__('lahore'):
        delete.append(key)
    if key[4].lower().__contains__('server 22'):
        delete.append(key)
    if key[4].lower().__contains__('server 25'):
        delete.append(key)
    # if key[1].lower().__contains__('local'):
    #     delete.append(key)

for key in delete:
    if from_line.__contains__(key):
        from_line.remove(key)
    else:
        print('Alert Not Found')
        print(key)

otrs = ['ye-mtn', 'af-mtn', 'sy-mtn', 'glo-ng', 'starlink', 'newco', 'mtn-c', 'gosoft', 'dna-finland', 'atm', 'bjmtn', 'gc-mtn','mtnliberia', 'gh-mtn', 'telecelBF', 'mtnsouthsudan', 'globenin', 'Datora', 'mtnzambia', 'mtnci', 'mtnbissau','gloghana','glo-gh', 'swazimobile','mtn-gb','mtn-benin','mtn-sy','mtn zambia','mtn-southsudan','sudan-mtn','ci@mtn']
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
        if(flag == 0):
            for key1 in otrs:
                temp = key[0].lower()
                if (temp.__contains__(key1)):
                    key[1] = key1
                    flag = 1
                    break
        else:
            flag = 0
    else:
        flag = 0

site_f = ['MTN Yemen', 'MTN Afghanistan', 'MTN Syria', 'Glo Nigeria', 'Starlink Qatar', 'NewCo Bahamas', 'MTN Congo',
 'Gosoft Thailand', 'DNA Finland', 'SE BANK SYSTEM', 'MTN Benin', 'MTN GC', 'MTN Liberia', 'MTN Ghana',
 'Telecel Burkina Faso', 'MTN South Sudan', 'Glo Benin', 'Datora Brazil', 'MTN Zambia', 'MTN Ivory Coast', 'MTN Bissau',
 'Glo Ghana', 'Swazi Mobile', 'MTN Bissau', 'MTN Sudan']

site_r = ['ye-mtn', 'af-mtn', ['sy-mtn', 'mtn-sy'], 'glo-ng', 'starlink', 'newco', 'mtn-c', 'gosoft', 'dna-finland', 'atm',
 ['bjmtn', 'mtn-benin'], 'gc-mtn', 'mtnliberia', 'gh-mtn', 'telecelBF', ['mtnsouthsudan', 'mtn-southsudan'], 'globenin',
 'Datora', ['mtnzambia','mtn zambia'], ['mtnci','ci@mtn'], 'mtnbissau', ['gloghana', 'glo-gh'], 'swazimobile', 'mtn-gb', 'sudan-mtn']
for key in from_line:
    for key1 in range(len(site_r)):
        if site_r[key1].__contains__(key[1]):
            key[1]=site_f[key1]




noc_dict = {
    'mtn_yemen':[],
    'mtn_afghanistan':[],
    'mtn_syria':[],
    'glo_nigeria':[],
    'starlink_qatar':[],
    'newco_bahamas':[],
    'mtn_congo':[],
    'gosoft_thailand':[],
    'dna_finland':[],
    'se_bank_system':[],
    'mtn_benin':[],
    'mtn_gc':[],
    'mtn_liberia':[],
    'mtn_ghana':[],
    'telecel_burkina_faso':[],
    'mtn_south_sudan':[],
    'glo_benin':[],
    'datora_brazil':[],
    'mtn_zambia':[],
    'mtn_ivory_coast':[],
    'glo_ghana':[],
    'swazi_mobile':[],
    'mtn_bissau':[],
    'mtn_sudan':[],
    'critical':[],
    'warning':[],
    'unknown':[]
}
for key in from_line:
    if key[6].lower().__contains__('critical'):
        noc_dict['critical'].append(key)
    if key[6].lower().__contains__('unknown'):
        noc_dict['unknown'].append(key)
    if key[6].lower().__contains__('warning'):
        noc_dict['warning'].append(key)
    try:
        noc_dict[key[1].lower().replace(' ','_')].append(key)
    except KeyError:
        print('Alert Not Added')
        print(key)
        f = input('Enter anything to pass.')
        from_line.remove(key)
# for key in range(len(noc_dict['mtn_bissau'])-1):
#     for key1 in range(key,len(noc_dict['mtn_bissau'])):
#         if key1<len(noc_dict['mtn_bissau']):
#             if noc_dict['mtn_bissau'][key][3]==noc_dict['mtn_bissau'][key1][3]:
#                 if noc_dict['mtn_bissau'][key][4]==noc_dict['mtn_bissau'][key1][4]:
#                     if noc_dict['mtn_bissau'][key][6]==noc_dict['mtn_bissau'][key1][6]:
#                         if noc_dict['mtn_bissau'][key][7].day==noc_dict['mtn_bissau'][key1][7].day:
#                             if noc_dict['mtn_bissau'][key][7].month==noc_dict['mtn_bissau'][key1][7].month:
#                                 if noc_dict['mtn_bissau'][key][7].year==noc_dict['mtn_bissau'][key1][7].year:
#                                     if noc_dict['mtn_bissau'][key][7].hour==noc_dict['mtn_bissau'][key1][7].hour:
#                                         if (noc_dict['mtn_bissau'][key][7].minute-noc_dict['mtn_bissau'][key1][7].minute)*-1<=5:
#                                             noc_dict['mtn_bissau'].remove(noc_dict['mtn_bissau'][key1])
#         else:
#             break



store_all_data(from_line)

book = xlsxwriter.Workbook(name_dest)
sheet = book.add_worksheet()



cell_format = book.add_format()
date_format = book.add_format({'num_format':'dd/mmm/yy'})
cell_format.set_bg_color('#000000')
cell_format.set_font_color('#FFFFFF')


sheet.write(0, 0, 'Site / Region',cell_format)
sheet.write(0, 1, 'Service',cell_format)
sheet.write(0, 2, 'Alert Type',cell_format)
sheet.write(0, 3, 'Alert',cell_format)
sheet.write(0, 4, 'Host',cell_format)
sheet.write(0, 5, 'Address',cell_format)
sheet.write(0, 6, 'Date',cell_format)

row = 1

site_list = get_site_by_count_desc()

for key1 in site_list:
    sheet1 = None
    row1=1
    for key in noc_dict[key1[0].lower().replace(' ','_')]:
        if sheet1 is not None:
            sheet1.write(row1, 0, key[1])
            sheet1.write(row1, 1, key[3])
            sheet1.write(row1, 2, key[6])
            sheet1.write(row1, 3, filter_characters(key[0]))
            sheet1.write(row1, 4, key[4])
            sheet1.write(row1, 5, key[5])
            sheet1.write(row1, 6, key[7], date_format)
            row1+=1
        else:
            sheet1= book.add_worksheet(name=key1[0])
            sheet1.write(0, 0, 'Site / Region', cell_format)
            sheet1.write(0, 1, 'Service', cell_format)
            sheet1.write(0, 2, 'Alert Type', cell_format)
            sheet1.write(0, 3, 'Alert', cell_format)
            sheet1.write(0, 4, 'Host', cell_format)
            sheet1.write(0, 5, 'Address', cell_format)
            sheet1.write(0, 6, 'Date', cell_format)
            sheet1.write(row1, 0, key[1])
            sheet1.write(row1, 1, key[3])
            sheet1.write(row1, 2, key[6])
            sheet1.write(row1, 3, filter_characters(key[0]))
            sheet1.write(row1, 4, key[4])
            sheet1.write(row1, 5, key[5])
            sheet1.write(row1, 6, key[7], date_format)
            row1+=1
        sheet.write(row, 0, key[1])
        sheet.write(row, 1, key[3])
        sheet.write(row, 2, key[6])
        sheet.write(row, 3, filter_characters(key[0]))
        sheet.write(row, 4, key[4])
        sheet.write(row, 5, key[5])
        sheet.write(row, 6, key[7],date_format)
        row += 1
book.close()


delete_data()
f = input('Everyhting was fine. Enter to close.')