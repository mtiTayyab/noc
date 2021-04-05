import os
import xlsxwriter
from math import floor
from datetime import datetime, date
from db import store_all_data, delete_all_data, filter_characters, get_all_zabix_alerts, get_all_vpn_alerts, \
    get_site_by_count_desc, filter_characters, get_zabix_alerts_by_issue, get_vpn_alerts_by_site

path = '.\\source\\'

date_s = date.today()
date_s = floor(date_s.day / 10) * 10
if date_s < 10:
    week = 'Week 3'
elif 10 <= date_s < 20:
    week = 'Week 1'
else:
    week = 'Week 2'
name_dest = '.\\Weekly KDC Report ' + week + '.xlsx'
l = os.listdir(path)
txt_name = []
for key in l:
    if key.lower().__contains__('zabbix') or key.lower().__contains__('vpn') or key.lower().__contains__('problem'):
        txt_name.append(key)
    # else:
    #     print(key)
    # else:
    #     print()
    # os.remove(path+key)

final = []
flag = 0
for key in txt_name:
    data_f = open(path + key, encoding='UTF-8')
    data = data_f.read()
    if (data.lower().__contains__('subject:') and data.lower().__contains__('from:') and data.lower().__contains__(
            'date:')):
        final.append(key)
    else:
        flag = 1
        os.remove(path + key)
    if not data_f._checkClosed():
        data_f.close()
        if (flag == 1):
            os.remove(path + key)
            flag = 0

subject = []
_type = []
_from = []
date = []
site = []
t = ''
for key in range(len(final)):
    data_f = open(path + final[key], encoding='UTF-8')
    sub_f = False
    _f_f = False
    _t_f = False
    d_f = False
    s_f = False
    for key1 in data_f.readlines():
        if (key1.lower().__contains__('subject:') and sub_f is False):
            t = key1.replace('Subject:', '')
            t = filter_characters(t)
            subject.append(t)
            sub_f = True
        if (key1.lower().__contains__('average') or key1.lower().__contains__('high') or key1.lower().__contains__(
                'information')) and _t_f is False:
            t = key1.split(":")[1]
            t = filter_characters(t)
            site.append(t)
            s_f = True
        if (key1.lower().__contains__('problem') or key1.lower().__contains__('critcal') or key1.lower().__contains__(
                'warning')) and _t_f is False:
            t = key1.split(":")[2]
            t = filter_characters(t)
            _type.append(t)
            _t_f = True
        if (key1.lower().__contains__('from:') and _f_f is False):
            t = key1.replace('From:', '')
            t = filter_characters(t)
            _from.append(t)
            _f_f = True
            stat_f = True
        if (key1.lower().__contains__('date: ') and d_f is False):
            t = key1.split('Date:')[1]
            t = filter_characters(t)
            # t = t.split(" ")
            if t.split()[2].lower() == 'pm':
                if int(t.split()[1].split(':')[0]) == 12:
                    t = datetime(day=int(t.split('/')[1]), month=int(t.split('/')[0]),
                                 year=int(t.split('/')[2].split()[0]), hour=int(t.split()[1].split(':')[0]) - 12,
                                 minute=int(t.split()[1].split(':')[1]))

                else:
                    t = datetime(day=int(t.split('/')[1]), month=int(t.split('/')[0]),
                                 year=int(t.split('/')[2].split()[0]), hour=int(t.split()[1].split(':')[0]) + 12,
                                 minute=int(t.split()[1].split(':')[1]))
            else:
                t = datetime(day=int(t.split('/')[1]), month=int(t.split('/')[0]), year=int(t.split('/')[2].split()[0]),
                             hour=int(t.split()[1].split(':')[0]), minute=int(t.split()[1].split(':')[1]))
            # t=t[0]
            date.append(t)
            d_f = True
    if (sub_f is False):
        subject.append('')
    if (_t_f is False):
        _type.append('')
    if (_f_f is False):
        _from.append('')
    if (d_f is False):
        date.append('')
    if (s_f is False):
        site.append('')

final_data = []

for key in range(len(final)):
    temp = {'final': final[key]}
    temp.update({'type': _type[key]})
    temp.update({'site': site[key]})
    temp.update({'from': _from[key]})
    temp.update({'date': date[key]})
    if _type[key]:
        temp.update({'subject': subject[key].split(_type[key])[1]})
    else:
        temp.update({'subject': subject[key]})
    if final[key].lower().__contains__('zabbix') or _from[key].lower().__contains__('zabbix'):
        temp.update({'category': 'zabbix'})
    elif final[key].lower().__contains__('vpn'):
        temp.update({'category': 'vpn'})
        temp.update({'subject': subject[key]})
    final_data.append(temp)

delete = []
for key in final_data:
    if key['subject'].lower().__contains__('ok'):
        delete.append(key)

for key in delete:
    if final_data.__contains__(key):
        final_data.remove(key)

site_f = ['Glo Nigeria', 'MTN Afganistan', 'MTN Benin', 'MTN IvoryCoast', 'MTN Nigeria', 'MTN Rwanda', 'MTN SouthSudan',
          'MTN Sudan', 'MTN Yemen', 'MTN Zambia', 'NewCo Bahamas', 'Swazi Mobile', 'KDC Data Centre']

noc_dict = {
    'glo_nigeria': [],
    'mtn_afganistan': [],
    'mtn_benin': [],
    'mtn_ivorycoast': [],
    'mtn_nigeria': [],
    'mtn_rwanda': [],
    'mtn_southsudan': [],
    'mtn_sudan': [],
    'mtn_yemen': [],
    'mtn_zambia': [],
    'newco_bahamas': [],
    'swazi_mobile': [],
    'kdc_data_centre': [],
    'average': [],
    'high': [],
    'information': []
}
otrs = [['glo nigeria', 'ng glo'], 'mtn afganistan', 'mtn benin', 'mtn ivorycoast', 'mtn nigeria', 'mtn rwanda',
        'mtn southsudan', ['mtn sudan', 'mtnsudan'], 'mtn yemen', 'mtn zambia', 'newco bahamas', 'swazi mobile', 'sds']

for key in final_data:
    for key1 in range(len(otrs)):
        if type(otrs[key1]) is list:
            for key2 in range(len(otrs[key1])):
                if key['subject'].lower().__contains__(otrs[key1][key2]):
                    key['site'] = site_f[key1]
                    break
        else:
            if key['subject'].lower().__contains__(otrs[key1]):
                key['site'] = site_f[key1]
                break

for key in final_data:
    try:
        noc_dict[key['site'].lower().replace(' ', '_')].append(key)
    except KeyError:
        print(key)

book = xlsxwriter.Workbook(name_dest)
try:
    store_all_data(final_data,delete)

    date_format = book.add_format({'num_format': 'dd/mmm/yy hh:mm:ss', 'align': 'center'})
    date_format.set_border(style=1)
    border_format = book.add_format({'align': 'center'})
    border_format.set_border(style=1)
    cell_format = book.add_format({'bg_color': '#000000', 'font_color': '#FFFFFF', 'align': 'center'})
    formats = {
        'date': date_format,
        'border': border_format,
        'heading': cell_format
    }
    sheet0 = book.add_worksheet()

    result = get_all_zabix_alerts()
    sheet = book.add_worksheet('Zabbix')
    sheet.hide_gridlines(2)

    sheet.write(0, 0, 'Severity', formats['heading'])
    sheet.write(0, 1, 'Types', formats['heading'])
    sheet.write(0, 2, 'Alerts', formats['heading'])
    sheet.write(0, 3, 'From', formats['heading'])
    sheet.write(0, 4, 'Date', formats['heading'])

    row = 1
    for key in result:
        sheet.write(row, 0, key[0], formats['border'])
        sheet.write(row, 1, key[1], formats['border'])
        sheet.write(row, 2, key[2], formats['border'])
        sheet.write(row, 3, key[3], formats['border'])
        sheet.write(row, 4, key[4], formats['date'])
        row += 1

    result = get_zabix_alerts_by_issue()

    sheet = book.add_worksheet('Zabbix_Graph')
    sheet.hide_gridlines(2)
    severity = {}
    issue = {}
    severity_row = 1
    issue_row = 1
    sheet.write(0, 0, 'Issue/Severity', formats['heading'])
    for key in result:
        if not severity.__contains__(key[0]):
            severity.update({key[0]: severity_row})
            sheet.write(severity_row, 0, key[0], formats['border'])
            severity_row += 1
        if not issue.__contains__(key[1]):
            issue.update({key[1]: issue_row})
            sheet.write(0, issue_row, key[1], formats['heading'])
            issue_row += 1
        sheet.write(severity[key[0]], issue[key[1]], key[2])

    sheet.write(severity_row, 0, 'Total', formats['heading'])
    sheet.write(severity_row, 1, '=SUM(' + chr(65 + 1) + str(2) + ':' + chr(65 + 1) + str(severity_row) + ')',
                formats['heading'])
    sheet.write(0, issue_row, 'Grand Total', formats['heading'])
    sheet.write(1, issue_row, '=SUM(' + chr(65 + 1) + str(2) + ':' + chr(65 + issue_row - 1) + str(2) + ')',
                formats['border'])

    chart = book.add_chart({'type': 'bar'})
    for key2 in range(1, issue_row):
        chart.add_series({
            'values': ('=Zabbix_Graph!$' + chr(65 + key2) + '$2:$' + chr(65 + key2) + '$' + str(severity_row)),
            'name': ('=Zabbix_Graph!$' + chr(65 + key2) + '$1'),
            'categories': ('=Zabbix_Graph!$A$2:$A$' + str(severity_row)),
            'data_labels': {
                'value': True,
                'font': {
                    'color': '#000000'
                }
            }
        })
    chart.set_x_axis({'visible': False})
    sheet.insert_chart('G8', chart)

    result = get_all_vpn_alerts()
    sheet = book.add_worksheet('Root')
    sheet.hide_gridlines(2)

    sheet.write(0, 0, 'Site', formats['heading'])
    sheet.write(0, 1, 'Issue', formats['heading'])
    sheet.write(0, 2, 'From', formats['heading'])
    sheet.write(0, 3, 'Date', formats['heading'])

    row = 1
    for key in result:
        sheet.write(row, 0, key[0], formats['border'])
        sheet.write(row, 1, key[1], formats['border'])
        sheet.write(row, 2, key[2], formats['border'])
        sheet.write(row, 3, key[3], formats['date'])
        row += 1

    result = get_vpn_alerts_by_site()
    sheet = book.add_worksheet('Root_Graph')
    sheet.hide_gridlines(2)

    sheet.write(0, 0, 'Sites', formats['heading'])
    sheet.write(0, 1, 'Count', formats['heading'])

    row = 1
    for key in result:
        sheet.write(row, 0, key[0], formats['border'])
        sheet.write(row, 1, key[1], formats['border'])
        row += 1

    sheet.write(row, 0, 'Total', formats['heading'])
    sheet.write(row, 1, '=SUM(Root_Graph!$B$' + str(1) + ':$B' + str(row) + ')', formats['heading'])

    chart = book.add_chart({'type': 'bar'})

    chart.add_series({
        'values': ('=Root_Graph!$B$2:$B$' + str(row)),
        'name': '=Root_Graph!$B$1',
        'categories': ('=Root_Graph!$A$2:$A$' + str(row)),
        'data_labels': {
            'value': True,
            'font': {
                'color': '#000000'
            }
        }
    })
    chart.set_x_axis({'visible': False})
    sheet.insert_chart('D5', chart)

    sheet0.hide_gridlines(2)

    sheet0.write(0, 0, 'Site/Type', formats['heading'])
    sheet0.write(0, 1, 'Severity', formats['heading'])
    sheet0.write(0, 2, 'Subject', formats['heading'])
    sheet0.write(0, 3, 'From', formats['heading'])
    sheet0.write(0, 4, 'Date', formats['heading'])
    row = 1
    site_list = get_site_by_count_desc()
    for key1 in site_list:
        for key in noc_dict[key1[0].lower().replace(' ', '_')]:
            sheet0.write(row, 0, key['site'], formats['border'])
            sheet0.write(row, 1, key['type'], formats['border'])
            sheet0.write(row, 2, filter_characters(key['subject']), formats['border'])
            sheet0.write(row, 3, key['from'], formats['border'])
            sheet0.write(row, 4, key['date'], formats['date'])
            row += 1


finally:
    delete_all_data()
    book.close()
input('Enter to Close.')
