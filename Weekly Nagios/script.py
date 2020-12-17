import os
import xlsxwriter
from datetime import datetime, date
from db import get_site_by_count_desc, store_all_data, delete_data, get_service_host_by_site, get_service_by_site, \
    get_alerts_by_type_and_site, get_alert_by_site, get_alert_by_alert_type, get_alert_by_team, filter_characters
from math import floor

path = './/source//'

date_s = date.today()
date_s = floor(date_s.day / 10) * 10
if date_s < 10:
    week = 'Week 3'
elif 10 <= date_s < 20:
    week = 'Week 1'
else:
    week = 'Week 2'
name_dest = './/Weekly Nagios Report ' + week + '.xlsx'
l = os.listdir(path)
txt_name = []
for key in l:
    if (key.__contains__('CRITICAL') or key.__contains__('PROBLEM') or key.__contains__('WARNING') or key.__contains__(
            'UNKNOWN') or key.__contains__('DOWN') or key.__contains__('Forwarded') or key.__contains__("SMS")):
        txt_name.append(key)
    # else:
    #     print()
    # os.remove(path+key)

final = []
flag = 0
for key in txt_name:
    data_f = open(path + key, encoding='UTF-8')
    data = data_f.read()
    if key.__contains__('DOWN'):
        if data.__contains__('Notification Type:') and data.__contains__('Host:') and data.__contains__('State:'):
            final.append(key)
        else:
            flag = 1
            # os.remove(path + key)
    elif data.__contains__('Service:') or (data.__contains__('Host:') and data.__contains__('State:')):
        final.append(key)
    else:
        flag = 1
    if not data_f._checkClosed():
        data_f.close()
        if flag == 1:
            os.remove(path + key)
            flag = 0

final_data = []
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
        if key1.__contains__('Subject:') and sub_f is False:
            t = key1.replace('Subject:', '')
            t = filter_characters(t)
            subject.append(t)
            sub_f = True
        if key1.__contains__('From:') and _f_f is False:
            t = key1.replace('From:', '')
            t = filter_characters(t)
            _from.append(t)
            _f_f = True
        if ((key1.__contains__('Service:') or (
                (final[key].__contains__("DOWN")) and key1.__contains__('Notification Type:'))) and serv_f is False):
            if key1.__contains__('Service:'):
                # if (final[key].lower().__contains__('forward')):
                #     final[key] = key1
                # t = key1.replace('Service:', '')
                t = key1.split('Service:')[1].split('Host:')[0]
                t = filter_characters(t)
            else:
                # t = key1.replace('Notification Type:', '')
                t = key1.split('Notification Type:')[1].split('State:')[0]
                t = filter_characters(t)
            service.append(t)
            serv_f = True
        if key1.__contains__('Host:') and h_f is False:
            if key1.lower().__contains__('se-bank-system') and key1.lower().__contains__('atm'):
                t = key1.split('Host:')[1].split('State:')[0]
                t = t.split()
                t[0] = filter_characters(t[0])
                t[1] = filter_characters(t[1])
                host.append(t[0])
                service.append(t[1])
                h_f = True
                serv_f = True
            elif key1.lower().__contains__('mtnz'):
                t = key1.split('Host:')[1].split('State:')[0]
                t = filter_characters(t)
                t = filter_characters(t)
                host.append(t)
                t = key1.split('MTNZ:')[1].split('Host:')[0]
                t = filter_characters(t)
                service.append(t)
                h_f = True
                serv_f = True
            else:
                t = key1.split('Host:')[1].split('State:')[0]
                t = filter_characters(t)
                host.append(t)
                h_f = True
        if key1.__contains__('Address:') and a_f is False:
            t = key1.replace('Address:', '')
            t = filter_characters(t)
            address.append(t)
            a_f = True
        if key1.__contains__('State:') and stat_f is False:
            t = key1.split('State:')[1].split('Date:')[0]
            t = filter_characters(t)
            state.append(t)
            stat_f = True
        if key1.__contains__('Date: ') and d_f is False:
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
    if sub_f is False:
        subject.append('')
    if _f_f is False:
        _from.append('')
    if serv_f is False:
        service.append('')
    if h_f is False:
        host.append('')
    if a_f is False:
        address.append('')
    if stat_f is False:
        state.append('')
    if d_f is False:
        date.append('')

temp = []

for key in range(len(final)):
    if final[key].__contains__("__"):
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
    final_data.append(temp)
    temp = []

delete = []
for key in final_data:
    if key[2].lower().__contains__('sds.noc@seamless.se'):
        delete.append(key)
    if key[0].lower().__contains__('acknowledgement'):
        delete.append(key)
    if key[0].lower().__contains__('recovery'):
        delete.append(key)
    if key[3].lower().__contains__('acknowledgement'):
        delete.append(key)
    if key[6].lower() == 'ok':
        delete.append(key)
    if key[6].lower() == 'up':
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
    if key[3].lower().__contains__('current'):
        delete.append(key)
    if key[3].lower().__contains__('total'):
        delete.append(key)
    if key[1].lower().__contains__(' local '):
        delete.append(key)

for key in delete:
    if final_data.__contains__(key):
        final_data.remove(key)

otrs = ['ye-mtn', 'af-mtn', 'sy-mtn', 'glo-ng', 'starlink', 'newco', 'mtn-c', 'gosoft', 'dna-finland', 'atm', 'bjmtn',
        'gc-mtn', 'mtnliberia', 'gh-mtn', 'telecelBF', 'mtnsouthsudan', 'globenin', 'Datora', 'mtnzambia', 'mtnci',
        'mtnbissau', 'gloghana', 'glo-gh', 'swazimobile', 'mtn-gb', 'mtn-benin', 'mtn-sy', 'mtn zambia',
        'mtn-southsudan', 'sudan-mtn', 'ci@mtn', 'mtn-lib', 'mtn lib', 'zm', 'lr mtn', 'syria', 'banksystem', 'mtnz',
        'evdnms', 'et sdt', 'mtnrw evd', 'mtnnevd', 'evd.ss', 'mtn-esw', 'mtnng', 'mtnrw', 'expressotelecom', 'zain-iraq', 'zain-ksa', 'tashicell',
        'btcl', 'ooa-drc', 'indosat', 'ooa-pr', 'za@mtn']

flag = 0
for key in final_data:
    flag = 0

    # if key[0].lower().__contains__("-sms from mtn-"):
    #   key[1] = "mtnbissau"
    #   flag = 1

    if key[0].lower().__contains__("-sms from evd-") or key[0].lower().__contains__("-sms from 761665-"):
        key[1] = "mtn lib"
        flag = 1

    for key1 in otrs:
        temp = key[1].lower()
        if temp.__contains__(key1):
            key[1] = key1
            flag = 1
            break
    if flag == 0:
        for key1 in otrs:
            temp = key[2].lower()
            if temp.__contains__(key1):
                key[1] = key1
                flag = 1
                break
        if flag == 0:
            for key1 in otrs:
                temp = key[0].lower()
                if temp.__contains__(key1):
                    key[1] = key1
                    flag = 1
                    break
            if flag == 0:
                for key1 in otrs:
                    temp = key[4].lower()
                    if temp.__contains__(key1):
                        key[1] = key1
                        flag = 1
                        break
            else:
                flag = 0
        else:
            flag = 0
    else:
        flag = 0

site_f = ['MTN_Yemen', 'MTN_Afghanistan', 'MTN_Syria', 'Glo_Nigeria', 'Starlink_Qatar', 'NewCo_Bahamas', 'MTN_Congo',
          'Gosoft_Thailand', 'DNA_Finland', 'SE_BANK_SYSTEM', 'MTN_Benin', 'MTN_GC', 'MTN_Liberia', 'MTN_Ghana',
          'MTN_South_Sudan', 'Glo_Benin', 'MTN_Zambia', 'MTN_Ivory_Coast', 'MTN_Bissau', 'Glo_Ghana', 'Eswatini_Mobile',
          'MTN_Sudan', 'SDT_Ethiopia', 'MTN_Rwanda', 'MTN_Nigeria', 'MTN_Eswatini', 'Expresso_Senegal', 'Zain_Iraq', 'Zain_KSA', 'Tashicell', 'BTCL',
          'OOA_DRC', 'Indosat','OOA-PR', 'MTN_South_Africa']

site_r = ['ye-mtn', 'af-mtn',
          ['sy-mtn', 'mtn-sy', 'syria'],
          'glo-ng', 'starlink', 'newco', 'mtn-c', 'gosoft', 'dna-finland',
          ['atm', 'banksystem'],
          ['bjmtn', 'mtn-benin'],
          'gc-mtn',
          ['mtnliberia', 'mtn-lib', 'mtn lib', 'lr mtn'],
          'gh-mtn',
          ['mtnsouthsudan', 'mtn-southsudan', 'evd.ss'],
          'globenin',
          ['mtnzambia', 'mtn zambia', 'zm', 'mtnz'],
          ['mtnci', 'ci@mtn'],
          ['mtnbissau', 'mtn-gb'],
          ['gloghana', 'glo-gh', 'evdnms'],
          'swazimobile', 'sudan-mtn', 'et sdt',
          ['mtnrw evd', 'mtnrw'],
          ['mtnnevd', 'mtnng'], 'mtn-esw', 'expressotelecom', 'zain-iraq', 'zain-ksa', 'tashicell', 'btcl', 'ooa-drc',
          'indosat', 'ooa-pr', 'za@mtn']
for key in final_data:
    for key1 in range(len(site_r)):
        if site_r[key1].__contains__(key[1]):
            key[1] = site_f[key1]

noc_dict = {
    'mtn_yemen': [],
    'mtn_afghanistan': [],
    'mtn_syria': [],
    'glo_nigeria': [],
    'starlink_qatar': [],
    'newco_bahamas': [],
    'mtn_congo': [],
    'gosoft_thailand': [],
    'dna_finland': [],
    'se_bank_system': [],
    'mtn_benin': [],
    'mtn_gc': [],
    'mtn_liberia': [],
    'mtn_ghana': [],
    'telecel_burkina_faso': [],
    'mtn_south_sudan': [],
    'glo_benin': [],
    'datora_brazil': [],
    'mtn_zambia': [],
    'mtn_ivory_coast': [],
    'glo_ghana': [],
    'eswatini_mobile': [],
    'mtn_bissau': [],
    'mtn_sudan': [],
    'mtn_rwanda': [],
    'sdt_ethiopia': [],
    'mtn_nigeria': [],
    'mtn_eswatini': [],
    'expresso_senegal': [],
    'zain_iraq': [],
    'zain_ksa': [],
    'tashicell': [],
    'btcl': [],
    'ooa_drc': [],
    'indosat': [],
    'ooa-pr': [],
    'mtn_south_africa': [],
    'critical': [],
    'warning': [],
    'unknown': [],
    'current': []
}
delete = []
for key in final_data:
    if key[6].lower().__contains__('critical'):
        noc_dict['critical'].append(key)
    if key[6].lower().__contains__('unknown'):
        noc_dict['unknown'].append(key)
    if key[6].lower().__contains__('warning'):
        noc_dict['warning'].append(key)
    try:
        noc_dict[key[1].lower().replace(' ', '_')].append(key)
    except KeyError:

        if key[0].lower().__contains__("-sms from 590-"):
            delete.append(key)
            continue

        if key[0].lower().__contains__("esg.com"):
            delete.append(key)
            continue

        # if key[0].lower().__contains__("-sms from mtnrw-"):
        #    delete.append(key)
        #    continue

        if key[0].lower().__contains__("-sms from 654233-"):
            delete.append(key)
            continue

        if key[0].lower().__contains__("-sms from 0944094044-"):
            delete.append(key)
            continue

        if key[0].lower().__contains__("-sms from 315688-"):
            delete.append(key)
            continue

        if key[0].lower().__contains__("-sms from 255822-"):
            delete.append(key)
            continue
        if key[6] not in ("CRITICAL", "WARNING", "UNKNOWN", "DOWN"):
            print(key)
            delete.append(key)
            continue

        print(key)
        site_name = input('Enter the site name of the this alert : ')
        if list(noc_dict.keys()).__contains__(site_name.lower()):
            noc_dict[site_name.lower().replace(' ', '_')].append(key)
            key[1] = site_name
            print('Alert Added')
        elif site_name:
            print('Enter any of the above:')
            print(site_f)
            site_name = input('Enter the site name of the this alert: ')
            if list(noc_dict.keys()).__contains__(site_name.lower()):
                noc_dict[site_name.lower().replace(' ', '_')].append(key)
                key[1] = site_name
                print('Alert Added')
            else:
                delete.append(key)
        else:
            delete.append(key)
for key in delete:
    if final_data.__contains__(key):
        final_data.remove(key)
#
# choice_f = 0
# while choice_f == 0:
#     choice = input('Enter [y\\n] to Filter SMS Duplication Alerts:')
#     if choice.lower() == 'y':
#         delete = []
#         for key2 in list(noc_dict.keys()):
#             if key2.lower().__contains__('bissau') or key2.lower().__contains__('critical')or key2.lower().__contains__('warning')or key2.lower().__contains__('unknown'):
#                 continue
#             for key in range(len(noc_dict[key2]) - 1):
#                 for key1 in range(key + 1, len(noc_dict[key2])):
#                     if noc_dict[key2][key][3] == noc_dict[key2][key1][3]:
#                         if noc_dict[key2][key][4] == noc_dict[key2][key1][4]:
#                             if noc_dict[key2][key][5] != noc_dict[key2][key1][5]:
#                                 if noc_dict[key2][key][6] == noc_dict[key2][key1][6]:
#                                     temp_v = noc_dict[key2][key][7] - noc_dict[key2][key1][7]
#                                     if temp_v.seconds >= 0 and temp_v.seconds <= 300:
#                                         print('key: '+str(key))
#                                         print('key1: '+str(key1))
#                                         print('key2: '+str(key2))
#                                         if noc_dict[key2][key1][5]=='' or not len(noc_dict[key2][key1][5].split('.'))==4:
#                                             delete.append(noc_dict[key2][key1])
#                                         else:
#                                             delete.append(noc_dict[key2][key])
#
#
#         for key in delete:
#             if noc_dict[key[1].lower()].__contains__(key):
#                 noc_dict[key[1].lower()].remove(key)
#             if final_data.__contains__(key):
#                 final_data.remove(key)
#         print('Duplicates Deleted')
#         choice_f = 1
#     elif choice.lower() == 'n':
#         break
#     else:
#         print('Please enter a proper option.')


book = xlsxwriter.Workbook(name_dest)

try:
    # Store all alerts data.
    store_all_data(final_data)
    sheet = book.add_worksheet()

    # Defining all the formats needed in the script
    cell_format = book.add_format({'align': 'center'})
    border_format = book.add_format({'align': 'center'})
    border_format.set_border(style=1)
    service_cell_format = book.add_format({'align': 'center'})
    date_format = book.add_format({'num_format': 'dd/mmm/yy hh:mm:ss', 'align': 'center'})
    date_format.set_border(style=1)
    cell_format.set_bg_color('#000000')
    service_cell_format.set_bg_color('#8db4e2')
    cell_format.set_font_color('#FFFFFF')
    service_cell_format.set_font_color('#000000')

    # Writting Headings on the first sheet of all the data
    sheet.write(0, 0, 'Site / Region', cell_format)
    sheet.write(0, 1, 'Service', cell_format)
    sheet.write(0, 2, 'Alert Type', cell_format)
    sheet.write(0, 3, 'Alert', cell_format)
    sheet.write(0, 4, 'Host', cell_format)
    sheet.write(0, 5, 'Address', cell_format)
    sheet.write(0, 6, 'Date', cell_format)

    row = 1
    sheet.hide_gridlines(2)

    # Starting per site analysis sheet

    site_list = get_site_by_count_desc()
    sheets = {}
    sheet2 = book.add_worksheet('Per_Site_Stats')
    sheet2.write(0, 0, 'Alert Type/Site', cell_format)
    sheet2.write(1, 0, 'Critical', cell_format)
    sheet2.write(2, 0, 'Warning', cell_format)
    sheet2.write(3, 0, 'Unknown', cell_format)
    sheet2.write(4, 0, '', cell_format)
    sheet2.hide_gridlines(2)

    # getting data of per site analysis first data table (alert type and site)
    result = get_alerts_by_type_and_site()
    sites, temp = {'unknown': {},
                   'critical': {},
                   'warning': {},
                   'down': {}
                   }, 0
    # Dividing them into all alert types
    for key3 in result:
        if list(sites[key3[1].lower()].keys()).__contains__(key3[0]):
            sites[key3[1].lower()][key3[0]] += key3[2]
        else:
            sites[key3[1].lower()].update({key3[0]: key3[2]})
    # Adding the down count to critical count
    for key in sites['down']:
        sites['critical'][key] += sites['down'][key]

    # writting data of per site analysis first data table (alert type and site)
    col = 0
    row = 4
    for key in range(len(site_list)):
        col = key + 1
        if list(sites['critical'].keys()).__contains__(site_list[key][0]):
            sheet2.write(1, col, sites['critical'][site_list[key][0]], border_format)
        else:
            sheet2.write(1, col, '', border_format)
        if list(sites['warning'].keys()).__contains__(site_list[key][0]):
            sheet2.write(2, col, sites['warning'][site_list[key][0]], border_format)
        else:
            sheet2.write(2, col, '', border_format)
        if list(sites['unknown'].keys()).__contains__(site_list[key][0]):
            sheet2.write(3, col, sites['unknown'][site_list[key][0]], border_format)
        else:
            sheet2.write(3, col, '', border_format)
        sheet2.write(row, col, '=SUM(' + chr(65 + col) + '2:' + chr(65 + col) + '4)', cell_format)
        sheet2.write(0, col, site_list[key][0].replace('_', ' '), cell_format)
    sheet2.write(0, col + 1, 'Total', cell_format)
    sheet2.write(1, col + 1, '=SUM(' + chr(65) + '2:' + chr(65 + col) + '2)', cell_format)
    sheet2.write(2, col + 1, '=SUM(' + chr(65) + '3:' + chr(65 + col) + '3)', cell_format)
    sheet2.write(3, col + 1, '=SUM(' + chr(65) + '4:' + chr(65 + col) + '4)', cell_format)
    sheet2.write(4, col + 1, '=SUM(' + chr(65 + col + 1) + '2:' + chr(65 + col + 1) + '4)', cell_format)
    chart_col = col - 4
    chart = book.add_chart({'type': 'column'})
    chart.add_series(
        {
            'values': ('=Per_Site_Stats!$' + chr(66) + '$2:$' + chr(65 + col) + '$2'),
            'categories': ('=Per_Site_Stats!$' + chr(66) + '$1:$' + chr(65 + col) + '$1'),
            'name': ('=Per_Site_Stats!$' + chr(65) + '$2'),
            'data_labels': {
                'value': True,
            }
        })
    chart.add_series(
        {
            'values': ('=Per_Site_Stats!$' + chr(66) + '$3:$' + chr(65 + col) + '$3'),
            'categories': ('=Per_Site_Stats!$' + chr(66) + '$1:$' + chr(65 + col) + '$1'),
            'name': ('=Per_Site_Stats!$' + chr(65) + '$3'),
            'data_labels': {
                'value': True,
            }
        })
    chart.add_series(
        {
            'values': ('=Per_Site_Stats!$' + chr(66) + '$4:$' + chr(65 + col) + '$4'),
            'categories': ('=Per_Site_Stats!$' + chr(66) + '$1:$' + chr(65 + col) + '$1'),
            'name': ('=Per_Site_Stats!$' + chr(65) + '$4'),
            'data_labels': {
                'value': True,
            }
        }
    )
    # chart.set_legend({'none': True})
    chart.set_y_axis({'visible': False})
    chart.set_title({'name': 'Per Site Stats(Severity Wise)'})

    # Adding chart for per site analysis first data table
    chart1_row = row + 4
    sheet2.insert_chart(chr(65 + 2) + str(chart1_row), chart)

    # getting data of per site analysis second data table (site)
    result = get_alert_by_site()
    col = 0
    row += 3
    row_s = row
    sheet2.write(row - 1, col, 'Site', cell_format)
    sheet2.write(row - 1, col + 1, 'Alert Count', cell_format)

    # writting data of per site analysis second data table (site)
    for key in result:
        sheet2.write(row, col, key[0].replace('_', ' '), border_format)
        sheet2.write(row, col + 1, key[1], border_format)
        row += 1

    sheet2.write(row, col, 'Total', cell_format)
    sheet2.write(row, col + 1, '=SUM(' + chr(65 + col + 1) + str(row_s) + ':' + chr(65 + col + 1) + str(row) + ')',
                 cell_format)

    # Adding chart for per site analysis second data table
    chart = book.add_chart({'type': 'column'})

    chart.add_series(
        {
            'values': ('=Per_Site_Stats!$' + chr(65 + col + 1) + '$' + str(row_s + 1) + ':$' + chr(
                65 + col + 1) + '$' + str(row)),
            'categories': (
                    '=Per_Site_Stats!$' + chr(65 + col) + '$' + str(row_s + 1) + ':$' + chr(65 + col) + '$' + str(
                row)),
            'name': ('=Per_Site_Stats!$' + chr(65 + col + 1) + '$1'),
            'data_labels': {
                'value': True,
            }
        })
    chart.set_y_axis({'visible': False})
    chart.set_title({'name': 'Per Site Alerts'})

    chart2_row = row
    sheet2.insert_chart(chr(65 + 2) + str(row), chart)

    # fetching data of per site analysis third data table (alert type)
    result = get_alert_by_alert_type()

    # Adding the down count to critical count
    sites = {'unknown': 0,
             'critical': 0,
             'warning': 0
             }
    for key3 in result:
        if key3[0].lower().__contains__('down'):
            sites['critical'] += key3[1]
        else:
            sites[key3[0].lower()] = key3[1]

    # Writting headings of per site analysis third data table
    col = 0
    row += 3
    row_s = row
    sheet2.write(row - 1, col, 'Alert Type', cell_format)
    sheet2.write(row - 1, col + 1, 'Count', cell_format)

    # writting data of per site analysis third data table (site)
    for key in result:
        if key[0].lower().__contains__('down'):
            continue
        sheet2.write(row, col, key[0], border_format)
        sheet2.write(row, col + 1, sites[key[0].lower()], border_format)
        row += 1

    sheet2.write(row, col, 'Total', cell_format)
    sheet2.write(row, col + 1, '=SUM(' + chr(65 + col + 1) + str(row_s) + ':' + chr(65 + col + 1) + str(row) + ')',
                 cell_format)

    # Adding chart for per site analysis third data table
    chart = book.add_chart({'type': 'pie'})

    chart.add_series(
        {
            'values': ('=Per_Site_Stats!$' + chr(65 + col + 1) + '$' + str(row_s + 1) + ':$' + chr(
                65 + col + 1) + '$' + str(row)),
            'categories': (
                    '=Per_Site_Stats!$' + chr(65 + col) + '$' + str(row_s + 1) + ':$' + chr(65 + col) + '$' + str(
                row)),
            'name': ('=Per_Site_Stats!$' + chr(65 + col + 1) + '$1'),
            'data_labels': {
                'value': True,
            }
        })
    chart.set_title({'name': 'Total Site Alerts'})
    sheet2.insert_chart(chr(65 + chart_col) + str(chart1_row), chart)

    # fetching data of per site analysis fourth data table (team)
    result = get_alert_by_team()

    # Writting headings of per site analysis fourth data table (team)
    col = 0
    row += 3
    row_s = row
    sheet2.write(row - 1, col, 'Teams', cell_format)
    sheet2.write(row - 1, col + 1, 'Count', cell_format)

    # writting data of per site analysis fourth data table (team)
    for key in result:
        sheet2.write(row, col, key[0], border_format)
        sheet2.write(row, col + 1, key[1], border_format)
        row += 1

    sheet2.write(row, col, 'Total', cell_format)
    sheet2.write(row, col + 1, '=SUM(' + chr(65 + col + 1) + str(row_s) + ':' + chr(65 + col + 1) + str(row) + ')',
                 cell_format)

    # Adding chart for per site analysis fourth data table (team)
    chart = book.add_chart({'type': 'pie'})

    chart.add_series(
        {
            'values': ('=Per_Site_Stats!$' + chr(65 + col + 1) + '$' + str(row_s + 1) + ':$' + chr(
                65 + col + 1) + '$' + str(row)),
            'categories': (
                    '=Per_Site_Stats!$' + chr(65 + col) + '$' + str(row_s + 1) + ':$' + chr(65 + col) + '$' + str(
                row)),
            'name': ('=Per_Site_Stats!$' + chr(65 + col + 1) + '$1'),
            'data_labels': {
                'value': True,
            }
        })
    chart.set_y_axis({'label_position': 'none'})
    chart.set_title({'name': 'Total Team Alerts'})
    sheet2.insert_chart(chr(65 + chart_col) + str(chart2_row), chart)

    # Writting data of all other sheets.
    row = 1
    for key1 in site_list:
        sheets[key1[0]] = [book.add_worksheet(name=key1[0])]
        sheets[key1[0]][0].hide_gridlines(2)
        result = get_service_host_by_site(key1[0])
        services = {}
        hosts = {}
        service_row = 1
        host_row = 1
        sheets[key1[0]][0].write(0, 0, 'Host/Service', service_cell_format)
        for key2 in result:
            for key3 in key2:
                if not services.__contains__(key3[0]):
                    services.update({key3[0]: service_row})
                    sheets[key1[0]][0].write(0, service_row, key3[0], service_cell_format)
                    service_row += 1
                if not hosts.__contains__(key3[1]):
                    hosts.update({key3[1]: host_row})
                    sheets[key1[0]][0].write(host_row, 0, key3[1], service_cell_format)
                    host_row += 1
                sheets[key1[0]][0].write(hosts[key3[1]], services[key3[0]], key3[2])

        sheets[key1[0]][0].write(0, service_row, 'Total', service_cell_format)
        sheets[key1[0]][0].write(1, service_row,
                                 '=SUM(' + chr(65 + 1) + str(2) + ':' + chr(65 + service_row - 1) + str(2) + ')',
                                 service_cell_format)
        sheets[key1[0]][0].write(host_row, 0, 'Total', service_cell_format)
        sheets[key1[0]][0].write(host_row, 1, '=SUM(' + chr(65 + 1) + str(2) + ':' + chr(65 + 1) + str(host_row) + ')',
                                 service_cell_format)

        result = get_service_by_site(key1[0])
        sheets[key1[0]][0].write(host_row + 2, 0, 'Services', cell_format)
        sheets[key1[0]][0].write(host_row + 2, 1, 'Count', cell_format)
        sheets[key1[0]].append([host_row, service_row])
        service_row1 = host_row + 3
        for key2 in result:
            sheets[key1[0]][0].write(service_row1, 0, key2[0], border_format)
            sheets[key1[0]][0].write(service_row1, 1, key2[1], border_format)
            service_row1 += 1

        sheets[key1[0]][0].write(service_row1, 0, 'Total', cell_format)
        sheets[key1[0]][0].write(service_row1, 1, '=SUM(' + chr(65 + 1) + str(host_row + 3) + ':' + chr(65 + 1) + str(
            service_row1) + ')', cell_format)
        sheets[key1[0]].append(service_row1)
        chart = book.add_chart({'type': 'column', 'subtype': 'percent_stacked'})
        for key4 in range(1, service_row):
            chart.add_series({
                'values': ('=' + key1[0] + '!$' + chr(65 + key4) + '$2:$' + chr(65 + key4) + '$' + str(host_row)),
                'categories': ('=' + key1[0] + '!$' + chr(65) + '$2:$' + chr(65) + '$' + str(host_row)),
                'name': ('=' + key1[0] + '!$' + chr(65 + key4) + '$1'),
                'data_labels': {
                    'value': True,
                    'font': {'color': '#FFFFFF'}
                }
            })
        # chart.set_legend({'none': True})
        chart.set_y_axis({'visible': False})
        chart.set_title({'name': key1[0].replace('_', ' ') + '(Per Host Alerts)'})
        sheets[key1[0]][0].insert_chart(chr(65 + 3) + str(host_row + 3), chart)
        chart = book.add_chart({'type': 'column'})
        chart.add_series({
            'values': ('=' + key1[0] + '!$' + chr(66) + '$' + str(host_row + 4) + ':$' + chr(66) + '$' + str(
                service_row1)),
            'categories': ('=' + key1[0] + '!$' + chr(65) + '$' + str(host_row + 4) + ':$' + chr(65) + '$' + str(
                service_row1)),
            'name': ('=' + key1[0] + '!$' + chr(66) + '$' + str(host_row + 3)),
            'data_labels': {
                'value': True
            }
        })
        chart.set_y_axis({'visible': False})
        chart.set_title({'name': key1[0].replace('_', ' ') + ' (Services)'})
        sheets[key1[0]][0].insert_chart(chr(65 + 12) + str(host_row + 3), chart)
        for key in noc_dict[key1[0].lower().replace(' ', '_')]:
            sheet.write(row, 0, key[1], border_format)
            sheet.write(row, 1, key[3], border_format)
            sheet.write(row, 2, key[6], border_format)
            sheet.write(row, 3, filter_characters(key[0]), border_format)
            sheet.write(row, 4, key[4], border_format)
            sheet.write(row, 5, key[5], border_format)
            sheet.write(row, 6, key[7], date_format)
            row += 1
finally:
    # print('Bye')
    # try:
    book.close()
    # delete_data()
input(' Enter to close.')
