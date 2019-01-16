import os
import xlsxwriter
from datetime import datetime
from miscellaneous import filter_characters
path = '.\\source\\'
name_dest = '.\\KDC_Final_Data.xlsx'
l = os.listdir(path)
txt_name = []
for key in l:
    if (key.lower().__contains__('zabbix') or key.lower().__contains__('vpn')):
        txt_name.append(key)
    else:
        print(key)
    # else:
    #     print()
        # os.remove(path+key)

final = []
flag = 0
for key in txt_name:
    data_f = open(path + key, encoding='UTF-8')
    data = data_f.read()
    if (data.lower().__contains__('subject:') and data.lower().__contains__('from:') and data.lower().__contains__('date:')):
        final.append(key)
    else:
        flag=1
        os.remove(path + key)
    if not data_f._checkClosed():
        data_f.close()
        if(flag==1):
            os.remove(path+key)
            flag=0

subject = []
_type = []
_from = []
date = []
t = ''
for key in range(len(final)):
    data_f = open(path + final[key], encoding='UTF-8')
    sub_f = False
    _f_f = False
    _t_f = False
    d_f = False
    for key1 in data_f.readlines():
        if (key1.lower().__contains__('subject:') and sub_f is False):
            t = key1.replace('Subject:', '')
            t = filter_characters(t)
            subject.append(t)
            sub_f = True
        if (key1.lower().__contains__('average:') or key1.lower().__contains__('critical:') or key1.lower().__contains__('warning:')) and _t_f is False:
            t = key1.split(":")[1]
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
    if(_t_f is False):
        _type.append('')
    if(_f_f is False):
        _from.append('')
    if(d_f is False):
        date.append('')


final_data = {
    'final':[],
    'subject':[],
    'type':[],
    'from':[],
    'date':[],
    'category':[]
}

for key in range(len(final)):
    final_data['final'].append(final[key])
    final_data['subject'].append(subject[key])
    final_data['type'].append(_type[key])
    final_data['from'].append(_from[key])
    final_data['date'].append(date[key])
    if final[key].lower().__contains__('zabbix'):
        final_data['category'].append('Zabbix')
    elif final[key].lower().__contains__('vpn'):
        final_data['category'].append('VPN')


# otrs = ['glo nigeria','benin','yemen','swazi mobile','mtn sudan','mtnsudan','ng glo','nigeria','newco bahamas','ivorycoast','zambia','afghanistan','suthsudan','rwanda','']
# for key in range(len(final_data['final'])):
#     flag=0
#     for key1 in otrs:
#         temp = final_data['subject'][key].lower()
#         if (temp.__contains__(key1)):
#             final_data['site'][key] = key1
#             flag = 1
#             break
    # if (flag == 0):
    #     for key1 in otrs:
    #         temp = final_data['from'][key].lower()
    #         if (temp.__contains__(key1)):
    #             key[1] = key1
    #             flag = 1
    #             break
    #     if(flag == 0):
    #         for key1 in otrs:
    #             temp = key[0].lower()
    #             if (temp.__contains__(key1)):
    #                 key[1] = key1
    #                 flag = 1
    #                 break
    #     else:
    #         flag = 0
    # else:
    #   flag = 0

site_f = [' Glo Nigeria',' MTN Afganistan','MTN Benin','MTN IvoryCoast','MTN Nigeria','MTN Rwanda','MTN SouthSudan','MTN Sudan','MTN Yemen','MTN Zambia','NewCo Bahamas','Swazi Mobile']

otrs = [['glo nigeria','ng glo'],'mtn afganistan','mtn benin','mtn ivorycoast','mtn nigeria','mtn rwanda','mtn southsudan',['mtn sudan','mtnsudan'],'mtn yemen','mtn zambia','newco bahamas','swazi mobile']

# otrs = ['glo nigeria','benin','yemen','swazi mobile','mtn sudan','mtnsudan','ng glo','nigeria','newco bahamas','ivorycoast','zambia','afghanistan','suthsudan','rwanda','']

# site_r = ['ye-mtn', 'af-mtn', ['sy-mtn', 'mtn-sy'], 'glo-ng', 'starlink', 'newco', 'mtn-c', 'gosoft', 'dna-finland', 'atm',
#  ['bjmtn', 'mtn-benin'], 'gc-mtn', ['mtnliberia','mtn-lib','mtn lib'], 'gh-mtn', ['mtnsouthsudan', 'mtn-southsudan'], 'globenin'
#  , ['mtnzambia','mtn zambia','zm'], ['mtnci','ci@mtn'], ['mtnbissau','mtn-gb'], ['gloghana', 'glo-gh'], 'swazimobile', 'sudan-mtn']
for key in range(len(final_data['subject'])):
    for key1 in range(len(otrs)):
        if type(otrs[key1]) is list:
            for key2 in range(len(otrs[key1])):
                if final_data['subject'][key].lower().__contains__(otrs[key1][key2]):
                    final_data['type'][key]=site_f[key1]
                    break
        else:
            if final_data['subject'][key].lower().__contains__(otrs[key1]):
                final_data['type'][key]=site_f[key1]
                break



book = xlsxwriter.Workbook(name_dest)

sheet = book.add_worksheet()
sheet.hide_gridlines(2)
date_format = book.add_format({'num_format': 'dd/mmm/yy hh:mm:ss'})
date_format.set_border(style=1)
border_format = book.add_format()
border_format.set_border(style=1)
cell_format = book.add_format({'bg_color':'#000000','font_color':'#FFFFFF'})

format = {
    'date':date_format,
    'final':border_format,
    'subject':border_format,
    'type':border_format,
    'from':border_format,
    'category':border_format,
    'heading': cell_format
    }


column = 0
for key in final_data:
    row = 1
    sheet.write(0, column, key,format['heading'])
    for key1 in range(len(final_data[key])):
        sheet.write(row,column,final_data[key][key1],format[key])
        row+=1

    column+=1
book.close()