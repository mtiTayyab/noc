import os
import xlsxwriter
from math import floor
from datetime import datetime,date
from miscellaneous import filter_characters
from db import store_all_data,delete_all_data,filter_characters
path = '.\\source\\'

date_s = date.today()
date_s = floor(date_s.day/10)*10
if date_s<10:
    week = 'Week 3'
elif 10<=date_s<20:
    week = 'Week 1'
else:
    week = 'Week 2'
name_dest = '.\\Weekly KDC Report '+week+'.xlsx'
l = os.listdir(path)
txt_name = []
for key in l:
    if key.lower().__contains__('zabbix') or key.lower().__contains__('vpn') or key.lower().__contains__('problem'):
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
        if (key1.lower().__contains__('average') or key1.lower().__contains__('high') or key1.lower().__contains__('information')) and _t_f is False:
            t = key1.split(":")[1]
            t = filter_characters(t)
            site.append(t)
            s_f = True
        if (key1.lower().__contains__('problem') or key1.lower().__contains__('critcal') or key1.lower().__contains__('warning')) and _t_f is False:
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
    if(s_f is False):
        site.append('')



final_data = []

for key in range(len(final)):
    temp={'final':final[key]}
    temp.update({'type':_type[key]})
    temp.update({'site':site[key]})
    temp.update({'from':_from[key]})
    temp.update({'date':date[key]})
    if _type[key]:
        temp.update({'subject': subject[key].split(_type[key])[1]})
    else:
        temp.update({'subject': subject[key]})
    if final[key].lower().__contains__('zabbix'):
        temp.update({'category':'Zabbix'})

    elif final[key].lower().__contains__('vpn'):
        temp.update({'category':'VPN'})
        temp.update({'subject': subject[key]})
    final_data.append(temp)

delete = []
for key in final_data:
    if key['subject'].lower().__contains__('ok'):
        delete.append(key)

for key in delete:
    if final_data.__contains__(key):
        final_data.remove(key)





site_f = ['Glo Nigeria','MTN Afganistan','MTN Benin','MTN IvoryCoast','MTN Nigeria','MTN Rwanda','MTN SouthSudan','MTN Sudan','MTN Yemen','MTN Zambia','NewCo Bahamas','Swazi Mobile','KDC Data Centre']

otrs = [['glo nigeria','ng glo'],'mtn afganistan','mtn benin','mtn ivorycoast','mtn nigeria','mtn rwanda','mtn southsudan',['mtn sudan','mtnsudan'],'mtn yemen','mtn zambia','newco bahamas','swazi mobile','sds']

for key in final_data:
    for key1 in range(len(otrs)):
        if type(otrs[key1]) is list:
            for key2 in range(len(otrs[key1])):
                if key['subject'].lower().__contains__(otrs[key1][key2]):
                    key['site']=site_f[key1]
                    break
        else:
            if key['subject'].lower().__contains__(otrs[key1]):
                key['site']=site_f[key1]
                break

book = xlsxwriter.Workbook(name_dest)
try:
    store_all_data(final_data)
    sheet = book.add_worksheet()
    sheet.hide_gridlines(2)
    date_format = book.add_format({'num_format': 'dd/mmm/yy hh:mm:ss'})
    date_format.set_border(style=1)
    border_format = book.add_format()
    border_format.set_border(style=1)
    cell_format = book.add_format({'bg_color':'#000000','font_color':'#FFFFFF'})

    format = {
        'date':date_format,
        'border':border_format,
        'heading': cell_format
        }

    sheet.write(0,0,'Site/Type', format['heading'])
    sheet.write(0,1,'Severity',format['heading'])
    sheet.write(0,2,'Subject',format['heading'])
    sheet.write(0,3,'From',format['heading'])
    sheet.write(0,4,'Date',format['heading'])
    row=1
    for key in final_data:
        sheet.write(row,0,key['site'],format['border'])
        sheet.write(row,1,key['type'],format['border'])
        sheet.write(row,2,filter_characters(key['subject']),format['border'])
        sheet.write(row,3,key['from'],format['border'])
        sheet.write(row,4,key['date'],format['date'])
        row+=1


finally:
    delete_all_data()
    book.close()