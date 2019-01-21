import xlsxwriter
import openpyxl
import csv



name_dest = 'Monthly KPI Repo rt.xlsx'
csv_reader, file_flag = 0,0
while file_flag==0:
    try:

        # file = input('Enter the path or name of you file: ')
        file = 'NOC_Created_2018-10-01_16-28.csv'
        data_file = open(file=file)
        csv_reader = csv.reader(data_file,delimiter=',')
        file_flag=1
    except FileNotFoundError:
        print('File Not Found.')
    except OSError:
        print('Enter a Valid Path or Name.')


data={}
flag = 0
header = []
for key in csv_reader:
    if flag ==0:
        for key1 in key:
            header.append(key1)
            data.update({key1:[]})
        flag=1
        continue
    for key1 in range(len(key)):
            data[header[key1]].append(key[key1])

ignore = ['Number','CustomerID','FirstResponseTimeWorkingTime','FirstResponseTime','Impact','Review Required','Decision Result','Decision Date','Due Date']
for key in ignore:
    data.__delitem__(key)


for key in range(len(data['Ticket#'])):
    data['Queue'][key]=data['Queue'][key].split(':')[-1]


book = xlsxwriter.Workbook(name_dest)

sheet = book.add_worksheet(name='Data')


col =0
header = {}
for key in list(data.keys()):
    sheet.write(0,col,key)
    header.update({key:col})
    col+=1

row = 1
for key in range(len(data['Ticket#'])):
    for key1 in list(header.keys()):
        try:
            sheet.write(row,header[key1],int(data[key1][key]))
        except ValueError:
            sheet.write(row,header[key1],data[key1][key])
    print(key)
    row+=1



book.close()