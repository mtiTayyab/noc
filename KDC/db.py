import pymysql

host = '192.168.1.220'
user = 'root'
database = 'noc_db'
password = 'refill'


def filter_characters(string):
    if not (string.lower().__contains__('am') or string.lower().__contains__('pm')):
        string = string.replace(":", "")

    string = string.replace(",", "")
    string = string.replace("_", " ")
    string = string.replace("-", " ")
    string = string.replace("</EOM>", "")
    string = string.replace(chr(10),"")
    if(string[-1]==" "):
        string= string[:-1]
    if(string[0]==" "):
        string= string[1:]
    return string

def store_all_data(data):
    db = pymysql.connect(host=host,user=user,password=password,db=database)
    cur = db.cursor()
    for key in data:
        if key['subject'].lower().__contains__('vpn'):
            key['subject'] = 'VPN'+key['subject'].split('VPN')[1]
        cur.execute("INSERT INTO kdc_data(subject,site,_from,alert_date) values (%s,%s,%s,%s);" ,[filter_characters(key['subject']),key['site'],key['from'],key['date']])

    db.commit()
    db.close()
    return 'successful'

def delete_all_data():
    db = pymysql.connect(host=host,user=user,password=password,db=database)
    cur = db.cursor()
    cur.execute('delete from kdc_data;')
    db.commit()
    db.close()
    return 'successful'
