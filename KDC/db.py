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


def get_site_by_count_desc():
    try:
        db = pymysql.connect(host=host, user=user, password=password, database=database)
        cur = db.cursor()
        query = "select site from kdc_data group by site order by count(*) desc;"
        cur.execute(query)
        return cur.fetchall()
    except pymysql.InterfaceError:
        raise pymysql.InterfaceError
def store_all_data(data):
    try:

        db = pymysql.connect(host=host,user=user,password=password,db=database)
        cur = db.cursor()
        for key in data:
            if key['subject'].lower().__contains__('vpn'):
                key['subject'] = 'VPN'+key['subject'].split('VPN')[1]
            cur.execute("INSERT INTO kdc_data(subject,site,severity,_from,alert_date,category) values (%s,%s,%s,%s,%s,%s);" ,[filter_characters(key['subject']),key['site'],key['type'],key['from'],key['date'],key['category']])

        db.commit()
        db.close()
        return 'successful'
    except pymysql.InterfaceError:
        raise pymysql.InterfaceError
def get_all_zabix_alerts():
    try:
        db = pymysql.connect(host=host, user=user, password=password, db=database)
        cur = db.cursor()
        cur.execute("select site,severity,subject,_from,alert_date from kdc_data where category='zabbix';")
        result = cur.fetchall()
        db.close()
        return result
    except pymysql.InterfaceError:
        raise pymysql.InterfaceError


def get_zabix_alerts_by_issue():
    try:
        db = pymysql.connect(host=host, user=user, password=password, db=database)
        cur = db.cursor()
        cur.execute(" select subject,site,count(*) from kdc_data where category='zabbix' group by subject,site;")
        result = cur.fetchall()
        db.close()
        return result
    except pymysql.InterfaceError:
        raise pymysql.InterfaceError


def get_vpn_alerts_by_site():
    try:
        db = pymysql.connect(host=host, user=user, password=password, db=database)
        cur = db.cursor()
        cur.execute("select site,count(*) from kdc_data where category='vpn' group by site;")
        result = cur.fetchall()
        db.close()
        return result
    except pymysql.InterfaceError:
        raise pymysql.InterfaceError
def get_all_vpn_alerts():
    try:
        db = pymysql.connect(host=host, user=user, password=password, db=database)
        cur = db.cursor()
        cur.execute("select site,subject,_from,alert_date from kdc_data where category='vpn';")
        result = cur.fetchall()
        db.close()
        return result
    except pymysql.InterfaceError:
        raise pymysql.InterfaceError
def delete_all_data():
    try:
        db = pymysql.connect(host=host,user=user,password=password,db=database)
        cur = db.cursor()
        cur.execute('delete from kdc_data;')
        db.commit()
        db.close()
        return 'successful'
    except pymysql.InterfaceError:
        raise pymysql.InterfaceError