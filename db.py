import pymysql
from miscellaneous import filter_characters
host='192.168.1.220'
user = 'root'
password = 'refill'
database = 'noc_db'



def store_all_data(data):
    db = pymysql.connect(host='192.168.1.220', user='root', password='refill', db='noc_db')
    cur = db.cursor()

    lahore = ['mtn_yemen', 'mtn_afghanistan', 'mtn_syria', 'glo_nigeria', 'starlink_qatar', 'newco_bahamas' , 'mtn_sudan']
    kolkata = ['gosoft_thailand', 'dna_finland', 'se_bank_system']
    accra = ['mtn_congo','mtn_ghana', 'mtn_south_sudan','mtn_benin', 'glo_benin', 'mtn_zambia', 'mtn_ivory_coast', 'mtn_bissau', 'glo_ghana', 'swazi_mobile']
    team = ''
    for key in data:
        if lahore.__contains__(key[1].lower()):
            team = 'OPS Lahore'
        if accra.__contains__(key[1].lower()):
            team = 'OPS Accra'
        if kolkata.__contains__(key[1].lower()):
            team = 'OPS Kolkata'
        cur.execute(
            "INSERT INTO data(site,service,alert_type,alert,host,address,alert_date,team) values(%s,%s,%s,%s,%s,%s,%s,%s);",
            [key[1], key[3], key[6], filter_characters(key[0]), key[4], key[5], key[7],team])
    db.commit()
    db.close()


def get_site_by_count_desc():
    try:
        db = pymysql.connect(host=host, user=user, password=password, database=database)
        cur = db.cursor()
        query = "select site from data group by site order by count(*) desc;"
        cur.execute(query)
        return cur.fetchall()
    except pymysql.InterfaceError:
        raise pymysql.InterfaceError


def get_service_host_by_site(site):
    try:
        db = pymysql.connect(host=host, user=user, password=password, database=database)
        cur = db.cursor()
        query = "select distinct(host) from data where site=%s;"
        cur.execute(query,[site])
        hosts = cur.fetchall()
        result = []
        for key in hosts:
            query = "select service, host, count(*) from data where site=%s and host=%s group by service,host;"
            cur.execute(query,[site,key[0]])
            result.append(cur.fetchall())
        return result
    except pymysql.InterfaceError:
        raise pymysql.InterfaceError


def get_service_by_site(site):
    try:
        db = pymysql.connect(host=host, user=user, password=password, database=database)
        cur = db.cursor()
        query = "select service,count(*) from data where site=%s group by service order by count(*) desc;"
        cur.execute(query,[site])
        return cur.fetchall()
    except pymysql.InterfaceError:
        raise pymysql.InterfaceError


def get_alerts_by_type_and_site():
    try:
        db = pymysql.connect(host=host, user=user, password=password, database=database)
        cur = db.cursor()
        query = "select site,alert_type,count(*) from data group by site,alert_type;"
        cur.execute(query)
        return cur.fetchall()
    except pymysql.InterfaceError:
        raise pymysql.InterfaceError

def get_alert_by_site():
    try:
        db = pymysql.connect(host=host, user=user, password=password, database=database)
        cur = db.cursor()
        query = "select site,count(*) from data group by site order by count(*) desc;"
        cur.execute(query)
        return cur.fetchall()
    except pymysql.InterfaceError:
        raise pymysql.InterfaceError

def get_alert_by_team():
    try:
        db = pymysql.connect(host=host, user=user, password=password, database=database)
        cur = db.cursor()
        query = "select team,count(*) from data group by team order by count(*) desc;"
        cur.execute(query)
        return cur.fetchall()
    except pymysql.InterfaceError:
        raise pymysql.InterfaceError


def get_alert_by_alert_type():
    try:
        db = pymysql.connect(host=host, user=user, password=password, database=database)
        cur = db.cursor()
        query = "select alert_type,count(*) from data group by alert_type order by count(*) desc;"
        cur.execute(query)
        return cur.fetchall()
    except pymysql.InterfaceError:
        raise pymysql.InterfaceError


def delete_data():
    try:
        db = pymysql.connect(host=host, user=user, password=password, database=database)
        cur = db.cursor()
        query = "Delete from data;"
        cur.execute(query)
    except pymysql.InterfaceError:
        raise pymysql.InterfaceError
