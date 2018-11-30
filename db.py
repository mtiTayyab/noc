import pymysql
from miscellaneous import filter_characters
host='192.168.1.220'
user = 'root'
password = 'refill'
database = 'noc_db'



def store_all_data(data):
    db = pymysql.connect(host='192.168.1.220', user='root', password='refill', db='noc_db')
    cur = db.cursor()
    for key in data:
        cur.execute(
            "INSERT INTO data(site,service,alert_type,alert,host,address,alert_date) values(%s,%s,%s,%s,%s,%s,%s);",
            [key[1], key[3], key[6], filter_characters(key[0]), key[4], key[5], key[7]])

    db.commit()
    db.close()


def get_site_by_count_desc():
    try:
        db = pymysql.connect(host=host, user=user, password=password, database=database)
        cur = db.cursor()
        query = "select site from data group by site order by count(*) desc;"
        cur.execute(query)
        result = cur.fetchall()
        return result
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



