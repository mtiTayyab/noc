import pymysql

host = '192.168.1.220'
user = 'root'
password = 'refill'
database = 'noc_db'


def filter_characters(string):
    string = string.replace(",", "")
    string = string.replace("_", " ")
    string = string.replace("-", " ")
    string = string.replace("  ", " ")
    string = string.replace("   ", " ")
    return string


def store_all_data(data):
    global db
    try:
        db = pymysql.connect(host=host, user=user, password=password, db=database)
        cur = db.cursor()
        for key in data:
            if not key['first response']:
                key['firstresponse'] = None
            if not key['close time']:
                # print(key)
                key['close time'] = None
                key['solutiontime'] = None
            cur.execute(
                "INSERT INTO monthly_kpi_data(age, title, created, close_time, state, priority, service, sla, type, owner, first_response, solution_time, customer_origin, severity, queue) values (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s);",
                [key['age'], key['title'], key['created'], key['close time'], key['state'], key['priority'],
                 key['service'], key['sla'], key['type'], key['agent/owner'], key['first response'],
                 key['solution time'],
                 key['customer origin'], key['severity'], key['queue']])
        db.commit()
        return 'successful'
    except pymysql.InterfaceError:
        raise pymysql.InterfaceError
    finally:
        db.close()


def delete_all_data():
    global db
    try:
        db = pymysql.connect(host=host, user=user, password=password, db=database)
        cur = db.cursor()
        cur.execute('delete from monthly_kpi_data;')
        return 'successful'
    except pymysql.InterfaceError:
        raise pymysql.InterfaceError
    finally:
        db.close()


def get_alerts_by_sla():
    global db
    try:
        db = pymysql.connect(host=host, user=user, password=password, db=database)
        cur = db.cursor()
        cur.execute('select distinct(sla),count(*) as Total from monthly_kpi_data group by sla order by Total desc;')
        return cur.fetchall()
    except pymysql.InterfaceError:
        raise pymysql.InterfaceError
    finally:
        db.close()


def get_alerts_by_state():
    global db
    try:
        db = pymysql.connect(host=host, user=user, password=password, db=database)
        cur = db.cursor()
        cur.execute(
            'select distinct(state),count(*) as Total from monthly_kpi_data group by state order by Total desc;')
        return cur.fetchall()
    except pymysql.InterfaceError:
        raise pymysql.InterfaceError
    finally:
        db.close()


def get_alerts_by_priority():
    global db
    try:
        db = pymysql.connect(host=host, user=user, password=password, db=database)
        cur = db.cursor()
        cur.execute(
            'select distinct(priority),count(*) as Total from monthly_kpi_data group by priority order by Total desc;')
        return cur.fetchall()
    except pymysql.InterfaceError:
        raise pymysql.InterfaceError
    finally:
        db.close()


def get_alerts_by_type():
    global db
    try:
        db = pymysql.connect(host=host, user=user, password=password, db=database)
        cur = db.cursor()
        cur.execute('select distinct(type),count(*) as Total from monthly_kpi_data group by type order by Total desc;')
        return cur.fetchall()
    except pymysql.InterfaceError:
        raise pymysql.InterfaceError
    finally:
        db.close()


def get_alerts_by_severity():
    global db
    try:
        db = pymysql.connect(host=host, user=user, password=password, db=database)
        cur = db.cursor()
        cur.execute(
            'select distinct(severity),count(*) as Total from monthly_kpi_data group by severity order by Total desc;')
        return cur.fetchall()
    except pymysql.InterfaceError:
        raise pymysql.InterfaceError
    finally:
        db.close()


def get_alerts_by_owner():
    global db
    try:
        db = pymysql.connect(host=host, user=user, password=password, db=database)
        cur = db.cursor()
        cur.execute(
            'select distinct(owner),count(*) as Total from monthly_kpi_data group by owner order by Total desc;')
        return cur.fetchall()
    except pymysql.InterfaceError:
        raise pymysql.InterfaceError
    finally:
        db.close()
