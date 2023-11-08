import sqlite3 as sq
from bot import bot
from datetime import date
import datetime


def sql_start():
    global base, cur
    base = sq.connect('clients.db')
    cur = base.cursor()
    if base:
        print("Database connected successfully")
    base.execute('CREATE TABLE IF NOT EXISTS data_users(user_id INTEGER, contact INTEGER, date TEXT)')
    base.commit()


async def add_new_user(user_id, contact, date):
    with base:
        cur.execute('INSERT INTO data_users VALUES (?, ?, ?)', (user_id, contact, date))


async def payment_succesfull(user_id, contact, date):
    with base:
        cur.execute('INSERT INTO data_users VALUES (?)', (user_id, contact, date))


async def check_subscription(user_id):
    with base:
        cur.execute('SELECT date FROM data_users WHERE user_id = ?', (user_id, ))

    if  date.today() > datetime.datetime.strptime(cur.fetchone()[0], '%Y-%m-%d').date():
        return False
    else: 
        return True


async def check_user(user_id):
    with base:
        cur.execute('SELECT EXISTS(SELECT * FROM data_users where user_id = ?)', (user_id, ))
    return cur.fetchone()[0]


async def add_subscription(user_id, date):
    with base:
        cur.execute('UPDATE data_users SET date=? WHERE user_id = ?', (date, user_id ))