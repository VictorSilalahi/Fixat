import sqlite3 as lite

con = None
cur = None

def connection():
    global con
    global cur
    
    con = lite.connect("data/fixat.db")
    cur = con.cursor()
    return cur