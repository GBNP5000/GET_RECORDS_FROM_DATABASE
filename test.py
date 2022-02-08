import clipboard as cb
import pyautogui as pag
import pyodbc

txt = """
<br><p>DON'T KNOW HOW TO OPEN THIS https://linksly.co/Cr9vQ<a href="https://youtu.be/MgwGAWTQH5c"> CLICK HERE -->  https://youtu.be/MgwGAWTQH5c</a></p><br><p>The Linksly.co referral program is a great way to spread the word of this great service and to earn even more money with your short links! Refer friends and receive 10% of their earnings for life!</p><p> <a href="https://linksly.co/ref/117357242480811365327">  Click Here --> https://linksly.co/ref/117357242480811365327 </a> </p> <br> <br><a href="https://linksly.co/Cr9vQ">https://linksly.co/Cr9vQ</a><br><br>Of the 7,16,662 workers that returned as a result of the pandemic, over 3 lakh came from the UAE, while 1.37 lakh returned from Saudi Arabia
"""


def GETREC(dbn, sql, mypwd):
    # set up some constants
    # MDB = 'c:/path/to/my.mdb'; DRV = '{Microsoft Access Driver (*.mdb)}'; PWD = 'pw'
    MDB = dbn
    DRV = '{Microsoft Access Driver (*.mdb)}'

    if not len(mypwd) == 0:
        # connect to db
        con = pyodbc.connect('DRIVER={};DBQ={};PWD={}'.format(DRV, MDB, mypwd))
    else:
        con = pyodbc.connect('DRIVER={};DBQ={}'.format(DRV, MDB))
    cur = con.cursor()
    # print(sql)
    # print("\n\n\n")
    cur.execute(sql)
    data = cur.fetchall()
    cur.close()
    con.close()
    return data


dbn = "E:/IMGS/RSSFEED.mdb"
frmr = 2
tor = 2
sql = "SELECT RSFED.TITLE, RSFED.CBDY, RSFED.COD FROM RSFED GROUP BY RSFED.TITLE, RSFED.CBDY, RSFED.COD HAVING (((RSFED.COD) Between " + str(
    frmr) + " And " + str(tor) + ")) ORDER BY RSFED.COD ASC"

recs = GETREC(dbn, sql, "")
print(recs[0][1])
