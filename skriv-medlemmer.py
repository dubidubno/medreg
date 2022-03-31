import openpyxl
from pathlib import Path
import sys
import mysql.connector
from mysql.connector import errorcode
import yaml
import io

try:
    configfile = io.open("config.yaml",'r',encoding='utf8')
except:
    sys.exit("failed to open configfile")
try:
    config = yaml.safe_load(configfile)
except:
    sys.exit("failed to read or parse configfile")

try:
    mydb = mysql.connector.connect(
        host=config['database-host'],
        user=config['database-user'],
        password=config['database-password'],
        database=config['database-name'],
        # ssl_disabled = True,
    )
except mysql.connector.Error as err:
    if err.errno == errorcode.ER_ACCESS_DENIED_ERROR:
        print("Something is wrong with your user name or password")
    elif err.errno == errorcode.ER_BAD_DB_ERROR:
        print("Database does not exist")
    else:
        print(err)
    sys.exit
else:
    print("DB connection ok")
mycursor = mydb.cursor(dictionary=True)

xlsx_file = Path('.', 'ny.xlsx');
wb_obj = openpyxl.Workbook()
sheet = wb_obj.active
sheet.append(['Ansattnr', 'Fanenr', 'Navn', 'Mobil', 'Epost', 'Adresse'])

mycursor.execute("SELECT * FROM medlemmer ORDER BY ansattnr")
myresult = mycursor.fetchall()

for x in myresult:
    sheet.append((x['ansattnr'], x['fanenr'], x['navn'], x['mobil'], x['epost'], x['adresse']))

wb_obj.save(filename = xlsx_file)
print("The end.")


