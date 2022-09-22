import mysql.connector
import openpyxl
import settings

db = mysql.connector.connect(
    host=settings.HOST,
    database=settings.DATABASE,
    user=settings.USER,
    password=settings.PASSWORD
)

query = input("enter the query you want to run: ")
filepath = input("where do you want to save the file? ")

cursor = db.cursor()
cursor.execute(query)
result = cursor.fetchall()
columns = cursor.column_names

wb = openpyxl.Workbook()  # creates a workbook object.
ws = wb.active  # creates a worksheet object.
ws.append(columns)
for row in result:
    ws.append(row)  # adds values to cells, each list is a new row.
wb.save(filepath)  # save to Excel file.
