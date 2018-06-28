import sqlite3
e =0
from xlrd import open_workbook
wb = open_workbook('sample.xlsx')
for s in wb.sheets():
    #print 'Sheet:',s.name
    values = []
    for row in range(s.nrows):
        col_value = []
        for col in range(s.ncols):
            value  = (s.cell(row,col).value)
            try : value = str(int(value))
            except : pass
            col_value.append(value)
        values.append(col_value)
print("data extracted from exel......",values)
lk = "http:"
for i in range(0,len(values)):
    for j in range(0,len(values)):
        if lk in values[i][j]:
            url = values[i][j]
print("the link is......",url)
######################################
from urllib.request import urlopen
from bs4 import BeautifulSoup
f=open("htmlfile.html","w")
file_handle=urlopen(url)
store=file_handle.read()
data=BeautifulSoup(store, "html.parser")
display=data.get_text()
for line in display.splitlines():
    f.write(line)
final = display.split()
##print(final)
f.close()
######################################
dem = []
for i in range(1,len(values)):
    for j in range(2,len(values)):
        dem.append(values[i][j])
##print(dem)
print("the words are......",dem)
game = list(set(final) & set(dem))
##print(game)
a =[]
for i in range(0,len(final)):
    e = e+1
for i in range(0,len(game)):
    r = game[i].__len__()
    f = (r/len(game))*10
    a.append(f)
##print(final)
##print(a)
dic = {key:value for key, value in zip(game, a)}
print("the final diotionary is.....",dic)
###################################################
###Database connectivity comes here################
con = sqlite3.connect("data.db")
con.execute("CREATE TABLE master(col1, col2)")
for i in range(len(game)):
    con.execute("INSERT INTO master (col1, col2)"
              " VALUES (?, ?)",
              (game[i], a[i]))
from xlsxwriter.workbook import Workbook
workbook = Workbook('output.xlsx')
worksheet = workbook.add_worksheet()
cursor = con.execute("SELECT col1, col2 from master")
for i, row in enumerate(cursor):
    worksheet.write(i, 0, row[0])
    worksheet.write(i, 1, row[1])
chart = workbook.add_chart({'type': 'pie'})
chart.add_series({
    'keywords': '=Sheet1!$A$1:$A$6',
    'values':     '=Sheet1!$B$1:$B$6',
    'points': [],})
chart.set_title({'name': 'Word density used in a live webpage'})
worksheet.insert_chart('D4', chart, {'x_offset': 25, 'y_offset': 10})
workbook.close()
con.commit()
con.close()
