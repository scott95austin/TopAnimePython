import pandas as pd
import numpy as np
import xlsxwriter
import requests
from urllib.request import urlopen
from bs4 import BeautifulSoup
from tabulate import tabulate

#pulls web page data
req = requests.get('https://myanimelist.net/topanime.php?type=bypopularity')
soup = BeautifulSoup(req.text, 'html.parser')

#retrieves table
table = soup.find('table', {'class': 'top-ranking-table'})
#retrieves ranking-list tr
rows = table.find_all('tr', attrs={'class': 'ranking-list'})

#gathers ranks
rankDF = []
for rank in rows:
    cells = rank.find_all('td', attrs={'class': 'rank ac'})
    ranks = cells[0].find('span').text
    rankDF.append((ranks))
    print(ranks)
#gathers titles
titleDF = []
for title in rows:
    cells = title.find_all('div', attrs={'class': 'di-ib clearfix'})
    titles = cells[0].find('a').text
    titleDF.append((titles))
    print(titles)

#gathers scores
scoreDF = []
for score in rows:
    cells = score.find_all('td', attrs={'class': 'score ac fs14'})
    scores = cells[0].find('span').text
    scoreDF.append((scores))
    print(scores)

#pandas dataframe
df_R = pd.DataFrame(rankDF, columns=['Rank'])
df_T = pd.DataFrame(titleDF, columns=['Title'])
df_S = pd.DataFrame(scoreDF, columns=['Score'])
#pandas to excel writer
writer = pd.ExcelWriter('test.xlsx', engine='xlsxwriter')
df_R.to_excel(writer, sheet_name='Sheet1', index=False)
df_T.to_excel(writer, sheet_name='Sheet1', startcol=1, index=False)
df_S.to_excel(writer, sheet_name='Sheet1', startcol=2, index=False)

#Assign Worksheet and Workbook
workbook = writer.book
worksheet = writer.sheets['Sheet1']

#Format the Workbook and Worksheet
scoreFormat = workbook.add_format({'num_format': '#,##0.00'})

#Format the Columns
worksheet.set_column('A:A', 10)
worksheet.set_column('B:B', 70)
worksheet.set_column('C:C', 10, scoreFormat)

#outputs excel file of previously specified name
writer.save()
