import sqlite3
import requests
from parse import update_info
from prettytable import PrettyTable
from lxml import html

conn = sqlite3.connect('base.db')
curs = conn.cursor()

session = requests.Session()

def compare(tickers):
    
    columns = [
        'Current Ratio',
        'Cash Ratio',
        'Accounts Payable Turnover Ratio',
        'Total Assets Turnover',
        'Working Capital',
        'Revenue per Employee',
        'Profit per Employee',
        'Gross Margin',
        'Net Income Margin',
        'Debt Ratio',
        'Return On Sales (ROS)',
        'Return On Assets (ROA)',
        'Earnings per Share Ratio (EPS)',
        'PE Ratio (TTM)'
    ]
    table = PrettyTable()
    column_list = [['Показатель']+columns]
    #table.add_column('Ratio',columns)
    mean = [0]*(len(columns)-1)
    pretty_value = lambda val: [f'{val[i]:.3f}' for i in range(len(columns)-1)]+[val[len(columns)-1] ]
    
    for ticker in tickers:
        
        update_info(ticker)
        fetch = list(curs.execute(f'SELECT * FROM [{ticker}_ratio]').fetchone())
        tree = html.fromstring(session.get('https://stockanalysis.com/stocks/'+ticker).content)
        pe = tree.xpath('//*[@id="main"]/div/div[2]/div[2]/table[1]/tbody/tr[6]/td[2]/text()')[0].replace(',','')
        if tickers.index(ticker)!=0:
            mean=[mean[i]+fetch[i] for i in range(len(mean))]
        fetch.append(pe)
        column_list.append([ticker]+fetch)
    mean = list(map(lambda m: m/(len(tickers)-1), mean))
    mean.append('-')
    column_list.append(['Среднее']+mean)

    column_list.append(['Расхождение']+[mean[i-1]-column_list[1][i] for i in range(1,len(mean))]+['-'])

    return column_list

