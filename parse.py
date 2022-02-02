import requests
from lxml import html
import sqlite3
import json
conn = sqlite3.connect('base.db')
curs = conn.cursor()

session = requests.Session()
ratio_columns = [
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
    'Earnings per Share Ratio (EPS)'
]
search_stock_api = lambda query, s: json.loads(session.get('https://api.stockanalysis.com/search?q='+query).text)[0][s]

def create_table_and_write_to_it(table_name, columns, rows):
    curs.execute(f'DROP TABLE IF EXISTS [{table_name}]')
    pretty_columns_str = "\""+"\" REAL, \"".join(columns)+"\" REAL"
    curs.execute(f'CREATE TABLE [{table_name}]({pretty_columns_str})')
    for row in rows:
        pretty_row_str = ", ".join([item.replace(",","").replace("%","") for item in row])
        curs.execute(f'INSERT INTO [{table_name}] VALUES({pretty_row_str})')
    conn.commit()

def parse_main_info(link):
    table = html.fromstring(session.get(link).content).xpath('//*[@id="financial-table"]/tbody')[0]
    columns = ['Year']+table.xpath('tr/td[1]//text()')
    
    pretty_num = lambda s: 'NULL' if s=='-' else s.replace('n/a','NULL')
    rows = [
        table.xpath('//*[@id="financial-table"]/thead/tr/th[2]/text()')+
        [pretty_num(row.xpath('td[2]/span')[0].get('title',default='NULL') if len(row.xpath('td[2]/span'))>0 else 'NULL') for row in table.xpath('tr')],
        table.xpath('//*[@id="financial-table"]/thead/tr/th[3]/text()')+
        [pretty_num(row.xpath('td[3]/span')[0].get('title',default='NULL') if len(row.xpath('td[3]/span'))>0 else 'NULL') for row in table.xpath('tr')]
        ]
    return columns, rows

def parse_and_save_other_info(quote):
    profile_link = f'https://stockanalysis.com/stocks/{quote}/company/'
    statistics_link = f'https://stockanalysis.com/stocks/{quote}/statistics/'
    curs.execute(f'DROP TABLE IF EXISTS [{quote}_other_info]')
    curs.execute(f'CREATE TABLE [{quote}_other_info]("Industry" TEXT, "Sector" TEXT, "Employees" INTEGER, "Shares Outstanding" REAL, "Market Cap" REAL, "Enterprise Value" REAL)')
    
    profile_text_list = html.fromstring(session.get(profile_link).text).xpath('//*[@id="main"]/div/div[2]/div[2]/div[1]/table/tbody//text()')
    industry = profile_text_list[profile_text_list.index('Industry')+1]
    sector = profile_text_list[profile_text_list.index('Sector')+1]
    employees = profile_text_list[profile_text_list.index('Employees')+1].replace(",","")
    
    statistics_col =  html.fromstring(session.get(statistics_link).content).xpath('//*[@id="main"]//table')
    
    shares_outstranding = statistics_col[2].xpath('tbody/tr[1]/td[2]')[0].get('title',default='NULL').replace(",","")
    market_cap = statistics_col[0].xpath('tbody/tr[1]/td[2]')[0].get('title',default='NULL').replace(",","")
    enterprise_value = statistics_col[0].xpath('tbody/tr[2]/td[2]')[0].get('title',default='NULL').replace(",","")
    strf = f'INSERT INTO [{quote}_other_info] VALUES("{industry}","{sector}",{employees},{shares_outstranding},{market_cap},{enterprise_value})'
    
    curs.execute(strf.replace('n/a','NULL'))
    conn.commit()

def tryfloat(value):
    try:
        return float(value)
    except:
        return 0
def tryDivision(a,b):
    try:
        return a/b
    except:
        return 0
def create_ratio_table(quote):
    curs.execute(f'DROP TABLE IF EXISTS [{quote}_ratio]')
    
    pretty_columns_str = "\""+"\" REAL, \"".join(ratio_columns)+"\" REAL"
    curs.execute(f'CREATE TABLE [{quote}_ratio]({pretty_columns_str})')
    
    cost_of_revenue, revenue, net_income = map(tryfloat, curs.execute(f'SELECT "Cost of Revenue", "Revenue", "Net Income" FROM [{quote}_income]').fetchone())
    total_liabilities, total_current_liabilities, total_assets, total_current_assets, accounts_payable, total_cash = map(tryfloat, curs.execute(f'SELECT "Total Liabilities", "Total Current Liabilities", "Total Assets", "Total Current Assets", "Accounts Payable", "Cash & Cash Equivalents" FROM [{quote}_balance_sheet]').fetchone())
    employees, shares_outstanding = map(tryfloat, curs.execute(f'SELECT "Employees","Shares Outstanding" FROM [{quote}_other_info]').fetchone())
    values = [
        tryDivision(total_current_assets,total_current_liabilities),
        tryDivision(total_cash,total_current_liabilities),
        tryDivision(cost_of_revenue,accounts_payable),
        tryDivision(revenue,total_assets),
        total_current_assets-total_current_liabilities,
        tryDivision(revenue,employees),
        tryDivision(net_income,employees),
        tryDivision((revenue-cost_of_revenue),(revenue)),
        tryDivision(net_income,revenue),
        tryDivision(total_liabilities,total_assets),
        tryDivision(net_income,revenue),
        tryDivision(net_income,total_assets),
        tryDivision(net_income,shares_outstanding)
    ]
    curs.execute(f'INSERT INTO [{quote}_ratio] VALUES({", ".join(map(str,values))})')
    conn.commit()

def update_info(quote):
    income_link = f'https://stockanalysis.com/stocks/{quote}/financials/'
    balance_sheet_link = income_link + 'balance-sheet/'
    cash_flow_link = income_link + 'cash-flow-statement/'
    create_table_and_write_to_it(quote+'_income',*parse_main_info(income_link))
    create_table_and_write_to_it(quote+'_balance_sheet',*parse_main_info(balance_sheet_link))
    create_table_and_write_to_it(quote+'_cash_flow',*parse_main_info(cash_flow_link))
    parse_and_save_other_info(quote)
    create_ratio_table(quote)
