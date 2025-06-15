import pandas as pd
import matplotlib.pyplot as plt
import numpy as np

df_bs = pd.read_excel('data.xls', sheet_name='BS', header=None).T
df_bs.columns = df_bs.iloc[0]
df_bs.columns = df_bs.columns.str.strip()
df_bs.drop(df_bs.index[0], inplace=True)
df_pl = pd.read_excel('data.xls', sheet_name='PL', header=None).T
df_pl.columns = df_pl.iloc[0]
df_pl.columns = df_pl.columns.str.strip()
df_pl.drop(df_pl.index[0], inplace=True)
df_pl['Description'] = df_pl['Description'].str.replace('~', '-').str.replace('.', '-')
df_pl['Description'] = pd.to_datetime(df_pl['Description'], errors='coerce').dt.strftime('%Y-%m')
df_pl.dropna(inplace=True)
df_bs['Description'] = df_bs['Description'].str.replace('~', '-').str.replace('.', '-')
df_bs['Description'] = pd.to_datetime(df_bs['Description'], errors='coerce').dt.strftime('%Y-%d')
df_bs.dropna(inplace=True)
df_bs = df_bs.loc[:, ~df_bs.columns.duplicated()]
df_pl = df_pl.loc[:, ~df_pl.columns.duplicated()]
df_bs = df_bs.reset_index(drop=True)
df_pl = df_pl.reset_index(drop=True)
# plt.figure(figsize=(12, 6))
# plt.plot(df_pl['Description'], df_pl['1.Sales'], marker='o')
# plt.title('Sales Over Time')
# plt.xlabel('Date')
# plt.ylabel('Sales')
# plt.xticks(rotation=45)
# for x, y in zip(df_pl['Description'], df_pl['1.Sales']):
#     plt.text(x, y, f'{y:,.0f}', ha='center', va='bottom', fontsize=8, rotation=0)
# plt.tight_layout()
# plt.show()

df_ex = pd.DataFrame()
df_ex['Date'] = pd.to_datetime(df_pl['Description']).dt.strftime('%Y-%m')
df_ex['Total Sales'] = pd.to_numeric(df_pl['1.Sales'])
df_ex['AR-AP'] = pd.to_numeric(df_pl['1.Sales'] - df_pl['COGS-Logistics'])
df_ex['% AR-AP'] = pd.to_numeric(df_ex['AR-AP']/df_ex['Total Sales'] * 100).round(2)
df_ex['Sales-COGS'] = pd.to_numeric(df_pl['1.Sales'] - df_pl['2.Cost of Goods Sold'])
df_ex['% Sales-COGS'] = pd.to_numeric(df_ex['Sales-COGS']/df_ex['Total Sales']).round(2)
df_ex['Operating Income'] = pd.to_numeric(df_pl['5.Operating Income **'])
df_ex['Net Income'] = pd.to_numeric(df_pl['15.Net Income for the Year **'])
df_ex['% Net Income'] = pd.to_numeric(df_ex['Net Income']/df_pl['1.Sales'] * 100).round(2)
df_ex['Current Ratio'] = pd.to_numeric(df_bs['1.Current Assets'] / df_bs['1.Current Liabilities'] * 100).round(2)
df_ex['Quick Ratio'] = pd.to_numeric((df_bs['Cash and Cash Equivalents'] + df_bs['Account Receivable'] )/ df_bs['1.Current Liabilities']).round(2)
df_ex['Asset Turnover Ratio'] = pd.to_numeric(df_pl['1.Sales'] / df_bs['2.Non-current Assets']).round(2)
df_ex['Inventory Turnover Ratio'] = 0
df_ex['Debt-to-equity Ratio'] = pd.to_numeric(df_bs['Liabilities'] / df_bs['Controlling Interest']).round(2)
df_ex['EPS'] = pd.to_numeric(df_pl['15.Net Income for the Year **'] / df_bs['số cổ phiếu']).round(2)
df_ex['P/E'] = 0
df_ex['ROE'] = pd.to_numeric(df_bs['** Net result: profit **'] / df_bs['số cổ phiếu']).round(2)
df_ex['ROA'] = pd.to_numeric(df_bs['** Net result: profit **'] / df_bs['Assets'] * 100).round(2)
df_ex['TNDN'] = pd.to_numeric(df_pl['75010101 Corporate Income Tax Expense-Corporate Tax'])
print(df_ex)