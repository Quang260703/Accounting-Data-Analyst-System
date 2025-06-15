import pandas as pd
import matplotlib.pyplot as plt

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
# df_pl.to_csv('export.csv', index=False)

print(df_bs['Description'])