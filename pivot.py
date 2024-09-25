import pandas as pd
df=pd.read_excel("pivot/Consolidate Cloud Expense December 2023.xlsx",sheet_name='December Detail')
print(df.columns)
pivot_columns = ['Tag: BU', 'Cloud', 'Subscription', 'Tag: Dept']
pivot_table = df.pivot_table(values='Cost', index=pivot_columns, aggfunc='sum')
print(pivot_table)
pivot_table.to_excel('pivot/table.xlsx')