import pandas as pd 

df=pd.read_csv('formulabot.csv')

df.to_excel('formulabot1.xlsx', index=False, engine='openpyxl')
print("Excel file creqtqed successfully: sample_data.xlsx")

