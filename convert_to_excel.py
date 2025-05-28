import pandas as pd 

df=pd.read_csv('sample_data.csv')

df.to_excel('sample_data.xlsx', index=False, engine='openpyxl')
print("Excel file creqtqed successfully: sample_data.xlsx")

