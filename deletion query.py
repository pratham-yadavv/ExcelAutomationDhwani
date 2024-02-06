import pandas as pd

path = r'C:/Pratham/python/ExcelAutomation/sheet1.xlsx'

# df = pd.read_excel(path,sheet_name='3B Program Farmer profiling')
df = pd.read_excel(path,sheet_name='3C Crop Card Farmer')
df = df.dropna(subset=['Transaction Id'])

transacid_str = ', '.join(["'" + str(transacid) + "'" for transacid in df['Transaction Id']])


print(transacid_str)