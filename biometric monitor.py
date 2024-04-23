import pandas as pd

df = pd.read_excel(r"C:\Users\bhara\Documents\attendance 16.4.24.xlsx")
df['name'] = ''

try:
    split_values = df['Name'].astype(str).str.split('', expand=True)
    df['name'] = split_values[0]

except ValueError:
    print("Warning: Skipping row with incomplete data due to invalid format.")

df.drop(columns=['SNo', 'E. Code', 'Shift', 'Work Dur.', 'OT', 'Tot.  Dur.', 'Status',],
        inplace=True)

df['Status'] = ''

for i in range(len(df)):
    if (pd.isnull(df.at[i, ' InTime'])):
        pass
    else:
        df[' InTime'] = df[' InTime'].astype(str).str.replace(":", "").astype(float)

df.to_excel(r"C:\Users\bhara\Desktop\main_out.xlsx", index=False)

df1 = pd.read_excel(r"C:\Users\bhara\Documents\storing sheet.xlsx")
df2 = pd.read_excel(r"C:\Users\bhara\Documents\storing sheet.xlsx")
df3 = pd.read_excel(r"C:\Users\bhara\Documents\storing sheet.xlsx")
df4 = pd.read_excel(r"C:\Users\bhara\Documents\storing sheet.xlsx")

k = 0
none = 'none'
min = 0

for i in range(len(df)):
    if pd.isnull(df.at[i, ' InTime']):

        df1.at[k, 'SNo'] = k + 1
        df1.at[k, 'Name'] = df.at[i, 'Name']

        df1.at[k, 'OutTime'] = ''
        df1.at[k, 'Status'] = 'NO PUNCH'

        k += 1



new_file_path = r'C:\Users\bhara\Desktop\Absent.xlsx'
df1.to_excel(new_file_path, index=False)
L=0
for I in range(len(df)):
    if ((df.at[I, ' InTime']) < 200000):
        df2.at[L, 'SNo'] = L + 1
        df2.at[L, 'Name'] = df.at[I, 'Name']
        df2.at[L, 'OutTime'] = ''
        df2.at[L, 'OutTime'] = df.at[I, ' InTime']
        df2.at[L, 'Status'] = 'Before 8:00pm'
        L+=1
for j in range(len(df)):
    if ((df.at[I, ' InTime']) > 204500):
        df2.at[L, 'SNo'] = L + 1
        df2.at[L, 'Name'] = df.at[j, 'Name']
        df2.at[L, 'OutTime'] = ''
        df2.at[L, 'OutTime'] = df.at[j, 'OutTime']
        df2.at[L, 'Status'] = 'After 8:45pm'
        k = k + 1

new_file_path = r'C:\Users\bhara\Desktop\Wrong Time.xlsx'
df2.to_excel(new_file_path, index=False)

M = 0
for J in range(len(df)):
    if (df.at[J, ' InTime'] >= 200000) & (df.at[J, ' InTime'] <= 204500):
        df3.at[M, 'SNo'] = M + 1
        df3.at[M, 'Name'] = df.at[J, 'Name']
        df3.at[M, 'OutTime'] = df.at[J, ' InTime']
        df3.at[M, 'Status'] = '8:00pm - 8:45pm'
        M+= 1

new_file_path = r'C:\Users\bhara\Desktop\Correct.xlsx'
df3.to_excel(new_file_path, index=False)

for Q in range(len(df)):
    if (df.at[Q, ' InTime'] >= 204500):
        df4.at[M, 'SNo'] = M + 1
        df4.at[M, 'Name'] = df.at[Q, 'Name']
        df4.at[M, 'OutTime'] = df.at[Q, ' InTime']
        df4.at[M, 'Status'] = 'After 8:45pm'
        M+=1

new_file_path = r'C:\Users\bhara\Desktop\Wrong Time1.xlsx'
df4.to_excel(new_file_path, index=False)
