import datetime
import pandas as pd
import pyexcelerate


data = {'id': [1, 2, 3, 4, 5, 6, 7],
        'Name': ['Alex', 'Justin', 'Set', 'Carlos', 'Gareth', 'John', 'Bob'],
        'Surname': ['Smur', 'Forman', 'Carey', "Carey", 'Chapman', 'James', 'James'],
        'Age': [21, 25, 35, 40, 19, 27, 25],
        'Job': ['Python Developer', 'Java Developer',
                'Project Manager', 'Enterprise architect', 'Python Developer', 'IOS Developer', 'Python Developer'],
        'Datetime': ['2022-01-01T09:45:12',
                     "2022-01-01T11:50:25", '2022-01-01T10:00:45',
                     '2022-01-01T09:07:36', '2022-01-01T11:54:10', '2022-01-01T09:56:40', '2022-01-01T09:52:45']}


df = pd.DataFrame(data)
df['Datetime'] = pd.to_datetime(df['Datetime'])
copy_df = df.copy()
# print(df_copy)
# print(df[(df['Age'] <= 21) & (df['Age'] > 18) & (df['Job'].str.contains('Developer'))] )
# print(datetime.time(hour=9, minute=0))

# df_dev = df[df['Job'].str.contains('Developer')]

# if df[(df['Job'].str.contains('Developer'))]:
#     if df[(df['Age'] <= 21) & (df['Age'] > 18)]:
#         df_copy['TimeToEnter'] = datetime.time(hour=9, minute=0)
#     else:
#         df_copy['TimeToEnter'] = datetime.time(hour=9, minute=15)
#
# print(df_copy)

df_dev = df[df['Job'].str.contains('Developer')]
copy_df['TimeToEnter'] = df_dev['Age'].apply(lambda x: datetime.time(hour=9, minute=0) if 18<x<=21 else datetime.time(hour=9, minute=15))
# print(copy_df)

workbook = pyexcelerate.Workbook()
worksheet = workbook.new_sheet('test')
column = 0
row = 0
origin = (1,1)
headers = copy_df.columns.values
print(headers)

row_num = copy_df.shape[0]
print(row_num)
col_num = copy_df.shape[1]
print(col_num)
ro = origin[0]+row
co = origin[1]+column
worksheet.range((ro, co), (ro, co+len(headers))).value = [[*headers]]
row = row + 1
ro = origin[0] + row
co = origin[1] + column
worksheet.range((ro, co), (ro+row_num,co+col_num)).value = copy_df.values.tolist()
for i in range(ro,co+col_num+1):
    worksheet.set_cell_style(i, 6, pyexcelerate.Style(format=pyexcelerate.Format('yyyy-mm-dd hh:mm')))
# worksheet.set_col_style(6, pyexcelerate.Style(format=pyexcelerate.Format('yyyy-mm-dd hh:mm')))
workbook.save('out2.xlsx')



