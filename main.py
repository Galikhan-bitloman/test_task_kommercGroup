import datetime
import pandas as pd
from pyexcelerate import Workbook
from pymongo import MongoClient
import json

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

df_first_condition = df.copy()
df_second_condition = df.copy()
df_third_condition = df.copy()

def set_time_by_condition(age, job):
    if 18 < age <= 21 and 'Developer' in job:
        return datetime.time(9, 0, 0)
    if 'Developer' in job:
        return datetime.time(9, 15, 0)


df_first_condition['TimetoEnter'] = df.apply(lambda x: set_time_by_condition(x['Age'], x['Job']), axis=1)


def second_task(age, job):
    if age >= 35 and ('Developer' and "Manager") in job:
        return datetime.time(11, 0, 0, 0)
    else:
        return datetime.time(11, 30, 0, 0)

df_second_condition['TimetoEnter'] = df.apply(lambda x: second_task(x['Age'], x['Job']), axis=1)


def third_task(job):
    if 'architect' in job:
        return datetime.time(10, 30, 0, 0)
    else:
        return datetime.time(10, 40, 0, 0)


df_third_condition['TimetoEnter'] = df.apply(lambda x: third_task(x['Job']), axis=1)


def from_df_to_xlsx(df, sheet_name, output_name):
    wb = Workbook()
    sheet = wb.new_sheet(sheet_name)
    origin = (1,1)
    column_length = 0
    row_length = 0
    columns = df.columns.tolist()
    row = origin[0] + row_length
    column = origin[1] + column_length
    sheet.range((row, column), (row, column+len(columns))).value = [[*columns]]
    row_length += 1
    df_row_num = df.shape[0]
    df_column_num = df.shape[1]
    row = origin[0] + row_length
    column = origin[1] + column_length
    sheet.range((row, column), (row+df_row_num, column + df_column_num)).value = df.values.tolist()
    sheet.range('F2', 'F8').style.format.format = 'yyyy-mm-dd hh:mm'
    sheet.range('G2', 'G8').style.format.format = 'hh:mm:ss'
    return wb.save(output_name)


from_df_to_xlsx(df_first_condition, 'first_sheet', 'output1.xlsx')
from_df_to_xlsx(df_second_condition, 'second_sheet', 'output2.xlsx')
from_df_to_xlsx(df_third_condition, 'third_sheet', 'output3.xlsx')



client = MongoClient('localhost', 27017)

def from_xlsx_to_mongodb(output_name, col_name):
    xlsx = pd.read_excel(output_name)
    mydb = client['Test_CommerceGroup1']
    mycol = mydb[col_name]
    docs = json.loads(xlsx.T.to_json()).values()

    mycol.insert_many(docs)

from_xlsx_to_mongodb('output2.xlsx', '18MoreAnd21andLess')

