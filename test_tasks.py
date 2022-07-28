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

all_col_first = df.copy()
all_col_second = df.copy()
all_col_third = df.copy()

# three main tasks
def first_task(age, job):
    if 18 < age <= 21 and 'Developer' in job:
        return datetime.time(9, 0, 0, 0)
    if 'Developer' in job:
        return datetime.time(9, 15, 00)


all_col_first['TimetoEnter'] = df.apply(lambda x: first_task(x['Age'], x['Job']), axis=1)


def second_task(age, job):
    if age >= 35 and ('Developer' and "Manager") in job:
        return datetime.time(11, 0, 0, 0)
    else:
        return datetime.time(11, 30, 0, 0)


all_col_second['TimetoEnter'] = df.apply(lambda x: second_task(x['Age'], x['Job']), axis=1)


def third_task(job):
    if 'architect' in job:
        return datetime.time(10, 30, 0, 0)
    else:
        return datetime.time(10, 40, 0, 0)


all_col_third['TimetoEnter'] = df.apply(lambda x: third_task(x['Job']), axis=1)




def from_df_to_xlsx(all_col, sheet_name, output_name, wb):
    var = [all_col.columns] + list(all_col.values)
    sheet = wb.new_sheet(sheet_name, data=var)
    sheet.range('F2', 'F8').style.format.format = 'hh/mm/ss'
    sheet.range('G2', 'G8').style.format.format = 'hh/mm/ss'

    # TODO use byteIO not to create intermediate xlsx file

    return wb.save(output_name)


from_df_to_xlsx(all_col_first, 'first_sheet', 'output1.xlsx', Workbook())
from_df_to_xlsx(all_col_second, 'second_sheet', 'output2.xlsx', Workbook())
from_df_to_xlsx(all_col_third, 'third_sheet', 'output3.xlsx', Workbook())


client = MongoClient('localhost', 27017)

def from_xlsx_to_mongodb(output_name, col_name):
    xlsx = pd.read_excel(output_name)
    mydb = client['Test_CommerceGroup']
    mycol = mydb[col_name]
    docs = json.loads(xlsx.T.to_json()).values()

    mycol.insert_many(docs)

from_xlsx_to_mongodb('output1.xlsx', '18MoreAnd21andLess')
from_xlsx_to_mongodb('output2.xlsx', '35AndMore')
from_xlsx_to_mongodb('output3.xlsx', 'ArchitectEnterTime')
