from pymongo import MongoClient
import pandas as pd
import json
import datetime
from pyexcelerate import Workbook


data = {
        'id': [1, 2, 3, 4, 5, 6, 7],
        'Name': ['Alex', 'Justin', 'Set', 'Carlos', 'Gareth', 'John', 'Bob'],
        'Surname': ['Smur', 'Forman', 'Carey', "Carey", 'Chapman', 'James', 'James'],
        'Age': [21, 25, 35, 40, 19, 27, 25],
        'Job': ['Python Developer', 'Java Developer',
                'Project Manager', 'Enterprise architect', 'Python Developer', 'IOS Developer', 'Python Developer'],
        'Datetime': ['2022-01-01T09:45:12',
                     "2022-01-01T11:50:25", '2022-01-01T10:00:45',
                     '2022-01-01T09:07:36', '2022-01-01T11:54:10',
                     '2022-01-01T09:56:40', '2022-01-01T09:52:45']
        }

df = pd.DataFrame(data)
df['Datetime'] = pd.to_datetime(df['Datetime'])

all_col_second = df.copy()


def second_state(age, job):
    if age >= 35 and ('Developer' and "Manager")  in job:
        return datetime.time(11, 0, 0, 0)
    else:
        return datetime.time(11, 30, 0, 0)

all_col_second['TimetoEnter'] = df.apply(lambda x: second_state(x['Age'], x['Job']), axis=1)

# from dataframe to excel
values = [all_col_second.columns] + list(all_col_second.values)
wb = Workbook()
ws = wb.new_sheet('sheet name', data=values)
changed_datetime = ws.range('F2', 'F8').style.format.format = 'hh/mm/ss'
changed_timetoenter = ws.range('G2', 'G8').style.format.format = 'hh/mm/ss'
# TODO use byteIO not to create intermediate xlsx file
wb.save('output1.xlsx')

# from excel to mongodb
client = MongoClient('localhost', 27017)
xlsx = pd.read_excel('output1.xlsx')

mydb = client['Test_CommercGroup']
mycol = mydb['35AndMore']
docs = json.loads(xlsx.T.to_json()).values()
mycol.insert_many(docs)



