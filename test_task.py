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

# condition by the first task
all_col_first = df.copy()

def first_state(age, job):
    if 18<age<=21 and 'Developer' in job:
        return datetime.time(9,0,0,0)
    if 'Developer' in job:
        return datetime.time(9,15,00)

all_col_first['TimetoEnter'] = df.apply(lambda x: first_state(x['Age'], x['Job']), axis=1)

# from dataframe to excel
values = [all_col_first.columns] + list(all_col_first.values)
wb = Workbook()
ws=wb.new_sheet('sheet name', data=values)
changed_datetime = ws.range('F2', 'F8').style.format.format='hh/mm/ss'
changed_timetoenter = ws.range('G2', 'G8').style.format.format='hh/mm/ss'
#TODO use byteIO not to create intermediate xlsx file
wb.save('output.xlsx')

# connect and insert data into mongodb
client = MongoClient('localhost', 27017)
xlsx = pd.read_excel('output.xlsx')

mydb = client['Test_CommercGroup']
mycol = mydb['18MoreAnd21andLess']
docs = json.loads(xlsx.T.to_json()).values()
mycol.insert_many(docs)