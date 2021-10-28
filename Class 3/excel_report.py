import pandas as pd
import os
from file_management import get_files

target_folder = r'C:\Users\acer\Desktop\Bootcamp\Class 3\excel'

def get_type(file):
    file_type = os.path.splitext(file)[-1]
    return file_type

def get_data(file_list):
    data_list = []
    for file in file_list:
        file_type = get_type(file)
        if file_type == '.csv':
            data = pd.read_csv(os.path.join(target_folder,file),index_col=[0])
        elif file_type == '.xlsx':
            data = pd.read_excel(os.path.join(target_folder,file),index_col=[0])
        data_list.append(data)
    return data_list

def sum_all(data_list):
    sum_result = pd.DataFrame()
    for data in data_list:
        sum_result = sum_result.add(data,fill_value=0)
    return sum_result

def column_sum(sum_result):
    sum_result['Total Qtr'] = sum_result.sum(axis=1)
    sum_result.loc['Total Product'] = sum_result.sum(axis=0)

def main():
    file_list = get_files(target_folder)
    data_list = get_data(file_list)
    sum_result = sum_all(data_list)
    column_sum(sum_result)
    sum_result.to_excel('Result.xlsx')

if __name__=='__main__':
    target_folder = input('Please input your directory: ')
    main()
