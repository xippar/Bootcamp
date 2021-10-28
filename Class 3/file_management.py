import os
import shutil

target_folder = os.path.join(os.getcwd(),'file_management')

def get_files(target_folder):
    file_list = os.listdir(target_folder)
    return file_list

def check_type(file_list):
    type_set = set()
    for file in file_list:
        type = os.path.splitext(file)[-1]
        if type == '':
            pass
        else:
            type_set.add(type)
    return type_set

def make_dir(type_set):
    for t in type_set:
        folder_path = os.path.join(target_folder,t)
        os.mkdir(folder_path)

def copy_file(file_list):
    for file in file_list:
        type = os.path.splitext(file)[-1]
        if type == '':
            pass
        else:
            file_path = os.path.join(target_folder,file)
            destination_path = os.path.join(target_folder,type)
            shutil.copy(file_path,destination_path)

def main():
    target_folder = input('Please input your directory: ')
    file_list = get_files()
    type_set = check_type(file_list)
    make_dir(type_set)
    copy_file(file_list)

if __name__=='__main__':
    main()
