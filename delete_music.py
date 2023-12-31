import openpyxl
import os
import sys

def get_excel_value(ws, row, column):
    return ws.cell(row = row, column = column).value

music_download_folder = 'downloaded_music_files'
music_url_list_excel_file = 'sample_music_url_list.xlsx'
to_delete_filepaths = []

def get_deleted_music_filepaths():
    wb = openpyxl.load_workbook(music_url_list_excel_file)
    ws = wb["Sheet1"]
    for x in range(2, 10000):
        if get_excel_value(ws, x, 1) == None: #No url
            break
        if get_excel_value(ws, x, 8) != 1: #mark_delete != 1 then skip to next row
            continue
        title = get_excel_value(ws, x, 2)
        filepath = os.path.join(music_download_folder,title + '.mp3')
        to_delete_filepaths.append(filepath)

get_deleted_music_filepaths()

def remove_invalid_char(filename):
    invalid = "<>:\"/\|?*,\'|"
    for char in invalid:
	    valid_filename = filename.replace(char, '').replace("'", '')
    return valid_filename

for filepath in to_delete_filepaths:
    print(filepath)
    valid_filepath = remove_invalid_char(filepath)
    print(valid_filepath)
    if os.path.exists(valid_filepath):
        os.remove(valid_filepath)
        print("Deleted")
    else:
        print("The file does not exist")