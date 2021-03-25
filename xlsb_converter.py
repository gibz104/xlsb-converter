# Created by: Ross Gibson

import win32com.client as win32
import os

def getFiles(directory):
    'given a directory, will return list of all files in given directory'
    fileList = []
    if directory[-1] != '\\':
        directory = directory + '\\'
    for item in os.scandir(directory):
        if item.is_file():
            fileList.append(directory + item.name)
    return fileList

def convertFiles(fileList):
    'given a list of file paths, will save .xls, .xlsx, .xlsm, or .csv files as .xlsb in same directory'
    for file in fileList:
        if os.path.splitext(file)[1] in ['.xls', '.xlsx', '.xlsm', '.csv']:
            tgtPath = os.path.splitext(file)[0] + '.xlsb'
            xlApp = win32.Dispatch('Excel.Application')
            xlApp.Visible = False
            xlApp.ScreenUpdating = False
            xlApp.DisplayAlerts = False
            try:
                wb = xlApp.Workbooks.Open(Filename=file, ReadOnly=True)
            except:
                print(f'Could not open {file}')
                continue
            wb.SaveAs(Filename=tgtPath, FileFormat=50)

            xlApp.Quit()
            print(f'Saved {tgtPath} from {file}')

def main(directory):
    'given a windows formatted folder path, will convert all .xls, .xlsx, .xlsm, and .csv files in that directory
    to .xlsb format'
    fileList = getFiles(directory)
    convertFiles(fileList)
    print('Done!')

