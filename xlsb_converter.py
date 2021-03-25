# Created by: Ross Gibson

import win32com.client as win32
import os

def getFiles(directory):
    '''given a directory, will return list of all files in given directory'''
    fileList = []  # place holder used for accumulation
    if directory[-1] != '\\':  # check if path doesn't end with a back slash
        directory = directory + '\\'  # add backslash to path
    for item in os.scandir(directory):  # loop through each item
        if item.is_file():  # check if item is a file
            fileList.append(directory + item.name)  # append path and file name to accumulator list
    return fileList  # return accumulated list

def convertFiles(fileList):
    '''given a list of file paths, will save .xls, .xlsx, .xlsm, or .csv files as .xlsb in same directory'''
    for file in fileList:  # loop through each file in provided list
        if os.path.splitext(file)[1] in ['.xls', '.xlsx', '.xlsm', '.csv']:  # check if excel or csv file
            tgtPath = os.path.splitext(file)[0] + '.xlsb'  # sets target path as .xlsb file extension
            xlApp = win32.Dispatch('Excel.Application')  # create Excel object
            xlApp.Visible = False  # hide Excel window
            xlApp.ScreenUpdating = False  # don't update Excel window (no window flashes)
            xlApp.DisplayAlerts = False  # do not display alerts like updating links
            try:
                wb = xlApp.Workbooks.Open(Filename=file, ReadOnly=True)  # try opening the file in read-only mode
            except:
                print(f'Could not open {file}')  # print file if cannot be opened
                continue
            wb.SaveAs(Filename=tgtPath, FileFormat=50)  # save file as xlsb file format
            wb.Close(False)  # closes Excel workbook without saving
            xlApp.Quit()  # kill Excel process
            print(f'Saved {tgtPath} from {file}')  # print confirmation

def main(directory):
    '''given a windows formatted folder path, will convert all .xls, .xlsx, .xlsm, and .csv files in that directory
    to .xlsb format'''
    fileList = getFiles(directory)  # get list of files in provided directory
    convertFiles(fileList)  # convert Excel/CSV files in file list
