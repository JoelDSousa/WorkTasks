#This script has the following objectives:

#To access the file containing information about the work schedule for the whole year
#To identify what day it is at the time of running the script
#To collect the relevant information for the next seven days, such as:
#   Clients to visit;
#   Time scheduled to visit;
#   What type of measurements are to be taken (CQI, CQE, CQE.EXTRA or TA)
#   What equipments are to be evaluated (INTRA, ORTO, ORTOCEF, ORTOCBCT, ORTOCBCTCEF,...)
#   The address of the client
#After collecting such data, a .txt file, by the name of "Week_Schedule.txt" is to be generated
#The purpose of this file is to facilitate the worker navigation, by creating a file easy to copy to apps like google maps or calendar

#    LIBS
from datetime import date
import pandas as pd #requires pyarrow
import re
import os
from os import listdir
import xlrd
import io
import msoffcrypto
import openpyxl



def get_days():
    today = date.today()

    #converting datetime day into int

    day_start = int(today.strftime("%d"))
    month_start = int(today.strftime("%m"))
    year = int(today.strftime("%Y"))
    day_end = int(today.strftime("%d"))+7
    
    bigMonths = [1,3,5,7,8,10,12]
    smallMonths = [4,6,9,11]

# Adjusting values when the week is present in 2 seperate months

    if month_start==2 and year%4==0:
        if day_end>29:
            month_end = month_start+1
            day_end = day_end-29
    elif month_start==2 and year%4!=0:
        if day_end > 28:
            month_end = month_start+1
            day_end = day_end-28
    elif month_start in smallMonths:
        if day_end>30:
            month_end = month_start+1
            day_end = day_end-30
    elif month_start in bigMonths:
        if day_end>31:
            month_end = month_start+1
            day_end = day_end-31


    #Creating a dictionary for easy access to information on dates 
    dateDict = {
        "year": year,
        "start_month" : month_start,
        "end_month": month_end,
        "start_day": day_start,
        "end_day": day_end
    }

    return dateDict

# List all files' names in a directory
def listFiles(cwd):
    allFiles = []
    allFiles = listdir(cwd)
    return allFiles



# will search for the excel file that has the current Year schedule
def matched_Excel(currentYear):
    fullCWD = os.getcwd()#Current Working Directory Full path
    allFiles = listFiles(fullCWD)
    currentYear=str(currentYear)
    regex = r"(.*)"+currentYear+" - PRO\.APL\.\d{,3} - AGENDA"+"(.*)"
    for file_name in allFiles:
        matches= re.search(regex,file_name)
        if matches is not None:
            name = matches.group(1).strip()
            file_matched = matches.group(0)
            break
        else:
            continue    
    return file_matched, name


# will open the corresponding excel and send the data via dataframe format

def open_Excel(matchedExcel):
    decrypted_workbook = io.BytesIO()
    with open(matchedExcel, 'rb') as file:
        office_file = msoffcrypto.OfficeFile(file)
        office_file.load_key(password='321')
        office_file.decrypt(decrypted_workbook)

    wb = openpyxl.load_workbook(filename=decrypted_workbook, data_only=True)
    return wb

def sheetDF(wb, dateDict):
    if len(str(dateDict['start_month']))<2:
        dateDict['start_month'] = '0'+str(dateDict['start_month'])
    else:
        dateDict['start_month'] = str(dateDict['start_month'])
    sheetName = str(dateDict['year'])[-2:] + dateDict['start_month']
    ws = wb[sheetName]
    xlDF = pd.DataFrame(ws.values) 

    return xlDF

def main():
    dateDict = get_days()
    print(dateDict)
    fileMatched, userName = matched_Excel(dateDict['year'])
    wb = open_Excel(fileMatched)
    xlDF = sheetDF(wb,dateDict)
    print(xlDF)


main()