# This script has the following objectives:
# - Access the file containing information about the work schedule for the whole year
# - Identify the current day when the script is run
# - Collect relevant information for the next seven days, such as:
#   - Clients to visit
#   - Scheduled visit times
#   - Type of measurements (CQI, CQE, CQE.EXTRA, TA)
#   - Equipment to be evaluated (INTRA, ORTO, ORTOCEF, ORTOCBCT, ORTOCBCTCEF, etc.)
#   - Client address
# - Generate a .txt file named "Week_Schedule.txt" with the collected data

import subprocess
import sys
import os
from datetime import date, datetime, timedelta
import re
import io
from os import listdir
# Install required packages if not already installed
required_packages = [
    'pandas',
    'xlrd',
    'openpyxl',
    'msoffcrypto-tool',
    'ics'
]

def install_and_import(package):
    try:
        __import__(package)
    except ImportError:
        subprocess.check_call([sys.executable, "-m", "pip", "install", package])

# Ensure all required packages are installed
for package in required_packages:
    install_and_import(package)

import pandas as pd
import xlrd
import msoffcrypto
import openpyxl
from ics import Calendar, Event

def get_days():
    today = date.today()

    # Converting datetime day into int
    day_start = int(today.strftime("%d"))
    month_start = int(today.strftime("%m"))
    year = int(today.strftime("%Y"))
    day_end = day_start + 7
    month_end = month_start
    bigMonths = [1, 3, 5, 7, 8, 10, 12]
    smallMonths = [4, 6, 9, 11]

    # Adjusting values when the week spans two months
    if month_start == 2 and year % 4 == 0:
        if day_end > 29:
            month_end = month_start + 1
            day_end -= 29
    elif month_start == 2 and year % 4 != 0:
        if day_end > 28:
            month_end = month_start + 1
            day_end -= 28
    elif month_start in smallMonths:
        if day_end > 30:
            month_end = month_start + 1
            day_end -= 30
    elif month_start in bigMonths:
        if day_end > 31:
            month_end = month_start + 1
            day_end -= 31

    # Creating a dictionary for easy access to information on dates
    dateDict = {
        "year": year,
        "start_month": month_start,
        "end_month": month_end,
        "start_day": day_start,
        "end_day": day_end
    }

    return dateDict

# List all files' names in a directory
def listFiles(cwd):
    return listdir(cwd)

# Search for the Excel file that has the current year's schedule
def matched_Excel(currentYear):
    fullCWD = os.getcwd()  # Current Working Directory Full path
    allFiles = listFiles(fullCWD)
    currentYear = str(currentYear)
    regex = r"(.*)" + currentYear + r" - PRO\.APL\.\d{1,3} - AGENDA.*"
    
    for file_name in allFiles:
        matches = re.search(regex, file_name)
        if matches is not None:
            name = matches.group(1).strip()
            file_matched = matches.group(0)
            return file_matched, name
    raise FileNotFoundError("No matching Excel file found.")

# Open the corresponding Excel file and return the data as a dataframe
def open_Excel(matchedExcel):
    decrypted_workbook = io.BytesIO()
    with open(matchedExcel, 'rb') as file:
        office_file = msoffcrypto.OfficeFile(file)
        office_file.load_key(password='321')
        office_file.decrypt(decrypted_workbook)

    wb = openpyxl.load_workbook(filename=decrypted_workbook, data_only=True)
    return wb

# Reuse code for extracting sheet info
def sheetinfo(wb, sheetName):
    ws = wb[sheetName]
    df = pd.DataFrame(ws.values)
    indices = [1, 2, 3, 4, 5, 6, 10]

    info = df.iloc[:, indices]
    info.columns = ['Date', 'Start', 'Finish', 'Client', 'Serv', 'Equips', 'Obs']
    
    return info

def sheetDF(wb, dateDict):
    dateDict['start_month'] = f"{dateDict['start_month']:02}"
    dateDict['end_month'] = f"{dateDict['end_month']:02}"

    sheetNameStart = str(dateDict['year'])[-2:] + dateDict['start_month']
    sheetNameEnd = str(dateDict['year'])[-2:] + dateDict['end_month']
    
    infoStart = sheetinfo(wb, sheetNameStart)
    infoEnd = sheetinfo(wb, sheetNameEnd)

    return infoStart, infoEnd

def getDateFormat(year, day, month):
    # Adjust string to get the correct expression, according to how the dataframe expresses date: "YYYY-MM-DD"
    return f"{year}-{month:02}-{day:02}"

def searchDate(infoStart, infoEnd, dateDict):
    startDate = getDateFormat(dateDict['year'], dateDict['start_day'], dateDict['start_month'])
    endDate = getDateFormat(dateDict['year'], dateDict['end_day'], dateDict['end_month'])

    indexStart = next((i for i, row in enumerate(infoStart.iloc[:, 0]) if startDate in str(row)), None)
    indexEnd = next((i for i, row in enumerate(infoEnd.iloc[:, 0]) if endDate in str(row)), None)

    if indexStart is None or indexEnd is None:
        raise ValueError("Start or end date not found in the data.")

    return indexStart, indexEnd

def filterInfo(indexStart, indexEnd, infoStart, infoEnd):
    # Merging both dataframes to include all the information in one dataframe
    if infoStart.equals(infoEnd):
        mergedInfo = infoStart.iloc[indexStart:indexEnd, :]
    else:
        startInterval = infoStart.iloc[indexStart:, :]
        endInterval = infoEnd.iloc[2:indexEnd, :]
        mergedInfo = pd.concat([startInterval, endInterval])

   #Changing the None value in the date to the date in the previous row
    for index in range(len(mergedInfo['Date'])):
       if mergedInfo.iloc[index,0]==None:
           mergedInfo.iloc[index,0]=mergedInfo.iloc[index-1,0]

    
    # Resetting the row index
    mergedInfo = mergedInfo.reset_index(drop=True)
    
    # Defining expressions that don't need to be included
    expressions = ["https", "None", "deslocação", "CLIENTE", "EQUIPs", "OBSERVAÇÕES TÉCNICO"]
    
    # Filtering the rows that include pointless information
    index_list = [row_index for row_index, row in mergedInfo.iterrows()
                  if any(expression in str(row['Client']) for expression in expressions)]
    
    finalInfo = mergedInfo.drop(index_list)
    
    return finalInfo

def writeICS(df):
    
    cal = Calendar()

    #Filtering the name field to get location info:

    for index, row in df.iterrows():
        event = Event()
        
        list_lines = str(row['Client']).split("\n")
        event.name = list_lines[0]


        if len(list_lines)>1:
            if len(list_lines[3])>5:
                event.location = list_lines[3]
            else:
                event.location = list_lines[2]
        
        if row['Start'] == None or pd.isna(row['Start']):
            event.begin = datetime(row['Date'].year, row['Date'].month,row['Date'].day)
            event.end = datetime(row['Date'].year, row['Date'].month,row['Date'].day)
            
        else:
            
            event.begin = datetime(int(row['Date'].year),int(row['Date'].month), int(row['Date'].day), int(row['Start'].hour), int(row['Start'].minute))
            
            event.end = datetime(int(row['Date'].year), row['Date'].month, row['Date'].day, row['Finish'].hour, row['Finish'].minute)
        
        
        event.description = f"Equipments: {row['Equips']}\nObservations: {row['Obs']}\nService: {row['Serv']}"
        cal.events.add(event)
        
    with open("Week_Schedule.ics", 'w', encoding="utf-8") as f:  # Encoding is specified to include special characters in the txt file
        
        f.writelines(cal)
    return(cal)

def main():
    dateDict = get_days()
    fileMatched, userName = matched_Excel(dateDict['year'])
    wb = open_Excel(fileMatched)
    infoStart, infoEnd = sheetDF(wb, dateDict)
    indexStart, indexEnd = searchDate(infoStart, infoEnd, dateDict)
    df = filterInfo(indexStart, indexEnd, infoStart, infoEnd)
    #print(df)
    cal = writeICS(df)
    #print(cal.serialize())
   

if __name__ == "__main__":
    main()
