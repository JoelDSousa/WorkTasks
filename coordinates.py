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


def get_days():
    today = date.today()

    #converting datetime day into int

    day_start = int(today.strftime("%d"))
    month = int(today.strftime("%m"))
    year = int(today.strftime("%Y"))
    day_end = int(today.strftime("%d"))+7

    #Creating a dictionary for easy access to information on dates 
    dateDict = {
        "year": year,
        "month" : month,
        "start_day": day_start,
        "end_day": day_end
    }
    return dateDict

# will search for the excel file that has the current Year schedule
def open_Excel(currentYear):
    currentYear=str(currentYear)
    regex = r"(.*)"+currentYear+" - PRO.APL.* - AGENDA"
    test_str = "JOEL 2024 - PRO.APL.010 - AGENDA"
    matches= re.search(regex,test_str)
    if matches.group(1).strip():
        return matches.group(1).strip()
    return 0





def main():
    dateDict = get_days()
    name = open_Excel(dateDict['year'])
    print(name)
    

main()