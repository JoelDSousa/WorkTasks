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
import time
from datetime import date, datetime
from pathlib import Path

import re
from os import listdir
import tkinter as tk
from tkinter import *
from tkinter import messagebox,ttk
# Install required packages if not already installed
required_packages = [
    'pandas',
    'pathlib',
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
from ics import Calendar, Event

def get_days(delta):
    today = date.today()

    # Converting datetime day into int
    day_start = int(today.strftime("%d"))
    month_start = int(today.strftime("%m"))
    year = int(today.strftime("%Y"))
    day_end = day_start + delta
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

    # Filtering for sheet names of interest in the Excel file
    if month_end == month_start:
        sheetName = str(year-2000)+str(month_start)
        
    else:
        sheetName = [str(year-2000)+str(month_start), str(year-2000)+str(month_end)]
        
    # Creating a dictionary for easy access to information on dates
    dateDict = {
        "year": year,
        "start_month": month_start,
        "end_month": month_end,
        "start_day": day_start,
        "end_day": day_end,
        "sheets":sheetName
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



# Filter data to generate a dataframe only with collumns of interest
def filterCollumns(df):
    
    indices = [1, 2, 3, 4, 5, 6, 10]

    info = df.iloc[:, indices]
    info.columns = ['Date', 'Start', 'Finish', 'Client', 'Serv', 'Equips', 'Obs']
    
    return info



def readExcel(file_path, sheets_to_read):
    if type(sheets_to_read) is str:
        df_interest = pd.read_excel(file_path, sheet_name=sheets_to_read)
        df_interest = filterCollumns(df_interest)
    else:
        dataframes = {sheet: pd.read_excel(file_path, sheet_name=sheet) for sheet in sheets_to_read}
        df_full = pd.concat(dataframes.values(), ignore_index=True) #Returns the result in a form of a single df. To standardize the next processing steps
        df_interest = filterCollumns(df_full)
    return df_interest

def getDateFormat(year, day, month):
    # Adjust string to get the correct expression, according to how the dataframe expresses date: "YYYY-MM-DD"
    return f"{year}-{month:02}-{day:02}"

def searchDate(df, dateDict):
    startDate = getDateFormat(dateDict['year'], dateDict['start_day'], dateDict['start_month'])
    endDate = getDateFormat(dateDict['year'], dateDict['end_day'], dateDict['end_month'])

    start_index = df[df['Date'] == startDate].index
    end_index = df[df['Date'] == endDate].index
    return start_index, end_index

def dst_check(userName):
    root = tk.Tk()
    root.withdraw()
    response = messagebox.askyesno(f"Agenda {userName}", "Horário de Verão?")
    return response

def writeICS(df, userName):
    i = 1
    cal = Calendar()

    # Verificação de horário de verão (assumindo que dst_check está definido)
    dst_in_effect = dst_check(userName=userName)

  
    
    for index, row in df.iterrows():
        start_time = time.time()
        if pd.isna(row['Date']) or pd.isna(row['Start']) or pd.isna(row['Finish']):
            print(i)
            i+=1
            
            continue  # Ignorar linhas inválidas
        
        # Criar um novo evento
        event = Event()
        
        # Formatar a data do evento
        row['Date'] = row['Date'].strftime('%Y-%m-%d')
        
        # Configurar propriedades do evento
        list_client = row['Client'].split('\n')
        

        event.name = list_client[0]  # Nome do evento baseado no cliente
        if len(list_client)>3:
            event.location = list_client[3]
            event.description = f"{row['Serv']}\n{row['Equips']}\n{row['Obs']}\n{list_client[2]}"
        elif len(list_client)>2:
            event.location = list_client[2]
            event.description = f"{row['Serv']}\n{row['Equips']}\n{row['Obs']}"
        else:
            event.description = f"{row['Serv']}\n{row['Equips']}\n{row['Obs']}"


        # Criar datas e horas de início e fim do evento
        start_datetime = datetime.strptime(f"{row['Date']} {row['Start']}", "%Y-%m-%d %H:%M:%S")
        end_datetime = datetime.strptime(f"{row['Date']} {row['Finish']}", "%Y-%m-%d %H:%M:%S")

        #start_datetime = datetime.combine(row['Date'], row['Start'])
        #end_datetime = datetime.combine(row['Date'], row['Start'])
        
        event.begin = start_datetime
        event.end = end_datetime

        # Adicionar evento ao calendário
        cal.events.add(event)
        end_time = time.time()
        print(f"Tempo da iteração {index}: {end_time - start_time:.4f} segundos")
        print(i)
        i+=1


    desktop_path = Path.home() / "Desktop"
    file_path = desktop_path / "Semana.ics"
    print('path')
    with open(file_path, 'w', encoding="utf-8") as f:
        f.writelines(cal)

    print('ics')

    return cal


# USER SELECTS FOLDER FOR ICS SCHEDULE
def list_folders_in_common_path(year):
    global common_path
    common_path = "//192.168.9.14/e/GY/AGENDA/"+str(year)
    user_folders = [folder for folder in os.listdir(common_path)]
    return user_folders, common_path


def item_selecionado(event):
    selecionado = combo.get()
    new_directory = common_path+'\\' + selecionado 
    os.chdir(new_directory)
    print(f"Você selecionou: {selecionado}")


def get_user_choice(folders):
    root = tk.Tk()
    root.title("Escolha a pasta")

    # Create a dropdown menu
    global combo
    
    combo = ttk.Combobox(root, values=folders)
    combo.bind("<<ComboboxSelected>>",item_selecionado)
    combo.pack(pady=20)

    # Set the default value to the first user folder
    if folders:
        combo.set(folders[0])

    # Create a button to select the folder
    select_button = tk.Button(root, text="Selecionar", command=lambda: root.quit())
    select_button.pack(pady=20)

    # Start the event loop
    root.mainloop()







def main(delta):
    dateDict = get_days(delta)
    user_folders, common_path = list_folders_in_common_path(dateDict['year'])
    selected_folder = get_user_choice(user_folders)
    
    if selected_folder:
        new_directory = common_path+'\\' + selected_folder 
        os.chdir(new_directory)

        
    
    fileMatched, userName = matched_Excel(dateDict['year'])
    dataframe = readExcel(fileMatched,dateDict['sheets'])
    

    start_index, end_index = searchDate(dataframe, dateDict)
    
    df = dataframe.iloc[start_index[0]:end_index[0]]
    


    # Fill down the Date, Start and Finish columns to propagate dates where NaT is present
    df['Date'].fillna(method='ffill', inplace=True)
    

    # Filter out NaN and 'deslocação values
    filtered_df = df[(df['Client'] != 'deslocação') & (df['Client'].notna()) & (df['Client'] != 'Almoço') & (df['Client'] != 'chegada')]

    

    writeICS(filtered_df,userName)
    
   

if __name__ == "__main__":
    root = Tk()
    root.title('Agenda Gyrad')
    root.geometry("400x400")
    def number():
        try:
            delta = int(my_box.get())
            main(delta)
        except ValueError:
            answer.config(text="Numero nao inserido!")

    my_label = Label(root, text="insira o número de dias para a agenda")
    my_label.pack(pady=20)

    my_box = Entry(root)
    my_box.pack(pady=20)

    my_button = Button(root,text="Inserir Número",command=number)
    my_button.pack(pady=20)

    answer = Label(root,text='')
    answer.pack(pady=20)

    root.mainloop()
