import subprocess
import sys
import os
import time
from datetime import date, datetime
from pathlib import Path
import sys
import re
import pandas as pd
from ics import Calendar, Event
import tkinter as tk
from tkinter import messagebox, ttk

# Instalar pacotes necessários se não estiverem instalados
def install_and_import(packages):
    for package in packages:
        try:
            __import__(package)
        except ImportError:
            subprocess.check_call([sys.executable, "-m", "pip", "install", package])

required_packages = ['pandas', 'pathlib', 'ics']
install_and_import(required_packages)

def get_days(delta):
    today = date.today()
    day_start, month_start, year = today.day, today.month, today.year
    day_end = day_start + delta
    big_months = {1, 3, 5, 7, 8, 10, 12}
    small_months = {4, 6, 9, 11}
    month_end = month_start
    # Ajusta a data para transição de mês
    if month_start == 2:
        if (year % 4 == 0 and day_end > 29) or (year % 4 != 0 and day_end > 28):
            day_end -= 29 if year % 4 == 0 else 28
            month_start += 1
    elif month_start in small_months and day_end > 30 or month_start in big_months and day_end > 31:
        day_end -= 30 if month_start in small_months else 31
        month_end = month_start + 1

    sheet_name = f"{year-2000}{month_start}" if month_start == month_end else [f"{year-2000}{month_start}", f"{year-2000}{month_start + 1}"]

    return {"year": year, "start_month": month_start, "end_month": month_end, "start_day": day_start, "end_day": day_end, "sheets": sheet_name}

def list_files(cwd):
    return [f for f in os.listdir(cwd) if os.path.isfile(os.path.join(cwd, f))]

def matched_excel(current_year):
    regex = fr"(.*){current_year} - PRO\.APL\.\d{{1,3}} - AGENDA.*"
    files = list_files(os.getcwd())
    for file in files:
        if match := re.search(regex, file):
            return match.group(0), match.group(1).strip()
    raise FileNotFoundError("No matching Excel file found.")

def filter_columns(df):
    df = df.iloc[:, [1, 2, 3, 4, 5, 6, 10]]
    df.columns = ['Date', 'Start', 'Finish', 'Client', 'Serv', 'Equips', 'Obs']
    return df

def read_excel(file_path, sheets_to_read):
    print('STOP')
    if isinstance(sheets_to_read, str):
        
        return filter_columns(pd.read_excel(file_path, sheet_name=sheets_to_read))
    dataframes = {sheet: pd.read_excel(file_path, sheet_name=sheet) for sheet in sheets_to_read}
    df_full = pd.concat(dataframes.values(), ignore_index=True)
    return filter_columns(df_full)

def get_date_format(year, day, month):
    return f"{year}-{month:02}-{day:02}"

def search_date(df, date_dict):
    start_date = get_date_format(date_dict['year'], date_dict['start_day'], date_dict['start_month'])
    end_date = get_date_format(date_dict['year'], date_dict['end_day'], date_dict['end_month'])
    return df[df['Date'] == start_date].index[0], df[df['Date'] == end_date].index[0]

def dst_check(user_name):
    root = tk.Tk()
    root.withdraw()
    return messagebox.askyesno(f"Agenda {user_name}", "Horário de Verão?")

def write_ics(df, user_name):
    cal = Calendar()
    #dst_in_effect = dst_check(user_name)
    print('GMT')
    for _, row in df.iterrows():
        if pd.isna(row['Date']) or pd.isna(row['Start']) or pd.isna(row['Finish']):
            continue

        event = Event()
        client_info = row['Client'].split('\n')
        index = 3
        if len(client_info)==1:
            index = 0
        elif len(client_info)<4:
            index = 1
        elif len(client_info[3])<7:
            index=2

         
        event.name = client_info[0]
        event.location = client_info[index]  if index<3   else client_info[3][4:]

        print('STOP')
        if len(client_info)==1:
            event.description = f"{row['Serv']}\n{row['Equips']}\n{row['Obs']}\n"
        else:
            event.description = f"{row['Serv']}\n{row['Equips']}\n{row['Obs']}\n{client_info[2]}"



        start_datetime = datetime.combine(row['Date'],row['Start'])
        end_datetime = datetime.combine(row['Date'],row['Finish'])

        event.begin = start_datetime
        event.end = end_datetime

        cal.events.add(event)
       

    desktop_path = Path.home() / "Desktop"
    file_path = desktop_path / "Semana.ics"
    with open(file_path, 'w', encoding="utf-8") as f:
        f.writelines(cal)

    print('finished writing ics')
    

def list_folders_in_common_path(year):
    common_path = f"//192.168.9.14/e/GY/AGENDA/{year}"
    user_folders = [folder for folder in os.listdir(common_path) if os.path.isdir(os.path.join(common_path, folder))]
    return user_folders, common_path

# Corrigido para receber 'date_dict' como argumento
def get_user_choice(folders, date_dict):
    root = tk.Tk()
    root.title("Escolha a pasta")
    combo = ttk.Combobox(root, values=folders)
    combo.pack(pady=20)

    def select_folder():
        folder = combo.get()
        os.chdir(f"//192.168.9.14/e/GY/AGENDA/{date_dict['year']}/{folder}")
        root.quit()

    select_button = tk.Button(root, text="Selecionar", command=select_folder)
    select_button.pack(pady=20)
    root.mainloop()

def main(delta):
    date_dict = get_days(delta)
    user_folders, common_path = list_folders_in_common_path(date_dict['year'])
    get_user_choice(user_folders, date_dict)  # Passa date_dict aqui

    file_matched, user_name = matched_excel(date_dict['year'])
    df = read_excel(file_matched, date_dict['sheets'])
    start_idx, end_idx = search_date(df, date_dict)

    df = df.iloc[start_idx:end_idx]
    df['Date'].ffill(inplace=True)
    filtered_df = df[(df['Client'] != 'deslocação') & df['Client'].notna() & (df['Client'] != 'Almoço') & (df['Client'] != 'chegada')]

    write_ics(filtered_df, user_name)
    sys.exit()

if __name__ == "__main__":
    root = tk.Tk()
    root.title('Agenda Gyrad')
    root.geometry("400x400")

    def number():
        try:
            delta = int(my_box.get())
            main(delta)
        except ValueError:
            answer.config(text="Número não inserido!")

    my_label = tk.Label(root, text="Insira o número de dias para a agenda")
    my_label.pack(pady=20)

    my_box = tk.Entry(root)
    my_box.pack(pady=20)

    my_button = tk.Button(root, text="Inserir Número", command=number)
    my_button.pack(pady=20)

    answer = tk.Label(root, text='')
    answer.pack(pady=20)

    root.mainloop()
