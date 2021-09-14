from tkinter import *
from tkinter import ttk
import pandas as pd
import io
import requests
from openpyxl import load_workbook
from tkinter import filedialog

window = Tk()
window.geometry("200x200")
#window.iconphoto(False, icon)
window.title("Weather Downloader app")
window.iconbitmap("icon.ico")   

global filepath_var
filepath_var = StringVar()
file_path_entry = Entry(window, width = 60, bg = "grey85", textvariable = filepath_var)
file_path_entry.pack()

def browse():
    global file_path
    file_path = filedialog.askopenfilename(filetypes=(("excel","*.xlsx"),("All files","*.*")))
    filepath_var.set(file_path)

month_variable = StringVar(window)
monthoptions = ["Select Month", "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12"]
month_menu = ttk.OptionMenu(window, month_variable, *monthoptions)
month_menu.pack()

year_var = StringVar(window)
yearoption = ["Select Year", "2020", "2021", "2022", "2023", "2024", "2025"]
year_menu = ttk.OptionMenu(window, year_var, *yearoption)
year_menu.pack()


def update():
    
    month = month_variable.get()
    year = year_var.get()
    df = pd.read_excel(file_path, "Weather Input 2021")
    url = "https://dd.weather.gc.ca/climate/observations/daily/csv/ON/climate_daily_ON_6158355_"+year+"-"+month+"_P1D.csv"
    s = requests.get(url).content
    c = pd.read_csv(io.StringIO(s.decode('cp1252')))
    c["Date/Time"] = pd.to_datetime(c["Date/Time"])
    
    new_df = pd.concat([df,c]).drop_duplicates().reset_index(drop = True)
    
    book = load_workbook(file_path)
    writer = pd.ExcelWriter(file_path, engine = "openpyxl")
    writer.book = book
    std=book.get_sheet_by_name("Weather Input 2021")
    book.remove_sheet(std)
    new_df.to_excel(writer, "Weather Input 2021", index = False)
    writer.save()
    
    message.config(text = "Success!")
    
    
browse = Button(window, text = "Browse file", command = browse)
browse.pack()

update = Button(window, text = "Update Weather Data in tool", command = update)
update.pack()  

message = Label(window)
message.pack()
window.mainloop()
