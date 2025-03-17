import customtkinter
from CTkDatePicker import CTkDatePicker
import pandas as pd
from db import ConnectionHandler
from logic import logic
from customtkinter import filedialog
from tkinter import messagebox

def selectfile():
    global filename
    filename = filedialog.askdirectory()
    if filename:
        label_location.configure(text=f"Location: {filename}")

def get_data(date: str, adult: int):
    try:
        
        con = ConnectionHandler()

        data_ro = con.execute_procedure("reporting.dbo.HotelPrices_compare", f"{date};{adult}", db="RO")
        data_samo = con.execute_procedure("reporting.dbo.HotelPrices_compare", f"{date};{adult}", db="SAMO")
        df_ro = pd.DataFrame(data_ro)
        df_samo = pd.DataFrame(data_samo)
        if df_ro.empty or df_samo.empty:
            messagebox.showerror("Couldn't Get Data From DB", f"Please wait 5 minutes before executing.")
    except Exception as e:
        messagebox.showerror("Couldn't Get Data From DB", f"Error: {str(e)}")
    con.close_connection("RO")
    con.close_connection("SAMO")

    print("reporting.dbo.HotelPrices_compare", f"{date};{adult}")
    return df_ro, df_samo

def button_callback():
    raw_date = date_picker.get_date()
    if not raw_date:
        messagebox.showerror("Input Error", "Please select a date.")
        return

    try:
        date_parts = raw_date.split('/')
        date = int(date_parts[2] + date_parts[1] + date_parts[0])
    except Exception:
        messagebox.showerror("Input Error", "Invalid date format. Please select a valid date.")
        return

    title_date = raw_date.replace("/", "-")

    adult = entry.get()
    if not adult:
        messagebox.showerror("Input Error", "Please enter the number of adults.")
        return
    
    try:
        adult = int(adult)
    except ValueError:
        messagebox.showerror("Input Error", "Number of adults must be a valid number.")
        return

    if 'filename' not in globals() or not filename:
        messagebox.showerror("Input Error", "Please choose a save location.")
        return

    df_ro, df_samo = get_data(date, adult)
    try:
        logic(df_ro, df_samo, title_date, adult, filename)
        messagebox.showinfo("Success", f"File saved successfully at:\n{filename}")
    except Exception as e:
        messagebox.showerror("Couldn't Save File", f"Please close the Excel file.\n\nError: {str(e)}")
        

app = customtkinter.CTk()
app.title("Vietnam Price Comparison")
app.iconbitmap("Logo.ico")
app.geometry("400x350")
customtkinter.set_appearance_mode("dark")
customtkinter.set_default_color_theme("green")

date_picker = CTkDatePicker(app)
date_picker.set_date_format("%d/%m/%Y")
date_picker.set_allow_manual_input(True)
date_picker.pack(pady=10)

entry = customtkinter.CTkEntry(app, placeholder_text="Adults")
entry.pack(pady=10)

button_to_select = customtkinter.CTkButton(app, text="Choose Save Location", command=selectfile)
button_to_select.pack(pady=10)

label_location = customtkinter.CTkLabel(app, text="Location: Not Selected")
label_location.pack(pady=5)

button = customtkinter.CTkButton(app, text="Download", command=button_callback)
button.pack(pady=10)

app.mainloop()
