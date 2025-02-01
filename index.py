import os
from openpyxl import Workbook
from openpyxl.styles import PatternFill
import random
import tkinter as tk
from tkinter import *
from tkinter import messagebox
from customtkinter import *
import tkinter.scrolledtext as st 
from PIL import Image, ImageTk, ImageDraw, ImageFont
import threading  # Import threading

# Fungsi buat GUI  
def create_widgets():

    # Logo
    my_image = Image.open("./assets/SMG-logo.png")
    my_image_resized = my_image.resize((103, 67))
    my_image_tk = ImageTk.PhotoImage(my_image_resized)
    root.image_label = CTkLabel(root, image=my_image_tk, text="",bg_color='transparent')
    root.image_label.place(relx=0.5, rely=0.14, anchor="center")  # Tengah secara horizontal dan sedikit di atas

    add_img = Image.open("./assets/add.png")
    add_img_resized = add_img.resize((22, 22))
    add_img_tk = ImageTk.PhotoImage(add_img_resized)

    # Label-input name
    root.labelname = CTkLabel(root, text="Input name:", bg_color='transparent')
    root.labelname.place(relx=0.09, rely=0.30, anchor="w")

    root.entryname = CTkEntry(root, width=120, textvariable=namevar)
    root.entryname.place(relx=0.09, rely=0.38, anchor="w")
    root.entryname.bind("<Return>", log_data)

    root.addButton = CTkButton(root, text='', image=add_img_tk, command=log_data, fg_color='transparent', hover_color="", width=10)
    root.addButton.place(relx=0.42, rely=0.38, anchor="w")

    # Log Label
    root.logLabel = CTkLabel(root, text="Names:")
    root.logLabel.place(relx=0.09, rely=0.47, anchor="w")

    root.logFrame = st.ScrolledText(root, wrap=tk.WORD,  font=("Montserrat", 16)) 
    root.logFrame.place(relx=0.09, rely=0.68, anchor="w", relwidth=0.45, relheight=0.32,)
    root.logFrame.configure(state ='disabled') 

    # Label-input date
    root.labeldays = CTkLabel(root, text="Number of Days:")
    root.labeldays.place(relx=0.58, rely=0.30, anchor="w")

    root.datecomBox = CTkComboBox(root, values=['28', '29', '30', '31'], width=120)
    root.datecomBox.place(relx=0.58, rely=0.38, anchor="w")

    # Label-input first day of the month
    root.labeldays = CTkLabel(root, text="First Day:")
    root.labeldays.place(relx=0.58, rely=0.47, anchor="w")

    root.daycomBox = CTkComboBox(root, values=['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'], width=120)
    root.daycomBox.place(relx=0.58, rely=0.55, anchor="w")

    # rules Button
    root.GenerateBTN = CTkButton(root, text="Rules", command='', width=120, fg_color="gray", corner_radius=32)
    root.GenerateBTN.place(relx=0.58, rely=0.71, anchor="w")

    # clear Button
    root.GenerateBTN = CTkButton(root, text="Clear", command=clear_data, width=120, fg_color="red", corner_radius=32)
    root.GenerateBTN.place(relx=0.58, rely=0.80, anchor="w")

    # generate Button
    root.GenerateBTN = CTkButton(root, text="Generate", command=start_generate_data_thread, width=300, fg_color="green", corner_radius=32)
    root.GenerateBTN.place(relx=0.5, rely=0.92, anchor="center")

# button add name func
def log_data(event=None):
    input_name = root.entryname.get().strip() 

    if input_name != '':  
        if input_name in inputted_names:
            messagebox.showerror("ERROR", f"Name '{input_name}' Already inserted!")
        else:
            inputted_names.append(input_name)
            root.logFrame.config(state='normal')
            root.logFrame.insert(tk.END, input_name + '\n') 
            root.logFrame.config(state='disabled')
            root.entryname.delete(0, tk.END) 
    else:
        messagebox.showerror("ERROR", "Please input a name!")

# Fungsi clear data
def clear_data():
    root.logFrame.config(state='normal')
    root.logFrame.delete('1.0', tk.END)
    root.logFrame.config(state='disabled')
    inputted_names.clear()

# Fungsi untuk memulai thread untuk generate data
def start_generate_data_thread():
    thread = threading.Thread(target=generate_data)
    thread.start()

# Fungsi untuk generate shift data
def generate_data():
    max_days = int(root.datecomBox.get())
    start_day = root.daycomBox.get()

    days_of_week = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
    start_index = days_of_week.index(start_day)
    week_days = days_of_week[start_index:] + days_of_week[:start_index]

    if inputted_names:
        # Shuffle sekali sebelum loop untuk efisiensi
        random.shuffle(inputted_names)

        shift_data = []
        total_names = len(inputted_names)

        # Dictionary untuk menghitung jumlah shift
        shift_count = {name: {"Malam": 0, "Pagi": 0, "Sore": 0} for name in inputted_names}

        for day in range(max_days):
            current_day = week_days[day % len(week_days)]

            # Ambil nama secara berurutan tanpa pengacakan tambahan
            day_shift = [inputted_names[i % total_names] for i in range(day * 3, (day + 1) * 3)]

            # Update shift_count untuk setiap shift
            shift_count[day_shift[0]]["Malam"] += 1
            shift_count[day_shift[1]]["Pagi"] += 1
            shift_count[day_shift[2]]["Sore"] += 1

            shift_data.append([current_day, day + 1] + day_shift)

        # Membuat workbook dan worksheet baru
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Shift Data"

        # Menulis header
        sheet.append(["Hari", "Tanggal", "Malam", "Pagi", "Sore"])

        # Warna merah muda untuk hari Sabtu dan Minggu
        pink_fill = PatternFill(start_color="FFC0CB", end_color="FFC0CB", fill_type="solid")

        # Menulis data shift
        for day_shift in shift_data:
            day_name = day_shift[0]
            sheet.append(day_shift)

            # Dapatkan baris terakhir yang baru saja ditambahkan
            last_row = sheet.max_row

            # Jika hari Sabtu atau Minggu, warnai baris tersebut dengan merah muda
            if day_name == 'Saturday' or day_name == 'Sunday':
                for cell in sheet[last_row]:
                    cell.fill = pink_fill

        # Menambahkan dua baris kosong
        sheet.append([])
        sheet.append([])

        # Menulis jumlah shift setiap orang
        for name, shifts in shift_count.items():
            result_row = f"{name}: Malam = {shifts['Malam']}, Pagi = {shifts['Pagi']}, Sore = {shifts['Sore']}"
            sheet.append([result_row])

        # Simpan file Excel
        filename = "shift.xlsx"
        workbook.save(filename)

        # Buka file Excel secara otomatis
        os.startfile(filename)

    else:
        messagebox.showerror("ERROR", "No names to save!")



def update_font(event):
    # Menghitung ukuran font berdasarkan lebar jendela
    font_size = int(root.winfo_width() / 32)  # Misalnya, ukuran font 1/20 dari lebar jendela

    # Update font untuk semua widget yang sesuai
    root.logFrame.config(font=("Helvetica", font_size))
    root.logFrame.config(font=("Helvetica", font_size))

# Buat objek class tk
root = CTk()

# Setting ukuran judul
root.title("Managed Service Shift Generator")
root.geometry("360x350")
root.resizable(False, False)  # Biarkan resize agar lebih fleksibel di berbagai resolusi
set_appearance_mode("light")
root.bind("<Configure>", update_font)

# Variable
namevar = StringVar()
inputted_names = []

create_widgets()
root.mainloop()