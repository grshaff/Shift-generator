import csv
import os
import random
import tkinter as tk
from tkinter import *
from tkinter import messagebox
from customtkinter import *
import tkinter.scrolledtext as st 
from PIL import ImageTk, Image
from datetime import datetime
import threading  # Import threading

# Fungsi buat GUI  
def create_widgets():
    # Logo
    my_image = Image.open("assets/SMG-logo.png")
    my_image_resized = my_image.resize((113, 77))
    my_image_tk = ImageTk.PhotoImage(my_image_resized)
    root.image_label = CTkLabel(root, image=my_image_tk, text="")
    root.image_label.place(x=140, y=20)

    add_img = Image.open("assets/add.png")
    add_img_resized = add_img.resize((30, 30))
    add_img_tk = ImageTk.PhotoImage(add_img_resized)

    # Label-input name
    root.labelname = CTkLabel(root, text="Input name:")
    root.labelname.place(x=30, y=100)

    root.entryname = CTkEntry(root, width=120, textvariable=namevar)
    root.entryname.place(x=30, y=130)

    root.addButton = CTkButton(root, text='', image=add_img_tk, command=log_data, fg_color='transparent', hover_color="", width=10)
    root.addButton.place(x=155, y=130)

    # Log Label
    root.logLabel = CTkLabel(root, text="Names:")
    root.logLabel.place(x=30, y=165)

    root.logFrame = st.ScrolledText(root, wrap=tk.WORD, width=18, height=6, font=("Montserrat", 16)) 
    root.logFrame.place(x=45, y=295)
    root.logFrame.configure(state ='disabled') 

    # Label-input date
    root.labeldays = CTkLabel(root, text="Number of Days:")
    root.labeldays.place(x=210, y= 100)

    root.datecomBox = CTkComboBox(root, values=['28', '29', '30', '31'], width=120)
    root.datecomBox.place(x=210, y=130)

    # Label-input first day of the month
    root.labeldays = CTkLabel(root, text="First Day:")
    root.labeldays.place(x=210, y= 165)

    root.daycomBox = CTkComboBox(root, values=['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'], width=120)
    root.daycomBox.place(x=210, y=195)

    # rules Button
    root.GenerateBTN = CTkButton(root, text="Rules", command='', width=120, fg_color="gray", corner_radius=32)
    root.GenerateBTN.place(x=210, y=235)

    # clear Button
    root.GenerateBTN = CTkButton(root, text="Clear", command=clear_data, width=120, fg_color="red", corner_radius=32)
    root.GenerateBTN.place(x=210, y=268)

    # generate Button
    root.GenerateBTN = CTkButton(root, text="Generate", command=start_generate_data_thread, width=300, fg_color="green", corner_radius=32)
    root.GenerateBTN.place(x=30, y=305)

# button add name func
def log_data():
    input_name = root.entryname.get().strip()  # Ambil teks dari entryname dan hapus spasi berlebih

    if input_name != '':  # Pastikan input tidak kosong
        if input_name in inputted_names:  # Periksa apakah nama sudah ada di inputted_names
            messagebox.showerror("ERROR", f"Name '{input_name}' Already inserted!")
        else:
            # Jika nama belum ada, masukkan ke daftar dan logFrame
            inputted_names.append(input_name)
            root.logFrame.config(state='normal')
            root.logFrame.insert(tk.END, input_name + '\n')  # Tambahkan nama ke logFrame
            root.logFrame.config(state='disabled')
            root.entryname.delete(0, tk.END)  # Bersihkan entry field setelah input
    else:
        messagebox.showerror("ERROR", "Please input a name!")

def clear_data():
    root.logFrame.config(state='normal')
    root.logFrame.delete('1.0', tk.END)
    root.logFrame.config(state='disabled')

# Fungsi untuk memulai thread untuk generate data
def start_generate_data_thread():
    thread = threading.Thread(target=generate_data)
    thread.start()

# Fungsi untuk generate shift data
def generate_data():
    # Ambil nilai dari combobox
    max_days = int(root.datecomBox.get())
    start_day = root.daycomBox.get() 

    # Daftar nama hari dalam urutan yang benar
    days_of_week = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
    
    # Menentukan hari pertama dalam urutan
    start_index = days_of_week.index(start_day)

    # Mengatur urutan hari untuk shift
    week_days = days_of_week[start_index:] + days_of_week[:start_index]

    if inputted_names:
        random.shuffle(inputted_names)

        shift_data = []
        last_names = [] 

        for day in range(max_days):
            day_shift = []
            current_day = week_days[day % len(week_days)]

            for shift in ["Malam", "Pagi", "Sore"]:
                # Cari nama yang belum muncul di shift sebelumnya
                available_names = [name for name in inputted_names if name not in last_names]
                if not available_names:
                    # Jika tidak ada nama yang tersisa, kembalikan ke daftar semua nama
                    available_names = inputted_names.copy()

                selected_name = random.choice(available_names)
                day_shift.append(selected_name)
                last_names.append(selected_name)

            # Mengatur nama-nama yang tidak boleh berurutan dalam hari yang sama
            while len(set(day_shift)) != len(day_shift):
                random.shuffle(day_shift)

            # Menambahkan shift data untuk hari ke-`day`
            shift_data.append([current_day] + day_shift)

        # Buat file CSV
        filename = "shift.csv"
        with open(filename, mode='w', newline='') as file:
            writer = csv.writer(file, delimiter=';')
            writer.writerow(["Hari", "Malam", "Pagi", "Sore"])  # Tulis header kolom

            # Menulis hasil shift ke dalam CSV dengan nomor urut
            for day_index, day_shift in enumerate(shift_data, 1):
                row = [f"Hari {day_index} ({day_shift[0]})"] + day_shift[1:]
                writer.writerow(row)

        # Buka file CSV secara otomatis
        os.startfile(filename)

    else:
        messagebox.showerror("ERROR", "No names to save!")

# Buat objek class tk
root = CTk()

# Setting ukuran judul
root.title("Managed Service Shift Generator")
root.geometry("360x350")
root.resizable(False, False)
set_appearance_mode("light")

# Variable
namevar = StringVar()
inputted_names = []

create_widgets()
root.mainloop()
