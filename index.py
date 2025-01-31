import tkinter as tk
from tkinter import *
from tkinter import messagebox
from customtkinter import *
import tkinter.scrolledtext as st 
from PIL import ImageTk, Image

# Fungsi buat GUI  
def create_widgets():
    
    # Logo
    my_image = ImageTk.PhotoImage(Image.open("assets/SMG-logo.png"))
    root.image_label = CTkLabel(root, image=my_image, text="")
    root.image_label.place(x=140, y=20)

    add_img = Image.open("assets/add.png")  # Membuka gambar
    add_img_resized = add_img.resize((30, 30))  # Mengubah ukuran gambar (misalnya 50x50 piksel)
    add_img_tk = ImageTk.PhotoImage(add_img_resized)  # Mengonversi ke PhotoImage

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

    root.daycomBox = CTkComboBox(root, values=['28', '29', '30', '31'], width=120)
    root.daycomBox.place(x=210, y=130)

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
    root.GenerateBTN = CTkButton(root, text="Generate", command='', width=300, fg_color="green", corner_radius=32)
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