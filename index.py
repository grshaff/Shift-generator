import tkinter as tk
from tkinter import *
from customtkinter import *
import tkinter.scrolledtext as st 
from PIL import ImageTk, Image

# Fungsi buat GUI
def create_widgets():

    # Logo
    my_image = ImageTk.PhotoImage(Image.open("assets/SMG-logo.png"))
    root.image_label = CTkLabel(root, image=my_image, text="")
    root.image_label.place(x=125, y=20)

    add_img = Image.open("assets/add.png")  # Membuka gambar
    add_img_resized = add_img.resize((30, 30))  # Mengubah ukuran gambar (misalnya 50x50 piksel)
    add_img_tk = ImageTk.PhotoImage(add_img_resized)  # Mengonversi ke PhotoImage

    # Label-input name
    root.labelname = CTkLabel(root, text="Input name:")
    root.labelname.place(x=30, y=100)

    root.entryname = CTkEntry(root, width=120, textvariable=namevar)
    root.entryname.place(x=30, y=130)

    root.addButton = CTkButton(root, text='', image=add_img_tk, command='', fg_color='transparent', hover_color="", width=10)
    root.addButton.place(x=155, y=130)

    # Log Label
    root.logLabel = CTkLabel(root, text="Names:")
    root.logLabel.place(x=30, y=165)

    root.logFrame = st.ScrolledText(root, wrap=tk.WORD, width=25, height=8, font=("Montserrat", 10)) 
    root.logFrame.place(x=45, y=295)
    root.logFrame.configure(state ='disabled') 

    # Label-input days
    root.labeldays = CTkLabel(root, text="Days:")
    root.labeldays.place(x=210, y= 100)

    root.daycomBox = CTkComboBox(root, values=['28', '29', '30', '31'], width=80,)
    root.daycomBox.place(x=210, y=130)

    # rules Button
    root.GenerateBTN = CTkButton(root, text="Rules", command='', width=120, fg_color="gray", corner_radius=32)
    root.GenerateBTN.place(x=190, y=195)

    # clear Button
    root.GenerateBTN = CTkButton(root, text="Clear", command='', width=120, fg_color="red", corner_radius=32)
    root.GenerateBTN.place(x=190, y=225)

    # generate Button
    root.GenerateBTN = CTkButton(root, text="Generate", command='', width=120, fg_color="green", corner_radius=32)
    root.GenerateBTN.place(x=190, y=255)

# Buat objek class tk
root = CTk()

# Setting ukuran judul
root.title("Managed Service Shift Generator")
root.geometry("340x320")
root.resizable(False, False)
set_appearance_mode("light")

# Variable
namevar = StringVar()


create_widgets()
root.mainloop()