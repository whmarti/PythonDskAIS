#*************************************************************************
# Develop of Courses Outline Hola Developer 1
# Auckland Institute of Studies
# Developers: William Martin, June , Sun....
# Date of creation 03/11/2021
#*************************************************************************
from tkinter.filedialog import askopenfile 
from tkinter import *
from tkinter.ttk import *

import tkinter as tk,time
from tkinter.constants import ANCHOR, NW, RIGHT, Y
from PIL import ImageTk, Image
import pandas as pd
import docx
from tkinter.messagebox import showinfo
from tkinter import ttk 
from tkinter import messagebox

import shutil,os



root = tk.Tk()
root.configure(background='#FFFFFF')
root.title("Course Outline Generator")
root.geometry('700x500')

logoImg = (Image.open("img/logo.jpg"))
resizedImg = logoImg.resize((120,50), Image.ANTIALIAS)
logoImg = ImageTk.PhotoImage(resizedImg)
label = tk.Label(root, image=logoImg).place(x=280, y=10)
heading = tk.Label(root, text="Course Outline Generator")
heading.config(font=('Arial', 18))
heading.pack(padx=50, pady=65)

#*************************************************************************
# Here start the fields definition:
ycbxPos=120
xCbxLbl=108

def trimester_changed(event):
    msg = f'You selected {trimester_cb.get()}!'
    showinfo(title='Result', message=msg)
# Trimester of year
months = ('1', '2', '3')
 #this creates 'Label' widget for Trimester.
label_trimester =tk.Label(root,text="Trimester", width=20,font=("bold",10))
label_trimester.place(x=xCbxLbl,y=ycbxPos) #x=48
xCbxLbl+=172

selected_month = tk.StringVar()
trimester_cb = ttk.Combobox(root, textvariable=selected_month,width=5)
trimester_cb['values'] = months
trimester_cb['state'] = 'readonly'  # normal
# trimester_cb.bind('<<ComboboxSelected>>', trimester_changed)
trimester_cb.current(1)
# entry_trimester=tk.Entry(root)
trimester_cb.place(x=xCbxLbl,y=ycbxPos) #x=220
xCbxLbl+=58

def year_changed(event):
    msg = f'You selected {year_cb.get()}!'
    showinfo(title='Result', message=msg)
# Years
years = ('2020', '2021', '2022', '2024')
 #this creates 'Label' widget for Year.
label_year =tk.Label(root,text="Year", width=20,font=("bold",10))
label_year.place(x=xCbxLbl,y=ycbxPos)  #278
xCbxLbl+=172

selected_year = tk.StringVar()
year_cb = ttk.Combobox(root, textvariable=selected_year,width=5)
year_cb['values'] = years
year_cb['state'] = 'readonly'  # normal
# year_cb.bind('<<ComboboxSelected>>', year_changed)
year_cb.current(1)
# entry_year=tk.Entry(root)
year_cb.place(x=xCbxLbl,y=ycbxPos) #x=450

w = tk.Canvas(root, width=1200, height=3)
# w.create_rectangle(0, 0, 1900, 2, fill="gray", outline = 'gray')
# w.pack()
w.place(y=160)

# Controls Margins:
xPosL=80
xPosF=260
yPos=180

lectHeading = tk.Label(root, text="Lecturer Information")
lectHeading.config(font=('Arial', 14))
lectHeading.place(x=50, y=yPos)
yPos+=50

#the variable 'var' mentioned here holds Integer Value, by deault 0
var=tk.StringVar()
#this creates 'Radio button' widget and uses place() method
tk.Radiobutton(root,text="Lecturer",padx= 5, variable= var, value="Lecturer").place(x=xPosL,y=yPos)
tk.Radiobutton(root,text="Course Coordinator",padx= 20, variable= var, value="Course Coordinator").place(x=xPosF,y=yPos)
yPos+=35

#this creates 'Label' widget for Fullname.
label_name =tk.Label(root,text="Full Name", width=20,font=("bold",10))
label_name.place(x=xPosL,y=yPos)

entry_name=tk.Entry(root)
entry_name.place(x=xPosF,y=yPos)
yPos+=30

#this creates 'Label' widget for Email.
label_email =tk.Label(root,text="Email", width=20,font=("bold",10))
label_email.place(x=xPosL,y=yPos)

#this will accept the input string text from the user.
entry_email=tk.Entry(root)
entry_email.place(x=xPosF,y=yPos)
yPos+=30
#*************************************************************************
# Here start the Upload file section:
# Course Descriptor section:
original=""
fileSize=tk.IntVar()
def open_file():
    
    file_path = askopenfile(mode='r', filetypes=[('Doc Files', '*docx')])
    if file_path is not None:
        global original,fileSize 
        original = file_path.name
        fileSize = os.path.getsize(original)
        print('Size del archivo: ' + str(fileSize))
    else:
        if original != "":
            messagebox.showerror(title="INFORMATION", message='No file chosen, Insert the file again!')
        # pass

def uploadFiles(): 
    if original != "":
        pb1 = Progressbar(
                root, 
                orient=HORIZONTAL, 
                length=300, 
                mode='determinate'
                )
        pb1.place(x=xPosF,y=yPos)
        labelPorc = ttk.Label(root,font=("bold",8))
        labelPorc.place(x=xPosF+146,y=yPos+2)
        for i in range(6):
            root.update_idletasks()
            pb1['value'] += 20
            labelPorc.config(text=str(int(pb1['value'])-20) + "%")  
            print(str(int(pb1['value'])-20))
            time.sleep(0.1)
        labelPorc.destroy()
        pb1.destroy()
        target = 'docs'
        try:
            shutil.copy(original, target)
            print("File copied successfully.")
            messagebox.showinfo(title="INFORMATION", message='File Uploaded Successfully!')
        # If source and destination are same
        except shutil.SameFileError:
            print("Source and destination represents the same file.")

        # If there is any permission issue
        except PermissionError:
            print("Permission denied.")

        # For other errors
        except:
            print("Error occurred while copying file.")
    else:
        messagebox.showerror(title="INFORMATION", message='Insert the File!')


yPos+=30

label_doc =tk.Label(root,text="Course Descriptor", width=20,font=("bold",10))
label_doc.place(x=xPosL,y=yPos)

adharbtn = Button(root, text ='Choose docx File', width=20,command = lambda:open_file()) 
adharbtn.place(x=xPosF,y=yPos)

yPos+=30

upld = tk.Button(root,text='Upload a Course Descriptor' , width=30,bg="blue",fg='white',activebackground='#0052cc', activeforeground='#aaffaa',command=uploadFiles )
upld.place(x=xPosF,y=yPos)

yPos+=30

#this creates button for submitting the details provides by the user
def show_entry_fields():
    fieldsDocx = ["+Lecturer Name+","+E-mail address+"]
    valuesDocx = [entry_name.get(), entry_email.get()]
    print("Trimestre: %s\nYear: %s" % (trimester_cb.get(), year_cb.get()))
    print("Name: %s\nEmail: %s" % (entry_name.get(), entry_email.get()))
    document = docx.Document('docs/TempleteCO.docx')
    for par in document.paragraphs:  # to extract the whole text
        for i in range(len(fieldsDocx)):
            if fieldsDocx[i] in par.text:
                if len(valuesDocx[i])>0:
                    tmp_text = par.text
                    print(tmp_text)
                    tmp_text = tmp_text.replace(fieldsDocx[i],valuesDocx[i])   
                    par.text=tmp_text
                    print(tmp_text)
                    break    
    document.save('docs/TempleteCO.docx')

tk.Button(root, text='Download a Course Outline' , width=30,bg="green",fg='white',activebackground='#0052cc', activeforeground='#aaffaa', command=show_entry_fields).place(x=xPosF,y=yPos)


#*************************************************************************
# Here start the handle of the document:
# Searching of fields inside:\
def replace_fields():
    doc = docx.Document('docs/TempleteCO.docx')



root.mainloop()



