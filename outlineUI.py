#*************************************************************************
# Develop of Courses Outline Hola Developer 1 esto se aprobo.
# Auckland Institute of Studies
# Developers: William Martin, June , Sun....
# Date of creation 03/11/2021
#*************************************************************************
from tkinter.filedialog import askopenfile 
from tkinter import *
from tkinter.ttk import Progressbar,Button
from PIL import ImageTk, Image
from tkinter.messagebox import showinfo
from tkinter import ttk , messagebox
import tkinter as tk,time, pandas as pd
import docx,shutil,os,globalVars as gv

class Root(Tk):
    def __init__(self):
        super(Root, self).__init__()
        self.title("Course Outline Generator")
        self.geometry('700x500')
        self.minsize(640,400)

#<editor-fold desc="Functions">
    def uploadFiles(self): 
        if gv.original != "":
            pb1 = Progressbar(
                    root, 
                    orient=HORIZONTAL, 
                    length=300, 
                    mode='determinate'
                    )
            pb1.place(x=gv.xPosF+120,y=gv.yBtnPos)
            labelPorc = ttk.Label(root,font=("bold",8))
            labelPorc.place(x=gv.xPosF+260,y=gv.yBtnPos+2)
            for i in range(6):
                root.update_idletasks()
                pb1['value'] += 20
                labelPorc.config(text=str(int(pb1['value'])-20) + "%")  
                print(str(int(pb1['value'])-20))
                time.sleep(0.1)
            labelPorc.destroy()
            pb1.destroy()
            gv.target = 'docs'
            try:
                shutil.copy(gv.original, gv.target)
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

    def open_file(self):
        file_path = askopenfile(mode='r', filetypes=[('Doc Files', '*docx')])
        if file_path is not None:
            gv.original=file_path.name
            print (gv.original)
            gv.fileSize = os.path.getsize(gv.original)
            print('Size del archivo: ' + str(gv.fileSize))
            self.uploadFiles()
        else:
            if gv.original != "":
                messagebox.showerror(title="INFORMATION", message='No file chosen, Insert the file again!')

    def segmentLine(self):
        w = tk.Canvas(self, width=690, height=3)
        w.place(y=gv.yPos,x=0)
        w.config( background=gv.bgSL)

    def show_entry_fields(self):
        # fieldsDocx = ["+Lecturer Name+","+E-mail address+"]
        # valuesDocx = [entry_name.get(), entry_email.get()]
        print("Trimestre: %s\nYear: %s" % (trimester_cb.get(), year_cb.get()))
        # print("Name: %s\nEmail: %s" % (entry_name.get(), entry_email.get()))
        document = docx.Document('docs/TempleteCO.docx')
        # for par in document.paragraphs:  # to extract the whole text
        #     for i in range(len(fieldsDocx)):
        #         if gv.fieldsDocx[i] in par.text:
        #             if len(valuesDocx[i])>0:
        #                 tmp_text = par.text
        #                 print(tmp_text)
        #                 tmp_text = tmp_text.replace(fieldsDocx[i],valuesDocx[i])
        #                 par.text=tmp_text
        #                 print(tmp_text)
        #                 break
        # document.save('docs/TempleteCO.docx')
#</editor-fold>

#<editor-fold desc="Constructors">
    def createHeader(self):
        heading = tk.Label(self, text="Course Outline Generator")
        heading.config(font=(gv.tFont, gv.mtSize),fg=gv.stColor)
        heading.pack(padx=50, pady=65)
        uplHeading = tk.Label(root, text="Upload Course Descriptor")
        uplHeading.config(font=(gv.tFont, gv.stSize),fg=gv.stColor)
        uplHeading.place(x=gv.xPosL, y=gv.yPos)
        gv.yPos+=30
        upld = tk.Button(self,text='Upload a Course Descriptor' , width=30,bg="blue",fg='white',activebackground='#0052cc', activeforeground='#aaffaa',command = self.open_file )
        upld.place(x=gv.xPosL,y=gv.yPos)
        gv.yBtnPos = gv.yPos
        gv.yPos += 30
        label_upl = tk.Label(self,text="Accepted file types: .doc and .docx (2MB limit)", width=35,font=(gv.lbFont,gv.lbSize))
        label_upl.place(x=gv.xPosL,y=gv.yPos)
        gv.yPos += 30
        self.segmentLine()
        gv.yPos += 20

    def createPeriod(self):
        
        label_trimester =tk.Label(self,text="Trimester", width=8,font=(gv.lbFont,gv.lbSize))
        label_trimester.place(x=gv.xCbxLbl,y=gv.yPos)
        gv.xCbxLbl+=100
        selected_month = tk.StringVar()
        self.trimester_cb = ttk.Combobox(self, textvariable=selected_month,width=10)
        self.trimester_cb['values'] = gv.months
        self.trimester_cb['state'] = 'readonly'
        self.trimester_cb.current(1)
        self.trimester_cb.place(x=gv.xCbxLbl,y=gv.yPos)
        gv.xCbxLbl+=110

           
        label_year =tk.Label(root,text="Year", width=9,font=(gv.lbFont,gv.lbSize))
        label_year.place(x=gv.xCbxLbl,y=gv.yPos) 
        gv.xCbxLbl+=100
        selected_year = tk.StringVar()
        year_cb = ttk.Combobox(self, textvariable=selected_year,width=10)
        year_cb['values'] = gv.years
        year_cb['state'] = 'readonly' 
        year_cb.current(1)
        year_cb.place(x=gv.xCbxLbl,y=gv.yPos)
        gv.yPos+=33
        self.segmentLine()
        gv.yPos+=20

    def createLecturerInf(self):    
        lectHeading = tk.Label(self, text="Lecturer Information")
        lectHeading.config(font=(gv.tFont, gv.stSize),fg=gv.stColor)
        lectHeading.place(x=gv.xPosL, y=gv.yPos)
        gv.yPos+=30
        var=tk.StringVar()
        tk.Radiobutton(self,text="Lecturer",padx= 5, variable= var, value="Lecturer").place(x=gv.xPosL,y=gv.yPos)
        tk.Radiobutton(self,text="Course Coordinator",padx= 20, variable= var, value="Course Coordinator").place(x=gv.xPosF,y=gv.yPos)
        gv.yPos+=40
        label_name =tk.Label(self,text="Full Name", width=8,font=(gv.lbFont,gv.lbSize))
        label_name.place(x=gv.xPosL,y=gv.yPos)

        entry_name=tk.Entry(self)
        entry_name.config(width=40)
        entry_name.place(x=gv.xPosF,y=gv.yPos)
        gv.yPos+=30

        #this creates 'Label' widget for Email.
        label_email =tk.Label(self,text="Email", width=5,font=(gv.lbFont,gv.lbSize))
        label_email.place(x=gv.xPosL,y=gv.yPos)
        entry_email=tk.Entry(self)
        entry_email.config(width=40)
        entry_email.place(x=gv.xPosF,y=gv.yPos)
        gv.yPos+=35
        self.segmentLine()
        gv.yPos+=30

    def endForm(self):
        tk.Button(self, text='Download a Course Outline' , width=30,bg="green",fg='white',activebackground='#0052cc', activeforeground='#aaffaa', command=self.show_entry_fields).place(x=gv.xPosF,y=gv.yPos)
#</editor-fold>

if __name__ == '__main__':
    root = Root()
    logoImg = (Image.open("img/logo.jpg"))
    resizedImg = logoImg.resize((120,50), Image.ANTIALIAS)
    logoImg = ImageTk.PhotoImage(resizedImg)
    label = tk.Label(root, image=logoImg).place(x=280, y=10)

   #Built of UI:
    root.createHeader()
    root.createPeriod()
    root.createLecturerInf()
    root.endForm()

    root.mainloop()
