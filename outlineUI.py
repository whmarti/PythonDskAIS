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

#<editor-fold desc="Logic Functions">
    def uploadFiles(self): 
        if gv.fName != "":
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
            # print (gv.original)
            pos = gv.original.rfind("/")
            gv.fName = gv.original[pos+1:len(gv.original)+1]
            print (gv.fName)
            gv.fileSize = os.path.getsize(gv.original)
            print('Size del archivo: ' + str(gv.fileSize))
            gv.state = True
            self.ableControls(gv.state)
            self.uploadFiles()
        else:
            if gv.original == "":
                gv.state = False
                self.ableControls(gv.state)
                messagebox.showerror(title="INFORMATION", message='No file chosen, Insert the file again!')

    def segmentLine(self):
        w = tk.Canvas(self, width=690, height=3)
        w.place(y=gv.yPos,x=0)
        w.config( background=gv.bgSL)

    def update_document(self):
        valuesDocx = [gv.entry_name.get(), gv.entry_email.get()]
        print("Trimestre: %s\nYear: %s" % (gv.trimester_cb.get(), gv.year_cb.get()))
        print("Name: %s\nEnail: %s" % (gv.entry_name.get(), gv.entry_email.get())) 
        res = self.validateForm()
        if res:
            document = docx.Document('docs/'+gv.fName)
            for par in document.paragraphs:  # to extract the whole text
                for i in range(len(gv.fieldsDocx)):
                    if gv.fieldsDocx[i] in par.text:
                        if len(valuesDocx[i])>0:
                            tmp_text = par.text
                            print(tmp_text)
                            tmp_text = tmp_text.replace(gv.fieldsDocx[i],valuesDocx[i])
                            par.text=tmp_text
                            print(tmp_text)
                            break
            document.save('docs/'+gv.fName)
            messagebox.showinfo(title="INFORMATION", message='File updated, Course descriptor loaded.')
        # else:
        #      messagebox.showinfo(title="INFORMATION", message='There are inconsistencies in the information provided, please verify.')

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
        gv.upld = tk.Button(self,text='Upload a Course Descriptor' , width=30,bg="blue",fg='white',activebackground='#0052cc', activeforeground='#aaffaa',command = self.open_file )
        gv.upld.place(x=gv.xPosL,y=gv.yPos)
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
        gv.trimester_cb = ttk.Combobox(self)
        gv.trimester_cb['values'] = gv.months
        gv.trimester_cb['state'] = 'readonly'
        gv.trimester_cb.current(1)
        gv.trimester_cb.config(width=10)
        gv.trimester_cb.place(x=gv.xCbxLbl,y=gv.yPos)
        gv.xCbxLbl+=110

        label_year =tk.Label(root,text="Year", width=9,font=(gv.lbFont,gv.lbSize))
        label_year.place(x=gv.xCbxLbl,y=gv.yPos) 
        gv.xCbxLbl+=100
        gv.year_cb = ttk.Combobox(self)
        gv.year_cb['values'] = gv.years
        gv.year_cb['state'] = 'readonly'
        gv.year_cb.config(width=10)
        gv.year_cb.current(0)
        gv.year_cb.place(x=gv.xCbxLbl,y=gv.yPos)
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

        gv.entry_name=tk.Entry(self)
        gv.entry_name.config(width=40)
        gv.entry_name.place(x=gv.xPosF,y=gv.yPos)
        gv.yPos+=30

        #this creates 'Label' widget for Email.
        label_email =tk.Label(self,text="Email", width=5,font=(gv.lbFont,gv.lbSize))
        label_email.place(x=gv.xPosL,y=gv.yPos)
        gv.entry_email=tk.Entry(self)
        gv.entry_email.config(width=40)
        gv.entry_email.place(x=gv.xPosF,y=gv.yPos)
        gv.yPos+=35
        self.segmentLine()
        gv.yPos+=30

    def endForm(self):
        gv.state = False
        self.ableControls(gv.state)
        tk.Button(self, text='Download a Course Outline' , width=30,bg="green",fg='white',activebackground='#0052cc', activeforeground='#aaffaa', command=self.update_document).place(x=gv.xPosF+70,y=gv.yPos)
#</editor-fold>

#<editor-fold desc="Form Functions">
    def ableControls(self,state):
        if state:
            gv.trimester_cb['state'] = tk.NORMAL
            gv.year_cb['state'] = tk.NORMAL  
            gv.entry_name['state'] = tk.NORMAL
            gv.entry_email['state'] = tk.NORMAL
        else:
            gv.trimester_cb['state'] = tk.DISABLED
            gv.year_cb['state'] = tk.DISABLED
            gv.entry_name['state'] = tk.DISABLED
            gv.entry_email['state'] = tk.DISABLED

    def validateForm(self):
        try:
            style = ttk.Style()
            style.configure("TCombobox", fieldbackground="yellow")
            res = False
            if gv.fName == "":
                messagebox.showerror(title="INFORMATION", message='The File was not chosen, please choose the Course Descriptor!')
                gv.trimester_cb.focus_set()
                return False
            elif gv.trimester_cb.get()=='Select >':
                # gv.trimester_cb.config(fg = 'blue')
                messagebox.showerror(title="INFORMATION", message='The Trimester was not chosen, choose the Trimester!')
                gv.trimester_cb.focus_set()
                return False
            elif gv.year_cb.get()=='Select >':
                messagebox.showinfo(title="INFORMATION", message='The Year was not chosen, choose the Year!')
                gv.year_cb.focus_set()
            elif len(gv.entry_name.get().strip()) == 0:
                messagebox.showinfo(title="INFORMATION", message='The Full name is empty, please enter It!')
                gv.entry_name.focus_set()
                return False
            elif len(gv.entry_email.get().strip()) == 0:
                messagebox.showinfo(title="INFORMATION", message='The Email is empty, please enter It!')
                gv.entry_email.focus_set()
                return False
            else:
                return True
        except Exception as ep:
            messagebox.showerror('Error: ', ep)
        return res


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
