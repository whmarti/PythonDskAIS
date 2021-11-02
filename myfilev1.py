import tkinter as tk
from tkinter.constants import ANCHOR, NW, RIGHT, Y
from PIL import ImageTk, Image
import pandas as pd
import docx


def read_docx_table(doc, table_num=1, nhader=1):
    table = doc.tables[table_num-1]
    
    data = [[cell.text for cell in row.cells] for row in table.rows]
    pd.set_option('display.max_rows', None)
    pd.set_option('display.max_columns', None)
    pd.set_option('display.width', None)
    # pd.set_option('display.max_colwidth', -1)
    df = pd.DataFrame(data)

    if nhader == 1 :
        df = df.rename(columns=df.iloc[0]).drop(df.index[0]).reset_index(drop=True)
    elif nheader == 2:
        outside_col, inside_col = df.iloc[0], df.iloc[1]
        hier_index = pd.MultiIndex.from_tuples(list(zip(outside_col, inside_col)))
        df = pd.DataFrame(data, columns=hier_index).drop(df.index[[0,1]]).reset_index(drop=True)
    elif nheader > 2:
        print("Not supported")
        df = pd.DataFrame()
    return df


root = tk.Tk()
root.configure(background='#FFFFFF')
root.title("Course Outline Generator")
root.geometry('700x500')

doc = docx.Document('TempleteCD(COMP503).docx')
# sb = tk.Scrollbar(root)
# sb.pack(side= RIGHT, fill= Y)

#text.config(state="normal")
#text.config(state="disable")

#canvas = tk.Canvas(root, width=700, height=500)
#canvas.grid(columnspan=3)

logoImg = (Image.open("logo.jpg"))
resizedImg = logoImg.resize((120,50), Image.ANTIALIAS)
logoImg = ImageTk.PhotoImage(resizedImg)
label = tk.Label(root, image=logoImg).place(x=280, y=10)
heading = tk.Label(root, text="Course Outline Generator")
heading.config(font=('Arial', 18))
heading.pack(padx=50, pady=65)
text = tk.Text(width=100, height=30)
text.pack(padx=20, pady=80)





table_num=1
nheader=0
df = read_docx_table(doc,table_num,nheader)

#Get the first Column
firstColumn = pd.Series(df[:][0], name="s")

#Filter rows that their values contain 'Learning Outcomes'
lo = df[firstColumn.isin(['Learning\nOutcomes']) == True]

#Get rows out of 'lo' series starts its second row because the first low is 'The learners will be able to:'
text.insert(tk.INSERT, lo.iloc[1:, 1:3])
print()

root.mainloop() 