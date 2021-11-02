import tkinter as tk
import pandas as pd
import docx

root = tk.Tk()
doc = docx.Document('TempleteCD(COMP503).docx')

canvas = tk.Canvas(root, width=700, height=500)
canvas.grid(columnspan=3)

def read_docx_table(doc, table_num=1, nhader=1):
    table = doc.tables[table_num-1]
    
    data = [[print(cell.text) for cell in row.cells] for row in table.rows]
    
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


table_num=1
nheader=0
df = read_docx_table(doc,table_num,nheader)

#Print selected row
print(df)

#print(df.iloc[2])

root.mainloop() 