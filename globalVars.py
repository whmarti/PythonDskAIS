#*************************************************************************
# Develop of Courses Outline 
# Auckland Institute of Studies
# Developers: William Martin, June , Sun....
# Date: 03/11/2021
# File: Definition of Global variables
#*************************************************************************
#File of variables
from tkinter.constants import FALSE


global fileSize,target,targetDoc
targetDoc = 'TempleteCO.docx'
original = ''
fName = ''

#Variables of position of controls --------------------------------------

# Controls Margins:
xPosL=50  #Labels
xPosF=160 #Fields
yPos=110  #Vertical
yBtnPos=110  #Button Upload


#Comboboxes:
# ycbxPos=120  selected_month = tk.StringVar()
xCbxLbl=120
lblFile = ''    #File name
#-------------------------------------------------------------------------

#Variables of controls ----------------------------------------------------

# Titles:
tFont='Arial'
mtSize=18       # Main Title
stSize=14       # Sub Title
mtColor='Black' # Main Color
stColor='Black' # Sub Color
# Labels:
lbFont='Arial'
lbSize=10
#SegmentLine
bgSL='#5a7684'   # Seg Line Color
#-------------------------------------------------------------------------

#Variables of logic ----------------------------------------------------
# Logic:
state = FALSE
# Controls:
years = ('Select >','2020', '2021', '2022', '2023')
months = ('Select >','1', '2', '3')
fieldsDocx = ["+Lecturer Name+","+E-mail address+"]

#-------------------------------------------------------------------------

#Form Variables ----------------------------------------------------------

# Controls:
global upld
global trimester_cb, year_cb, entry_name, entry_email


#-------------------------------------------------------------------------
