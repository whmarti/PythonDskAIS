#*************************************************************************
# Develop of Courses Outline 
# Auckland Institute of Studies
# Developers: William Martin, June , Sun....
# Date: 03/11/2021
# File: Definition of Global variables
#*************************************************************************
#File of variables
from tkinter.constants import FALSE


global fileSize,target,originalDoc,targetDoc
originalDoc = 'docOrigin/TempleteCO.docx'
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
years = ('Select >', '2021', '2022', '2023')
months = ('Select >','1', '2', '3')
fieldsDocx = ["+Lecturer Name+","+Room #+","+815-1717+ ext.#   ++Delete if adjunct++","+E-mail address+","+Contact time for this course+"]
# Final values:
nameF = ''
roomF = ''
phoneF = ''
hourF = ''
#-------------------------------------------------------------------------

#Form Variables ----------------------------------------------------------

# Controls:
global upld
global trimester_cb, year_cb , empty_ch
rbCoursePerson=''
global entry_name, entry_room, entry_phone1, entry_phone2, entry_ext, entry_email, entry_contHour, entry_contMinute
#Max. fields length:
mxNa=50
mxRo=10
mxph1=4
mxph2=10
mxph3=4
mxEm=70
# Validation Logic:
global vcmd   #Validate the integrity of the value.
emailRegex = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'   #Validate the integrity of the email.
textRegex = r'^[A-Za-z@. ]*$'   #Validate the integrity of text fields.
alphanumRegex = r'^[A-Za-z0-9 ]*$'   #Validate the integrity of alphanumeric fields.
hourRegex = r'^(0[0-9]|1[0-9]|2[0-3])$'
minRegex = r'^[0-5][0-9]$'
#-------------------------------------------------------------------------
