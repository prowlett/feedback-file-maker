import os,xlrd

def excel_col_to_index_num(col_name):
    """
    Converts Excel column letter to zero-indexed number.
    e.g. "A" - > 0, "Z" -> 25, "AA" -> 26, etc.
    """
    col_num = 0
    for i in range(0,len(col_name)):
        col_num += (ord(col_name[len(col_name)-i-1])-64)*26**i
    return col_num-1

def deal_with_it(err, msg):
    """
    Prints an error message and quits.
    """
    if err == "format":
        print("Please check the format of feedback.xlsx.",end=" ")
    elif err == "file":
        print("Something has gone wrong importing feedback.xlsx.",end=" ")
    elif err == "student":
        print("Please correct error in student data.",end=" ")
    print(msg)
    import sys
    sys.exit()

# open file
fbfile = 'feedback.xlsx'
try:
    fbwb = xlrd.open_workbook(fbfile)
except FileNotFoundError:
    deal_with_it("file","Check there is a file feedback.xlsx in the same folder as this script.")
except xlrd.biffh.XLRDError:
    deal_with_it("format","Something is wrong with the feedback.xlsx file. Please check this file.")

# get the two worksheets - doesn't care what they are called, but the order matters
configsheet = fbwb.sheet_by_index(0)
try:
    fbsheet = fbwb.sheet_by_index(1)
except IndexError:
    deal_with_it("format","Does the file have two sheets - the first for configuration and the second for feedback?")

# checking format - are the first three columns Surname, Forename and Username?
try:
    if fbsheet.cell_value(0,0) != "Surname" or fbsheet.cell_value(0,1) != "Forename" or fbsheet.cell_value(0,2) != "Username":
        raise IndexError
except IndexError: # either because one of the cells checked is empty or I've thrown my own because they aren't the right column headings
    deal_with_it("format","Are the first three columns in sheet 2 'Surname', 'Forename' and 'Username' in that order?")

try:
    assignment_info = []
    for i in range(2,10):

        # Module (C3): appears on feedback pages: assignment_info[0]
        # Module code (C4): appears on feedback pages: assignment_info[1]
        # Academic year (C5): appears on feedback pages: assignment_info[2]
        # Assignment title (C6): appears on feedback pages: assignment_info[3]
        # Staff (C7): appears on feedback pages: assignment_info[4]
        # Code from Advanced Assignment tool (C8): important for Bb! Download from the Advanced Assignment Tool. The spreadsheet top line has a code in cell B1. Copy it here: assignment_info[5]
        # Bb file name (C9): important for Bb! This is the filename of the spreadsheet you got from the Advanced Assignment Tool: assignment_info[6]
        # Column holding overall marks (C10): important for Bb! Which column contains the overall marks?: assignment_info[7]
        assignment_info.append(configsheet.cell_value(i,2))
except IndexError:
    deal_with_it("Please check the format of feedback.xlsx. Is the configuration information in sheet 1 in the correct cells?")

grade_col = excel_col_to_index_num(assignment_info[7]) # column that the final grade is in

# check for missing essential data
error_flag = False
error_msg = ""
for i in range(2,fbsheet.nrows):
    surname_error_flag = False
    forename_error_flag = False
    username_error_flag = False
    grade_error_flag = False
    if fbsheet.cell_value(i,0) == "": # Surname
        surname_error_flag = True
    if fbsheet.cell_value(i,1) == "": # Forename
        forename_error_flag = True
    if fbsheet.cell_value(i,2) == "": # Username
        username_error_flag = True
    if fbsheet.cell_value(i,grade_col) == "": # grade
        grade_error_flag = True
    if surname_error_flag or forename_error_flag or username_error_flag or grade_error_flag:
        error_flag = True
        error_msg = "{}\nStudent on row {}: missing".format(error_msg,i+1)
        if surname_error_flag:
            error_msg = "{} surname".format(error_msg)
        if forename_error_flag:
            error_msg = "{} forename".format(error_msg)
        if username_error_flag:
            error_msg = "{} username".format(error_msg)
        if grade_error_flag:
            error_msg = "{} grade".format(error_msg)
if error_flag:
    deal_with_it("student",error_msg)

# check usernames are unique
usernames = []
for i in range(2,fbsheet.nrows):
    usernames.append(fbsheet.cell_value(i,2))
if len(usernames) != len(set(usernames)):
    deal_with_it("student","Please make sure all usernames are unique.")

# cycle through the worksheet processing data into
# header_info: information about the header row
# fb_data: feedback data in HTML format
#    - Text in row 1 and h in this column for a student creates this as a <h2>.
#    - Text in row 1 and hh in this column for a student creates this as a <h3>.
#    - "x" in row 1 and text in this column for a student includes this text as a <p>.
#    - Text in row 1 and y in this column for a students includes the text from row 1 as a <p>.
#    - "no" in row 1 skips this column (use it for notes to yourself).
# grade_data: just surname, forename, username and grade
mode="header"
header_info = []
fb_data = []
grade_data = []
for i in range(fbsheet.nrows):
    if i==1:
        mode="skip"
    elif i>1:
        mode="body"
        fb_data.append(list())
        grade_data.append(list())
    for j in range(fbsheet.ncols):
        if mode == "header":
            header_info.append(fbsheet.cell_value(i,j))
        elif mode == "body":
            if header_info[j] == "Surname":
                fb_data[i-2].append(fbsheet.cell_value(i,j))
                grade_data[i-2].append(fbsheet.cell_value(i,j))
            elif header_info[j] == "Forename":
                fb_data[i-2].append(fbsheet.cell_value(i,j))
                grade_data[i-2].append(fbsheet.cell_value(i,j))
            elif header_info[j] == "Username":
                fb_data[i-2].append(fbsheet.cell_value(i,j))
                grade_data[i-2].append(fbsheet.cell_value(i,j))
            elif header_info[j] == "x" and fbsheet.cell_value(i,j)!="":
                fb_data[i-2].append("<p>{}</p>".format(fbsheet.cell_value(i,j)))
            elif fbsheet.cell_value(i,j) == "h":
                fb_data[i-2].append("<h2>{}</h2>".format(header_info[j]))
            elif fbsheet.cell_value(i,j) == "hh":
                fb_data[i-2].append("<h3>{}</h3>".format(header_info[j]))
            elif fbsheet.cell_value(i,j) == "y":
                fb_data[i-2].append("<p>{}</p>".format(header_info[j]))
            elif header_info[j] == "no" or fbsheet.cell_value(i,j) == "":
                pass
            else:
                fb_data[i-2].append("<p><strong>{}</strong>: {}</p>".format(header_info[j],fbsheet.cell_value(i,j)))
            if j == grade_col:
                grade_data[i-2].append(fbsheet.cell_value(i,j))

print("{} students imported".format(len(fb_data)))

# create feedback directory if it doesn't exist
try:
    os.mkdir(os.path.join(os.getcwd(),"feedback"))
except FileExistsError: # already exists
    pass

fb_files = []

# create feedback HTML file per student
for student in fb_data:
    file_loc = "fb_{}.html".format(student[2])
    fb_files.append(file_loc)
    f = open(os.path.join(os.getcwd(),"feedback",file_loc),'w')
    f.write("""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Feedback: {} - {} {} ({})</title>
    <style>""".format(assignment_info[3],student[1],student[0],student[2]))
    f.write("""
         body {
            background-color: #c0dbff;
            text-align: center;
            font-family: Tahoma, Geneva, sans-serif;
            height: 100%;
        }
        #header, #feedback {
            width: 47%;
            margin-left: auto;
            margin-right: auto;
            margin-top: 1em;
            margin-bottom: 1em;
            background-color: #edf4fe;
            text-align: left;
            border: thin solid #002251;
        }
        h1 {
            font-size: 1.7em;
            font-weight: normal;
            color: #eee;
            background-color: #002251;
            width: 96%;
            padding: 1% 2%;
        }
        h2 {
            font-size: 1.2em;
            font-weight: normal;
            color: #eee;
            background-color: #002251;
            width: 96%;
            padding: 1% 2%;
        }
        h3 {
            font-size: 1em;
            color: #002251;
            font-weight: bold;
            margin-left: 2%;
            margin-right: 2%;
        }
        p {
            font-size: 0.8em;
            margin-left: 2%;
            margin-right: 2%;
        }
        #header p {
            font-size: 1.1em;
            margin-left: 2%;
            margin-right: 2%;
        }
        a {
            color: #155ab9;
            text-decoration: none;
        }
        a:hover {
            text-decoration: underline;
        }
        @media screen and (max-width: 1310px) {
            #header, #feedback {
                width: 70%;
            }
        }
        @media screen and (max-width: 670px) {
            #header, #feedback {
                width: 95%;
            }
        }
        @media screen and (max-width: 350px) {
            #header, #feedback {
                width: 100%;
            }
        }
""")
    f.write("""    </style>
</head>

<body>
<div id="header">
    <h1>Feedback: {}</h1>
    <p><i>{}: {}</i> ({}),<br>{}</p>
    <p>Student: {} {} ({})</p>
</div>
<div id="feedback">
""".format(assignment_info[3],assignment_info[1],assignment_info[0],assignment_info[2],assignment_info[4],student[1],student[0],student[2]))
    for i in range(3,len(student)):
        f.write("\t{}\n".format(student[i]))
    f.write("</div>\n</body>\n</html>")
    f.close()

# Bb bit

# xls version - open file for editing
import xlwt
wb = xlwt.Workbook()
worksheet = wb.add_sheet('Sheet 1')

# xlsx version - open file for editing
# N.B. I made this work before I realised Bb tool wants an xls, not an xlsx. Leaving it here in case this changes.
# import xlsxwriter
# workbook = xlsxwriter.Workbook('grades.xlsx') # this would need to change to match file_loc bit below
# worksheet = workbook.add_worksheet()


# essential first row
worksheet.write(0,0,'NB: Do not change or delete the information in this row.')
worksheet.write(0,1,assignment_info[5])

# column headings
worksheet.write(1,0,"Username")
worksheet.write(1,1,"First Name")
worksheet.write(1,2,"Last Name")
worksheet.write(1,5,"Grade")
worksheet.write(1,6,"Feedback")

# Fill in spreadsheet of id info and grades for each student
for i in range(2,len(grade_data)+2):
    worksheet.write(i,0,grade_data[i-2][2]) # Username
    worksheet.write(i,1,grade_data[i-2][1]) # First Name
    worksheet.write(i,2,grade_data[i-2][0]) # Last name
    worksheet.write(i,5,grade_data[i-2][3]) # Grade
    worksheet.write(i,6,"Please see attached file")

# xls - output file
if assignment_info[6][-4:] == ".xls":
    file_loc = assignment_info[6]
else:
    file_loc = "{}.xls".format(assignment_info[6])
wb.save(os.path.join(os.getcwd(),"feedback",file_loc))
fb_files.append(file_loc)

# xlsx - finish
#workbook.close()

# now zip up ./feedback/* into ./feedback.zip

from zipfile import ZipFile

zip = ZipFile('feedback.zip','w')
os.chdir("./feedback")
for file in fb_files:
    zip.write(file)
