# Markdown output
# Adapted from the HTML version 13/02/2021

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
    for i in range(2,7):

        # Module (C3): appears on feedback pages: assignment_info[0]
        # Module code (C4): appears on feedback pages: assignment_info[1]
        # Academic year (C5): appears on feedback pages: assignment_info[2]
        # Assignment title (C6): appears on feedback pages: assignment_info[3]
        # Staff (C7): appears on feedback pages: assignment_info[4]
        # C8-C10: ignored in this version - used by Bb and this file is not targetting Bb.
        assignment_info.append(configsheet.cell_value(i,2))
except IndexError:
    deal_with_it("Please check the format of feedback.xlsx. Is the configuration information in sheet 1 in the correct cells?")

# check for missing essential data
error_flag = False
error_msg = ""
for i in range(2,fbsheet.nrows):
    surname_error_flag = False
    forename_error_flag = False
    username_error_flag = False
    if fbsheet.cell_value(i,0) == "": # Surname
        surname_error_flag = True
    if fbsheet.cell_value(i,1) == "": # Forename
        forename_error_flag = True
    if fbsheet.cell_value(i,2) == "": # Username
        username_error_flag = True
    if surname_error_flag or forename_error_flag or username_error_flag:
        error_flag = True
        error_msg = "{}\nStudent on row {}: missing".format(error_msg,i+1)
        if surname_error_flag:
            error_msg = "{} surname".format(error_msg)
        if forename_error_flag:
            error_msg = "{} forename".format(error_msg)
        if username_error_flag:
            error_msg = "{} username".format(error_msg)
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
# fb_data: feedback data in Markdown format
#    - Text in row 1 and h in this column for a student creates this as a ##.
#    - Text in row 1 and hh in this column for a student creates this as a ###.
#    - "x" in row 1 and text in this column for a student includes this text as a paragraph.
#    - Text in row 1 and y in this column for a students includes the text from row 1 as a paragraph.
#    - "no" in row 1 skips this column (use it for notes to yourself).
mode="header"
header_info = []
fb_data = []
for i in range(fbsheet.nrows):
    if i==1:
        mode="skip"
    elif i>1:
        mode="body"
        fb_data.append(list())
    for j in range(fbsheet.ncols):
        if mode == "header":
            header_info.append(fbsheet.cell_value(i,j))
        elif mode == "body":
            if header_info[j] == "Surname":
                fb_data[i-2].append(fbsheet.cell_value(i,j))
            elif header_info[j] == "Forename":
                fb_data[i-2].append(fbsheet.cell_value(i,j))
            elif header_info[j] == "Username":
                fb_data[i-2].append(fbsheet.cell_value(i,j))
            elif header_info[j] == "x" and fbsheet.cell_value(i,j)!="":
                fb_data[i-2].append("{}\n".format(fbsheet.cell_value(i,j)))
            elif fbsheet.cell_value(i,j) == "h":
                fb_data[i-2].append("## {}\n".format(header_info[j]))
            elif fbsheet.cell_value(i,j) == "hh":
                fb_data[i-2].append("### {}\n".format(header_info[j]))
            elif fbsheet.cell_value(i,j) == "y":
                fb_data[i-2].append("{}\n".format(header_info[j]))
            elif header_info[j] == "no" or fbsheet.cell_value(i,j) == "":
                pass
            else:
                fb_data[i-2].append("**{}**: {}\n".format(header_info[j],fbsheet.cell_value(i,j)))

print("{} students imported".format(len(fb_data)))

# create feedback directory if it doesn't exist
try:
    os.mkdir(os.path.join(os.getcwd(),"feedback"))
except FileExistsError: # already exists
    pass

fb_files = []

# create feedback Markdown file per student
for student in fb_data:
    file_loc = "fb_{}.md".format(student[2])
    fb_files.append(file_loc)
    f = open(os.path.join(os.getcwd(),"feedback",file_loc),'w')
    f.write("""# Feedback: {}\n\n*{}: {}* ({}),  \n{}\n\nStudent: {} {} ({})\n\n""".format(assignment_info[3],assignment_info[1],assignment_info[0],assignment_info[2],assignment_info[4],student[1],student[0],student[2]))
    for i in range(3,len(student)):
        f.write("{}\n".format(student[i]))
    f.close()
