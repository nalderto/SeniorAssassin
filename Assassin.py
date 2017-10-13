import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import openpyxl
from tkinter import *
from tkinter.filedialog import askopenfilename
import os
import random
import copy
seniors_array = []
wb = None
ws = None



#Tkinter Window Setup
main = Tk()
main.title("Senior Assassin")
main.geometry("420x300")

#Excel Sheet Label
excel_sheet_label = Label(main, text = "File Name:")
excel_sheet_label.place(x = 5, y = 50)

#Excel Sheet Name Label
excel_sheet_name_label = Label(main, text = "No File Initialized")
excel_sheet_name_label.place(x = 80, y = 50)

#Status Label
status_label = Label(main, text = "")
status_label.pack(side = TOP)

#Assignment Scroll Table
frame = Frame(main)
frame.place(x = 5, y = 80)
assassin_list = Listbox(frame, width = 43)
assassin_list.pack(side = "left", fill = "y")
scrollbar = Scrollbar(frame, orient="vertical")
scrollbar.pack(side = RIGHT, fill=Y)
scrollbar.config(command = assassin_list.yview)
assassin_list.config(yscrollcommand=scrollbar.set)

assassin_list.insert(END, "There is currently no data")


#Function for closing window
def close_window():
    main.destroy()

#openpyxl Workbook Setup
def workbook_setup():
    global seniors_array
    global wb
    global ws
    file_path = askopenfilename()
    extension = os.path.splitext(file_path)[1]
    basename = os.path.basename(file_path)
    if extension ==".xlsx":
        wb = openpyxl.load_workbook(file_path)
        ws = wb.worksheets[0]
        if ws.cell(row = 1, column = 2).value == "First Name":
            status_label.config(text="Workbook Successfully Opened")
            excel_sheet_name_label.config(text=basename)
            assassin_list.delete(0, END)
            for row in range(2, ws.max_row + 1):
                current_person = []
                for column in range(2, 5):
                    current_person.append(ws.cell(row=row, column=column).value)
                seniors_array.append(current_person)
            selection_pool = copy.deepcopy(seniors_array)

            for current in range(0, len(seniors_array)):
                random_pick = random.choice(selection_pool)
                while seniors_array[current] == random_pick:
                    random_pick = random.choice(selection_pool)
                seniors_array[current].append(random_pick[0])
                seniors_array[current].append(random_pick[1])
                selection_pool.remove(random_pick)

                #Update Listbox
                assassin_list.insert(END, str(seniors_array[current][0]) + " " + str(seniors_array[current][1]) + " - " + str(seniors_array[current][3]) + " " + str(seniors_array[current][4]) + " - " + str(seniors_array[current][2]))

        else:
            status_label.config(text="Error: Not a Valid File")
            wb = None
            ws = None
    else:
        status_label.config(text="Error: Not a Valid File")
        wb = None
        ws = None
    print(seniors_array)

def save():
    #Create New Workbook
    master_wb = openpyxl.Workbook()
    master_ws = master_wb.active
    master_ws.cell(row=1, column=1).value = "Assassin First"
    master_ws.cell(row=1, column=2).value = "Assassin Last"
    master_ws.cell(row=1, column=3).value = "Assassin Email"
    master_ws.cell(row=1, column=4).value = "Target First"
    master_ws.cell(row=1, column=5).value = "Target Last"
    currentDirectory = os.getcwd()
    for current in range(0, len(seniors_array)):
        row = current + 2
        master_ws.cell(row=row, column=1).value = seniors_array[current][0]
        master_ws.cell(row=row, column=2).value = seniors_array[current][1]
        master_ws.cell(row=row, column=3).value = seniors_array[current][2]
        master_ws.cell(row=row, column=4).value = seniors_array[current][3]
        master_ws.cell(row=row, column=5).value = seniors_array[current][4]
        master_wb.save(str(currentDirectory) + "/Master_List.xlsx")

def email():
    if wb == None or ws == None:
        status_label.config(text="Error: You Must Initiate a File")

    else:
        for current in range(0, len(seniors_array)):

            #To, From, Message
            fromaddr = " " #Enter a Gmail email address in the string
            toaddr = seniors_array[current][2]
            msg = MIMEMultipart()
            msg['From'] = fromaddr
            msg['To'] = toaddr
            msg['Subject'] = "Senior Assassin"

            #Message Body
            body = seniors_array[current][0] + " " + seniors_array[current][1]+ ",\n" + "You need to assassinate " + seniors_array[current][3] + " " + seniors_array[current][4]
            msg.attach(MIMEText(body, 'plain'))

            #SMTP Settings
            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.starttls()
            server.login(fromaddr, " " ) #Enter the password for the Gmail account in the string
            text = msg.as_string()
            server.sendmail(fromaddr, toaddr, text)
            server.quit()


#Initiate button
init_button = Button(main, text="Initiate Workbook", command = workbook_setup)
init_button.place(x = 70, y = 265)

#Email Button
email_button = Button(main, text="Email", command = email)
email_button.place(x = 210, y = 265)

#Save Button
save_button = Button(main, text="Save", command = save)
save_button.place(x = 270, y = 265)

#Quit Button
quit_button = Button(main, text="Quit", command = close_window)
quit_button.place(x = 340, y = 265)

main.mainloop()
