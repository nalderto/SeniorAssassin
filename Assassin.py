import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import csv
from tkinter import *
from tkinter.filedialog import askopenfilename
import os
import random
import re

class Senior:
    def __init__(self, first_name: str, last_name: str, email: str):
        self.first_name = first_name
        self.last_name = last_name
        self.email = email

    def setTarget(self, first_name: str, last_name: str):
        self.target_first_name = first_name
        self.target_last_name = last_name

    def isTarget(self, senior: object) -> bool:
        try:
            return self.target_first_name == senior.first_name and self.target_last_name == senior.last_name
        except AttributeError:
            return False

    def getCsvRow(self) -> list:
        return [
            self.first_name,
            self.last_name,
            self.email,
            self.target_first_name,
            self.target_last_name
        ]

    def __eq__(self, other: object) -> bool:
        if isinstance(other, Senior):
            return self.email == other.email
    
    def __str__(self):
        return self.first_name + " " + self.last_name + " - " + self.email


seniors_array = []
from_address = ""
password = ""

# Tkinter Window Setup
main = Tk()
main.title("Senior Assassin")
main.geometry("530x340")

header_frame = Frame(main)
header_frame.grid(row = 0, column = 1, sticky = W, pady = 2)

# Excel Sheet Label
file_label = Label(header_frame, text = "File Name:")
file_label.grid(row = 0, column = 0, sticky = W, pady = 2)

# Excel Sheet Name Label
file_name_label = Label(header_frame, text = "No File Initialized")
file_name_label.grid(row = 0, column = 1, sticky = W, pady = 2)

# Status Label
status_label = Label(main, text = "")
status_label.grid(row = 1, column = 1, sticky = W, pady = 2)

from_address_label = Label(main, text = "From Address")
from_address = Entry(main)
from_address_label.grid(row = 2, column = 0, sticky = W, pady = 2)
from_address.grid(row = 2, column = 1, sticky="nsew", pady = 2)

password_label = Label(main, text = "Password")
password = Entry(main, show = "*", width=43)
password_label.grid(row = 3, column = 0, sticky = W, pady = 2)
password.grid(row = 3, column = 1, sticky="nsew", pady = 2)

# Assignment Scroll Table
assignment_frame = Frame(main)
assignment_frame.grid(row = 4, column = 1, sticky="nsew", pady = 2)

assassin_list = Listbox(assignment_frame, width = 43)
assassin_list.pack(side = "left", fill = "y")
scrollbar = Scrollbar(assignment_frame, orient="vertical")
scrollbar.pack(side = RIGHT, fill=Y)
scrollbar.config(command = assassin_list.yview)
assassin_list.config(yscrollcommand = scrollbar.set)

assassin_list.insert(END, "There is currently no data")


# Function for closing window
def close_window():
    main.destroy()

def csv_setup():
    global seniors_array
    seniors_array.clear()
    file_path = askopenfilename()
    extension = os.path.splitext(file_path)[1]
    basename = os.path.basename(file_path)
    if extension == ".csv":
        with open(file_path, 'r', newline='') as file:
            reader = csv.DictReader(file, delimiter=',',)
            status_label.config(text="CSV Successfully Opened")
            file_name_label.config(text=basename)
            for row in reader:
                try:
                    current_person = Senior(
                        row["First Name"],
                        row["Last Name"],
                        row["Email Address"]
                    )
                    seniors_array.append(current_person)
                except KeyError:
                    status_label.config(text="Error: Not a Valid File")

            selection_pool = list(range(len(seniors_array)))
            assassin_list.delete(0, END) # Clear ListBox

            for current in seniors_array:
                random_pick = random.choice(selection_pool)
                while current == seniors_array[random_pick] or seniors_array[random_pick].isTarget(current):
                    random_pick = random.choice(selection_pool)
                current.setTarget(seniors_array[random_pick].first_name, seniors_array[random_pick].last_name)
                selection_pool.remove(random_pick)

                # Update Listbox
                assassin_list.insert(END, str(current.first_name) + " " + str(current.last_name) + " - " + str(current.target_first_name) + " " + str(current.target_last_name) + " - " + str(current.email))

    else:
        status_label.config(text="Error: Not a Valid File")

def save():
    with open('targetList.csv', 'w', newline='') as file:
        writer = csv.writer(file)
        fields = ["First Name", "Last Name", "Email", "Target First Name", "Target Last Name"]
        writer.writerow(fields)
        for senior in seniors_array:
            writer.writerow(senior.getCsvRow())

    status_label.config(text="targetList.csv Saved!")

def email():
    if not seniors_array:
        status_label.config(text="Error: You Must Initiate a File")

    if not isEmailValid(from_address.get()):
        status_label.config(text="Error: Invalid Formatted From Email Address")

    if password == "":
        status_label.config(text="Error: Password empty")

    else:
        # SMTP Settings
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()

        try:
            server.login(from_address.get(), password.get())
        except:
            status_label.config(text="Error: Unable to authenticate.  Double check your from email address and password.")

        for current in seniors_array:
            to_address = current.email
            msg = MIMEMultipart()
            msg['From'] = from_address.get()
            msg['To'] = to_address
            msg['Subject'] = "Senior Assassin"

            # Message Body
            body = current.first_name + " " + current.last_name + ",\n" + "You have been assigned  " + current.target_first_name + " " + current.target_last_name
            msg.attach(MIMEText(body, 'plain'))

            text = msg.as_string()
            server.sendmail(from_address.get(), to_address, text)
            server.quit()

# Determine if email is valid
def isEmailValid(email):
    regex = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,7}\b'
    return re.fullmatch(regex, email)

buttons = Frame(main)
buttons.grid(row = 5, column = 1, sticky=E)

# Initiate button
init_button = Button(buttons, text = "Initiate CSV", command = csv_setup)
init_button.grid(row = 0, column = 1, padx = 2)

# Email Button
email_button = Button(buttons, text = "Email", command = email)
email_button.grid(row = 0, column = 2, padx = 2)

# Save Button
save_button = Button(buttons, text = "Save", command = save)
save_button.grid(row = 0, column = 3, padx = 2)

# Quit Button
quit_button = Button(buttons, text = "Quit", command = close_window)
quit_button.grid(row = 0, column = 4, padx = 2)

main.mainloop()
