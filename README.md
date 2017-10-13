# SeniorAssassin
### What is Senior Assassin?
Senior Assassin is a game, typically played in the ending months of the a student's senior year of high school, where students are assigned to "kill" their assigned target.  Usually, squirt guns or nerf guns are used to "kill" a target.  Each participating student is assinged a target, however the target isn't necessarily assigned their assassin.  This creates a chain of targets and assasins.  Students must "kill" their target in the given period of time, or they will be eliminated from the game.  Students that successfully kill their target, and avoid being "killed" continue into the next round.  The rounds continue until a winner is determined.  Additional rules are typically added to limit when and where a person can be assassinated.      
### What does this program do?
Essentially this program accepts an Excel file (.xlsx) with the participating students first names, last names, and email addresses.  It then randomly assigns the assassins to a unique target.  Once the targets have been assigned to the assassins, it then emails the assassins, notifying them of who their target is.
### How do I run this program?
All that is needed is the Assassin.py file, Python 3 Interpreter, the Excel file of the participating students, and a Gmail account (or another email account, however the program needs additional configuration).  **I emphasize that the Excel file used must look like Example Responses.xlsx as seen below.  This means the first names must begin in cell B2 and continue down the B column, the last names begin in cell C2 and continue down the C column, and the email addresses start in cell D2 and continue down column D.**  The reason for this formatting choice is to match the Excel file format produced by Google Forms.  For Senior Assassin games, I highly encourage using Google Forms to collect the information of participating students.  In addition, so slight configuration of the Assassin.py file is necessary.  **The Gmail address and password of the sending account needs to be added on lines 120 and 134 respectively.  Be sure to add the Gmail address and password within the quotation marks.**  Once the configuration of the Assassin.py file is done, all that needs to be done is change the current working directory in terminal to the location of the Assassin.py file and run `python3 Assassin.py`.
### What do the buttons do?
#### Initiate Workbook
This will initiate the Excel file that the user selects through the file browser.  If the initiated file does not exactly match the format specified, then the error "Not a Valid File" will appear.  If the file has been accepted, then the assignments will automatically appear in the Listbox.
#### Email
This will immediately email the assassins their target using the sender's email from the Python file.  With hundreds of participants, the process can take a while, so be patient.  
#### Save
This will create a Master_List.xlsx file that stores who is assigned to who.  This file will be saved in the same directory as the Assassin.py.  
#### Quit
This will quit this program.  
