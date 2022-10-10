import sys
import win32com.client as win32
from datetime import date, datetime
from dateutil.relativedelta import relativedelta
from tkinter import *
from tkinter import ttk

#this variable creates the last month so that the forms can be submitted for the previous month
d = date.today()
changeMonth = d.replace(day=1) - relativedelta(months=1)
changeMonth2 = d.replace(day=1) + relativedelta(months=1)
lastMonth = changeMonth.strftime("%B %Y")
nextMonth = changeMonth2.strftime("%B %Y")

#dictionary containing the names of & HTML documents for email bodies
bodyTypes = {
    0: "Exit",
    1: "passcode.html",
    2: "scheduleUpdate.html",
    3: "monthlyeport.html"
}

#dictionary of subject lines to be used with corresponding email bodies
subjectTypes = {
    1: "Pass Code Updated {:%m/%d/%y}",
    2: f"Schedule Update for {nextMonth}", 
    3: f"{lastMonth} Report"
}

#dictionary for distribution lists used for sending emails for To:
distributionLists = {
    1: "Office Work Team; Hybrid Work Team",
    2: "Remote Work Team",
    3: "john.smith@example.com"
}

#selection of a body or a to exit the program
def submitButton(inputBody):
    
    if inputBody > 0:
        selectBody = bodyTypes[inputBody]
        selectSubject = subjectTypes[inputBody]
        selectTo = distributionLists[inputBody]
    else:
        sys.exit("Program Exited")

    #reference for Outlook to pull text from relevant HTML document selected
    with open(selectBody,"r", encoding="utf-8") as f:
        emailBody = f.read()

    #connecting to Outlook application and namespace
    olApp = win32.Dispatch("Outlook.Application")
    olNS = olApp.GetNameSpace("MAPI")

    #creating the email using information choose above
    #the date format works when there is a date for today present in the subject line
    newMail = olApp.CreateItem(0)
    newMail.To = selectTo
    newMail.Subject = selectSubject.format(date.today())
    newMail.BodyFormat = 1 
    newMail.HTMLBody = emailBody
    newMail.Display(True)

#***tkinter GUI application for selecting email to send***

#create window
root = Tk()
root.title("Email Template Selector")
frm = ttk.Frame(root,padding= 10).grid()

inputBody = 0

#basic button & label
olLabel = ttk.Label(frm, text="Outlook Form").grid(row=0, column=0, pady=10)
#ttk.Button(frm, text="Quit", command=root.destroy).pack()

#labels for email options
ttk.Label(frm, text="Pass Code Email Template").grid(row=1,column=0, padx=5, pady=5)
button1 = ttk.Button(frm, text="Select", width=10, command=lambda:submitButton(1)).grid(row=1,column=1, padx=5, pady=5)
ttk.Label(frm, text="Update Schedule Email Template").grid(row=2,column=0, padx=5, pady=5)
button2 = ttk.Button(frm, text="Select", width=10, command=lambda:submitButton(2)).grid(row=2,column=1, padx=5, pady=5)
ttk.Label(frm, text="Monthly Email Template").grid(row=3,column=0, padx=5, pady=5)
button3 = ttk.Button(frm, text="Select", width=10, command=lambda:submitButton(3)).grid(row=3,column=1, padx=5, pady=5)

button0 = ttk.Button(frm, text="Exit Program", command=lambda:submitButton(0)).grid(row=9,column=1, pady=10)
#buttonSubmit = ttk.Button(frm, text="Submit", command=lambda:submitButton(inputBody)).grid(row=9,column=1, pady=10)

root.mainloop()