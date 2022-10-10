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
lastMonth = changeMonth.strftime("%B")
nextMonth = changeMonth2.strftime("%B")

#dictionary containing the names of & HTML documents for email bodies
bodyTypes = {
    0: "Exit",
    1: "scheduleUpdate.html",
    2: "passcode.html",
    3: "monthlyeport.html"
}

#dictionary of subject lines to be used with corresponding email bodies
subjectTypes = {
    1: f"Schedule Update for {nextMonth}",
    2: "Pass Code Updated {:%m/%d/%y}", 
    3: f"PR #3720 Hardin County {lastMonth} Janitorial Sign In/Out"
}

#dictionary for distribution lists used for sending emails for To:
distributionLists = {
    1: "Remote Work Team",
    2: "Office Work Team; Hybrid Work Team",
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
ttk.Label(frm, text="On Call Email Template").grid(row=1,column=0)
button1 = ttk.Button(frm, text="Select On Call Email", width=30, command=lambda:submitButton(1)).grid(row=1,column=1, padx=5, pady=5)
ttk.Label(frm, text="Door Code Email Template").grid(row=2,column=0)
button2 = ttk.Button(frm, text="Select Door Code Email", width=30, command=lambda:submitButton(2)).grid(row=2,column=1, padx=5, pady=5)
ttk.Label(frm, text="Safety Walkthrough Email Template").grid(row=3,column=0)
button3 = ttk.Button(frm, text="Select Safety Walkthrough Email", width=30, command=lambda:submitButton(3)).grid(row=3,column=1, padx=5, pady=5)
ttk.Label(frm, text="RMS Personnel Update Email Template").grid(row=4,column=0)
button4 = ttk.Button(frm, text="Select Personnel Update Email", width=30, command=lambda:submitButton(4)).grid(row=4,column=1, padx=5, pady=5)
ttk.Label(frm, text="RMS Schedule Email Template").grid(row=5,column=0)
button5 = ttk.Button(frm, text="Select Schedule Email", width=30, command=lambda:submitButton(5)).grid(row=5,column=1, padx=5, pady=5)
ttk.Label(frm, text="Janitor Sign In/Out Email Template").grid(row=6,column=0)
button6 = ttk.Button(frm, text="Select Sign In/Out Email", width=30, command=lambda:submitButton(6)).grid(row=6,column=1, padx=5, pady=5)
ttk.Label(frm, text="Bimonthly Janitorial Inspection Email Template").grid(row=7,column=0)
button7 = ttk.Button(frm, text="Select Inspection Email", width=30, command=lambda:submitButton(7)).grid(row=7,column=1, padx=5, pady=5)
ttk.Label(frm, text="Pest Control Invoice & Report Email Template").grid(row=8,column=0)
button8 = ttk.Button(frm, text="Select Pest Control Email", width=30, command=lambda:submitButton(8)).grid(row=8,column=1, padx=5, pady=5)

button0 = ttk.Button(frm, text="Exit Program", command=lambda:submitButton(0)).grid(row=9,column=1, pady=10)
#buttonSubmit = ttk.Button(frm, text="Submit", command=lambda:submitButton(inputBody)).grid(row=9,column=1, pady=10)

root.mainloop()