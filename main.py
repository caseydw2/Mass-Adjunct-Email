import win32com.client as win32
import pandas as pd

adjunct_data = pd.read_excel("Adjunct Appreciation Reception Invite List.xlsx")
leftovers = pd.read_excel("Leftovers.xlsx")

print(leftovers.head)

body = """Hello, [FIRST_NAME]!\n
On behalf of the Center for Teaching and Learning and the [DIVISION] division, I want to let you know that your work as an instructor for [DEPARTMENT] is invaluable.  Our students’ experience at Longview is directly impacted by your good work, so we want to thank you by honoring you at our Adjunct Appreciation Reception.
Please see the attachment for your invitation to an afternoon honoring YOU.
We are excited to see you there!
Casey
"""

def body_substitute(body: str, firstName, division, department):
    body_sub = body.replace("[FIRST_NAME]",firstName).replace("[DIVISION]", division).replace("[DEPARTMENT]",department)
    return body_sub


def sendEmail(df:pd.DataFrame ,body):
    for index,series in df.iterrows():
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.Importance = 2
        #mail.SentOnBehalfOfName = "Casey.Wheaton-Werle@mcckc.edu"
        mail.To = series["Email"]
        mail.Subject = "Important: Adjunct Appreciation Reception Invitation"
        mail.Attachments.Add(r"C:\Users\E1448105\OneDrive - Metropolitan Community College - Kansas City\Programs\Mass Adjunct Email\Adjunct Appreciation Reception .jpg")
        mail.Body = body_substitute(body,series["First Name"],series["Division"],series["Department"])
        mail.Send()

