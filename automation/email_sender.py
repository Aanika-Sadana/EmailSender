import csv

import win32com.client

#############################################################
# ENTER EXACT FILE NAME BELOW & SAVE BEFORE RUNNING
file_name = "test_batch.csv"
#############################################################

print("> Finding file:", file_name, "\n")
file = open(file_name, "r")
batch = list(csv.reader(file, delimiter=","))
file.close()

name_index = 0
company_index = batch[0].index("Organization")
email_index = batch[0].index("Work Emails (Unlocked)")
email_list = []

for row in batch:
    if row[email_index]:
        if "@" in row[email_index]:
            emails = row[email_index].split(", ")
            email_list.append((row[name_index], row[company_index], emails[0]))
    else:
        if "@" in row[email_index + 1]:
            emails = row[email_index + 1].split(", ")
            email_list.append((row[name_index], row[company_index], emails[0]))
print("> Now sending [", len(email_list), "] emails:")
print(email_list, "\n")

# generate emails
for record in email_list:
    ol = win32com.client.Dispatch("Outlook.Application")
    olmailitem = 0x0
    newmail = ol.CreateItem(olmailitem)

    name = " Sir/Madam"
    if record[0]:
        name = " " + record[0].split()[0]

    recipients = newmail.Recipients.Add(record[2])
    # to CC/BCC multiple emails, separate email addresses with semicolon (;)
    # for example, 'email1@example.com;email2@example.com'

    newmail.CC = "aparot@greenhillsventures.com;asadana@greenhillsventures.com;ptan@greenhillsventures.com;sbryant@greenhillsventures.com"
    newmail.BCC = ""

    subject = "Introduction Call - " + record[1].upper() + " ACTION NEEDED"
    html_msg = (
        '<p style="font-family:Calibri;font-size:14px">Hi'
        + name
        + ", <br><br>My name is Ashley Parot, Associate at GreenHills Ventures, a Private Investment Company located in New York City. Our firm seeks to make direct minority investment positions in companies incorporated in the United States. We received a brief description of your company which piqued our interest. <br><br> Our firm invests $1.0 - $25.0 million in start-ups to early-stage companies in industries we have industry experience to lead Pre-Seed, Seed or co-lead follow on Series A investments. Please send us a Non-Confidential updated brief description or summary of the company for our review and schedule a call with our management team. <br><br>Thank you in advance and look forward to our meeting. <br><br><br>Sincerely yours, <br><br><b>Ashley Parot</b> | Assistant Corporate Business Strategy Team <br><b><font color=\"blue\">GreenHills Ventures</font></b><br>Graybar Building | 420 Lexington Avenue, 3<sup>rd</sup> Floor, Suite 300 | New York, N.Y. 10170<br>Chrysler Building | 405 Lexington Avenue | New York, N.Y. 10174 <br>T: (212) 794-4027 | E: <b><a href='mailto:aparot@greenhillsventures.com'>aparot@greenhillsventures.com</a></b> <br>New York | Madrid | Ireland | Singapore | Luxembourg <br>Website: <b><a href='https://www.greenhillsventures.com/'>GreenHills Ventures</a></b> | <b><a href='https://www.greenhillsventures.com/GH_Company%20Overview%202019.pdf'>GreenHills Overview</a></b></p><p style=\"font-family:Calibri;font-size:10px;color:lightgray\">**************************************************************************************************************************************************************************************************************************************************************************************************************************** <br><br> </p><p style=\"font-family:Calibri;font-size:11px\">NOTE: THIS MESSAGE IS INTENDED ONLY FOR THE USE OF THE RECIPIENT TO WHOM IT IS ADDRESSED, AND MAY CONTAIN INFORMATION THAT IS PRIVILEGED, CONFIDENTIAL AND EXEMPT FROM DISCLOSURE UNDER APPLICABLE LAW. IF THE READER OF THIS MESSAGE IS NOT THE INTENDED RECIPIENT, OR THE EMPLOYEE OR AGENT RESPONSIBLE FOR DELIVERING THE MESSAGE TO THE INTENDED RECIPIENT, YOU ARE HEREBY NOTIFIED THAT ANY DISSEMINATION, DISTRIBUTION OR COPYING OF THIS COMMUNICATION IS STRICTLY PROHIBITED. IF YOU HAVE RECEIVED THIS COMMUNICATION IN ERROR, PLEASE DESTROY IT AND NOTIFY US IMMEDIATELY BY TELEPHONE OR RETURN EMAIL. <br><br> Pursuant to IRS Circular 230, we hereby inform you that any U.S. federal tax advice set forth herein was not intended or written by GreenHills Ventures, LLC,, or any other wholly owned subsidiary to be used, and cannot be used, by you or any taxpayer, for the purpose of avoiding any penalties that may be imposed on you or any other person under the Internal Revenue Code.</p>"
    )
    newmail.Subject = subject
    newmail.HTMLBody = html_msg
    # attach = 'C:\\Users\\admin\\Desktop\\Python\\Sample.xlsx'
    # newmail.Attachments.Add(attach)

    # newmail.Display()
    if newmail.Recipients.ResolveAll():
        newmail.Send()

print("> Done")
