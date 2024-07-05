import csv
import win32com.client

# ENTER EXACT FILE NAME BELOW
file_name = "test_batch.csv"

print("> Finding file:", file_name, "\n")
file = open(file_name, "r")
batch = list(csv.reader(file, delimiter = ","))
file.close()

email_index = batch[0].index("Work Emails (Unlocked)")
email_list = []

for row in batch:
    if(row[email_index]):
        if('@' in row[email_index]):
            emails = row[email_index].split(', ')
            for e in emails:
                email_list.append(e)
    else:
        if('@' in row[email_index + 1]):
            emails = row[email_index + 1].split(', ')
            for e in emails:
                email_list.append(e)
print('> Now sending [', len(email_list), "] emails:")
print(email_list, "\n")

emailstr = ';'.join(email_list)

# generate emails
for email in email_list:
    ol = win32com.client.Dispatch('Outlook.Application')
    olmailitem = 0x0
    newmail = ol.CreateItem(olmailitem)
    newmail.Subject = 'Introduction Call with GreenHills'

    newmail.To = email
    newmail.CC = ''
    newmail.BCC = ''

    # newmail.Body = ''
    newmail.HTMLBody = "<p>My name is Ashley Parot, and I am an associate at Greenhills Ventures. We recently came across your company and after a preliminary cursory review of your company\'s mission statement and would like to learn more. <br><br>To know more about our company, please visit <a href='http://www.greenhillsventures.com/'>GreenHills Ventures</a> and our company overview <a href='http://www.greenhillsventures.com/GH_Company%20Overview%202019.pdf'>GH_Company Overview</a>. Please send us a Non-Confidential Executive Summary or Investor Deck to familiarize ourselves with your company and possibly have a brief introductory call with our management team. <br><br>Thank you in advance and we hope to hear from you soon. Enjoy the rest of your day.</p>"
    # attach = 'C:\\Users\\admin\\Desktop\\Python\\Sample.xlsx'
    # newmail.Attachments.Add(attach)

    # newmail.Display()
    newmail.Send()

print("> Done")