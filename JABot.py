


'''
Setting up a secure, encrypted SMTP connection.

creates a secure connection with Gmail’s SMTP server, using the SMTP_SSL() 
of smtplib to initiate a TLS-encrypted connection. The default context of ssl validates the host name and its certificates and optimizes the security of the connection. 
'''

import smtplib, ssl
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import xlrd # excel manipulations
from datetime import datetime 
import openpyxl # excel manipulations
import getpass #secure password input
from email_sender import send_mail
from VerifyEmailAddress import verifier
'''import re
import smtplib
import dns.resolver'''

# open book
book = xlrd.open_workbook("resultfortune100.xlsx")
sh = book.sheet_by_index(0)
wb = openpyxl.load_workbook(filename = 'resultfortune100.xlsx')
ws = wb.worksheets[0]

row = int(sh.cell_value(rowx=1, colx=13))
#print(row)

while (row < sh.nrows):
  # who to send a message to
  firstName = str(sh.cell_value(rowx=row, colx=2))
  lastName = str(sh.cell_value(rowx=row, colx=3))
  at_index = (str(sh.cell_value(rowx=row, colx=4)).find(" at "))
  company = str(sh.cell_value(rowx=row, colx=4))[at_index+4:]

  if (at_index == -1):
    row += 1
    continue

  combination0 = firstName[0]+lastName+"@"+company+".com"
  combination1 = firstName+"."+lastName+"@"+company+".com"
  combination2 = firstName+lastName+"@"+company+".com"
  combination3 = lastName+"@"+company+".com"
  combination4 = firstName+"_"+lastName+"@"+company+".com"
  combination5 = firstName[0]+"_"+lastName+"@"+company+".com"
  combination6 = firstName+lastName[0]+"@"+company+".com"
  combination7 = firstName+"@"+company+".com"
  combination8 = lastName+firstName[0]+"@"+company+".com"
  combination9 = lastName+"."+firstName+"@"+company+".com"
  combination10 = lastName[0]+firstName+"@"+company+".com"
  combination11 = lastName+firstName+"@"+company+".com"

  # try various emails
  email_combinations = {"0":combination1, "1":combination2, "2": combination2, "3": combination3, "4": combination4, "5": combination5, "6": combination6, "7": combination7, "8": combination8, "9": combination9, "10": combination10,"11": combination11}
  for i in range(12):
    receiver_email = (email_combinations[str(i)]).lower()
    #print (receiver_email)

    send_mail (receiver_email, firstName, lastName)
    
  # updating excel sheet
  print("updating")
  #ws.cell(row=row+1, column=4).value = receiver_email
  ws.cell(row=row+1, column=10).value = "Emailed"
  ws.cell(row=row+1, column=11).value = datetime.now()
  wb.save("resultfortune100.xlsx")

  book = xlrd.open_workbook("resultfortune100.xlsx")
  sh = book.sheet_by_index(0)
  wb = openpyxl.load_workbook(filename = 'resultfortune100.xlsx')
  ws = wb.worksheets[0]
  row += 1

ws.cell(row=2, column=12).value = row
wb.save("resultfortune100.xlsx")


print ("Done current emailing cycle")'''