import urllib2
from bs4 import BeautifulSoup


'''
Setting up a secure, encrypted SMTP connection.

creates a secure connection with Gmail’s SMTP server, using the SMTP_SSL() 
of smtplib to initiate a TLS-encrypted connection. The default context of ssl validates the host name and its certificates and optimizes the security of the connection. 
'''

import smtplib, ssl

port = 465  # For SSL
send_email = input("Type your email and press enter: ")
password = input("Type your password and press enter: ")

receiver_email = input("Type the email you want to send an email to: ")

# Create a secure SSL context
context = ssl.create_default_context()

with smtplib.SMTP_SSL("smtp.gmail.com", port, context=context) as server:
    server.login("hrithikchamp@gmail.com", password)
    server.sendmail(sender_email, receiver_email, message)