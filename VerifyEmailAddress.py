# THIS WAS MADE BY Mikael Davranche. THANK YOU FOR MAKING THIS FREE

import re
import smtplib
import dns.resolver

def verifier (email):

	# Address used for SMTP MAIL FROM command  
	fromAddress = 'hrithikshah00@gmail.com'

	print(email)

	# Simple Regex for syntax checking
	regex = '^[_a-z0-9-]+(\.[_a-z0-9-]+)*@[a-z0-9-]+(\.[a-z0-9-]+)*(\.[a-z]{2,})$'

	# Email address to verify
	addressToVerify = email

	# Syntax check
	match = re.match(regex, addressToVerify)
	if match == None:
		print('Bad Syntax')
		return ""
		#raise ValueError('Bad Syntax')
	else:
		# Get domain for DNS lookup
		splitAddress = addressToVerify.split('@')
		domain = str(splitAddress[1])
		#print('Domain:', domain)

		# MX record lookup
		records = dns.resolver.query(domain, 'MX')
		mxRecord = records[0].exchange
		mxRecord = str(mxRecord)


		# SMTP lib setup (use debug level for full output)
		server = smtplib.SMTP()
		server.set_debuglevel(0)

		# SMTP Conversation
		server.connect(mxRecord)
		server.helo(server.local_hostname) ### server.local_hostname(Get local server hostname)
		server.mail(fromAddress)
		code, message = server.rcpt(str(addressToVerify))

		#print(code)
		#print(message)

		# Assume SMTP response 250 is success
		if code == 250:
			print('Success: '+email)
			return email
		else:
			print('Bad')