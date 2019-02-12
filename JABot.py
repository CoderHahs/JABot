


'''
Setting up a secure, encrypted SMTP connection.

creates a secure connection with Gmail’s SMTP server, using the SMTP_SSL() 
of smtplib to initiate a TLS-encrypted connection. The default context of ssl validates the host name and its certificates and optimizes the security of the connection. 
'''

import email, smtplib, ssl
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import getpass #secure password input
import xlrd # excel manipulations
from datetime import datetime 
import openpyxl

port = 465  # For SSL
# open book
book = xlrd.open_workbook("recruiters.xlsx")
sh = book.sheet_by_index(0)

sender_email = "hrithikshah00@gmail.com"
#password = getpass.getpass("Type your password and press enter: ")
print ("Starting emailing....")
receiver_email = str(sh.cell_value(rowx=1, colx=4))

# message stuff
message = MIMEMultipart()
message["Subject"] = "Application for Co-op Position"
message["From"] = sender_email
message["To"] = receiver_email

# who to send a message to
firstName = str(sh.cell_value(rowx=1, colx=0))
lastName = str(sh.cell_value(rowx=1, colx=1))
prefix = str(sh.cell_value(rowx=1, colx=2))
company = str(sh.cell_value(rowx=1, colx=3))

'''# updating excel sheet
wb = openpyxl.load_workbook(filename = 'recruiters.xlsx')
ws = wb.worksheets[0]
ws.cell(row=2, column=6).value = "Emailed"
ws.cell(row=2, column=7).value = datetime.now()
wb.save("recruiters.xlsx")'''


text = """\

Dear """+ prefix + " " + firstName + " " + lastName+ """

I hope you have had a wonderful start to the new year!

 

My name is Hrithik Shah and I am an undergrad student pursuing Software Engineering (CO-OP) at the University of Ottawa. As part of my undergraduate program I am looking to complete an internship for Summer 2019 and am pleased to present to you my resume (attached with this email) as consideration for the following role(s):

Software Developer Intern
Data Analyst Intern
As someone who enjoys technology and working on software development projects, I believe I would learn a lot through an internship and would prove to be an immediate contributor to your firm.

 

I welcome an opportunity to further discuss my qualifications with you and learn more about the internship role(s) at your earliest convenience. Thank you very much for your time and consideration.

 

Best Regards,

 

Hrithik Shah
hrithikshah.com"""

html = """\
<html>
  <body>
    <div class="">
   <div class="aHl"></div>
   <div id=":16i" tabindex="-1"></div>
   <div id=":16v" class="ii gt">
      <div id=":16w" class="a3s aXjCH ">
         <div dir="ltr">
            <p class="MsoNormal" style="margin:0cm 0cm 0.0001pt;line-height:normal;font-size:11pt;font-family:Calibri,sans-serif"><span style="font-family:&quot;Calibri Light&quot;,sans-serif;font-size:11pt">Dear """+ prefix + " "+firstName + " " + lastName+""",</span><br></p>
            <p class="MsoNormal" style="margin:0cm 0cm 0.0001pt;line-height:normal;font-size:11pt;font-family:Calibri,sans-serif"><span style="font-family:&quot;Calibri Light&quot;,sans-serif">&nbsp;</span></p>
            <p class="MsoNormal" style="margin:0cm 0cm 0.0001pt;line-height:normal;font-size:11pt;font-family:Calibri,sans-serif"><span style="font-family:&quot;Calibri Light&quot;,sans-serif">I hope you have had a wonderful start to the new year!</span></p>
            <p class="MsoNormal" style="margin:0cm 0cm 0.0001pt;line-height:normal;font-size:11pt;font-family:Calibri,sans-serif"><span style="font-family:&quot;Calibri Light&quot;,sans-serif">&nbsp;</span></p>
            <p class="MsoNormal" style="margin:0cm 0cm 0.0001pt;line-height:normal;font-size:11pt;font-family:Calibri,sans-serif"><span style="font-family:&quot;Calibri Light&quot;,sans-serif">My name is Hrithik Shah and I am an undergrad student pursuing&nbsp;<strong>Software Engineering (CO-OP)</strong>&nbsp;at the University of Ottawa. As part of my undergraduate program I am looking to complete an internship for&nbsp;<strong>Summer 2019</strong>&nbsp;and am pleased&nbsp;</span><span style="font-family:&quot;Calibri Light&quot;,sans-serif;color:black">to present to you my resume (attached with this email) as consideration for the following role(s):</span></p>
            <ul>
               <li style="margin-left:15px"><span style="color:rgb(0,0,0);font-family:&quot;Calibri Light&quot;,sans-serif"><span style="font-size:14.6667px"><strong>Software Developer Intern</strong></span></span></li>
               <li style="margin-left:15px"><font color="#000000" face="Calibri Light, sans-serif"><span style="font-size:14.6667px"><b>Data Analyst Intern</b></span></font></li>
            </ul>
            <p class="MsoNormal" style="margin:0cm 0cm 0.0001pt;line-height:normal;font-size:11pt;font-family:Calibri,sans-serif"><span style="font-family:&quot;Calibri Light&quot;,sans-serif">As someone who enjoys technology and working on software development projects, I believe I would learn a lot through an internship and would prove to be an immediate contributor to your firm.</span></p>
            <p class="MsoNormal" style="margin:0cm 0cm 0.0001pt;line-height:normal;font-size:11pt;font-family:Calibri,sans-serif"><span style="font-family:&quot;Calibri Light&quot;,sans-serif">&nbsp;</span></p>
            <p class="MsoNormal" style="margin:0cm 0cm 0.0001pt;line-height:normal;font-size:11pt;font-family:Calibri,sans-serif"><span style="font-family:&quot;Calibri Light&quot;,sans-serif">I welcome an opportunity to further discuss my qualifications with you and learn more about the internship role(s) at your earliest convenience. Thank you very much for your time and consideration.</span></p>
            <p class="MsoNormal" style="margin:0cm 0cm 0.0001pt;line-height:normal;font-size:11pt;font-family:Calibri,sans-serif"><span style="font-family:&quot;Calibri Light&quot;,sans-serif">&nbsp;</span></p>
            <p class="MsoNormal" style="margin:0cm 0cm 0.0001pt;line-height:normal;font-size:11pt;font-family:Calibri,sans-serif"><span style="font-family:&quot;Calibri Light&quot;,sans-serif">Best Regards,</span></p>
            <p class="MsoNormal" style="margin:0cm 0cm 0.0001pt;line-height:normal;font-size:11pt;font-family:Calibri,sans-serif"><span style="font-family:&quot;Calibri Light&quot;,sans-serif">&nbsp;</span></p>
            <p class="MsoNormal" style="margin:0cm 0cm 0.0001pt;line-height:normal;font-size:11pt;font-family:Calibri,sans-serif"><strong><span style="font-family:&quot;Calibri Light&quot;,sans-serif">Hrithik Shah</span></strong></p>
            <span style="font-size:11pt;line-height:15.6933px;font-family:&quot;Calibri Light&quot;,sans-serif"><a href="http://hrithikshah.com/" rel="noopener" target="_blank" data-saferedirecturl="https://www.google.com/url?q=http://hrithikshah.com/&amp;source=gmail&amp;ust=1550005700791000&amp;usg=AFQjCNEeb9hMp3-juOdRunCjvEmufdarcA">hrithikshah.com</a></span>
            <div class="yj6qo"></div>
            <div class="adL">&nbsp;&nbsp;<br></div>
         </div>
         <div class="adL"></div>
      </div>
   </div>
   <div id=":15v" class="ii gt" style="display:none">
      <div id=":15u" class="a3s aXjCH undefined"></div>
   </div>
   <div class="hq gt a10" id=":162">
      <div class="hp"></div>
      <div class="a3I">Attachments area</div>
      <div id=":15z"></div>
      <div class="aQH" id=":160">
         <span class="aZo N5jrZb" download_url="application/pdf:resume.pdf:https://mail.google.com/mail/u/0/https://mail.google.com/mail/u/0?ui=2&amp;ik=1067aada98&amp;attid=0.1&amp;permmsgid=msg-a:r4263844881900792160&amp;th=1682a9bec71005c9&amp;view=att&amp;disp=safe&amp;realattid=f_jqmy58zp0" draggable="true">
            <a id=":16g" target="_blank" role="link" class="aQy aZr e" href="https://mail.google.com/mail/u/0?ui=2&amp;ik=1067aada98&amp;attid=0.1&amp;permmsgid=msg-a:r4263844881900792160&amp;th=1682a9bec71005c9&amp;view=att&amp;disp=inline&amp;realattid=f_jqmy58zp0" data-tooltip-align="t,c" data-tooltip-class="a1V" tabindex="0">
               <span class="a3I" id=":168">Preview attachment resume.pdf</span>
               <div aria-hidden="true">
                  <div class="aSG"></div>
                  <div class="aVY aZn">
                     <div class="aZm"></div>
                  </div>
                  <div class="aSH">
                     <img class="aQG aYB" id=":16f" src="https://mail.google.com/mail/u/0?ui=2&amp;ik=1067aada98&amp;attid=0.1&amp;permmsgid=msg-a%3Ar4263844881900792160&amp;th=1682a9bec71005c9&amp;view=snatt&amp;realattid=f_jqmy58zp0&amp;disp=thd&amp;safe=1&amp;sz=w360-h240-p-nu" style="position: absolute; top: 0px; width: 100%;">
                     <div id=":16d" class="aYv" style="display: none;"><img class="aYw aZB" src="//ssl.gstatic.com/ui/v1/icons/mail/images/cleardot.gif"></div>
                     <div id=":169" class="aYy">
                        <div class="aYA"><img id=":16b" class="aSM" src="//ssl.gstatic.com/docs/doclist/images/mediatype/icon_3_pdf_x32.png" title="PDF"></div>
                        <div class="aYz">
                           <div class="a12">
                              <div class="aQA"><span class="aV3 zzV0ie" id=":16c">resume.pdf</span></div>
                              <div class="aYp"><span id=":16a" class="SaH2Ve">95 KB</span></div>
                           </div>
                        </div>
                     </div>
                  </div>
                  <div class="aSI">
                     <div id=":16e" class="aSJ" style="border-color: #fb4c2f"></div>
                  </div>
               </div>
            </a>
            <div class="aQw">
               <div id=":165" class="T-I J-J5-Ji aQv T-I-ax7 L3" role="button" aria-label="Download attachment resume.pdf" data-tooltip-class="a1V" aria-disabled="false" style="user-select: none;" tabindex="0" data-tooltip="Download">
                  <div class="aSK J-J5-Ji aYr"></div>
               </div>
               <div id=":164" class="T-I J-J5-Ji aQv T-I-ax7 L3" role="button" aria-label="Save attachment to Drive resume.pdf" data-tooltip-class="a1V" aria-hidden="false" style="user-select: none;" data-tooltip="Save to Drive" tabindex="0">
                  <div class="wtScjd J-J5-Ji aYr aQu">
                     <div class="T-aT4" style="display: none;">
                        <div></div>
                        <div class="T-aT4-JX"></div>
                     </div>
                  </div>
               </div>
               <div id=":163" class="T-I J-J5-Ji aQv T-I-ax7 L3 T-I-JE aVZ" role="button" aria-label="Add to contacts" data-tooltip-class="a1V" aria-disabled="true" aria-hidden="true" style="user-select: none; display: none;" data-tooltip="Add to contacts">
                  <div class="aVX J-J5-Ji aYr"></div>
               </div>
            </div>
         </span>
         <div class="aZK"></div>
      </div>
   </div>
   <div class="hi"></div>
</div>
  </body>
</html>
"""

part1 = MIMEText(text, "plain")
part2 = MIMEText(html, "html")

message.attach(part1)
message.attach(part2)


filename = "resume.pdf"  # In same directory as script

# Open PDF file in binary mode
with open(filename, "rb") as attachment:
    # Add file as application/octet-stream
    # Email client can usually download this automatically as attachment
    resume = MIMEBase("application", "octet-stream")
    resume.set_payload(attachment.read())

# Encode file in ASCII characters to send by email    
encoders.encode_base64(resume)

# Add header as key/value pair to attachment part
resume.add_header(
    "Hrithik Shah resume",
    f"attachment; filename= {filename}",
)

# Add attachment to message and convert message to string
message.attach(resume)
text = message.as_string()

# Create a secure SSL context
#context = ssl.create_default_context()

with smtplib.SMTP_SSL("smtp.gmail.com", port, context=context) as server:
    server.login(sender_email, password)
    print ("Successfully logged in!")
    server.sendmail(sender_email, receiver_email, text)
    print ("Email sent to "+firstName+ ", " +lastName+ ", " +company+ " with email " +receiver_email+"!")