import email, smtplib, ssl
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import getpass #secure password input

def send_mail (receiver_email, f, l):
  port = 465  # For SSL

  sender_email = "hrithikshah00@gmail.com"
  password = getpass.getpass("Type your password and press enter: ")
  print ("Starting emailing....")

  # message stuff
  message = MIMEMultipart()
  message["Subject"] = "Application for Co-op Position"
  message["From"] = "Hrithik Shah"
  message["To"] = receiver_email

  html = """\
      <html>
        <body>
          <div class="">
         <div class="aHl"></div>
         <div id=":16i" tabindex="-1"></div>
         <div id=":16v" class="ii gt">
            <div id=":16w" class="a3s aXjCH ">
               <div dir="ltr">
                  <p class="MsoNormal" style="margin:0cm 0cm 0.0001pt;line-height:normal;font-size:11pt;font-family:Calibri,sans-serif"><span style="font-family:&quot;Calibri Light&quot;,sans-serif;font-size:11pt">Dear """+f+" "+l+""",</span><br></p>
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
      </div>
        </body>
      </html>
      """

  part1 = MIMEText(html, "html")

  message.attach(part1)

  filename = "resume.pdf"  
  filenametwo = "coverletter.pdf"

  # Open PDF file in binary mode
  with open(filename, "rb") as attachment:
      # Add file as application/octet-stream
      # Email client can usually download this automatically as attachment
      resume = MIMEBase("application", "octet-stream")
      resume.set_payload(attachment.read())

  with open(filenametwo, "rb") as attachment:
      # Add file as application/octet-stream
      # Email client can usually download this automatically as attachment
      coverletter = MIMEBase("application", "octet-stream")
      coverletter.set_payload(attachment.read())

  # Encode file in ASCII characters to send by email    
  encoders.encode_base64(resume)
  encoders.encode_base64(coverletter)

  # Add header as key/value pair to attachment part
  resume.add_header('Content-Disposition', 'attachment', filename=filename)

  coverletter.add_header('Content-Disposition', 'attachment', filename=filenametwo)

  # Add attachment to message and convert message to string
  message.attach(resume)
  message.attach(coverletter)
  text = message.as_string()

  # Create a secure SSL context
  context = ssl.create_default_context()
  with smtplib.SMTP_SSL("smtp.gmail.com", port, context=context) as server:
    server.login(sender_email, password)
    server.sendmail(sender_email, receiver_email, text)