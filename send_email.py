import os
import smtplib
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email import encoders
def send_selenium_report():
  msg = MIMEMultipart()
  #Config Reading Part
  fromaddr = 'sender_email_here'
  toaddr = 'receiver_email_here'
  password = 'sender_email_login_password_here'
  #Email Sending Part
  msg['From'] = fromaddr
  msg['To'] = toaddr
  msg['Subject'] = "Test Result Automation"
  body = MIMEText("Hi All, <br> <br> Test for this count<br> <br> Thank You", 'html', 'utf-8')
  msg.attach(body)  # add message body (text or html)
  yourpath = 'test_result_folder_path_here'
  for subdir, dirs, files in os.walk(yourpath):
      for filename in files:
          attachment = open(yourpath+filename, "rb")
          part = MIMEBase('application', 'octet-stream')
          part.set_payload((attachment).read())
          encoders.encode_base64(part)
          part.add_header('Content-Disposition', "attachment; filename= %s" % filename)
          msg.attach(part)
  server = smtplib.SMTP('smtp.gmail.com', 587)
  server.starttls()
  server.login(fromaddr, password)
  text = msg.as_string()
  server.sendmail(fromaddr, toaddr, text)
  print("Email Sent Successfully")
  server.quit()
