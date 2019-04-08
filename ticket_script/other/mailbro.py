import smtplib
from email.mime.multipart import MIMEMultipart 
from email.mime.text import MIMEText 
from email.mime.base import MIMEBase 
from email import encoders 

# textfile = 'amazing.txt'

# fp = open(textfile, 'rb')

# msg = MIMEText(fp.read())
# fp.close()

# me = 'ticket-dispatcher@virtusa.com'
# you = 'sratnappuli@virtusa.com'

# msg['Subject'] = 'The contents of %s' % textfile
# msg['From'] = me
# msg['To'] = you

# s = smtplib.SMTP('10.62.65.37')
# s.sendmail(me, [you], msg.as_string())
# s.quit()


## ----------------

fromaddr = 'fss@virtusa.com'
toaddr = 'sratnappuli@virtusa.com'
   
# instance of MIMEMultipart 
msg = MIMEMultipart() 
  
# storing the senders email address   
msg['From'] = fromaddr 
  
# storing the receivers email address  
msg['To'] = toaddr 
  
# storing the subject  
msg['Subject'] = "Subject of the Mail this Rock"

textfile = 'amazing.txt'

fp = open(textfile, 'rb')
mybody = fp.read()
# string to store the body of the mail 
body = mybody
  
# attach the body with the msg instance 
msg.attach(MIMEText(body, 'plain')) 
  
# open the file to be sent  
filename = "Delegated_Ticket_List_20190326_0324.xlsx"
attachment = open("C:\\Temp\\pkg\\reports\\Delegated_Ticket_List_20190326_0324.xlsx", "rb") 
  
# instance of MIMEBase and named as p 
p = MIMEBase('application', 'octet-stream') 
  
# To change the payload into encoded form 
p.set_payload((attachment).read()) 
  
# encode into base64 
encoders.encode_base64(p) 
   
p.add_header('Content-Disposition', "attachment; filename= %s" % filename) 
  
# attach the instance 'p' to instance 'msg' 
msg.attach(p) 
  
# creates SMTP session 
s = smtplib.SMTP('10.62.65.37') 
  
# start TLS for security 
######s.starttls() 
  
# Authentication 
#####s.login(fromaddr, "Password_of_the_sender") 
  
# Converts the Multipart msg into a string 
text = msg.as_string() 
  
# sending the mail 
s.sendmail(fromaddr, toaddr, text) 
  
# terminating the session 
s.quit()