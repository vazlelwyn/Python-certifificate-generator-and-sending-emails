# -*- coding: utf-8 -*-
"""
Created on Sun Dec  6 18:34:02 2020

@author: Hardwell
"""

import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import smtplib
from email.mime.multipart import MIMEMultipart 
from email.mime.text import MIMEText
from email.mime.base import MIMEBase 
from email import encoders  

data = pd.read_excel (r'C:\Users\Hardwell\Desktop\Certificates\20201205111252_2988.xls') 

fromaddr = "example@gmail.com"

# creates SMTP session 
s = smtplib.SMTP("smtp.gmail.com", 587)

#s.connect()
s.ehlo()
s.starttls() 

# Authentication 
s.login("example@gmail.com", "password")

for i in range(len(data['First Name'])):
    im = Image.open(r'C:\Users\Hardwell\Desktop\Certificates\IMG-20201205-WA0004.jpg')
    d = ImageDraw.Draw(im)
    location = (431, 517)
    text_color = (0, 0, 0)
    font = ImageFont.truetype("timesbd.ttf", 60)
    d.text(location, data['First Name'][i], fill = text_color, font = font)
    im.save(data['First Name'][i]+".pdf")
    
    toaddr = data['Email'][i]
    
    name = data['First Name'][i]

    # instance of MIMEMultipart 
    msg = MIMEMultipart() 

    # storing the senders email address 
    msg['From'] = fromaddr 

    # storing the receivers email address 
    msg['To'] = toaddr 

    # storing the subject 
    msg['Subject'] = "Certificate"

    # string to store the body of the mail 
    body = "Thank you for participating " +name+"."+"\n"+"Certificate of Participation is attached in this email."+"\n\n"+"Regards"
     # attach the body with the msg instance 
    msg.attach(MIMEText(body, 'plain')) 
    
    # open the file to be sent 
    filename = data['First Name'][i]+".pdf"
    attachment = open(data['First Name'][i]+".pdf", "rb") 

    # instance of MIMEBase and named as p 
    p = MIMEBase('application', 'octet-stream') 

    # To change the payload into encoded form 
    p.set_payload((attachment).read()) 

    # encode into base64 
    encoders.encode_base64(p) 

    p.add_header('Content-Disposition', "attachment; filename= %s" % filename) 

    # attach the instance 'p' to instance 'msg' 
    msg.attach(p) 





    # Converts the Multipart msg into a string 
    text = msg.as_string() 

    # sending the mail 
    s.sendmail(fromaddr, toaddr, text) 

# terminating the session 
s.quit() 


