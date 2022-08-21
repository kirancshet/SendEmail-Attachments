#-------------------------------------------------------------------------------
# Name:        send_email
# Purpose:
# Author:      Kiran Chandrashekhar
# Created:     09-08-2022
#-------------------------------------------------------------------------------

import ssl
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders


import conf

class SendEmail:
    def __init__(self):
        self.username   = conf.gmail_username
        self.password   = conf.gmail_app_password
        self.host       = conf.gmail_smtp_host
        self.port       = conf.port

    #---------------------------------------------------------------
    #   Send Email along with Attachments to the user via Gmail
    #   Recipient Email address is a Single Email address
    #   If you want to send to multiple email address, please use "," 
    #   seperated email address
    #---------------------------------------------------------------
    def send_email_with_attachment(self, recipient_email_address,  email_message, attachment_lst=[]):

        try:           
         
            msg = MIMEMultipart()
            msg['From']     = self.username
            msg['To']       = recipient_email_address
            msg['Subject']  = "Test Email with Attachment"

            msg.attach(MIMEText(email_message, 'html'))
 
            server = smtplib.SMTP(self.host, self.port)
            
            for ind_file in attachment_lst:
                part = MIMEBase('application', "octet-stream")
                part.set_payload(open(ind_file, "rb").read())
                encoders.encode_base64(part)
                part.add_header('Content-Disposition', 'attachment; filename="%s"' % os.path.basename(ind_file))                        
                msg.attach(part)

            context = ssl.create_default_context()
            server.starttls(context=context) # Secure the connection
            server.login(self.username, self.password)   # Use your Gmail username and App password here              
            server.sendmail(self.username, recipient_email_address, msg.as_string())

        except Exception as e:
            print(str(e))

        finally:
            server.quit()


def main():
    email_message = '''
    <p>Hi Test user, </p>
    <p>Thi is a Test email sent via an automated process in Gmail along with 2 attachments</p>
    '''

    cur_dir = os.getcwd()
    attachment_lst = []

    attachment_lst.append(f"{cur_dir}/files/sample_excel_file.xlsx")
    attachment_lst.append(f"{cur_dir}/files/python_tutorial.pdf")

    
    obj = SendEmail()
    obj.send_email_with_attachment('kiran.cshet@gmail.com', email_message, attachment_lst)
    
if __name__ == '__main__':
    main()
    print("Done")
