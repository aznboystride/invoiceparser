#!/usr/bin/env python3
import smtplib
import imaplib
import email
from sys import exit
from os.path import basename
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import COMMASPACE, formatdate


class IMAPEmailer:
    
    def __init__(self, email, password, server):
        self.connection = imaplib.IMAP4_SSL(server)
        
        self.connection.login(email, password)

    def retrieveMostRecentFileWithExt(self, ext, path, person):
        self.connection.select(readonly=False)
        result, data = self.connection.search(None, '(FROM "{}")'.format(person))
        inbox_item_list = data[0].split()
            
        for item in reversed(inbox_item_list):
            result, item_bytes = self.connection.fetch(item, '(RFC822)')
            decoded_item = item_bytes[0][1].decode("utf-8")
            message = email.message_from_string(decoded_item)
            for part in message.walk():
                filename = part.get_filename()
                if filename is not None and ext in filename:
                    with open(path, 'wb') as fp:
                        fp.write(part.get_payload(decode=True))
                        return path
    def close(self):
        self.connection.close()

        return None

class SMTPEmailer:

    def __init__(self, email, password, server):
        self.server = server
        self.login(email, password)

    def login(self, email, password):
        self.email = email
        self.password = password
        self.smtp = smtplib.SMTP(self.server, 587)
        self.smtp.ehlo()
        self.smtp.starttls()
        self.smtp.login(email, password)
    
    def sendattachment(self, sub, to, filename):
        self.sendmail(sub, to, filename) 

    def sendmail(self, sub, to, filename):
        
        msg = MIMEMultipart()
        msg["From"] = self.email
        msg["To"] = to
        msg["Date"] = formatdate(localtime=True)
        msg["Subject"] = sub

        with open(filename, "rb") as fp:
            part = MIMEApplication(fp.read(), Name=basename(filename))

        part["Content-Disposition"] = 'attachment; filename="{}"'.format(basename(filename))
        msg.attach(part)

        self.smtp.sendmail(self.email, to, msg.as_string())

    def close(self):
        self.smtp.close()
def main():
    pass

if __name__ == '__main__':
    main()

