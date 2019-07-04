"""
Author: RPR

Script sending email to recipients.

Arguments to script:
- Build-tag (string): Build-tag to be used in Subject of email.
- Recipients (string): Representation of list with recipients.


info on mail account:

Account = QT
Email = QATeam@televic.com

Last Name = Team
First Name = QA

Type user = Miscellaneous
BU = Televic Rail
OU = OU=Miscellaneous,OU=Users,OU=Televic Rail,DC=televic,DC=com

Password = mfmfTdt8!
Mailbox created = yes
» Email sending
To send emails using Office365 server enter these details:
SMTP Host: smtp.office365.com
SMTP Port: 587
TLS Protocol: ON
SMTP Username: (your Office365 username)
SMTP Password: (your Office365 password)

IMAP Host: outlook.office365.com
IMAP Port: 993
Encryption: SSL
IMAP Username: (your Office365 username)
IMAP Password: (your Office365 password)

"""

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import optparse
import ast

usage = ("usage: %prog <arg1> <arg2>\n\n"
         "Arguments:\n"
         "\t<arg1> : build-tag\n"
         "\t<arg2> : recipients")

parser = optparse.OptionParser(usage=usage)
(options, args) = parser.parse_args()

if len(args) != 2:
    parser.error("incorrect number of arguments")
    exit(1)

if None in args:
    parser.error("in arguments")
    exit(1)

build_tag = args[0]
recipients = args[1]
recipients = ast.literal_eval(recipients)

# Email Sender address and password
#sender_address = 'jenkins.QATestFramework@gmail.com'
#sender_pass = 'civelet1970'
sender_address = 'QATeam@televic.com'
sender_pass = 'mfmfTdt8!'

try:
    #smtp_server = smtplib.SMTP('smtp.gmail.com', 587)
    smtp_server = smtplib.SMTP('smtp.office365.com', 587)
    smtp_server.starttls()
    smtp_server.login(sender_address, sender_pass)  # login with mail_id and password

    for receiver_address in recipients:
        # Setup the MIME
        message = MIMEMultipart()
        message['From'] = sender_address
        message['To'] = receiver_address
        message['Subject'] = 'New TestFramework %s available !' % build_tag  # The subject line

        mail_content = '''Hello,

        Update the TestFramework to %s.

        How to:
        
        - If not yet present, create folder 'C:\\TestFramework' on the host.
        - Unzip 'S:\\R&D\Ontwikkelingen\\33.96.7535\\F.studie\\4.Software\\%s.zip' to folder 'C:\\TestFramework' on the host.
        - Create shortcut on desktop of the host from the file 'C:\\TestFramework\\%s\\33.96.7535\\TestFramework.bat'.
   
        Regards,
        Test Team

        PS: This is an automated message - please do not reply directly to this email.
        ''' % (build_tag, build_tag, build_tag)

        # The body and the attachments for the mail
        message.attach(MIMEText(mail_content, 'plain'))

        text = message.as_string()
        smtp_server.sendmail(sender_address, receiver_address, text)

    smtp_server.close()

    print 'Notification Email sent!'
except Exception as err:
    print 'Something went wrong: ', err
