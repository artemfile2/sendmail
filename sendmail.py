import os
import re
import sys
import smtplib
import fnmatch
import codecs
import mimetypes
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from email.header import Header
from email.utils import formataddr
from email.message import EmailMessage
from urllib.parse import quote
from time import sleep

def send_mail(smtp_addr,
              fromaddr,
              password,
              toaddr,
              sender=None,
              subject=None,
              text=None,
              data=None,
              ):
    from_addr = formataddr((sender, fromaddr))
    msg = MIMEMultipart()
    msg['From'] = from_addr
    msg['To'] = toaddr
    msg['Subject'] = subject or ''

    if text:
        msg.attach(MIMEText(*text))

    if data:
        for filename, body in data.items():
            attachment = MIMEBase(mimetypes.guess_type(filename)[0].split('/')[0],
                                  mimetypes.guess_type(filename)[0].split('/')[1])
            # attachment.set_payload(body)
            attachment.set_payload(open(filename, "rb").read())
            encoders.encode_base64(attachment)

            attachment.add_header('Content-Disposition',
                                  'attachment',
                                  # filename=quote(filename))
                                  filename=('utf_8', '', filename))
            msg.attach(attachment)

    server = smtplib.SMTP_SSL('smtp.mail.ru', 465)
    server.login(fromaddr, password)
    text = msg.as_string()

    try:
        server.send_message(msg)
    except (smtplib.SMTPRecipientsRefused, smtplib.SMTPSenderRefused) as err:
        sys.stderr.write("Проблема с отправкой письма. Причина: %s" % err)
    finally:
        server.quit()


def prepare():

    RootFoldersFile = open('from.lst')
    mail_from = RootFoldersFile.readline()
    mail_from.join(mail_from.replace('\n', ''))
    RootFoldersFile.close()

    mail_adr = mail_from.split('\\')[0]
    mail_pass = mail_from.split('\\')[1]

    mails_to = []
    mailslist = codecs.open('mails.lst', "r", "utf-8")
    for line in mailslist:
        mails_to.append(line.replace('\n', ''))
    mailslist.close()

    i = 0
    for mail in mails_to:
        i += 1
        f_to = mail.split('\\')[0]
        m_to = mail.split('\\')[1]

        m_to = re.sub("^\s+|\n|\r|\s+$", '', m_to)

        smtp_addr = 'smtp.mail.ru', 465
        fromaddr = mail_adr
        password = mail_pass
        toaddr = m_to
        subject = "146 форма отчетности"
        sender = "Макс-М экономический отдел"
        text = 'Здравствуйте, высылаем Вам 146 форму'

        data = {}
        path = r"C:\sendmail"
        fil = '146_' + f_to + '*.xls'

        matches = []
        for root, dirnames, filenames in os.walk(path):
            for filename in fnmatch.filter(filenames, fil):
                matches.append(os.path.join(root, filename))
        with open(matches[0], "rb") as f:
            data[filename] = f.read()

        send_mail(smtp_addr,
                  fromaddr,
                  password,
                  toaddr,
                  sender=sender,
                  subject=subject,
                  text=(text, 'plain', 'utf-8'),
                  data=data,
                  )

        print(f"{i}| файл {f_to} на почтовый ящик {m_to} отправленн успешно...")
        sleep(2)

    input("Для выхода из программы нажмите Enter")


if __name__ == "__main__":
    prepare()