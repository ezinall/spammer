import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from time import sleep
from random import randrange
import requests
import socks
import socket
from stem import Signal
from stem.control import Controller


def connectTor():
   print('Connect to Tor...')
   socks.setdefaultproxy(socks.PROXY_TYPE_SOCKS5, "127.0.0.1", 9150, True)
   socket._real_socket = socket.socket
   socket.socket = socks.socksocket
   socks.wrapmodule(smtplib)
   print(' new IP', requests.get("http://icanhazip.com/").text)

    
def newIdentify():
   socks.setdefaultproxy()
   with Controller.from_port(port=9151) as controller:
       controller.authenticate()
       controller.signal(Signal.NEWNYM)
   connectTor()

   
def connectMail():
   global SERVER
   SERVER = smtplib.SMTP_SSL(SERVER_NAME, SERVER_PORT)
   # SERVER = smtplib.SMTP(SERVER_NAME, SERVER_PORT)  # if you use TLS
   SERVER.set_debuglevel(0)
   # SERVER.starttls()  # if you use TLS
   SERVER.login(USER_NAME, USER_PASSWORD)


SERVER_NAME = 'srv_name'
SERVER_PORT = 465
USER_NAME = 'login'
USER_PASSWORD = 'pass'
FROMADDR = 'ur_email'

TOADDRLIST = open('toaddr.txt').read().splitlines()

MSG_SUBJ = 'Hello, World'
MSG_TEXT = open('text.html').read()
MSG = MIMEMultipart()
MSG['Subject'] = MSG_SUBJ
MSG['From'] = FROMADDR
MSG.attach(MIMEText(MSG_TEXT, 'html', 'utf-8'))

connectTor()
connectMail()

COUNTER = 0
for TOADDR in TOADDRLIST:
    if COUNTER//10 == COUNTER/10 and COUNTER!=0:
        SERVER.quit()
        newIdentify()
        connectMail()
    MSG['To'] = TOADDR
    SERVER.sendmail(FROMADDR, TOADDR, MSG.as_string())
    print(COUNTER, 'Message send to', TOADDR)
    sleep(randrange(5, 10, 1))
    COUNTER += 1
print("Sending completed")

SERVER.quit()
socks.setdefaultproxy()

'''
print(' new IP', requests.get("http://httpbin.org/ip").text)
 from email.mime.base import MIMEBase
import email.encoders
MSG_FILE = MIMEBase('application', "octet-stream")
MSG_FILE.set_payload(open('success-baby.jpg', 'rb').read())
email.encoders.encode_base64(MSG_FILE)
MSG_FILE.add_header('Content-Disposition', 'attachment', filename="123")
 from email.mime.image import MIMEImage
MSG_FILE = open('success-baby.jpg', 'rb').read()
MSG_FILE = MIMEImage(MSG_FILE)
MSG_FILE.add_header('Content-ID', '<image1>', filename="123")
    #MSG.attach(MIMEImage(MSG_FILE))
    MSG.attach(MSG_FILE)
'''
