#!/usr/bin/python
# -*- coding: utf-8 -*-

import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from time import sleep
import csv
import configparser
import string
import re
# import requests
# import socks
# import socket
# from stem import Signal
# from stem.control import Controller


# def connectTor():
#     print('Connect to Tor...')
#     socks.setdefaultproxy(socks.PROXY_TYPE_SOCKS5, "127.0.0.1", 9150, True)
#     #socket._real_socket = socket.socket
#     #socket.socket = socks.socksocket
#     socks.wrapmodule(smtplib)
#     print(' new IP', requests.get("http://icanhazip.com/").text)


# def newIdentify():
#     socks.setdefaultproxy()
#     with Controller.from_port(port = 9151) as controller:
#         controller.authenticate()
#         controller.signal(Signal.NEWNYM)
#     connectTor()


def get_server_list():
    return [label for label in server_config_ini.sections() if 'smtp' in label.lower()]


def get_connected_servers(server_label_list):
    server_con_list = []
    for server_label in server_label_list:
        connect = connect_mail_server(server_label)
        if connect:
            server_con_list.append((connect, server_label))
    return server_con_list


def connect_mail_server(server_label):
    server_name = server_config_ini[server_label]['server_name']
    server_port = server_config_ini[server_label]['server_port']
    user_name = server_config_ini[server_label]['user_name']
    user_password = server_config_ini[server_label]['user_password']
    print('Connecting to mail server %s ...' % server_name, end=' ')
    try:
        if int(server_config_ini[server_label]['ssl']):
            mail_server = smtplib.SMTP_SSL(server_name, server_port)  # if you use SSL
        else:
            mail_server = smtplib.SMTP(server_name, server_port)  # if you use TLS
            mail_server.starttls()
        mail_server.set_debuglevel(0)
        mail_server.login(user_name, user_password)
    except Exception as e:
        print('Not connected!')
        print('%s: %s' % (e.__class__.__name__, e))
    else:
        print('Connected')
        return mail_server


def get_message(from_, to):
    global msg
    msg = MIMEMultipart()
    msg['From'] = from_
    msg['To'] = to
    msg['Subject'] = msg_subj
    msg.attach(MIMEText(msg_text, 'html', 'utf-8'))


try:
    open('config_smtp.ini').read()
except FileNotFoundError:
    print('File "config_smtp.ini" not found.\nFile create. Specify server settings.')
    with open('config_smtp.ini', 'w') as config_smtp:
        config_smtp.write('[smtp1]\n'
                          'SERVER_NAME = \n'
                          'SERVER_PORT =\n'
                          'USER_NAME =\n'
                          'USER_PASSWORD =\n'
                          'SSL = 0\n'
                          'FROM_ADR =')
finally:
    server_config_ini = configparser.ConfigParser()
    server_config_ini.read('config.ini')

# to_adr_list = open('toadr.txt').read().splitlines()
try:
    with open('*.csv', 'r') as f:
        to_adr_list = [r[0] for r in csv.reader(f) if r]
except FileNotFoundError:
    print('Not found address list!')
    to_adr_list = []

msg_subj = 'Subject'
msg_text = open('text.html').read()

try:
    black_list = open('blacklist.txt').read().splitlines()
except FileNotFoundError:
    black_list = []

if __name__ == '__main__':

    server_conn_list = get_connected_servers(get_server_list())

    print('Start sending')

    for counter, to_adr in enumerate(to_adr_list):
        to_adr_ascii = ''.join(i for i in to_adr if i in string.ascii_letters + string.digits + '@._-')
        if not re.search(r'(\w+@[a-zA-Z_]+?\.[a-zA-Z]{2,6})', to_adr_ascii):
            print(counter, "Message DON'T send to %s. Incorrect format!" % to_adr)
            continue
        elif to_adr_ascii in black_list:
            print(counter, "Message DON'T send. %s in black list!" % to_adr)
            continue
        for server, name in server_conn_list:
            server.helo()
            get_message(server_config_ini[name]['from_adr'], to_adr_ascii)
            try:
                server.sendmail(server_config_ini[name]['from_adr'], to_adr_ascii, msg.as_string())
            except smtplib.SMTPServerDisconnected as exc:
                print('Server %s disconnected' % name)
                print('%s: %s' % (exc.__class__.__name__, exc))
                server_conn_list.remove((server, name))
                continue
            except Exception as exc:
                print(counter, "Message DON'T send to", to_adr)
                print('%s: %s' % (exc.__class__.__name__, exc))
                break
            else:
                server_conn_list.remove((server, name))
                server_conn_list.append((server, name))
                print(counter, 'Message send to', to_adr_ascii)
                sleep(13 / len(server_conn_list) if len(server_conn_list) else 13)
                break
        else:
            print('No servers to send')
            break
    else:
        print('Sending completed')
    for server, name in server_conn_list:
        print('Disconnecting from mail server %s' % name)
        server.quit()
