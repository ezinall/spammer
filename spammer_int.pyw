#!/usr/bin/python3
# -*- coding: utf-8 -*-

import tkinter as tk
from tkinter import filedialog
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import csv
import configparser
import string
import re
import smtplib
from time import sleep


class Application(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master.title('JMail alpha v0.0.1')
        self.master.geometry('1024x768')

        self.config_ini = None
        self.msg_subj = None
        self.msg = MIMEMultipart()
        self.msg_text = None
        self.msg_subj_entry = None

        self.server_conn_list = []
        self.to_adr_list = []
        self.black_list = []

        self.conf_identify()
        self.main_ui()

    def conf_identify(self):
        try:
            open('config.ini').read()
        except FileNotFoundError:
            with open('config.ini', 'w') as config_smtp:
                config_smtp.write(
                    '\n[smtp1]\n'
                    'SERVER_NAME = \n'
                    'SERVER_PORT =\n'
                    'USER_NAME =\n'
                    'USER_PASSWORD =\n'
                    'SSL = 0\n'
                    'FROM_ADR =\n')
        finally:
            self.config_ini = configparser.ConfigParser()
            self.config_ini.read('config.ini')

    def main_ui(self):
        menu_bar = tk.Menu(self)
        self.master.config(menu=menu_bar)

        file_menu = tk.Menu(menu_bar, tearoff=0)
        file_menu.add_command(label="Exit", command=self.on_exit)
        menu_bar.add_cascade(label="File", menu=file_menu)

        settings_menu = tk.Menu(menu_bar, tearoff=0)
        settings_menu.add_command(label='Servers', command=self.servers_set_ui)
        menu_bar.add_cascade(label='Settings', menu=settings_menu)

        to_label = tk.Label(self.master, text='To: ')
        to_label.grid(row=0, column=0, sticky='w')
        to_button = tk.Button(self.master, text="To", command=self.to_adr_list_open, width=10)
        to_button.grid(row=0, column=1, sticky='w')

        from_label = tk.Label(self.master, text='From: ')
        from_label.grid(row=1, column=0, sticky='w')
        from_str = '; '.join(self.config_ini[label]['from_adr'] for label in self.config_ini.sections() if 'smtp' in label)
        from_label = tk.Label(self.master, text=from_str)
        from_label.grid(row=1, column=1, sticky='w')

        subject_label = tk.Label(self.master, text='Subject: ')
        subject_label.grid(row=2, column=0, sticky='w')
        subject = tk.StringVar()
        subject.set(self.msg_subj if self.msg_subj else 'Subject')
        self.msg_subj_entry = tk.Entry(self.master, text=subject, textvariable=subject, width=75)
        self.msg_subj_entry.grid(row=2, column=1, sticky='w')
        msg_subj_button = tk.Button(self.master, text="Save", command=self.msg_subj_open, width=10)
        msg_subj_button.grid(row=2, column=2, sticky='w')

        mag_label = tk.Label(self.master, text='Mail: ')
        mag_label.grid(row=3, column=0, sticky='w')
        msg_button = tk.Button(self.master, text="Mail", command=self.msg_open, width=10)
        msg_button.grid(row=3, column=1, sticky='w')

        start_button = tk.Button(self.master, text="Connect", command=self.con_servers, width=10)
        start_button.grid(row=4, column=0, sticky='w')
        start_button = tk.Button(self.master, text="Start", command=self.star_sending, width=10)
        start_button.grid(row=4, column=1, sticky='e')
        start_button = tk.Button(self.master, text="Stop", command='', width=10)
        start_button.grid(row=4, column=2, sticky='w')

    def msg_subj_open(self):
        self.msg_subj = self.msg_subj_entry.get()

    def to_adr_list_open(self):
        dlg = filedialog.Open(self, filetypes=[('CSV files', '*.csv'), ('All files', '*')])
        to_file = dlg.show()
        if to_file:
            with open('*.csv', 'r') as f:
                self.to_adr_list = [r[0] for r in csv.reader(f) if r]

    def msg_open(self):
        dlg = filedialog.Open(self, filetypes=[('HTML files', '*.html'), ('All files', '*')])
        msg_file = dlg.show()
        if msg_file:
            with open(msg_file, "r") as f:
                self.msg_text = f.read()

    def on_exit(self):
        self.quit()

    def servers_set_ui(self):
        servers_set_win = tk.Toplevel(self.master)
        servers_set_win.geometry('640x480')
        servers_set_win.title('Servers settings')
        server_label_list = [label for label in self.config_ini.sections() if 'smtp' in label.lower()]
        for count, server_label in enumerate(server_label_list):
            tk.Label(servers_set_win, text=str(count)+'.').grid(row=count, column=0, sticky='w')
            tk.Label(servers_set_win, text=self.config_ini[server_label]['server_name']).grid(row=count, column=1, sticky='w')
            tk.Button(servers_set_win, text='Edit', command=self.server_set_ui, width=10).grid(row=count, column=2, sticky='w')
            tk.Button(servers_set_win, text='Delete', command=self.server_delete, width=10).grid(row=count, column=3, sticky='w')

    def servers_set_save(self):
        pass

    def servers_set_cancel(self):
        pass

    def server_set_ui(self):
        server_set_win = tk.Toplevel(self.master)

    def server_set_save(self):
        pass

    def server_set_cancel(self):
        pass

    def server_delete(self):
        pass

    def con_servers(self):
        self.server_conn_list = []
        server_label_list = [label for label in self.config_ini.sections() if 'smtp' in label.lower()]
        for server_label in server_label_list:
            connect = self.connect_mail_server(server_label)
            if connect:
                self.server_conn_list.append((connect, server_label))

    def connect_mail_server(self, server_label):
        server_name = self.config_ini[server_label]['server_name']
        server_port = self.config_ini[server_label]['server_port']
        user_name = self.config_ini[server_label]['user_name']
        user_password = self.config_ini[server_label]['user_password']
        print('Connecting to mail server %s ...' % server_name, end=' ')
        try:
            if int(self.config_ini[server_label]['ssl']):
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

    def star_sending(self):
        for counter, to_adr in enumerate(self.to_adr_list):
            to_adr_ascii = ''.join(i for i in to_adr if i in string.ascii_letters + string.digits + '@._-')
            if not re.search(r'(\w+@[a-zA-Z_]+?\.[a-zA-Z]{2,6})', to_adr_ascii):
                print(counter, "Message DON'T send to %s. Incorrect format!" % to_adr)
                continue
            elif to_adr_ascii in self.black_list:
                print(counter, "Message DON'T send. %s in black list!" % to_adr)
                continue
            for server, name in self.server_conn_list:
                server.helo()
                self.get_message(self.config_ini[name]['from_adr'], to_adr_ascii)
                try:
                    server.sendmail(self.config_ini[name]['from_adr'], to_adr_ascii, self.msg.as_string())
                except smtplib.SMTPServerDisconnected as exc:
                    print('Server %s disconnected' % name)
                    print('%s: %s' % (exc.__class__.__name__, exc))
                    self.server_conn_list.remove((server, name))
                    continue
                except Exception as exc:
                    print(counter, "Message DON'T send to", to_adr)
                    print('%s: %s' % (exc.__class__.__name__, exc))
                    break
                else:
                    self.server_conn_list.remove((server, name))
                    self.server_conn_list.append((server, name))
                    print(counter, 'Message send to', to_adr_ascii)
                    sleep(13 / len(self.server_conn_list) if len(self.server_conn_list) else 13)
                    break
            else:
                print('No servers to send')
                break
        else:
            print('Sending completed')
        for server, name in self.server_conn_list:
            print('Disconnecting from mail server %s' % name)
            server.quit()

    def get_message(self, from_, to):
        self.msg['From'] = from_
        self.msg['To'] = to
        self.msg['Subject'] = self.msg_subj
        self.msg.attach(MIMEText(self.msg_text, 'html', 'utf-8'))


root = tk.Tk()
app = Application(master=root)
app.mainloop()
