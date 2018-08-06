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
        # self.master.geometry('1024x768')

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
            pass
        else:
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

        frame_mail = tk.Frame(self.master)
        frame_mail.grid(sticky='w')

        to_label = tk.Label(frame_mail, text='To: ')
        to_label.grid(row=0, column=0, sticky='w')
        to_button = tk.Button(frame_mail, text="To", command=self.to_adr_list_open, width=10)
        to_button.grid(row=0, column=1, sticky='w')

        from_label = tk.Label(frame_mail, text='From: ')
        from_label.grid(row=1, column=0, sticky='w')
        if self.config_ini:
            from_str = '; '.join(self.config_ini[label]['from_adr'] for label in self.config_ini.sections() if 'smtp' in label)
        else:
            from_str = ''
        from_label = tk.Label(frame_mail, text=from_str)
        from_label.grid(row=1, column=1, sticky='w')

        subject_label = tk.Label(frame_mail, text='Subject: ')
        subject_label.grid(row=2, column=0, sticky='w')
        subject = tk.StringVar()
        subject.set(self.msg_subj if self.msg_subj else 'Subject')
        self.msg_subj_entry = tk.Entry(frame_mail, text=subject, textvariable=subject, width=75)
        self.msg_subj_entry.grid(row=2, column=1, sticky='w')
        msg_subj_button = tk.Button(frame_mail, text="Save", command=self.msg_subj_open, width=10)
        msg_subj_button.grid(row=2, column=2, sticky='w')

        msg_label = tk.Label(frame_mail, text='Mail: ')
        msg_label.grid(row=3, column=0, sticky='w')
        msg_button = tk.Button(frame_mail, text="Mail", command=self.msg_open, width=10)
        msg_button.grid(row=3, column=1, sticky='w')

        action_frame = tk.Frame(self.master)
        action_frame.grid(sticky='w')

        con_button = tk.Button(action_frame, text="Connect", command=self.connect_servers, width=10)
        con_button.grid(row=4, column=0, sticky='w')
        disconnect_button = tk.Button(action_frame, text="Disconnect", command=self.disconnect_servers, width=10)
        disconnect_button.grid(row=4, column=1, sticky='w')

        start_button = tk.Button(action_frame, text="Start", command=self.star_sending, width=10)
        start_button.grid(row=5, column=0, sticky='w')
        stop_button = tk.Button(action_frame, text="Stop", command='', width=10)
        stop_button.grid(row=5, column=1, sticky='w')

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
        servers_set_frame = tk.Toplevel(self.master)
        # servers_set_win.geometry('640x480')
        servers_set_frame.title('Servers settings')

        servers_frame = tk.Frame(servers_set_frame)
        servers_frame.grid()

        if self.config_ini:
            server_label_list = [label for label in self.config_ini.sections() if 'smtp' in label.lower()]
            for count, server_label in enumerate(server_label_list):
                tk.Label(servers_frame, text=str(count+1)+'.').grid(row=count, column=0, sticky='w')
                tk.Label(servers_frame, text=self.config_ini[server_label]['server_name']).grid(row=count, column=1, sticky='w')
                tk.Button(servers_frame, text='Edit', command=self.server_set_ui, width=10).grid(row=count, column=2, sticky='w')
                tk.Button(servers_frame, text='Delete', command=self.server_delete, width=10).grid(row=count, column=3, sticky='w')

        add_server_frame = tk.Frame(servers_set_frame)
        add_server_frame.grid(sticky='w')
        tk.Button(add_server_frame, text='Add', command=self.server_set_ui, width=10).grid(sticky='w')

        action_frame = tk.Frame(servers_set_frame)
        action_frame.grid(sticky='e')
        save_button = tk.Button(action_frame, text="Save", command='', width=10).grid(row=0, column=0)
        cancel_button = tk.Button(action_frame, text="Cancel", command='', width=10).grid(row=0, column=1)

    def servers_set_save(self):
        pass

    def servers_set_cancel(self):
        pass

    def server_set_ui(self):
        server_set_frame = tk.Toplevel(self.master)
        server_set_frame.title('Server configure')

        config_frame = tk.Frame(server_set_frame)
        config_frame.grid()
        tk.Label(config_frame, text='Server name: ').grid(row=0, column=0, sticky='w')
        server_name_entry = tk.Entry(config_frame).grid(row=0, column=1, sticky='w')
        tk.Label(config_frame, text='Server port: ').grid(row=1, column=0, sticky='w')
        server_port_entry = tk.Entry(config_frame).grid(row=1, column=1, sticky='w')
        tk.Label(config_frame, text='User name: ').grid(row=2, column=0, sticky='w')
        user_name_entry = tk.Entry(config_frame).grid(row=2, column=1, sticky='w')
        tk.Label(config_frame, text='User password: ').grid(row=3, column=0, sticky='w')
        user_password_entry = tk.Entry(config_frame).grid(row=3, column=1, sticky='w')
        tk.Label(config_frame, text='SSL: ').grid(row=4, column=0, sticky='w')
        ssl_check = tk.Checkbutton(config_frame).grid(row=4, column=1, sticky='w')
        tk.Label(config_frame, text='From address: ').grid(row=5, column=0, sticky='w')
        from_adr_entry = tk.Entry(config_frame).grid(row=5, column=1, sticky='w')

        action_frame = tk.Frame(server_set_frame)
        action_frame.grid(sticky='e')
        ok_button = tk.Button(action_frame, text="Ok", command='', width=10).grid(row=0, column=0)
        cancel_button = tk.Button(action_frame, text="Cancel", command='', width=10).grid(row=0, column=1)

    def server_set_save(self):
        pass

    def server_set_cancel(self):
        pass

    def server_delete(self):
        pass

    def connect_servers(self):
        self.server_conn_list = []
        if self.config_ini:
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

    def disconnect_servers(self):
        for server, name in self.server_conn_list:
            print('Disconnecting from mail server %s' % name)
            server.quit()

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

    def get_message(self, from_, to):
        self.msg['From'] = from_
        self.msg['To'] = to
        self.msg['Subject'] = self.msg_subj
        self.msg.attach(MIMEText(self.msg_text, 'html', 'utf-8'))


root = tk.Tk()
app = Application(master=root)
app.mainloop()
