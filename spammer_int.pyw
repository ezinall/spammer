#!/usr/bin/python3
# -*- coding: utf-8 -*-

import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import csv
import configparser
import string
import re
import smtplib
from time import sleep
import uuid


class Application(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master.title('JMail alpha v0.0.1')
        # self.master.geometry('1024x768')

        self.config_ini = None
        self.from_str = tk.StringVar()
        self.msg_subj = tk.StringVar()
        self.msg = MIMEMultipart()
        self.msg_text = None

        self.server_name = tk.StringVar()
        self.server_port = tk.StringVar()
        self.user_name = tk.StringVar()
        self.user_password = tk.StringVar()
        self.ssl = tk.IntVar()
        self.from_adr = tk.StringVar()

        self.server_conn_list = []
        self.to_adr_list = []
        self.black_list = []

        self.servers_set_win = None
        self.server_add_win = None
        self.server_edit_win = None

        self.conf_identify()
        self.main_ui()

    def conf_identify(self):
        try:
            open('config.ini').read()
        except FileNotFoundError:
            open('config.ini', 'w')
        finally:
            self.config_ini = configparser.ConfigParser()
            self.config_ini.read('config.ini')

    def main_ui(self):
        menu_bar = tk.Menu(self)
        self.master.config(menu=menu_bar)

        file_menu = tk.Menu(menu_bar, tearoff=0)
        file_menu.add_command(label="Exit", command=lambda: self.quit())
        menu_bar.add_cascade(label="File", menu=file_menu)

        settings_menu = tk.Menu(menu_bar, tearoff=0)
        settings_menu.add_command(label='Servers', command=self.servers_set_ui)
        menu_bar.add_cascade(label='Settings', menu=settings_menu)

        frame_mail = tk.Frame(self.master)
        frame_mail.grid(sticky='w')

        tk.Label(frame_mail, text='To: ').grid(row=0, column=0, sticky='w')
        tk.Button(frame_mail, text="Open", command=self.to_adr_list_open, width=10).grid(row=0, column=1, sticky='w')

        tk.Label(frame_mail, text='From: ').grid(row=1, column=0, sticky='w')
        if self.config_ini:
            self.from_str.set('; '.join(self.config_ini[label]['from_adr'] for label in self.config_ini.sections() if 'smtp' in label))
        else:
            self.from_str.set('')
        tk.Label(frame_mail, textvariable=self.from_str).grid(row=1, column=1, sticky='w')

        tk.Label(frame_mail, text='Subject: ').grid(row=2, column=0, sticky='w')
        try:
            self.msg_subj.set(self.config_ini['config']['msg_subj'])
        except (KeyError, TypeError):
            self.msg_subj.set('Message subject')
        tk.Entry(frame_mail, textvariable=self.msg_subj, width=75).grid(row=2, column=1, sticky='w')
        tk.Button(frame_mail, text="Save", command=self.msg_subj_save, width=10).grid(row=2, column=2, sticky='w')

        tk.Label(frame_mail, text='Mail: ').grid(row=3, column=0, sticky='w')
        tk.Button(frame_mail, text="Open", command=self.msg_open, width=10).grid(row=3, column=1, sticky='w')

        ttk.Separator(frame_mail).grid(row=4, column=0, sticky="ew", columnspan=3)

        action_frame = tk.Frame(self.master)
        action_frame.grid(sticky='w')

        tk.Button(action_frame, text="Connect", command=self.connect_servers, width=10).grid(row=0, column=0, sticky='w')
        tk.Button(action_frame, text="Disconnect", command=self.disconnect_servers, width=10).grid(row=0, column=1, sticky='w')

        tk.Button(action_frame, text="Start", command=self.star_sending, width=10).grid(row=1, column=0, sticky='w')
        tk.Button(action_frame, text="Stop", command='', width=10).grid(row=1, column=1, sticky='w')

        # send_progress = ttk.Progressbar(action_frame, mode='determinate', value=10).grid(row=2, column=0)

    def msg_subj_save(self):
        try:
            self.config_ini.add_section('config')
        except configparser.DuplicateSectionError:
            pass
        finally:
            self.config_ini.set('config', 'msg_subj', self.msg_subj.get())
        with open('config.ini', 'w') as config_ini:
            self.config_ini.write(config_ini)

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

    def servers_set_ui(self):
        self.servers_set_win = tk.Toplevel(self.master)
        # servers_set_win.geometry('640x480')
        self.servers_set_win.title('Servers settings')
        self.servers_set_win.grab_set()

        servers_frame = tk.Frame(self.servers_set_win)
        servers_frame.grid()

        if self.config_ini:
            server_label_list = [label for label in self.config_ini.sections() if 'smtp' in label.lower()]
            for count, server_label in enumerate(server_label_list):
                tk.Label(servers_frame, text=str(count+1)+'.').grid(row=count, column=0, sticky='w')
                tk.Label(servers_frame, text=self.config_ini[server_label]['server_name']).grid(row=count, column=1, sticky='w')
                tk.Button(servers_frame, text='Edit', command=lambda label=server_label: self.server_edit_ui(label), width=10).grid(row=count, column=2, sticky='w')
                tk.Button(servers_frame, text='Delete', command=lambda label=server_label: self.server_delete(label), width=10).grid(row=count, column=3, sticky='w')

        add_server_frame = tk.Frame(self.servers_set_win)
        add_server_frame.grid(sticky='w')

        tk.Button(add_server_frame, text='Add', command=self.server_add_ui, width=10).grid(sticky='w')

        action_frame = tk.Frame(self.servers_set_win)
        action_frame.grid(sticky='e')

        tk.Button(action_frame, text="Save", command='', width=10).grid(row=0, column=0)
        tk.Button(action_frame, text="Cancel", command=lambda: self.servers_set_win.destroy(), width=10).grid(row=0, column=1)

    def servers_set_save(self):
        pass

    def server_add_ui(self):
        self.server_add_win = tk.Toplevel(self.master)
        self.server_add_win.title('Server add')
        self.server_add_win.grab_set()

        config_frame = tk.Frame(self.server_add_win)
        config_frame.grid()

        self.server_name.set('')
        self.server_port.set('')
        self.user_name.set('')
        self.user_password.set('')
        self.ssl.set(0)
        self.from_adr.set('')

        tk.Label(config_frame, text='Server name: ').grid(row=0, column=0, sticky='w')
        tk.Entry(config_frame, textvariable=self.server_name).grid(row=0, column=1, sticky='w')
        tk.Label(config_frame, text='Server port: ').grid(row=1, column=0, sticky='w')
        tk.Entry(config_frame, textvariable=self.server_port).grid(row=1, column=1, sticky='w')
        tk.Label(config_frame, text='User name: ').grid(row=2, column=0, sticky='w')
        tk.Entry(config_frame, textvariable=self.user_name).grid(row=2, column=1, sticky='w')
        tk.Label(config_frame, text='User password: ').grid(row=3, column=0, sticky='w')
        tk.Entry(config_frame, textvariable=self.user_password, show='*').grid(row=3, column=1, sticky='w')
        tk.Label(config_frame, text='SSL: ').grid(row=4, column=0, sticky='w')
        tk.Checkbutton(config_frame, variable=self.ssl).grid(row=4, column=1, sticky='w')
        tk.Label(config_frame, text='From address: ').grid(row=5, column=0, sticky='w')
        tk.Entry(config_frame, textvariable=self.from_adr).grid(row=5, column=1, sticky='w')

        action_frame = tk.Frame(self.server_add_win)
        action_frame.grid(sticky='e')
        tk.Button(action_frame, text="Save", command=lambda: self.server_add_save(), width=10).grid(row=0, column=0)
        tk.Button(action_frame, text="Cancel", command=lambda: self.server_add_win.destroy(), width=10).grid(row=0, column=1)

    def server_add_save(self):
        server_label = 'smtp-'+uuid.uuid4().hex[:5]
        self.config_ini.add_section(server_label)
        self.config_ini.set(server_label, 'server_name', self.server_name.get())
        self.config_ini.set(server_label, 'server_port', self.server_port.get())
        self.config_ini.set(server_label, 'user_name', self.user_name.get())
        self.config_ini.set(server_label, 'user_password', self.user_password.get())
        self.config_ini.set(server_label, 'ssl', str(self.ssl.get()))
        self.config_ini.set(server_label, 'from_adr', self.from_adr.get())
        with open('config.ini', 'w') as config_ini:
            self.config_ini.write(config_ini)
        self.server_add_win.destroy()
        self.servers_set_win.destroy()
        self.servers_set_ui()
        self.from_str.set('; '.join(self.config_ini[label]['from_adr'] for label in self.config_ini.sections() if 'smtp' in label))

    def server_edit_ui(self, server_label):
        self.server_edit_win = tk.Toplevel(self.master)
        self.server_edit_win.title('Server configure')
        self.server_edit_win.grab_set()

        config_frame = tk.Frame(self.server_edit_win)
        config_frame.grid()

        self.server_name.set(self.config_ini[server_label]['server_name'])
        self.server_port.set(self.config_ini[server_label]['server_port'])
        self.user_name.set(self.config_ini[server_label]['user_name'])
        self.user_password.set(self.config_ini[server_label]['user_password'])
        self.ssl.set(int(self.config_ini[server_label]['ssl']))
        self.from_adr.set(self.config_ini[server_label]['from_adr'])

        tk.Label(config_frame, text='Server name: ').grid(row=0, column=0, sticky='w')
        tk.Entry(config_frame, textvariable=self.server_name).grid(row=0, column=1, sticky='w')
        tk.Label(config_frame, text='Server port: ').grid(row=1, column=0, sticky='w')
        tk.Entry(config_frame, textvariable=self.server_port).grid(row=1, column=1, sticky='w')
        tk.Label(config_frame, text='User name: ').grid(row=2, column=0, sticky='w')
        tk.Entry(config_frame, textvariable=self.user_name).grid(row=2, column=1, sticky='w')
        tk.Label(config_frame, text='User password: ').grid(row=3, column=0, sticky='w')
        tk.Entry(config_frame, textvariable=self.user_password, show='*').grid(row=3, column=1, sticky='w')
        tk.Label(config_frame, text='SSL: ').grid(row=4, column=0, sticky='w')
        tk.Checkbutton(config_frame, variable=self.ssl).grid(row=4, column=1, sticky='w')
        tk.Label(config_frame, text='From address: ').grid(row=5, column=0, sticky='w')
        tk.Entry(config_frame, textvariable=self.from_adr).grid(row=5, column=1, sticky='w')

        action_frame = tk.Frame(self.server_edit_win)
        action_frame.grid(sticky='e')
        tk.Button(action_frame, text="Save", command=lambda label=server_label: self.server_edit_save(label), width=10).grid(row=0, column=0)
        tk.Button(action_frame, text="Cancel", command=lambda: self.server_edit_win.destroy(), width=10).grid(row=0, column=1)

    def server_edit_save(self, server_label):
        self.config_ini.set(server_label, 'server_name', self.server_name.get())
        self.config_ini.set(server_label, 'server_port', self.server_port.get())
        self.config_ini.set(server_label, 'user_name', self.user_name.get())
        self.config_ini.set(server_label, 'user_password', self.user_password.get())
        self.config_ini.set(server_label, 'ssl', str(self.ssl.get()))
        self.config_ini.set(server_label, 'from_adr', self.from_adr.get())
        with open('config.ini', 'w') as config_ini:
            self.config_ini.write(config_ini)
        self.server_edit_win.destroy()
        self.servers_set_win.destroy()
        self.servers_set_ui()
        self.from_str.set('; '.join(self.config_ini[label]['from_adr'] for label in self.config_ini.sections() if 'smtp' in label))

    def server_delete(self, server_label):
        self.config_ini.remove_section(server_label)
        with open('config.ini', 'w') as config_ini:
            self.config_ini.write(config_ini)
        self.servers_set_win.destroy()
        self.servers_set_ui()
        self.from_str.set('; '.join(self.config_ini[label]['from_adr'] for label in self.config_ini.sections() if 'smtp' in label))

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
            print('Disconnecting from mail server %s' % self.config_ini[name]['server_name'])
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
        self.msg.attach(MIMEText(self.msg_text.get(), 'html', 'utf-8'))


root = tk.Tk()
app = Application(master=root)
app.mainloop()
