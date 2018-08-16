#!/usr/bin/python3
# -*- coding: utf-8 -*-

import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import csv
import configparser
import string
import re
import smtplib
from time import sleep
import uuid
from tkinter import simpledialog


class Application(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master.title('JMail alpha v0.0.1')

        self.config_ini = None
        self.from_str = tk.StringVar()
        self.msg_subj = tk.StringVar()
        self.msg = MIMEMultipart()
        self.msg_text = None

        self.server_label = tk.StringVar()
        self.server_name = tk.StringVar()
        self.server_port = tk.StringVar()
        self.user_name = tk.StringVar()
        self.user_password = tk.StringVar()
        self.ssl = tk.IntVar()
        self.from_adr = tk.StringVar()

        self.server_conn_list = []
        self.to_adr_list = []
        self.to_adr_num = 0
        self.black_list = []

        self.servers_set_win = None
        self.black_list_win = None

        self.progress_send = None
        self.but_server_del = None
        self.combo_servers = None
        self.box_black_list = None

        self.stop = False

        self.conf_identify()
        self.main_ui()

        self.master.update_idletasks()
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        frm_width = self.master.winfo_rootx() - self.master.winfo_x()
        titlebar_height = self.master.winfo_rooty() - self.master.winfo_y()
        win_width = self.master.winfo_width() + 2 * frm_width
        win_height = self.master.winfo_height() + titlebar_height + frm_width
        self.master.geometry('+%s+%s' % (int(screen_width/2 - win_width/2), int(screen_height/2 - win_height/2)))

    def conf_identify(self):
        try:
            open('config.ini').read()
        except FileNotFoundError:
            open('config.ini', 'w')
        finally:
            self.config_ini = configparser.ConfigParser()
            self.config_ini.read('config.ini')

        try:
            with open('black_list.txt', 'r') as black_list:
                self.black_list = [adr for adr in black_list.read().splitlines() if adr]
        except FileNotFoundError:
            pass

    def main_ui(self):
        menu_bar = tk.Menu(self)
        self.master.config(menu=menu_bar)
        file_menu = tk.Menu(menu_bar, tearoff=0)
        file_menu.add_command(label="Exit", command=lambda: self.quit())
        menu_bar.add_cascade(label="File", menu=file_menu)
        settings_menu = tk.Menu(menu_bar, tearoff=0)
        settings_menu.add_command(label='Servers', command=self.servers_set_ui)
        settings_menu.add_command(label='Black list', command=self.black_list_ui)
        menu_bar.add_cascade(label='Settings', menu=settings_menu)

        frame_mail = tk.Frame(self.master)
        frame_mail.grid(row=0, column=0, sticky='w', padx=5)
        tk.Label(frame_mail, text='To:', width=10, anchor='w').grid(row=0, column=0, sticky='w')
        tk.Button(frame_mail, text="Browse...", command=self.to_adr_list_open, width=10).grid(row=0, column=1, sticky='w')
        tk.Label(frame_mail, text='From:', width=10, anchor='w').grid(row=1, column=0, sticky='w')
        if self.config_ini:
            self.from_str.set('; '.join(self.config_ini[label]['from_adr'] for label in self.config_ini.sections() if 'smtp' in label))
        else:
            self.from_str.set('')
        tk.Label(frame_mail, textvariable=self.from_str).grid(row=1, column=1, sticky='w')
        tk.Label(frame_mail, text='Subject:', width=10, anchor='w').grid(row=2, column=0, sticky='w')
        try:
            self.msg_subj.set(self.config_ini['config']['msg_subj'])
        except (KeyError, TypeError):
            self.msg_subj.set('Message subject')
        tk.Entry(frame_mail, textvariable=self.msg_subj, width=75).grid(row=2, column=1, sticky='w')
        tk.Button(frame_mail, text="Save", command=self.msg_subj_save, width=10).grid(row=2, column=2, sticky='w')
        tk.Label(frame_mail, text='Mail:', width=10, anchor='w').grid(row=3, column=0, sticky='w')
        tk.Button(frame_mail, text="Browse...", command=self.msg_open, width=10).grid(row=3, column=1, sticky='w')

        ttk.Separator(self.master).grid(row=1, column=0, columnspan=1, sticky="ew", padx=5, pady=5)

        frame_action = tk.Frame(self.master)
        frame_action.grid(row=2, column=0, sticky='w', padx=5)
        tk.Button(frame_action, text="Connect", command=self.connect_servers, width=10).grid(row=0, column=0, sticky='w')
        tk.Button(frame_action, text="Disconnect", command=self.disconnect_servers, width=10).grid(row=0, column=1, sticky='w', padx=5)
        tk.Button(frame_action, text="Start", command=lambda: self.star_sending(), width=10).grid(row=1, column=0, sticky='w')
        tk.Button(frame_action, text="Stop", command=lambda: self.stop_sending(), width=10).grid(row=1, column=1, sticky='w', padx=5)

        self.progress_send = ttk.Progressbar(self.master, mode='determinate')
        self.progress_send.grid(row=3, column=0, columnspan=1, sticky="ew", padx=5, pady=5)

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
            with open(to_file, 'r') as f:
                self.to_adr_list = [r[0] for r in csv.reader(f) if r]
                self.progress_send['maximum'] = len(self.to_adr_list)

    def msg_open(self):
        dlg = filedialog.Open(self, filetypes=[('HTML files', '*.html'), ('All files', '*')])
        msg_file = dlg.show()
        if msg_file:
            with open(msg_file, "r") as f:
                self.msg_text = f.read()

    def servers_set_ui(self):
        self.servers_set_win = tk.Toplevel(self.master)
        self.servers_set_win.geometry('+%s+%s' % (self.master.winfo_x()+20, self.master.winfo_y()+20))
        self.servers_set_win.title('Servers settings')
        self.servers_set_win.grab_set()

        frame_servers = tk.Frame(self.servers_set_win)
        frame_servers.grid(row=0, column=0, sticky='nsew', padx=5, pady=5)
        tk.Label(frame_servers, text='Server').grid(row=0, column=0)
        self.combo_servers = ttk.Combobox(frame_servers, values=[label for label in self.config_ini.sections() if 'smtp' in label])
        self.combo_servers.grid(row=0, column=1, padx=5)
        self.combo_servers.current(0) if [label for label in self.config_ini.sections() if 'smtp' in label] else None
        self.combo_servers.bind('<<ComboboxSelected>>', lambda _: self.server_edit(self.combo_servers.get()))
        tk.Button(frame_servers, text='New', command=lambda: self.server_edit(), width=7).grid(row=0, column=2, sticky='w')
        self.but_server_del = tk.Button(frame_servers, text='Del', command=lambda: self.server_delete(), width=7)
        self.but_server_del.grid(row=0, column=3, sticky='w', padx=5)

        self.server_edit(self.combo_servers.get() if self.combo_servers.current() != -1 else None)

        frame_server = tk.LabelFrame(self.servers_set_win, text='Server')
        frame_server.grid(row=1, column=0, sticky='w', padx=5, ipadx=2, ipady=2)
        tk.Label(frame_server, text='Address: ', width=13, anchor='w').grid(row=0, column=0, sticky='w')
        tk.Entry(frame_server, textvariable=self.server_name, width=35).grid(row=0, column=1, sticky='w')
        tk.Label(frame_server, text='Port: ', width=13, anchor='w').grid(row=1, column=0, sticky='w')
        tk.Entry(frame_server, textvariable=self.server_port, width=35).grid(row=1, column=1, sticky='w')

        frame_user = tk.LabelFrame(self.servers_set_win, text='User')
        frame_user.grid(row=2, column=0, sticky='w', padx=5, pady=5, ipadx=2, ipady=2)
        tk.Label(frame_user, text='Name: ', width=13, anchor='w').grid(row=0, column=0, sticky='w')
        tk.Entry(frame_user, textvariable=self.user_name, width=35).grid(row=0, column=1, sticky='w')
        tk.Label(frame_user, text='Password: ', width=13, anchor='w').grid(row=1, column=0, sticky='w')
        tk.Entry(frame_user, textvariable=self.user_password, width=35, show='*').grid(row=1, column=1, sticky='w')
        tk.Label(frame_user, text='SSL: ', width=13, anchor='w').grid(row=2, column=0, sticky='w')
        tk.Checkbutton(frame_user, variable=self.ssl).grid(row=2, column=1, sticky='w')
        tk.Label(frame_user, text='From address: ', width=13, anchor='w').grid(row=3, column=0, sticky='w')
        tk.Entry(frame_user, textvariable=self.from_adr, width=35).grid(row=3, column=1, sticky='w')

        frame_action = tk.Frame(self.servers_set_win)
        frame_action.grid(row=3, column=0, sticky='e', padx=5, pady=5)
        tk.Button(frame_action, text="Save", command=lambda: self.servers_set_save(), width=10).grid(row=0, column=0, padx=5)
        tk.Button(frame_action, text="Cancel", command=lambda: self.servers_set_win.destroy(), width=10).grid(row=0, column=1)

    def server_edit(self, server_label=None):
        self.combo_servers.set(server_label if server_label else '')
        self.but_server_del['state'] = 'active' if server_label else 'disable'
        self.server_label.set(server_label if server_label else 'smtp-' + uuid.uuid4().hex[:5])
        self.server_name.set(self.config_ini[server_label]['server_adr'] if server_label else '')
        self.server_port.set(self.config_ini[server_label]['server_port'] if server_label else '')
        self.user_name.set(self.config_ini[server_label]['user_name'] if server_label else '')
        self.user_password.set(self.config_ini[server_label]['user_password'] if server_label else '')
        self.ssl.set(int(self.config_ini[server_label]['ssl']) if server_label else 0)
        self.from_adr.set(self.config_ini[server_label]['from_adr'] if server_label else '')

    def server_delete(self):
        if messagebox.askokcancel(self.config_ini[self.server_label.get()]['server_adr'], 'Would you like to delete %s?' % self.config_ini[self.server_label.get()]['server_adr']):
            self.config_ini.remove_section(self.server_label.get())
            self.combo_servers['values'] = [label for label in self.config_ini.sections() if 'smtp' in label]
            self.combo_servers.current(0) if [label for label in self.config_ini.sections() if 'smtp' in label] else None
            self.server_edit(self.combo_servers.get() if self.combo_servers.current() != -1 else None)
            with open('config.ini', 'w') as config_ini:
                self.config_ini.write(config_ini)
            self.from_str.set('; '.join(self.config_ini[label]['from_adr'] for label in self.config_ini.sections() if 'smtp' in label))

    def servers_set_save(self):
        if self.server_label.get() not in self.config_ini.sections():
            self.config_ini.add_section(self.server_label.get())
        self.config_ini.set(self.server_label.get(), 'server_adr', self.server_name.get())
        self.config_ini.set(self.server_label.get(), 'server_port', self.server_port.get())
        self.config_ini.set(self.server_label.get(), 'user_name', self.user_name.get())
        self.config_ini.set(self.server_label.get(), 'user_password', self.user_password.get())
        self.config_ini.set(self.server_label.get(), 'ssl', str(self.ssl.get()))
        self.config_ini.set(self.server_label.get(), 'from_adr', self.from_adr.get())
        with open('config.ini', 'w') as config_ini:
            self.config_ini.write(config_ini)
        self.combo_servers['values'] = [label for label in self.config_ini.sections() if 'smtp' in label]
        self.server_edit(self.server_label.get())
        self.from_str.set('; '.join(self.config_ini[label]['from_adr'] for label in self.config_ini.sections() if 'smtp' in label))

    def black_list_ui(self):
        self.black_list_win = tk.Toplevel(self.master)
        self.black_list_win.geometry('+%s+%s' % (self.master.winfo_x() + 20, self.master.winfo_y() + 20))
        self.black_list_win.title('Black list')
        self.black_list_win.grab_set()

        frame_list = tk.Frame(self.black_list_win)
        frame_list.grid(row=0, column=0, sticky='w', padx=5, pady=5)
        self.box_black_list = tk.Listbox(frame_list, selectmode='extended', width=40)
        self.box_black_list.grid(row=0, column=0)
        for adr in self.black_list:
            self.box_black_list.insert('end', adr)
        scroll = tk.Scrollbar(frame_list, command=self.box_black_list.yview)
        scroll.grid(row=0, column=1, sticky='ns')
        self.box_black_list.config(yscrollcommand=scroll.set)

        frame_edit = tk.Frame(self.black_list_win)
        frame_edit.grid(row=0, column=1, sticky='n', padx=5, pady=5)
        tk.Button(frame_edit, text='Add', command=lambda: self.black_list_add(), width=7).grid(row=0, column=0)
        but_del_adr = tk.Button(frame_edit, text='Delete', state='disable', command=lambda: self.black_list_delete(), width=7)
        but_del_adr.grid(row=1, column=0, pady=5)

        self.box_black_list.bind('<<ListboxSelect>>', lambda _: but_del_adr.config(state='active') if self.box_black_list.curselection() else None)

        frame_action = tk.Frame(self.black_list_win)
        frame_action.grid(row=1, column=0, sticky='e', columnspan=2, padx=5, pady=5)
        tk.Button(frame_action, text="Save", command=lambda: self.black_list_save(), width=10).grid(row=0, column=0, padx=5)
        tk.Button(frame_action, text="Cancel", command=lambda: self.black_list_win.destroy(), width=10).grid(row=0, column=1)

    def black_list_add(self):
        answer = simpledialog.askstring('Black address', 'Please enter email.', parent=self.black_list_win)
        if answer:
            self.box_black_list.insert('end', answer)

    def black_list_delete(self):
        selected_adr = [self.box_black_list.get(i) for i in self.box_black_list.curselection()]
        if messagebox.askokcancel('Delete address', 'Would you like to delete %s?' % ', '.join(selected_adr)):
            for i in reversed(self.box_black_list.curselection()):
                self.box_black_list.delete(i)
            self.black_list = [adr for adr in self.box_black_list.get(0, 'end')]
            with open('black_list.txt', 'w') as black_list:
                for i in self.box_black_list.get(0, 'end'):
                    black_list.write("%s\n" % i)

    def black_list_save(self):
        self.black_list = [adr for adr in self.box_black_list.get(0, 'end')]
        if self.black_list:
            with open('black_list.txt', 'w') as black_list:
                for i in self.box_black_list.get(0, 'end'):
                    black_list.write("%s\n" % i)

    def connect_servers(self):
        self.server_conn_list = []
        if self.config_ini:
            server_label_list = [label for label in self.config_ini.sections() if 'smtp' in label.lower()]
            for server_label in server_label_list:
                connect = self.connect_mail_server(server_label)
                if connect:
                    self.server_conn_list.append((connect, server_label))

    def connect_mail_server(self, server_label):
        server_name = self.config_ini[server_label]['server_adr']
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
            print('Disconnecting from mail server %s' % self.config_ini[name]['server_adr'])
            server.quit()

    def stop_sending(self):
        self.stop = True

    def star_sending(self):
        if self.stop:
            self.stop = False
            return
        if not self.server_conn_list:
            print('No servers to send')
            return
        if self.to_adr_num + 1 > len(self.to_adr_list):
            return
        to_adr = self.to_adr_list[self.to_adr_num]
        to_adr_ascii = ''.join(i for i in to_adr if i in string.ascii_letters + string.digits + '@._-')
        if not re.search(r'(\w+@[a-zA-Z_]+?\.[a-zA-Z]{2,6})', to_adr_ascii):
            # print(counter, "Message DON'T send to %s. Incorrect format!" % to_adr)
            self.to_adr_num += 1
            self.master.after(10, lambda: self.star_sending())
            return
        elif to_adr_ascii in self.black_list:
            # print(counter, "Message DON'T send. %s in black list!" % to_adr)
            self.to_adr_num += 1
            self.master.after(10, lambda: self.star_sending())
            return
        for server, name in self.server_conn_list:
            server.helo()
            self.get_message(self.config_ini[name]['from_adr'], to_adr_ascii)
            try:
                server.sendmail(self.config_ini[name]['from_adr'], to_adr_ascii, self.msg.as_string())
            except smtplib.SMTPServerDisconnected as exc:
                # print('Server %s disconnected' % name)
                # print('%s: %s' % (exc.__class__.__name__, exc))
                self.server_conn_list.remove((server, name))
                continue
            except Exception as exc:
                # print(counter, "Message DON'T send to", to_adr)
                # print('%s: %s' % (exc.__class__.__name__, exc))
                return
            else:
                self.server_conn_list.remove((server, name))
                self.server_conn_list.append((server, name))
                # print(counter, 'Message send to', to_adr_ascii)
                # sleep(13 / len(self.server_conn_list) if len(self.server_conn_list) else 13)
                self.to_adr_num += 1
                self.progress_send['value'] = self.to_adr_num
                self.master.after(10, lambda: self.star_sending())
                break

    def get_message(self, from_, to):
        self.msg['From'] = from_
        self.msg['To'] = to
        self.msg['Subject'] = self.msg_subj.get()
        self.msg.attach(MIMEText(self.msg_text, 'html', 'utf-8'))


root = tk.Tk()
app = Application(master=root)
app.mainloop()
