from mimetypes import init

from email import message
import smtplib

import sys
import tkinter as tk
from tkinter import filedialog

import pandas as pd


class Application():
    def __init__(self):
        self.root = tk.Tk()
        self.root.title(u"メール自動送信")
        self.root.geometry("500x400")
        self.data = {'smpt_host': 'smtp.office365.com', 'smpt_port': 587, 'from': '',
                     'pass': '', 'username': 'gbs54850@nuc.kwansei.ac.jp', 'content': '', 'csv_path': ''}
        self.MainMenu()

    def MainMenu(self):
        lbl_port = tk.Label(text='件名')
        lbl_port.grid(row=0, column=0)
        self.txt_subject = tk.Entry(width=30)
        self.txt_subject.grid(row=0, column=1, pady=10)
        #self.txt_port.insert(0, self.data['subject'])

        lbl_path = tk.Label(text='CSVファイル')
        lbl_path.grid(row=1, column=0)
        self.txt_path = tk.Entry(width=30)
        self.txt_path.grid(row=1, column=1, pady=10)
        button_path = tk.Button(text=u'参照', command=self.FileSelect)
        button_path.grid(row=1, column=2, pady=20)

        lbl_smtphost = tk.Label(text='サーバー名')
        lbl_smtphost.grid(row=2, column=0)
        self.txt_smtphost = tk.Entry(width=30)
        self.txt_smtphost.grid(row=2, column=1, padx=10, pady=10)
        self.txt_smtphost.insert(0, self.data['smpt_host'])

        lbl_port = tk.Label(text='ポート番号')
        lbl_port.grid(row=3, column=0)
        self.txt_port = tk.Entry(width=30)
        self.txt_port.grid(row=3, column=1, padx=10, pady=10)
        self.txt_port.insert(0, self.data['smpt_port'])

        lbl_myaddress = tk.Label(text='送信元メールアドレス')
        lbl_myaddress.grid(row=4, column=0)
        self.txt_myaddress = tk.Entry(width=30)
        self.txt_myaddress.grid(row=4, column=1, padx=10, pady=10)
        self.txt_myaddress.insert(0, self.data['from'])

        lbl_pass = tk.Label(text='パスワード')
        lbl_pass.grid(row=5, column=0)
        self.txt_pass = tk.Entry(show='*', width=30)
        self.txt_pass.grid(row=5, column=1, padx=10, pady=10)
        self.txt_pass.insert(0, self.data['pass'])

        lbl_username = tk.Label(text='ユーザ名')
        lbl_username.grid(row=6, column=0)
        self.txt_username = tk.Entry(width=30)
        self.txt_username.grid(row=6, column=1, padx=10, pady=10)
        self.txt_username.insert(0, self.data['username'])

        button_create = tk.Button(text=u'メール生成', command=self.ConfirmMenu)
        button_create.grid(row=7, column=1, padx=10, pady=10)
        self.root.mainloop()

    def ConfirmMenu(self):
        self.data = {'smpt_host': self.txt_smtphost.get(), 'smpt_port': self.txt_port.get(), 'from': self.txt_myaddress.get(
        ), 'pass': self.txt_pass.get(), 'username': self.txt_username.get(), 'subject': self.txt_subject.get(), 'csv_path': self.txt_path.get()}
        self.mail = AutoMail(self.data)
        output = self.mail.CreateContent(0)
        print(self.data)
        self.window_confirm = tk.Toplevel(self.root)
        self.window_confirm.title(u"送信内容")
        self.window_confirm.geometry("500x600")

        # モーダルにする設定
        self.window_confirm.grab_set()        # モーダルにする
        self.window_confirm.focus_set()       # フォーカスを新しいウィンドウをへ移す

        # Canvas Widget を生成
        canvas = tk.Canvas(self.window_confirm)

        # Top Widget上に Scrollbar を生成して配置
        bar = tk.Scrollbar(self.window_confirm, orient=tk.VERTICAL)
        bar.pack(side=tk.RIGHT, fill=tk.Y)
        bar.config(command=canvas.yview) # ScrollbarでCanvasを制御

        # Canvas Widget をTopWidget上に配置
        canvas.config(yscrollcommand=bar.set) # Canvasのサイズ変更をScrollbarに通知
        canvas.config(scrollregion=(0,0,400,1200)) #スクロール範囲
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Frame Widgetを 生成
        frame = tk.Frame(canvas)

        # Frame Widgetを Canvas Widget上に配置（）
        canvas.create_window((0,0), window=frame, anchor=tk.NW, width=500, height=500)

        lbl_subject = tk.Label(frame, text='件名：{}'.format(
            self.data['subject']), padx=10, pady=10, wraplength=400, anchor='w', justify='left')
        lbl_subject.grid(row=0, column=0, columnspan=2)

        lbl_content = tk.Label(frame, text=output, padx=10, pady=10, wraplength=400, anchor='w', justify='left')
        lbl_content.grid(row=1, column=0, columnspan=2)

        button_cancel = tk.Button(frame, text=u'キャンセル', command=self.window_confirm.destroy)
        button_cancel.grid(row=2, column=1, pady=20)

        button_send = tk.Button(
            frame, text=u'送信', command=self.SendMail)
        button_send.grid(row=2, column=0, pady=20)

        print(output)

    def SendMail(self):
        self.mail.SendMail()
        self.window_confirm.destroy()

    def FileSelect(self):
        idir = 'C:'
        filetype = [("CSVファイルb", "*.csv"), ("すべて", "*")]
        self.data['csv_path'] = tk.filedialog.askopenfilename(
            filetypes=filetype, initialdir=idir)
        self.txt_path.insert(0, self.data['csv_path'])


class AutoMail():
    def __init__(self, data):
        self.smtp_host = data['smpt_host']
        self.smtp_port = data['smpt_port']
        self.from_email = data['from']
        self.username = data['username']
        self.password = data['pass']
        self.subject = data['subject']
        self.csv_input = pd.read_csv(
            data['csv_path'], encoding='utf_8', sep=',')

        self.Get_to()

    def Get_to(self):
        self.to_email = self.csv_input['メールアドレス']
        self.name = self.csv_input['名前']
        self.number = self.csv_input['ID']
        self.pswd = self.csv_input['パスワード']

    def CreateContent(self, i):
        return f'''{self.name[i]}さん，はじめまして．\n
4月11日からの授業開始に伴い，課題の提出ページについての案内とセットアップについて連絡します．\n
\n

・セットアップについて
初回授業の前日までにセットアップとして以下の3つの項目を済ませておいてください．

1. ログイン
以下のIDとパスワードで各個人の課題提出ページにログインしてください．

ID : 学生番号の下4桁
PW: {self.pswd[i]}
※ ログイン後は必ず，学籍番号と氏名が自分のものであるか確認してください．

2.パスワードの変更
課題提出ページの"パスワード変更"から新しいパスワードに変更してください．
※ 初期パスワードのままにしないようにお願いします．

ログインやパスワードの変更，課題提出のテストについて上手くいかない場合は，お手数ですがTAまでメールにて連絡お願いします．
'''

    def SendMail(self):
        for i in range(len(self.csv_input)):
            self.content = self.CreateContent(i)

            # メッセージ内容
            msg = message.EmailMessage()
            msg.set_content(self.content)
            msg['Subject'] = self.subject
            msg['From'] = self.from_email
            msg['To'] = self.to_email[i]

            server = smtplib.SMTP(self.smtp_host, self.smtp_port)
            server.ehlo()
            server.starttls()
            server.ehlo()
            server.login(self.username, self.password)
            server.send_message(msg)
            print(f'{self.name[i]}さんに送信が完了しました')


if __name__ == '__main__':
    app = Application()
    #data = {'smpt_host':'smtp.office365.com', 'smpt_port':587, 'from':'', 'pass':'', 'username':'gbs54850@nuc.kwansei.ac.jp'}
    #mail = AutoMail(data)
    # mail.Get_to()
