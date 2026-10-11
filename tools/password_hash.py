"""Run locally. Enter the chosen shared password without putting it in history."""
from getpass import getpass
from werkzeug.security import generate_password_hash
import subprocess
import sys
import os
from pathlib import Path

def windows_password_window():
    import tkinter as tk
    from tkinter import ttk, messagebox

    root = tk.Tk()
    root.title('Hurda Analiz Sistemi — site şifresi')
    root.resizable(False, False)
    frame = ttk.Frame(root, padding=24)
    frame.pack(fill='both', expand=True)
    ttk.Label(frame, text='Site şifrenizi iki kez girin (en az 8 karakter).').pack(anchor='w')
    ttk.Label(frame, text='Yazdığınız karakterler yıldız olarak görünür.').pack(anchor='w', pady=(0, 12))
    ttk.Label(frame, text='Şifre:').pack(anchor='w')
    first = ttk.Entry(frame, show='*', width=48)
    first.pack(pady=(0, 8))
    ttk.Label(frame, text='Şifre tekrar:').pack(anchor='w')
    second = ttk.Entry(frame, show='*', width=48)
    second.pack(pady=(0, 12))
    status = ttk.Label(frame, text='', wraplength=400)

    def prepare():
        password = first.get()
        if len(password) < 8:
            messagebox.showerror('Şifre', 'Şifre en az 8 karakter olmalı.', parent=root)
            return
        if password != second.get():
            messagebox.showerror('Şifre', 'Şifreler uyuşmuyor.', parent=root)
            return
        hashed = generate_password_hash(password)
        powershell = Path(os.environ['SystemRoot']) / 'System32/WindowsPowerShell/v1.0/powershell.exe'
        try:
            subprocess.run([str(powershell), '-NonInteractive', '-NoProfile', '-Command',
                            'Set-Clipboard -Value ([Console]::In.ReadToEnd().Trim())'],
                           input=hashed, text=True, check=True,
                           creationflags=subprocess.CREATE_NO_WINDOW)
        except subprocess.SubprocessError:
            messagebox.showerror('Pano', 'Panoya kopyalanamadı. Tekrar deneyin.', parent=root)
            return
        first.delete(0, tk.END)
        second.delete(0, tk.END)
        status.config(text='Şifre özeti panoya kopyalandı. Render’daki APP_PASSWORD_HASH satırının Value alanına Ctrl+V ile yapıştırın, ardından Deploy web service düğmesine basın. Siteye girişte belirlediğiniz şifreyi kullanacaksınız.')

    ttk.Button(frame, text='Şifre özetini panoya kopyala', command=prepare).pack(fill='x')
    status.pack(anchor='w', pady=(12, 0))
    first.focus_set()
    root.mainloop()


if __name__=='__main__' and sys.platform=='win32':
    windows_password_window()
elif __name__=='__main__':
    password=getpass('Site erişim şifresi (en az 8 karakter): ')
    if len(password)<8:raise SystemExit('Şifre en az 8 karakter olmalı.')
    if password!=getpass('Şifre tekrar: '):raise SystemExit('Şifreler uyuşmuyor.')
    hashed=generate_password_hash(password)
    print('APP_PASSWORD_HASH değeri:')
    print(hashed)
