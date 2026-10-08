"""Run locally. Enter the chosen shared password without putting it in history."""
from getpass import getpass
from werkzeug.security import generate_password_hash
import subprocess
import sys

if __name__=='__main__':
    password=getpass('Site erişim şifresi (en az 12 karakter): ')
    if len(password)<12:raise SystemExit('Şifre en az 12 karakter olmalı.')
    if password!=getpass('Şifre tekrar: '):raise SystemExit('Şifreler uyuşmuyor.')
    hashed=generate_password_hash(password)
    if sys.platform=='win32':
        subprocess.run(['powershell','-NoProfile','-Command','$input | Set-Clipboard'],input=hashed,text=True,check=True)
        print('Şifre özeti panoya kopyalandı. Render APP_PASSWORD_HASH alanına Ctrl+V ile yapıştırın.')
        print('Düz şifrenizi sohbete yazmayın. Siteye girişte belirlediğiniz düz şifreyi kullanacaksınız.')
    else:
        print('APP_PASSWORD_HASH değeri:')
        print(hashed)
