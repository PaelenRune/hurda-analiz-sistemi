"""Run locally. Enter the chosen shared password without putting it in history."""
from getpass import getpass
from werkzeug.security import generate_password_hash

if __name__=='__main__':
    password=getpass('Site erişim şifresi (en az 12 karakter): ')
    if len(password)<12:raise SystemExit('Şifre en az 12 karakter olmalı.')
    if password!=getpass('Şifre tekrar: '):raise SystemExit('Şifreler uyuşmuyor.')
    print('APP_PASSWORD_HASH değeri:')
    print(generate_password_hash(password))
