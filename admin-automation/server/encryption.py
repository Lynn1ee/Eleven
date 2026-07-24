"""SMTP 密码加密/解密，使用 Fernet (AES-128-CBC + HMAC)"""
import os
from cryptography.fernet import Fernet

DATA_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "data")
KEY_FILE = os.path.join(DATA_DIR, "secret.key")


def _get_or_create_key():
    os.makedirs(DATA_DIR, exist_ok=True)
    if os.path.exists(KEY_FILE):
        with open(KEY_FILE, "rb") as f:
            return f.read()
    key = Fernet.generate_key()
    with open(KEY_FILE, "wb") as f:
        f.write(key)
    return key


def encrypt(plain_text):
    key = _get_or_create_key()
    f = Fernet(key)
    return f.encrypt(plain_text.encode("utf-8")).decode("utf-8")


def decrypt(encrypted_text):
    key = _get_or_create_key()
    f = Fernet(key)
    return f.decrypt(encrypted_text.encode("utf-8")).decode("utf-8")
