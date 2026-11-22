#!/usr/bin/env python3
"""Encrypt data.json for the public repo. Usage: FIN_PASSWORD=... python3 encrypt.py data.json data.enc.json
AES-256-GCM with a PBKDF2-SHA256 key (600,000 iterations). The page decrypts in the browser with WebCrypto.
The password is never written to the repo."""
import sys, os, json, base64
from cryptography.hazmat.primitives.ciphers.aead import AESGCM
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
from cryptography.hazmat.primitives import hashes
ITER = 600000
pw = os.environ["FIN_PASSWORD"].encode()
plain = open(sys.argv[1], "rb").read()
salt, iv = os.urandom(16), os.urandom(12)
key = PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=salt, iterations=ITER).derive(pw)
ct = AESGCM(key).encrypt(iv, plain, None)
b = lambda x: base64.b64encode(x).decode()
json.dump({"v": 1, "iter": ITER, "salt": b(salt), "iv": b(iv), "ct": b(ct)}, open(sys.argv[2], "w"))
print("encrypted", len(plain), "bytes")
