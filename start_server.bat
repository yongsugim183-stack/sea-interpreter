@echo off
cd /d C:\Users\김용수\sea-interpreter
set PYTHONIOENCODING=utf-8
set HF_HUB_DISABLE_SSL_VERIFICATION=1
set CURL_CA_BUNDLE=
set REQUESTS_CA_BUNDLE=
"C:\Users\김용수\AppData\Local\Programs\Python\Python314\python.exe" server.py
