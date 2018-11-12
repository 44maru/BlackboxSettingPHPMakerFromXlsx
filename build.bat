@echo off

call venv\Scripts\activate
rem pyInstaller --onefile --noconsole phpMaker.py
pyInstaller phpmaker.spec
rem move dist\jsonMaker.exe dist\phpMaker.exe

pause