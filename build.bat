@echo off

rem call venv\Scripts\activate
call py -m pipenv shell
rem pyInstaller --onefile --noconsole phpMaker.py
pyInstaller phpmaker.spec
rem move dist\jsonMaker.exe dist\phpMaker.exe

pause