@echo off
cd /d C:\IDD
"C:\Program Files\Git\mingw64\bin\git.exe" add .
"C:\Program Files\Git\mingw64\bin\git.exe" commit -m "Update"
"C:\Program Files\Git\mingw64\bin\git.exe" push
echo.
echo GitHub upload complete!
pause
