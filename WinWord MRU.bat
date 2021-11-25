@echo off
c:
cd %APPDATA%\mru

reg export "hkcu\Software\Microsoft\Office\14.0\Word\File MRU" "WinWord.reg" /y >nul
reg query "hkcu\Software\Microsoft\Office\15.0\Word\User MRU" |findstr /r "AD_" > newmru

set /p clean= <newmru
del newmru
set target=HKEY_CURRENT_USER\Software\Microsoft\Office\14.0\Word

"\\WESC-FP-01\IT Updates\Scripts\Office2013\Production\BatchSubstitute.bat" "%target%" "%clean%" WinWord.reg > NUL && reg import WinWord.reg && Exit
