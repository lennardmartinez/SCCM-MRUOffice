@ECHO OFF
c:
cd %APPDATA%\mru

reg export "hkcu\Software\Microsoft\Office\14.0\Excel\File MRU" "Excel.reg" /y >nul
reg query "hkcu\Software\Microsoft\Office\15.0\Excel\User MRU" |findstr /r "AD_" > newmru

set /p clean= <newmru
del newmru
set target=HKEY_CURRENT_USER\Software\Microsoft\Office\14.0\Excel

"\\WESC-FP-01\IT Updates\Scripts\Office2013\Production\BatchSubstitute.bat" "%target%" "%clean%" Excel.reg >NUL && reg import Excel.reg && EXIT 

