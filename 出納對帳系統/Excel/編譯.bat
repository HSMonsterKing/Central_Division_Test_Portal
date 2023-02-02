@taskkill /t /f /im excel.exe
cd "C:\出納對帳測試\Excel\Temp"
del /F /Q *.*
%windir%\Microsoft.NET\Framework64\v4.0.30319\aspnet_compiler.exe -v / -p "C:\出納對帳測試" "C:\出納對帳系統" -f
copy /v /y "C:\出納對帳系統\Web.Release.config" "C:\出納對帳系統\Web.config"
mkdir "C:\出納對帳系統\Excel\Temp"
pause
