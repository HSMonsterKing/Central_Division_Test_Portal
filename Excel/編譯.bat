@chcp 65001
@taskkill /t /f /im excel.exe
cd "C:\零用金測試\Excel\Temp"
del /F /Q *.*
%windir%\Microsoft.NET\Framework64\v4.0.30319\aspnet_compiler.exe -v / -p "C:\零用金測試" "C:\零用金" -f
copy /v /y "C:\零用金\Web.Release.config" "C:\零用金\Web.config"
mkdir "C:\零用金\Excel\Temp"
pause
