@taskkill /t /f /im excel.exe
cd "C:\大宗測試\Excel\Temp"
del /F /Q *.*
%windir%\Microsoft.NET\Framework64\v4.0.30319\aspnet_compiler.exe -v / -p "C:\大宗測試" "C:\大宗郵件" -f
copy /v /y "C:\大宗郵件\Web.Release.config" "C:\大宗郵件\Web.config"
mkdir "C:\大宗郵件\Excel\Temp"
pause
