@taskkill /t /f /im excel.exe
cd "C:\�j�v����\Excel\Temp"
del /F /Q *.*
%windir%\Microsoft.NET\Framework64\v4.0.30319\aspnet_compiler.exe -v / -p "C:\�j�v����" "C:\�j�v�l��" -f
copy /v /y "C:\�j�v�l��\Web.Release.config" "C:\�j�v�l��\Web.config"
mkdir "C:\�j�v�l��\Excel\Temp"
pause
