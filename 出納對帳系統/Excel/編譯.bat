@taskkill /t /f /im excel.exe
cd "C:\�X�ǹ�b����\Excel\Temp"
del /F /Q *.*
%windir%\Microsoft.NET\Framework64\v4.0.30319\aspnet_compiler.exe -v / -p "C:\�X�ǹ�b����" "C:\�X�ǹ�b�t��" -f
copy /v /y "C:\�X�ǹ�b�t��\Web.Release.config" "C:\�X�ǹ�b�t��\Web.config"
mkdir "C:\�X�ǹ�b�t��\Excel\Temp"
pause
