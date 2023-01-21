echo off
for /f "delims=" %%a in ('wmic OS Get localdatetime  ^| find "."') do set dt=%%a
set YYYY=%dt:~0,4%
set MM=%dt:~4,2%
set DD=%dt:~6,2%
set HH=%dt:~8,2%
set Min=%dt:~10,2%
set Sec=%dt:~12,2%

set stamp=%YYYY%%MM%%DD%_%HH%%Min%%Sec%

ROBOCOPY "C:\Users\Gear PC\Desktop\ruta-origen\prueba.txt" "C:\Users\Gear PC\Desktop\ruta-destino\prueba_%stamp%.txt"


$datetime = get-date -f MMddyy-hhmmtt 

rename-item C:\Users\Gear PC\Desktop\ruta-destino\prueba.txt -newname ($datetime + ".txt")