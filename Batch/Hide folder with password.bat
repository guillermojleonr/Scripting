:: Proteger una carpeta con contraseña
:: Edita el archivo "proteger carpeta con contraseña" cambiando la clave "youtube" por la clave que quieres
:: Guardalo con una extension .bat
:: Ejecuta el archivo
:: Mete en la carpeta locker lo que deseas proteger
:: Coloca S + enter

cls 
@ECHO OFF 
title Folder Locker 
if EXIST "Control Panel.{21EC2020-3AEA-1069-A2DD-08002B30309D}" goto UNLOCK 
if NOT EXIST Locker goto MDLOCKER 
:CONFIRM 
echo Esta Seguro de que quiere proteger la Carpeta(S/N) 
set/p "cho=>" 
if %cho%==S goto LOCK 
if %cho%==s goto LOCK 
if %cho%==n goto END 
if %cho%==N goto END 
echo Invalid choice. 
goto CONFIRM 
:LOCK 
ren Locker "Control Panel.{21EC2020-3AEA-1069-A2DD-08002B30309D}" 
attrib +h +s "Control Panel.{21EC2020-3AEA-1069-A2DD-08002B30309D}" 
echo Folder locked 
goto End 
:UNLOCK 
echo Ingrese su Contraseña para proteger su carpeta 
set/p "pass=>" 
if NOT %pass%== AQUI VA LA CONTRASEÑA goto FAIL 
attrib -h -s "Control Panel.{21EC2020-3AEA-1069-A2DD-08002B30309D}" 
ren "Control Panel.{21EC2020-3AEA-1069-A2DD-08002B30309D}" Locker 
echo Folder Unlocked successfully 
goto End 
AIL 
echo Invalid password 
goto end 
:MDLOCKER 
md Locker 
echo Locker created successfully 
goto End 
:End 
