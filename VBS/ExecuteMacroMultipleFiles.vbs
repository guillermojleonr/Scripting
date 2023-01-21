'Option Explicit
'Opera una serie de archivos contenidos en una carpeta, ejecuta una macro contenida en libro personal y mueve el archivo a un repositorio

On Error Resume Next

Set ObjFSO = CreateObject("Scripting.FileSystemObject")
Set objExcel = CreateObject("Excel.Application")
Set Folder = objFso.GetFolder("G:\Mi unidad\04_CUENTAS\CASA LIVING\CUADRES DIARIOS") 'Carpeta a operar
Set FolderNoFlex = objFSO.GetFolder("G:\Mi unidad\04_CUENTAS\CASA LIVING\CUADRES DIARIOS\NO FLEX") 'Carpeta repositorio NO FLEX
Set FolderFlex = objFSO.GetFolder("G:\Mi unidad\04_CUENTAS\CASA LIVING\CUADRES DIARIOS\FLEX") 'Carpeta repositorio FLEX

ObjExcel.DisplayAlerts = False

For Each File In Folder.Files
	'Almacena el nombre del archivo, reemplaza los "-" y muestra el archivo a tratar
	sNewFile = File.Name
	sNewFile = Replace(sNewFile,"-","") 
	MsgBox(sNewFile)

	'Cambia el nombre del archivo en caso que sea necesario
	If sNewFile <> File.Name Then
	File.Move(Folder & "\" & sNewFile)
	End If

	'Ignora el archivo desktop.ini
	If InStr(1, sNewFile, "desktop.ini") > 0 Then
		WScript.Echo "Llegamos al desktop.ini"
		WScript.Quit
	End If
	
	'Opera la macro correspondiente dependiendo del tipo de archivo
	If InStr(1, sNewFile, "Rocket") > 0 Then
	Wscript.Sleep 5000
		Set objWorkbook = objExcel.Workbooks.Open(Folder & "\" & sNewFile)
		objExcel.Application.Run "'C:\Users\Gear PC\AppData\Roaming\Microsoft\Excel\XLSTART\PERSONAL.XLSB'!CL.Seteo_Cuadre_CLNOFLEX"
		objWorkbook.Save
		objExcel.ActiveWorkbook.Close
		File.Move(FolderNoFlex & "\" & sNewFile)
	Else
	Wscript.Sleep 5000
		Set objWorkbook = objExcel.Workbooks.Open(Folder & "\" & sNewFile)
		objExcel.Application.Run "'C:\Users\Gear PC\AppData\Roaming\Microsoft\Excel\XLSTART\PERSONAL.XLSB'!CL.Seteo_Cuadre_CLFLEX"
		objWorkbook.Save
		objExcel.ActiveWorkbook.Close
		File.Move(FolderFlex & "\" & sNewFile)
	End If

Next

ErrHandlr:

If Err.Number <> 0 Then ' Catch your error
   WScript.Echo "Error while opening: " & Err.Number & " " & Err.Description
   XLWkbk.Close False ' Close your workbook.
   xlApp.Quit ' Quit the excel program. 
   WScript.Quit 
   Err.Clear
End If

WScript.Echo "Finished."
WScript.Quit