'Opera una serie de archivos contenidos en una carpeta, ejecuta una macro contenida en libro personal y mueve el archivo a un repositorio

Set ObjFSO = CreateObject("Scripting.FileSystemObject")
Set objExcel = CreateObject("Excel.Application")
Set Folder = objFso.GetFolder("C:\Users\guill\Documents\Treatment") 'Carpeta a operar
Set Folder2 = objFSO.GetFolder("C:\Users\guill\Documents\Repository") 'Carpeta repositorio

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
	If InStr(9, sNewFile, "registro") > 0 Then
	Wscript.Sleep 3000
		Set objWorkbook = objExcel.Workbooks.Open(Folder & "\" & sNewFile)
		objExcel.Application.Run "'C:\Users\guill\AppData\Roaming\Microsoft\Excel\XLSTART\PERSONAL.XLSB'!AjusterReportCL.Seteo_Cuadre_2"
		objWorkbook.Save
		objExcel.ActiveWorkbook.Close
	Else
	Wscript.Sleep 3000
		Set objWorkbook = objExcel.Workbooks.Open(Folder & "\" & sNewFile)
		objExcel.Application.Run "'C:\Users\guill\AppData\Roaming\Microsoft\Excel\XLSTART\PERSONAL.XLSB'!AjusterReportCL.Seteo_Cuadre_1"
		objWorkbook.Save
		objExcel.ActiveWorkbook.Close
	End If

	'Mueve el archivo ya operado al directorio repositorio
	File.Move(Folder2 & "\" & sNewFile)
Next

WScript.Echo "Finished."
WScript.Quit