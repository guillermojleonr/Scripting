'Removes hyphen character from a collection of files.

Set ObjFSO = CreateObject("Scripting.FileSystemObject")

Set Folder = objFso.GetFolder("G:\Mi unidad\Imported") 'Carpeta a operar

For Each File In Folder.Files
	'Almacena el nombre del archivo, reemplaza los "-" y muestra el archivo a tratar
	sNewFile = File.Name
	sNewFile = Replace(sNewFile,"-","") 
	MsgBox(sNewFile)

	'Cambia el nombre del archivo en caso que sea necesario
	If sNewFile <> File.Name Then
	File.Move(Folder & "\" & sNewFile)
	End If
Next