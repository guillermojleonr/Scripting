'Changes a file name from a collection of files.

Set ObjFSO = CreateObject("Scripting.FileSystemObject")
Set Folder = objFso.GetFolder("G:\Mi unidad\Folder") 'Carpeta a operar

For Each File In Folder.Files
    

    sNewFile = File.Name
    If Instr(sNewFile,"gsheet") > 0 Then
        'MsgBox(sNewFile)
        sNewFile = Mid(sNewFile, 4,4) & "-" & Left(sNewFile,2) & " " & Mid(sNewFile,9,3) & ".gsheet"
        MsgBox(sNewFile)
        File.Move(Folder & "\" & sNewFile)
    Else
        'MsgBox(sNewFile)
        sNewFile = Mid(sNewFile, 4,4) & "-" & Left(sNewFile,2) & " " & Mid(sNewFile,9,3) & ".xlsx"
        MsgBox(sNewFile)
        File.Move(Folder & "\" & sNewFile)

    End If
Next