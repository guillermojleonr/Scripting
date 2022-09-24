Attribute VB_Name = "Import"
Option Compare Database
Option Explicit
'----------------------------------------------------------------------------------------
' Module    : Import
' Author    : Guillermo Leon
' Website   : https://savingl.
' Purpose   : Load data to the DB
'----------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------
' Procedure : importExcelSheets
' Purpose   : Imports  reports
'----------------------------------------------------------------------------------------

Public Function importExcelSheets(Directory As String, TableName As String) As Long

    Dim strDir, strFile, FileName, BBDDpath, BBDDdestination, BBDDname As String
    Dim i As Long
    Dim fso As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    'Creación de copia de seguridad antes de empezar importación
    BBDDpath = Application.CurrentProject.FullName
    BBDDname = Application.CurrentProject.Name
    BBDDdestination = "G:\Mi unidad\01_BBDD\COPIAS DE SEGURIDAD\"
    Call FileBackup(BBDDpath, BBDDdestination, BBDDname)

    'Verifica que al Path Directory se le agregue el \ al final
    If Left(Directory, 1) <> "\" Then
        strDir = Directory & "\"
    Else
        strDir = Directory
    End If

    'Establece el nombre del primer archivo dentro de strDir que sea EXCEL
    strFile = Dir(strDir & "*.XLSX")

    'Loopea en el directorio mientras no se vacíe la variable strFile
    i = 0
    While strFile <> ""
        i = i + 1
        strFile = strDir & strFile
        FileName = fso.GetFileName(strFile)
        DoCmd.TransferSpreadsheet acImport, , TableName, strFile, True 'Realiza la importacion. True = has columnheaders
        fso.MoveFile strFile, strDir & "\Imported\" & FileName 'Mueve los archivos importados a una nueva carpeta "Imported"
        Debug.Print "imported " & strFile
        strFile = Dir() 'Pasa al siguiente archivo del directorio actual

    Wend
    importExcelSheets = i
End Function

'---------------------------------------------------------------------------------------
' Procedure : ImportREC
' Purpose   : Imports  in-warehouse reception report
'----------------------------------------------------------------------------------------

Public Function importREC(Directory As String, TableName As String) As Long 'Importa los registros de recepción de nuestro transportista
'On Error Resume Next

 Dim strDir As String
 Dim strFile As String
 Dim FileName As String
 Dim BBDDpath As String
 Dim BBDDdestination As String
 Dim BBDDname As String
 Dim strSQL As String
 Dim strDate As String
 Dim i, C As Long
 Dim fso As Object
 Dim qdf As DAO.QueryDef
 
 Set fso = CreateObject("Scripting.FileSystemObject")
 
 'Creación de copia de seguridad antes de empezar importación
 BBDDpath = Application.CurrentProject.FullName
 BBDDname = Application.CurrentProject.Name
 BBDDdestination = "G:\Mi unidad\01_BBDD\COPIAS DE SEGURIDAD\"
 C = fso.GetFolder(BBDDdestination).Files.Count + 1
 'fso.CopyFile BBDDpath, BBDDdestination & C & BBDDname, False 'Copia archivo con nuevo nombre

 'Verifica que al Path Directory se le agregue el \ al final
 If Left(Directory, 1) <> "\" Then
     strDir = Directory & "\"
 Else
     strDir = Directory
 End If

 'Establece el nombre del primer archivo dentro de strDir que sea EXCEL
 strFile = Dir(strDir & "*.XLSX")

 'Loopea en el directorio mientras no se vacíe la variable strFile
 i = 0
 While strFile <> ""
     i = i + 1
     strFile = strDir & strFile
     FileName = fso.GetFileName(strFile)
     strDate = TransformDate(Left(FileName, 8))
     strSQL = "UPDATE Table1 SET FECHA = " & Chr(35) & strDate & Chr(35) & " WHERE ISNULL(FECHA);"
     
     DoCmd.TransferSpreadsheet acImport, , TableName, strFile, True 'Realiza la importacion. True = has columnheaders
     'DoEvents 'Para pausar un momento la ejecución, buscando resolver un bug que hace que se importe 3 veces.
     
     Set qdf = CurrentDb.QueryDefs("Qry")
     qdf.SQL = strSQL
     DoCmd.OpenQuery ("Qry")
     Debug.Print "imported " & strFile 'Confirma que se importó
     fso.MoveFile strFile, strDir & "\Imported\" & FileName 'Mueve los archivos importados a una nueva carpeta "Imported"
     strFile = Dir() 'Pasa al siguiente archivo del directorio actual
 Wend
 importREC = i
End Function


'---------------------------------------------------------------------------------------
' Procedure : ImportExcel
' Purpose   : Import any kind of Excel File (unifies ImportREC and importExcelSheets)
'----------------------------------------------------------------------------------------

Public Function ImportExcel(Directory As String, TableName As String, Optional SQLfunctionName As String) As Long

 Dim strDir As String
 Dim strFile As String
 Dim FileName As String
 Dim BBDDpath As String
 Dim BBDDdestination As String
 Dim BBDDname As String
 Dim strSQL As String
 Dim i As Long
 Dim fso As Object
 
 Set fso = CreateObject("Scripting.FileSystemObject")
 
 'Creación de copia de seguridad antes de empezar importación
    Call FileBackup( _
        "G:\Mi unidad\01_BBDD\BBDD.accdb", _
        "G:\Mi unidad\01_BBDD\COPIAS DE SEGURIDAD\", _
        "BBDD.accdb")
        
 'Verifica que al Path Directory se le agregue el \ al final
 If Left(Directory, 1) <> "\" Then
     strDir = Directory & "\"
 Else
     strDir = Directory
 End If

 'Establece el nombre del primer archivo dentro de strDir que sea EXCEL
 strFile = Dir(strDir & "*.XLSX")

 'Loopea en el directorio mientras no se vacíe la variable strFile
 i = 0
 While strFile <> ""
     i = i + 1
     strFile = strDir & strFile
     FileName = fso.GetFileName(strFile)
     
     If SQLfunctionName <> "" Then
        strSQL = Application.Run(SQLfunctionName, FileName)
     End If
     
     DoCmd.TransferSpreadsheet acImport, , TableName, strFile, True 'Realiza la importacion. True = has columnheaders
     
     If strSQL <> "" Then
        Call ExecuteQuery("qry", strSQL)
     End If
     
     Debug.Print "imported " & strFile 'Confirma que se importó
     fso.MoveFile strFile, strDir & "\Imported\" & FileName 'Mueve los archivos importados a una nueva carpeta "Imported"
     strFile = Dir() 'Pasa al siguiente archivo del directorio actual
 Wend
 ImportExcel = i
End Function