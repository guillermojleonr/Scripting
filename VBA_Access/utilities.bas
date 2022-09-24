Attribute VB_Name = "GlobalScope"
Option Compare Database
Option Explicit

'---------------------------------------------------------------------------------------
' Author    : Guillermo Leon
' Website   : https://savingl.cl
' Purpose   : Manage Constants, Variables, Functions and Procedures to be reused
'---------------------------------------------------------------------------------------

Sub FileBackup(OriginPath, DestinationPath, FileName As String)
'---------------------------------------------------------------------------------------
' Procedure : FileBackup
' Purpose   : Create a DB backup
'----------------------------------------------------------------------------------------

    Dim i, C As Long
    Dim fso As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    'Creación de copia de seguridad antes de empezar importación
    C = fso.GetFolder(DestinationPath).Files.Count + 1
    fso.CopyFile OriginPath, DestinationPath & C & FileName, False 'Copia archivo con nuevo nombre
End Sub


Sub CreateQuery(QueryName, strSQL As String)
'---------------------------------------------------------------------------------------
' Procedure : CreateQuery
' Purpose   : Creates a query object
'----------------------------------------------------------------------------------------

     Dim qdf As DAO.QueryDef
     
     Set qdf = CurrentDb.CreateQueryDef(QueryName, strSQL)
End Sub

Sub ExecuteQuery(QueryName, strSQL As String)
'---------------------------------------------------------------------------------------
' Procedure : ExecuteQuery
' Purpose   : Executes a query by passing the query object name and the SQL string to be executed
'----------------------------------------------------------------------------------------
     Dim qdf As DAO.QueryDef
     
     Set qdf = CurrentDb.QueryDefs(QueryName)
     
     qdf.SQL = strSQL
     DoCmd.OpenQuery (QueryName)
End Sub


Function SQLbuilder1(FileName As String) As String
'---------------------------------------------------------------------------------------
' Procedure : SQLbuilder1
' Purpose   : Build an SQL query string to be used in Table1 import procedure
'----------------------------------------------------------------------------------------

    Dim strSQL As String
    Dim strDate As String
    
    strDate = TransformDate(Left(FileName, 8))
    strSQL = "UPDATE Table1 SET FECHA = " & Chr(35) & strDate & Chr(35) & " WHERE ISNULL(FECHA);"
    
    SQLbuilder1 = strSQL
End Function

Function TransformDate(FileStr As String) As String
'---------------------------------------------------------------------------------------
' Procedure : TransformDate
' Purpose   : Extract and transform date from .xlsx file name.
' Input     : i.e: 20220115FileName
' Output    : i.e 2022-01-15
'----------------------------------------------------------------------------------------

' El formato de salida es YYYY-MM-DD (formato de fecha no ambiguo). Access tiene problemas al importar formatos de fecha ambiguos como DD-MM-YYYY o MM-DD-YYYY https://stackoverflow.com/questions/34662225/importing-into-access-dates-in-dd-mm-yyy-or-mm-dd-yyy-format-from-csv-file
' FileName No deben contener guiones porque VBScript tiene problemas para reconocer guiones.
' Ventajas: el año adelante se mantiene el orden de los archivos si se almacenan en una misma carpeta

    Dim DateStr As String
    Dim DateStr2 As String

    DateStr = Mid(FileStr, 1, 8) 'Extrae del string completo los primeros 8 caracteres: YYYYMMDD
    DateStr2 = Left(DateStr, 4) & "-" & _
            Left(Mid(DateStr, 5, 2), 2) & "-" & _
            Right(DateStr, 2)
    TransformDate = DateStr2
End Function



Function TransformDateToAmerican(FileStr As String) As String
'---------------------------------------------------------------------------------------
' Procedure : TransformDateToAmerican
' Purpose   : Converts latin date format to american date format to be used in SQL queries
' Input     : i.e: 05/08/2022
' Output    : i.e: 08/05/2022
'----------------------------------------------------------------------------------------

    Dim DateStr As String
    Dim DateStr2 As String

    DateStr = Mid(FileStr, 1, 10) 'Extrae del string completo los primeros 8 caracteres: YYYYMMDD
    DateStr2 = Mid(DateStr, 4, 3) & _
                Left(DateStr, 2) & _
                Right(DateStr, 5)
    TransformDateToAmerican = DateStr2
End Function

Sub scan_line_break()
'---------------------------------------------------------------------------------------
' Purpose   : Creates multiple SELECT queries to find records with line breaks in a table.
'----------------------------------------------------------------------------------------
    
    Dim strSQL As String
    Dim fields() As Variant
    Dim field As Variant
    
    Call FileBackup("C:/OriginPath", "C:/DestinationPath", "filename")
    
    'Array definition storing the table field names
    fields() = Array("Id", "Field1", "Field2", "Field3")
    
    'SQL definition
    For Each field In fields
        strSQL = "SELECT envios.*, IIf(InStr(1," & field & ",Chr(10))<>0,True,False) AS NewField " & _
                "FROM envios " & _
                "WHERE (((IIf(InStr(1," & field & ",Chr(10))<>0,True,False))=True));"
        Call CreateQuery(field, strSQL)
    Next field
End Sub