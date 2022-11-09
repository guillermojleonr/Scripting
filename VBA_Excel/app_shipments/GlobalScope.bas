Attribute VB_Name = "GlobalScope"
'---------------------------------------------------------------------------------------
' Module    : GlobalScope
' Author    : Guillermo Leon
' Website   : https://savingl.client
' Purpose   : Manage Constants, Variables, Functions and Procedures to be used across different modules in the library
'---------------------------------------------------------------------------------------

Option Explicit
Public Const vbDDQ As String = """" 'represents 1 double quote (") in a formula construction within a  VBA string, which needs to be represented as double-double quote
Public Const vbSQ As String = "'" 'represents 1 single quote

'---------------------------------------------------------------------------------------
' Procedure : CountUnique
' Purpose   : Counts unique values within a passed range
' Input     : Range object
' Output    : Count (integer)
'----------------------------------------------------------------------------------------

Public Function CountUnique(rng As Range) As Integer
    Dim dict As Dictionary
    Dim cell As Range
    Set dict = New Dictionary 'Object Requieres reference to de Microsoft Scripting Rutime
    For Each cell In rng.Cells
         If Not dict.Exists(cell.Value) Then
            dict.Add cell.Value, 0
        End If
    Next
    CountUnique = dict.Count - 1 'Como cuenta el valor null le resto 1 para que s�lo cuente nos valores no null.
End Function

'---------------------------------------------------------------------------------------
' Procedure : Col_Letter
' Purpose   : Gets the column letter from column number
' Input     : Column number
' Output    : Column letter
'----------------------------------------------------------------------------------------
Public Function Col_Letter(lngCol As Long) As String
    Dim vArr
    vArr = Split(Cells(1, lngCol).Address(True, False), "$")
    Col_Letter = vArr(0)
End Function

'---------------------------------------------------------------------------------------
' Procedure : ExtraeFecha
' Purpose   : Extracts the date from the client intern report file name
' Output    : YYY-MM-DD
' Comment   : In order to work properly the date format configuration at operative system level
'           : must be set unambiguosly (YYYY-MM-DD)
'----------------------------------------------------------------------------------------

Function ExtraeFecha(FileStr As String) As Date
    Dim DateStr As String
    Dim DateStr2 As String
    
    DateStr = Mid(FileStr, 1, 8)
    DateStr2 = Right(DateStr, 4) & "/" & Left(Mid(DateStr, 3, 2), 2) & "/" & Left(DateStr, 2)
    
    ExtraeFecha = CDate(DateStr2)
End Function

'---------------------------------------------------------------------------------------
' Procedure : AddNew
' Purpose   : Create a new workbook to store reports
'----------------------------------------------------------------------------------------

Sub AddNew(strFileName As String)
    Dim NewBook As Workbook
    Set NewBook = Workbooks.Add
        With NewBook
            .Title = "All Sales" 'You can modify this value.
            .Subject = "Sales" 'You can modify this value.
            .SaveAs FileName:=strFileName
        End With
End Sub
'---------------------------------------------------------------------------------------
' Procedure : DeleteSlicers
' Purpose   : Delete all Slicers in ActiveSheet
' Authoring : https://www.mrexcel.com/board/threads/delete-all-slicers-in-activesheet.755970/
'----------------------------------------------------------------------------------------

Sub DeleteSlicers()
Dim shp As Shape
For Each shp In ActiveSheet.Shapes
    If shp.Type = msoSlicer Then shp.Delete
Next shp
End Sub

Sub test()
    Debug.Print Col_Letter(14)
End Sub



