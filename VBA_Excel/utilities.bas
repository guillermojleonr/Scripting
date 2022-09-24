Attribute VB_Name = "utilities"
Option Explicit

'---------------------------------------------------------------------------------------
' Website   : https://savingl.cl
' Purpose   : Functions and Procedures compilation to be used across different modules in the library
'---------------------------------------------------------------------------------------

Public Function count_unique_values(rng As Range) As Integer
    '---------------------------------------------------------------------------------------
    ' Purpose   : Counts unique values in a range
    '----------------------------------------------------------------------------------------
    Dim dict As Dictionary
    Dim cell As Range
    Set dict = New Dictionary 'Dictionary object requieres reference to Microsoft Scripting Rutime library
    For Each cell In rng.Cells
         If Not dict.Exists(cell.Value) Then
            'Add range values to the key field in the dictionary, we can add only unique values. Repeated values are excluded.
            'In the dictonary value field add a 0 number just to fill it with something.
            dict.Add cell.Value, 0
        End If
    Next
    count_unique_values = dict.Count - 1
End Function

Public Function get_col_letter(lngCol As Long) As String
    '---------------------------------------------------------------------------------------
    ' Purpose   : Gets the column letter from column number
    ' Input     : Column number
    ' Output    : Column letter
    '----------------------------------------------------------------------------------------
    Dim vArr
    vArr = Split(Cells(1, lngCol).Address(True, False), "$")
    get_col_letter = vArr(0)
End Function

Public Function extract_date_from_string(str As String) As Date
    '---------------------------------------------------------------------------------------
    ' Purpose   : Extract the date substring from a string.
    ' Output    : YYYY-MM-DD unambiguous date format
    ' Comments  : - Operative system date-time configuration:
    '               The date format configuration at operative system level must be set to the unambiguo date format too (YYYY-MM-DD).
    '
    '             - Useful use case:
    '               Extract the date from file names.
    '
    '             - Unability to identify the date position inside the string:
    '               The function doesn't identify where the subtring date is inside the string, that position must be hardcoded inside the function.
    '
    '             - Only one Output date format supported:
    '               The output date format has to be hardcoded as well, this function doesn't support multiple formats.
    
    '----------------------------------------------------------------------------------------
    
    Dim date_substring As String
    Dim format_substring_date As String
    
    date_substring = Mid(str, 1, 8) 'Get the date substring
    format_substring_date = CDate(Right(date_substring, 4) & "/" & Left(Mid(date_substring, 3, 2), 2) & "/" & Left(date_substring, 2)) 'Apply format and cast the str to date format
    
    extract_date_from_string = format_substring_date
End Function

Public Sub create_new_workbook(strFilePathAndName As String)
    '---------------------------------------------------------------------------------------
    ' Purpose   : Create a new workbook
    ' Comment   : Must provide as argument the path and file name, example: C:
    '----------------------------------------------------------------------------------------
    Dim NewBook As Workbook
    Set NewBook = Workbooks.Add
        With NewBook
            .Title = "All Sales"
            .Subject = "Sales"
            .SaveAs FileName:=strFilePathAndName
        End With
End Sub

Public Sub delete_slicers()
    '---------------------------------------------------------------------------------------
    ' Purpose   : Delete all Slicers in ActiveSheet
    ' Authoring : https://www.mrexcel.com/board/threads/delete-all-slicers-in-activesheet.755970/
    '----------------------------------------------------------------------------------------
    Dim shp As Shape
    For Each shp In ActiveSheet.Shapes
        If shp.Type = msoSlicer Then shp.Delete
    Next shp
End Sub

Public Function get_last_row_number(rng As Range) As Integer
    '---------------------------------------------------------------------------------------
    ' Purpose   : Find the last row number from a given column range.
    '----------------------------------------------------------------------------------------
    Dim column_range As Range
    Dim last_row_number As Integer
    Dim cell As Range
    
    Set column_range = rng
    
    For Each cell In column_range
        If cell = "" Then
            last_row_number = cell.Row - 1
            Exit For
        End If
    Next
    
    get_last_row_number = last_row_number
    
End Function

