Attribute VB_Name = "testing"
Option Explicit

'---------------------------------------------------------------------------------------
' Website   : https://savingl.cl
' Purpose   : Test functions and procedures
'---------------------------------------------------------------------------------------

Public Sub count_unique_values_test()
    Dim wb As Workbook
    Dim ws As Worksheet
    
    Dim unique_values As Integer
    Dim sample_range As Range
    
    Set wb = Application.Workbooks("testing_workbook.xlsm")
    Set ws = wb.Sheets("sheet_test1")
    
    Set sample_range = Range("A2:A20")
    
    unique_values = count_unique_values(sample_range)
    
    Debug.Print unique_values
End Sub

Public Sub get_col_letter_test()
    
    Dim col_letter As String
    
    col_letter = get_col_letter(8)
    
    Debug.Print col_letter
End Sub

Public Sub extract_date_from_string_test()
    
    Dim date_extracted As String

    date_extracted = extract_date_from_string("01022022-SomeFileName.xlsx")
    
    Debug.Print (date_extracted)
    
End Sub

Public Sub create_new_workbook_test()
    create_new_workbook ("F:\Repositories\VBA-Excel\testFile.xlsx")
End Sub

Public Sub get_last_row_test()
    Dim last_row As Integer
    
    last_row = get_last_row_number(Range("A:A"))
    
    Debug.Print last_row
End Sub

Public Sub delete_blank_rows()
    Range("A:A", "H:H").Delete
End Sub
