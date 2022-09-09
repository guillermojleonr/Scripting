Attribute VB_Name = "reporting"
Option Explicit

Public Sub split_reports_from_pivot_table()
    '---------------------------------------------------------------------------------------
    ' Purpose   : From a pivot table and a linked slicer:
    '             Takes the selected slicer items
    '             Creates plain-text reports based on each slicer items selected
    '----------------------------------------------------------------------------------------
    Dim wb, wbt As Workbook
    Dim ws As Worksheet
    Dim sc As SlicerCache
    Dim scItem As SlicerItem
    Dim scDummy As SlicerItem
    Dim sclvl As SlicerCacheLevel
    
    Dim strArr() As Variant 'Array to store dates selected in Slicer
    Dim i As Integer 'counter
    Dim sheet_name As String 'sheet name
    
    Call create_new_workbook("G:\route\to\file\account.xlsx") 'Create new destination workbook
    
    Set wb = Workbooks("account.xlsx") 'origin workbook
    Set ws = wb.Worksheets("account") 'origin worksheet
    Set sc = wb.SlicerCaches(1) 'Slicer object
    Set wbt = Workbooks("account2.xlsx") 'destination workbook
    
    'Create Array and populate it with the slicer item names
    i = 0
    For Each sclvl In sc.SlicerCacheLevels ' SlicerCacheLevels is an object needed to loop through OLAP-related slicers (OLAP-related slicers are special)
        For Each scItem In sclvl.SlicerItems 'Looping through slicer items
            If scItem.Selected = True Then 'Checks Item Selected
                ReDim Preserve strArr(i)
                strArr(i) = scItem.Name 'Store the selected item in array
                i = i + 1
            End If
        Next scItem
    Next sclvl
    
    sc.VisibleSlicerItemsList = "[Calendario].[Date].&[2021-02-25T00:00:00]" 'Changes the slicer selection to prevent it to interfere with the following instructions
    
    'Extracting the reports
    i = 0
    For Each sclvl In sc.SlicerCacheLevels
        For Each scItem In sclvl.SlicerItems
            If scItem.Name = strArr(i) Then 'check for a slicer-item/array match
                sc.VisibleSlicerItemsList = strArr(i) 'Filter the pivot table
                ws.Copy After:=ws 'Create report copy
                Range("A:H").Copy
                Range("A:H").PasteSpecial Paste:=xlPasteValues 'Delete pivot table
                Application.CutCopyMode = False
                Call DeleteSlicers 'Delete slicer
                sheet_name = Replace(Range("B1").Value, "/", "-") 'Store sheet name
                ActiveSheet.Name = "Report_" & sheet_name 'Change sheet name
                ActiveSheet.Copy After:=wbt.Worksheets(1) 'Copy sheet to destination workbook
                'ActiveSheet.Move After:=wbt.Worksheets(1) 'Method not working, excel crashes
                If i < UBound(strArr) Then 'Prevents the counter to increase beyond the array Upper bound
                    i = i + 1
                End If
            End If
        Next scItem
    Next sclvl
    
    wb.Worksheets("diferences").Copy After:=wbt.Worksheets(wbt.Worksheets.Count)
    wbt.Worksheets(1).Delete 'Delete original sheet in destination workbook.
    
    'Delete all trash worksheets in origin workbook
    For Each ws In wb.Worksheets
        If (InStr(1, ws.Name, "Report_") > 0) Then
            ws.Delete
        End If
    Next ws
End Sub
