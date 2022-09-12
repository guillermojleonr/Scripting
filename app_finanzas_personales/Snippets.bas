Attribute VB_Name = "Snippets"
Private Sub Snippet_1()
If TextBoxMonto.Value > 0 Then
        'Si el TextBoxMonto es mayor que 0, implica que representa un ingreso
        ThisWorkbook.Sheets("TRANS").Cells(Nuevafila, 10).Value = 0
        ThisWorkbook.Sheets("TRANS").Cells(Nuevafila, 11).Value = Me.TextBoxMonto.Value
    Else
        'Si el TextBoxMonto es menor que 0, implica que representa un gasto
        ThisWorkbook.Sheets("TRANS").Cells(Nuevafila, 10).Value = Me.TextBoxMonto.Value
        ThisWorkbook.Sheets("TRANS").Cells(Nuevafila, 11).Value = 0
    End If

End Sub

    
Sub FindLastRowInExcelTable()
Dim lastRow1 As Long
Dim ws As Worksheet
Set ws = Sheets("TRANS")
'Assuming the name of the table is "TRANS"
lastRow1 = ws.ListObjects("TRANS").Range.Columns(12).Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Value

MsgBox "Last Row: " & lastRow1

End Sub

