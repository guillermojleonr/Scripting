Attribute VB_Name = "Formatting"
Option Explicit

'---------------------------------------------------------------------------------------------------
' Author    : Guillermo Leon
' Website   : https://savingl.cl
' Purpose   : Manage procedures related the apply format
'----------------------------------------------------------------------------------------------------

Public Sub quick_formatting()
    '----------------------------------------------------------------------------------------
    ' Purpose   : Formatting
    '----------------------------------------------------------------------------------------
    Columns("J:J").WrapText = True
    Columns("E:E").WrapText = True
    Columns("K:K").Font.Size = 1
End Sub

Public Sub copy_multiple_pivot_tables()
    '----------------------------------------------------------------------------------------t
    ' Purpose   : Creates multiple pivot table reports switching the filter field value
    '             and place the each result in multiple worksheets.
    '----------------------------------------------------------------------------------------
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim pivtbl As PivotTable
    Dim x As Integer
    Dim last_row As Integer
    Dim d As Dictionary
    Dim drv_range As Range
    Dim cell As Range
    Dim key As Variant 'Key must be always Variant.
    
    last_row = get_last_row_number(Range("A:A"))
    
    Set wb = Workbooks("testing_workbook.xlsm")
    Set ws = wb.Sheets("sheet_test1")
    Set drv_range = ws.Range("A2:A" & last_row)
    Set d = CreateObject("Scripting.Dictionary")
    
    'There are a couple unhandled exceptions:
    On Error Resume Next
    
    drv_range.SpecialCells(xlBlanks).EntireRow.Delete 'Delete blank rows from drv_range
   
    x = 1
    
    'Store unique name values
    For Each cell In drv_range
        d.Add cell.Value, x
        x = x + 1
    Next
    
    'Change to other worksheet
    Set ws = ActiveWorkbook.Worksheets("PivotTable")
    Set pivtbl = ws.PivotTables("TablaDinámica1")
    
    'Switch pivot table filters, copy the pivot table into a new sheet and rename it.
    For Each key In d.Keys
        pivtbl.PivotFields("NOMBRE").ClearAllFilters
        pivtbl.PivotFields("NOMBRE").CurrentPage = key
        Sheets("PivotTable").Copy After:=Sheets(3)
        ActiveSheet.Name = key
    Next
End Sub

Public Sub format_multiple_worksheets()
    '----------------------------------------------------------------------------------------
    ' Purpose   : Apply format to a file with multiple worksheets
    '----------------------------------------------------------------------------------------
    On Error Resume Next 'Unhandled exception at the last line.
    
    Dim ColName As Range
    Dim EntireRange As Range
    Dim TelefonoRange As Range
    Dim HeadersRange As Range
    
    Dim LastRow As Integer
    Dim x As Integer
    Dim y As Integer
    Dim CountFlex As Integer
    Dim CountNoFlex As Integer
    Dim LastColumn As Long
    
    Dim ws As Worksheet
    
    Dim cdtName As String
    Dim fecha As String
    Dim guias As String
    Dim puntos As String
    Dim patente As String
    Dim val As String
    
    For i = 1 To Application.Sheets.Count
        
        'Recopila datos
        fecha = Range("B13").Value
        guias = Range("B18").Value
        puntos = Range("B19").Value
        patente = Range("B21").Value
        cdtName = Range("I30").Value
        
        'Elimina foto
        ActiveSheet.Shapes.Range(Array("Picture 1")).Delete
        
        'Formato
        Range("2:26").EntireRow.Delete 'Elimina filas
        Cells.Font.Size = 8 'Cambia tamaño de letra
        ActiveSheet.PageSetup.Orientation = xlLandscape 'Orientación horizontal
        
        'Seteo de variables
        LastRow = Cells(Rows.Count, 2).End(xlUp).Row 'Última fila
        LastColumn = Cells(4, Columns.Count).End(xlToLeft).Column 'Ultima columna (numero)
        
        'Convierte los valores de columna en número y no en texto
        Columns("C:C").TextToColumns Destination:=Range("C1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
        
        Set HeadersRange = Range("A4:" & col_letter(LastColumn) & "4")
        
        'Eliminar filas no útiles
        HeadersRange.Select
        For x = LastColumn To 1 Step -1
            Set ColName = Cells(4, x)
            val = ColName.Value
            Select Case val
                Case "NOMBRE2", "NOMBRE3", "NOMBRE4"
                    ColName.EntireColumn.Delete ' Elimina columnas no necesarias
            End Select
        Next
        
        'Cambia nombres y formatea los encabezados
        For Each ColName In HeadersRange
            Select Case ColName.Value
                Case "TELEFONO"
                    Set TelefonoRange = Range(ColName.Offset(1, 0), ColName.End(xlDown))
            End Select
        Next
        
        ' Detalles y resumenes
        Cells(2, 2).Value = cdtName
        Cells(2, 4).Value = "Reporte" 'Titulo reporte
        
        'Aplica cálculos con fórmulas
        Cells(LastRow + 1, 4).Formula = Application.WorksheetFunction.Sum(Range("C5:C" & LastRow))
        Cells(LastRow + 1, 5).Formula = Application.WorksheetFunction.CountIf(TelefonoRange, "FLEX")
        Cells(LastRow + 1, 6).Formula = Application.WorksheetFunction.CountIf(TelefonoRange, "<>FLEX")
        Cells(2, 6).Value = fecha
        Cells(3, 2).Value = "test"
        Cells(3, 4).Value = "test2"
        
        'Amplia filas
        With Rows("3:3")
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlTop
            .RowHeight = 30
        End With
        
        'Formatea filas de resumen
         With Rows("1:4")
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlTop
            .Size = 10
            .Font.Bold = True
            .Font.ColorIndex = xlAutomatic
            .Shadow = False
            .Interior.Pattern = xlNone
            .Interior.TintAndShade = 0
            .Interior.PatternTintAndShade = 0
        End With
        
        'Insertar campo N°
        Columns("A:A").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromRightOrBelow
        Columns("A:A").EntireColumn.ColumnWidth = 2
        Range("A4").Value = "N°"
        
        'Introduce correlativo
        Range("A5").Value = "1"
        Range("A6").Value = "2"
        Range("A5:A6").AutoFill Destination:=Range("A5:A" & LastRow)
        
        'Coloca los bordes
        LastColumn = Cells(4, Columns.Count).End(xlToLeft).Column 'Actualización de última columna (ya hay menos columnas que antes)
        Set EntireRange = Range("A5:" & col_letter(LastColumn) & LastRow)
        With EntireRange
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        End With
        
        'Elimina palabras no deseadas de una columna de valores string
        Set DirRange = Range("E5:E" & LastRow)
        For Each cell In DirRange
            cell.Value = Replace(cell.Value, ", PALABRA", "")
            cell.Value = Replace(cell.Value, "PALABRA2 ", "")
            cell.Value = Replace(cell.Value, "PALABRA3 ", "")
        Next cell
        
        'Ajusta ancho de todas las columnas
        Cells.EntireColumn.AutoFit
        
        'Cambia nombres y formatea los encabezados
        Range("A4:T4").Select
        For Each ColName In Selection
            Select Case ColName.Value
                Case "NombreA"
                    ColName.Value = "NombreB"
                Case "NombreC"
                    ColName.Value = "NombreD"
                    ColName.EntireColumn.ColumnWidth = 15
                Case "NombreE"
                    ColName.Value = "NombreF"
                    ColName.EntireColumn.AutoFit
                Case "NombreG"
                    ColName.Value = "NombreH"
                Case "NombreI"
                    ColName.EntireColumn.ColumnWidth = 20
                    ColName.EntireColumn.WrapText = True
            End Select
        Next
        
        'Cambia el nombre a la hoja y pasa a la siguiente hoja
        ActiveSheet.Name = cdtName
        ActiveSheet.Next.Select
    Next
End Sub

'---------------------------------------------------------------------------------------
' Module    : CL
' Author    : Guillermo Leon
' Website   : https://savingl.cl
' Purpose   : Manage all procedures related to CL reporting
'---------------------------------------------------------------------------------------

Option Explicit
Dim fecha As String
Dim last_row As Integer
Dim FileName As String
Dim rng As Range
Dim EntireRange As Range

Public Sub formatting0()
    '---------------------------------------------------------------------------------------
    ' Purpose   : Set date and package type in the withdrawal CL report
    '----------------------------------------------------------------------------------------
    
    Dim OpRng As Range
    Dim NumRows As Integer

    'Quita el ajuste de texto de la primera columna
    With Columns("A:A")
        .WrapText = False
        .ShrinkToFit = False
    End With

    'Setea los headers
    Range("C1").FormulaR1C1 = "type"
    Range("B1").FormulaR1C1 = "subtype"
    Range("A1").FormulaR1C1 = "id"

    Columns("B:B").NumberFormat = "@" 'Formato fecha
    
    'Inicia la variable fecha
    fecha = InputBox("Introduce la fecha")
    NumRows = Range("A1", Range("A1").End(xlDown)).Rows.Count
    
    Range("A2:A" & NumRows).Select

    For Each OpRng In Selection
        OpRng.Select
        If InStr(1, OpRng.Value, "sender") > 0 Then
            ActiveCell.Offset(0, 1).Value = fecha
            ActiveCell.Offset(0, 1).Value = "type1"
        Else
            ActiveCell.Offset(0, 1).Value = fecha
            ActiveCell.Offset(0, 1).Value = "type2"
        End If
    Next
End Sub

Public Sub formatting1()
    '---------------------------------------------------------------------------------------
    ' Purpose   : apply format to a report
    '----------------------------------------------------------------------------------------
    last_row = get_last_row_number(Range("A:A"))
    
    Range("A:A", "H:H").Delete
    
    'Headers
    Range("A1").Value = "header1"
    Range("B1").Value = "header2"
    Range("C1").Value = "header3"
    Range("D1").Value = "header4"
    Range("E1").Value = "header5"
    Range("F1").Value = "header5"
    Range("G1").Value = "header6"
    Range("H1").FormulaR1C1 = "header7"
    Range("I1").Value = "header8"
    Range("J1").FormulaR1C1 = "header9"
    Range("K1").FormulaR1C1 = "header10"
    Range("L1").FormulaR1C1 = "header11"
    

    'Formato
    With Cells
        .Copy
        .PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        .EntireRow.AutoFit
        .EntireColumn.AutoFit
        .Font.Size = 8
    End With
    
    Application.CutCopyMode = False
    
    'Primera formula
    Range("H2").FormulaR1C1 = "=RIGHT(RC[-3],5)"
    Range("H2").AutoFill Destination:=Range("H2:H" & last_row)

    'Segunda formula
    Range("J2").FormulaR1C1 = "=RIGHT(RC[-1],5)"
    Range("J2").AutoFill Destination:=Range("J2:J" & last_row), Type:=xlFillDefault
    
    'Ajuste de tamaï¿½o de columnas
    Range("A:A, E:E, F:F, G:G, H:H, L:L, J:J").EntireColumn.AutoFit
    Range("B:B, C:C, D:D").EntireColumn.Hidden = True
    
    Columns("N:N").NumberFormat = "@" 'Cambio del tipo de dato a insertar en la columna de fecha
    
    'Seteo de fecha
    FileName = ActiveWorkbook.Name
    fecha = DateValue((ExtraeFecha(FileName)))
    Range("L2").FormulaR1C1 = fecha
    Range("L2").AutoFill Destination:=Range("L2:L" & last_row), Type:=xlFillCopy

    'Formato condicional
    Range("H:H,J:J").Select
    Selection.FormatConditions.AddUniqueValues
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    Selection.FormatConditions(1).DupeUnique = xlDuplicate
    With Selection.FormatConditions(1).Font
        .Color = -16383844
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    Columns("I:I").Select
    Selection.FormatConditions.AddUniqueValues
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    Selection.FormatConditions(1).DupeUnique = xlDuplicate
    With Selection.FormatConditions(1).Font
        .Color = -16383844
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    Range("A" & last_row + 1 & ":N1048576").Delete Shift:=xlUp 'Eliminacion de filas sobrantes
    Range("M:XFD").Delete 'Eliminacion de columnas sobrantes
End Sub


Public Sub sort_by_cell_color()
    '---------------------------------------------------------------------------------------
    ' Purpose   : Sort values on two columns based on the cell color
    '----------------------------------------------------------------------------------------
    Dim wb As Workbook
    Dim ws As Worksheet
    
    Set wb = Application.Workbooks("testing_workbook.xlsm")
    Set ws = wb.Sheets("sheet_test1")
    
    last_row = get_last_row_number(Range("A:A"))
    
        With ws.Sort
            .SortFields.Clear
            .SortFields.Add2 key:=Range("J2:J" & last_row), SortOn:=xlSortOnCellColor, Order:=xlAscending, DataOption:=xlSortNormal
            .SortFields.Add2 key:=Range("L2:L" & last_row), SortOn:=xlSortOnCellColor, Order:=xlAscending, DataOption:=xlSortNormal
            .SetRange Range("A1").CurrentRegion
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
    End If
End Sub
