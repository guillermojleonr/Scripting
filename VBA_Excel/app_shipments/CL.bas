Attribute VB_Name = "client"
'---------------------------------------------------------------------------------------
' Module    : client
' Author    : Guillermo Leon
' Website   : https://savingl.client
' Purpose   : Manage all procedures related to client reporting
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
' Procedure : AdjusterCtaclient2
' Purpose   : Set date in all report sheets: header and sheet name
'----------------------------------------------------------------------------------------
Option Explicit
Dim fecha As String
Dim Ultimafila As Integer
Dim FileName As String
Dim rng As Range
Dim EntireRange As Range

Sub AdjusterCtaclient2()
    On Error Resume Next
    Dim fecha As Date
    fecha = InputBox("Introduce la fecha")
    For i = 1 To Application.Sheets.Count - 1
        Range("A1").Value = Format(fecha, "Long Date")
        Application.Sheets(i).Name = Format(fecha, "dd-mm")
        ActiveSheet.Next.Select
        fecha = fecha + 1
    Next
End Sub

'---------------------------------------------------------------------------------------
' Procedure : AdjusterRECclient2
' Purpose   : Set date and package type in the withdrawal client report
'----------------------------------------------------------------------------------------

Sub AdjusterRECclient2()
    Dim OpRng As Range
    Dim NumRows As Integer

    'Quita el ajuste de texto de la primera columna
    With Columns("A:A")
        .WrapText = False
        .ShrinkToFit = False
    End With

    'Setea los headers
    Range("C1").FormulaR1C1 = "TIPO_ENVIO"
    Range("B1").FormulaR1C1 = "FECHA"
    Range("A1").FormulaR1C1 = "N_INT"

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

'---------------------------------------------------------------------------------------
' Procedure : MPW
' Purpose   : Asks for password and run functions related to client intern reports
'----------------------------------------------------------------------------------------

Sub MPW()
    Dim MyPassword
    MyPassword = InputBox("Please enter password", "Password Prompt", "********")

    If MyPassword = "3QNt4vvKL2Eyk2" Then
        UserFormSelectType.Show 0
        Exit Sub
    Else
        MsgBox "Acceso Denegado, consulta con el administrador", vbCritical, "Error"
        Exit Sub
    End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Seteo_Cuadre_1
' Purpose   : Formats 1 intern report, sets date and extra fields
'----------------------------------------------------------------------------------------
Sub Seteo_Cuadre_1()

    Ultimafila = Cells(Rows.Count, 4).End(xlUp).Row
    Range("A:A").Columns.Delete
    Range("H:H").Columns.Delete
    
    'Headers
    Range("A1").Value = "N"
    Range("B1").Value = "NO ENVIAR"
    Range("C1").Value = "ESCANEO"
    Range("D1").Value = "company"
    Range("E1").Value = "numero de seguimiento"
    Range("F1").Value = "COMUNA"
    Range("G1").Value = "VALOR"
    Range("H1").FormulaR1C1 = "client"
    Range("I1").Value = "N_INT"
    Range("J1").FormulaR1C1 = "company"
    Range("K1").FormulaR1C1 = "OBSERVACION"
    Range("L1").FormulaR1C1 = "FECHA"
    
    
    
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
    Range("H2").AutoFill Destination:=Range("H2:H" & Ultimafila)

    'Segunda formula
    Range("J2").FormulaR1C1 = "=RIGHT(RC[-1],5)"
    Range("J2").AutoFill Destination:=Range("J2:J" & Ultimafila), Type:=xlFillDefault
    
    'Ajuste de tama�o de columnas
    Range("A:A, E:E, F:F, G:G, H:H, L:L, J:J").EntireColumn.AutoFit
    Range("B:B, C:C, D:D").EntireColumn.Hidden = True
    
    Columns("N:N").NumberFormat = "@" 'Cambio del tipo de dato a insertar en la columna de fecha
    
    'Seteo de fecha
    FileName = ActiveWorkbook.Name
    fecha = DateValue((ExtraeFecha(FileName)))
    Range("L2").FormulaR1C1 = fecha
    Range("L2").AutoFill Destination:=Range("L2:L" & Ultimafila), Type:=xlFillCopy

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
    
    Range("A" & Ultimafila + 1 & ":N1048576").Delete Shift:=xlUp 'Eliminacion de filas sobrantes
    Range("M:XFD").Delete 'Eliminacion de columnas sobrantes
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Seteo_Cuadre_clientNOtype1
' Purpose   : Formats client type2 intern report, sets date and extra fields
'----------------------------------------------------------------------------------------

Sub Seteo_Cuadre_clientNOtype1()
    Range("A:A").Columns.Delete
    
    'Finding the last filled row
    Set EntireRange = Range("D:D")
    For Each rng In EntireRange
        If rng = "" Then
            Ultimafila = rng.Row - 1
            Exit For
        End If
    Next

    'Headers
    Range("A1").Value = "N"
    Range("B1").Value = "NO ENVIAR"
    Range("C1").Value = "INGRESA CODIGO BARRA"
    Range("D1").Value = "company"
    Range("E1").Value = "COMUNA"
    Range("F1").Value = "PRECIO"
    Range("G1").Value = "PAQUETES"
    Range("H1").Value = "CANTIDAD"
    Range("I1").Value = "HORA"
    Range("J1").FormulaR1C1 = "client"
    Range("K1").Value = "N_INT"
    Range("L1").FormulaR1C1 = "company"
    Range("M1").FormulaR1C1 = "OBSERVACION"
    Range("N1").FormulaR1C1 = "FECHA"
    
    'Formato
    With Cells
        .Copy
        .PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False 'Quita las formulas
        .EntireRow.AutoFit 'Ajusta filas
        .EntireColumn.AutoFit 'Ajusta columnas
        .Font.Size = 8 'Ajusta el tama�o fuente
    End With
    
    Application.CutCopyMode = False 'Sale del modo clientipboard

    'Primera formula
    Range("J2").FormulaR1C1 = "=RIGHT(RC[-6],5)"
    Range("J2").AutoFill Destination:=Range("J2:J" & Ultimafila)

    'Segunda formula
    Range("L2").FormulaR1C1 = "=RIGHT(RC[-1],5)"
    Range("L2").AutoFill Destination:=Range("L2:L" & Ultimafila), Type:=xlFillDefault
        
    'Ajuste de tama�o de columnas
    Columns("A:A").ColumnWidth = 2.29
    Range("C:C, D:D, E:E").ColumnWidth = 10
    Columns("F:F").ColumnWidth = 4.71
    Columns("G:G").ColumnWidth = 6.14
    Range("J:J,L:L").EntireColumn.AutoFit
    Range("B:B,C:C,H:H,I:I").EntireColumn.Hidden = True

    Columns("P:P").NumberFormat = "@" 'Cambio del tipo de dato a insertar en la columna de fecha
    
    'Seteo de fecha
    FileName = ActiveWorkbook.Name
    fecha = ExtraeFecha(FileName)
    Range("N2").FormulaR1C1 = fecha
    Range("N2").AutoFill Destination:=Range("N2:N" & Ultimafila), Type:=xlFillCopy
    
    'Formato condicional
    Set rng = Range("L:L,J:J")
    With rng
        .FormatConditions.AddUniqueValues
        .FormatConditions(rng.FormatConditions.Count).SetFirstPriority
        .FormatConditions(1).DupeUnique = xlDuplicate
    End With
    
    With rng.FormatConditions(1).Font
        .Color = -16383844
        .TintAndShade = 0
    End With
    
    With rng.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    
    rng.FormatConditions(1).StopIfTrue = False
    
    Set rng = Range("K:K")
    
    With rng
        .FormatConditions.AddUniqueValues
        .FormatConditions(rng.FormatConditions.Count).SetFirstPriority
        .FormatConditions(1).DupeUnique = xlDuplicate
    End With
    
    With rng.FormatConditions(1).Font
        .Color = -16383844
        .TintAndShade = 0
    End With
    
    With rng.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    
    rng.FormatConditions(1).StopIfTrue = False
    
    'Eliminacion de filas sobrantes
    Range("A" & Ultimafila + 1 & ":N1048576").Delete Shift:=xlUp
    
    'Eliminacion de columnas sobrantes
    Range("O:XFD").Delete
End Sub

'---------------------------------------------------------------------------------------
' Procedure : OrdenaPorColores
' Purpose   : Sort to the top of the list the mismatching records in the client intern report
'----------------------------------------------------------------------------------------

Sub OrdenaPorColores()
    Ultimafila = Cells(Rows.Count, 1).End(xlUp).Row
    If Range("H1").Value = "cantidad" Then
        With ActiveWorkbook.Worksheets("registro").Sort
            .SortFields.clientear
            .SortFields.Add2 key:=Range("J2:J" & Ultimafila), SortOn:=xlSortOnCellColor, Order:=xlAscending, DataOption:=xlSortNormal
            .SortFields.Add2 key:=Range("L2:L" & Ultimafila), SortOn:=xlSortOnCellColor, Order:=xlAscending, DataOption:=xlSortNormal
            .SetRange Range("A1").CurrentRegion
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
    Else
        With ActiveWorkbook.Worksheets("envio type1").Sort
            .SortFields.clientear
            .SortFields.Add2 key:=Range("H2:H" & Ultimafila), SortOn:=xlSortOnCellColor, Order:=xlAscending, DataOption:=xlSortNormal
            .SortFields.Add2 key:=Range("J2:J" & Ultimafila), SortOn:=xlSortOnCellColor, Order:=xlAscending, DataOption:=xlSortNormal
            .SetRange Range("A1").CurrentRegion
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
    End If
End Sub



'---------------------------------------------------------------------------------------
' Procedure : clientAccountFormatting
' Purpose   : Remove pivot tables and data connection and set formatting to send the report
'----------------------------------------------------------------------------------------

Sub clientAccountFormatting()
    Dim wb, wbt As Workbook
    Dim ws As Worksheet
    Dim sc As SlicerCache
    Dim scItem, scDummy As SlicerItem
    Dim sclientvl As SlicerCacheLevel
    Dim strArr() As Variant 'Array to store dates selected in Slicer
    Dim i As Integer 'counter
    Dim strName As String 'sheet name
    
    Call AddNew("G:\Mi unidad\04_CUENTAS\client\CUENTAclient.xlsx") 'Create new destination workbook
    
    Set wb = Workbooks("CUENTA.xlsx") 'wb: origin workbook
    Set ws = wb.Worksheets("CUENTA") 'origin worksheet
    Set sc = wb.SlicerCaches(1) 'Slicer object
    Set wbt = Workbooks("CUENTAclient.xlsx") 'wbt: destination workbook
    
    'Create Array and populate it with the slicer item names
    i = 0
    For Each sclientvl In sc.SlicerCacheLevels 'Needed to loop through OLAP related slicers
        For Each scItem In sclientvl.SlicerItems 'Looping through slicer items
            If scItem.Selected = True Then 'Checks Item Selected Property
                ReDim Preserve strArr(i)
                strArr(i) = scItem.Name
                i = i + 1
            End If
        Next scItem
    Next sclientvl
    
    sc.VisibleSlicerItemsList = "[Calendario].[Date].&[2021-02-25T00:00:00]" 'Changes the slicer selection in order to perform more instructions
    
    'Extracting the reports
    i = 0
    For Each sclientvl In sc.SlicerCacheLevels
        For Each scItem In sclientvl.SlicerItems
            If scItem.Name = strArr(i) Then 'check for a slicer-item/array match
                sc.VisibleSlicerItemsList = strArr(i) 'Filter the pivot table
                ws.Copy After:=ws 'Create report copy
                Range("A:H").Copy
                Range("A:H").PasteSpecial Paste:=xlPasteValues 'Delete pivot table
                Application.CutCopyMode = False
                Call DeleteSlicers 'Delete slicer
                strName = Replace(Range("B1").Value, "/", "-") 'Store sheet name
                ActiveSheet.Name = "Report_" & strName 'Change sheet name
                ActiveSheet.Copy After:=wbt.Worksheets(1) 'Copy sheet to destination workbook
                'ActiveSheet.Move After:=wbt.Worksheets(1) 'Method not working, excel crashes
                If i < UBound(strArr) Then 'Prevents the counter to increase further the array Upper bound
                    i = i + 1
                End If
            End If
        Next scItem
    Next sclientvl
    
    wb.Worksheets("DIFERENCIAS").Copy After:=wbt.Worksheets(wbt.Worksheets.Count)
    wbt.Worksheets(1).Delete 'Delete original sheet in destination workbook
    
    'Delete all trash worksheets in origin workbook
    For Each ws In wb.Worksheets
        If (InStr(1, ws.Name, "Report_") > 0) Then
            ws.Delete
        End If
    Next ws

End Sub

Sub test()
    MsgBox ExtraeFecha("12072022registrocompany.xlsx")
End Sub
