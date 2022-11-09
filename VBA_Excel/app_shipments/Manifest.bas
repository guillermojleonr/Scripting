Attribute VB_Name = "Manifest"
'---------------------------------------------------------------------------------------
' Module    : Manifest
' Author    : Guillermo Leon
' Website   : https://savingl.client
' Purpose   : Manage procedures related to fix the manifests format to be able to print them properly
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
' Procedure : AjusterManifest
' Purpose   : Fix the format in the current production manifest, called by an event
'----------------------------------------------------------------------------------------

Public Sub AjusterManifest()
    Columns("J:J").WrapText = True
    Columns("E:E").WrapText = True
    Columns("K:K").Font.Size = 1
End Sub

'---------------------------------------------------------------------------------------
' Procedure : FiltradorManifiestocompany
' Purpose   : Separates the pivot table report in multiple worksheets.
' Comments  : To execute it must activate the "IMP" worksheet
'----------------------------------------------------------------------------------------

Sub FiltradorManifiestocompany()
    Dim ws As Worksheet
    Dim pivtbl As PivotTable
    Dim x, Ultimafila As Integer
    Dim d As Dictionary 'Objeto diccionario, requiere la librer�a Microsoft Scripting Runtime
    Dim DrvRange As Range
    Dim key As Variant 'la key de un diccionario siempre debe ser de tipo Variant
    
    Ultimafila = Cells(Rows.Count, 3).End(xlUp).Row
    
    Set ws = ActiveWorkbook.Worksheets("IMP") 'Hoja de trabajo
    Set DrvRange = ws.Range("S2:S" & Ultimafila) 'Rango de trabajo
    Set d = CreateObject("Scripting.Dictionary") 'Estructura de datos con la que trabajaremos
    
    On Error Resume Next 'Para que cuando no existan filas en blanco para borrar pare la ejecuci�n. Tambi�n cuando se agreguen keys duplicadas al diccionario no pare la ejecuci�n
    Selection.EntireRow.SpecialCells(xlBlanks).EntireRow.Delete 'Borra filas en blanco
    DrvRange.Select 'S�lo para verificar que rango se est� seleccionando
    x = 1
    
    'Agrega las keys (unique) al diccionario
    For Each cell In DrvRange
        d.Add cell.Value, x 'Las keys duplicadas no se agregar�n porque no son �nicas, arrojar� un error que ser� saltado por On Error Resume Next, el Item x es s�lo para rellenar ya que s�lo nos interesa trabajar con las keys
        x = x + 1
    Next
    
    'Cambiamos a la hoja de planillas
    Set ws = ActiveWorkbook.Worksheets("PLANILLAS")
    Set pivtbl = ws.PivotTables("TablaDin�mica1")
    
    'Printea las keys para ver sobre qu� se est� trabajando
    For Each key In d.Keys
        Debug.Print key, d(key)
    Next
    
    'Aplica el filtro por cada conductor, copia la hoja y la renombra
    For Each key In d.Keys
        pivtbl.PivotFields("filter_field").clientearAllFilters
        pivtbl.PivotFields("filter_field").CurrentPage = key
        Sheets("PLANILLAS").Copy After:=Sheets(3)
        ActiveSheet.Name = key
    Next
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Adjusterrouting_manifest
' Purpose   : Applies format to the routing_manifest manifest
'----------------------------------------------------------------------------------------

Public Sub Adjusterrouting_manifest()

    On Error Resume Next 'La �ltima linea da error, lo skippeamos
    Dim ColName, EntireRange, TelefonoRange, DirRange, HeadersRange As Range
    Dim LastRow, x, y, Counttype1, CountNotype1 As Integer
    Dim LastColumn As Long
    Dim wks As Worksheet
    Dim cdtName, fecha, guias, puntos, patente, val As String
    
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
        Cells.Font.Size = 8 'Cambia tama�o de letra
        ActiveSheet.PageSetup.Orientation = xlLandscape 'Orientaci�n horizontal
        
        'Seteo de variables
        LastRow = Cells(Rows.Count, 2).End(xlUp).Row '�ltima fila
        LastColumn = Cells(4, Columns.Count).End(xlToLeft).Column 'Ultima columna (numero)
        
        'Convierte los valores de la comuna Q. en n�mero y no en texto
        Columns("C:C").TextToColumns Destination:=Range("C1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
        
        Set HeadersRange = Range("A4:" & Col_Letter(LastColumn) & "4")
        
        'Eliminar filas no �tiles
        HeadersRange.Select
        For x = LastColumn To 1 Step -1
            Set ColName = Cells(4, x)
            val = ColName.Value
            Select Case val
                Case "Q2", "Q1", "clientiente", "DESTINATARIO", "filter_field", "VEHICULOS", "tipo de paquete", "Q"
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
        Cells(2, 2).Value = cdtName 'filter_field
        Cells(2, 4).Value = "PLANIFICACI�N DE RUTA" 'Titulo reporte
        Cells(LastRow + 1, 4).Formula = Application.WorksheetFunction.Sum(Range("C5:C" & LastRow)) & " BULTOS"
        Cells(LastRow + 1, 5).Formula = Application.WorksheetFunction.CountIf(TelefonoRange, "type1") & " type1"
        Cells(LastRow + 1, 6).Formula = Application.WorksheetFunction.CountIf(TelefonoRange, "<>type1") & " type2"
        Cells(2, 6).Value = fecha
        Cells(2, 1).Formula = "N� " & patente
        Cells(3, 2).Value = "EFECTIVO"
        Cells(3, 4).Value = "TRANSFERENCIA"
        
        'Amplia fila de cobros
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
        
        'Insertar campo N�
        Columns("A:A").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromRightOrBelow
        Columns("A:A").EntireColumn.ColumnWidth = 2
        Range("A4").Value = "N�"
        
        'Introduce correlativo
        Range("A5").Value = "1"
        Range("A6").Value = "2"
        Range("A5:A6").AutoFill Destination:=Range("A5:A" & LastRow)
        
        'Coloca los bordes
        LastColumn = Cells(4, Columns.Count).End(xlToLeft).Column 'Actualizaci�n de �ltima columna (ya hay menos columnas que antes)
        Set EntireRange = Range("A5:" & Col_Letter(LastColumn) & LastRow)
        With EntireRange
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        End With
        
        'Acorta direcciones
        Set DirRange = Range("E5:E" & LastRow)
        For Each cell In DirRange
            cell.Value = Replace(cell.Value, ", CHILE", "")
            cell.Value = Replace(cell.Value, "AVENIDA ", "")
            cell.Value = Replace(cell.Value, "PASAJE ", "")
        Next cell
        
        'Ajusta todas las columnas
        Cells.EntireColumn.AutoFit
        
        'Cambia nombres y formatea los encabezados
        Range("A4:T4").Select
        For Each ColName In Selection
            Select Case ColName.Value
                Case "N� Documento"
                    ColName.Value = "GUIA"
                Case "Item"
                    ColName.Value = "NOMBRE"
                    ColName.EntireColumn.ColumnWidth = 15
                Case "Cantidad"
                    ColName.Value = "Q"
                    ColName.EntireColumn.AutoFit
                Case "Direcci�n"
                    ColName.Value = "DIRECCION"
                Case "OBSERVACION"
                    ColName.EntireColumn.ColumnWidth = 20
                    ColName.EntireColumn.WrapText = True
            End Select
        Next
        
        'Cambia el nombre a la hoja y pasa a la siguiente hoja
        ActiveSheet.Name = cdtName
        ActiveSheet.Next.Select
    Next
End Sub

Public Sub test()

    Range(Range("E4").Offset(1, 0), Range("E4").End(xlDown)).Select
End Sub



