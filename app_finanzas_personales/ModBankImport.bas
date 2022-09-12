Attribute VB_Name = "ModBankImport"
Option Explicit

Sub BankImport()

Dim RangoDestino, RangoOrigen, RowOp As Range
Dim wsTNuevafila, wstUltFila, wsBUltFila, ID_Value As Integer
Dim wb As Workbook
Dim wsT, wsB As Worksheet
Dim Cta As String
Dim FirstItem As Boolean
Dim TodayDate As Date

Set wb = ActiveWorkbook
Set wsT = wb.Sheets("TRANS") 'hoja destino
Set wsB = wb.Sheets("BDO") 'hoja origen
Set RangoOrigen = wsB.Range("A1").CurrentRegion
wsBUltFila = RangoOrigen.Rows.Count 'numero de ultima fila origen
Cta = InputBox("Introduce el nombre de la cuenta") 'nombre de la cuenta bancaria a operar

wsB.Range("C2").AddComment ("Ultimo registro de la cartola, importación fecha:" & Date) 'comenta el ultimo registro de la cartola
wsB.Range("A1").EntireColumn.NumberFormat = "dd/mm/yyyy;@" 'formatea correctamente la fecha en el origen

'Ordena datos por fecha
Range("A1").AutoFilter 'Coloca el filtro

With wsB.AutoFilter.Sort.SortFields
    .Clear 'Limpia cualquier filtro aplicado
    .Add2 _
    Key:=Range("A1:A" & wsBUltFila), _
    SortOn:=xlSortOnValues, _
    Order:=xlAscending 'Establece los parametros del filtro
End With

With wsB.AutoFilter.Sort
    .Header = xlYes 'Toma en cuenta el header
    .Apply 'Ejecuta el filtro
End With

wsB.Range("A2:A" & wsBUltFila).Select 'Establece el columna-rango de origen para operar como índice

For Each RowOp In Selection

    'Establece las CurrentRegion Destino para cada iteracion, necesario para recalcular cada vez la wstUltFila
    Set RangoDestino = wsT.Range("A1").CurrentRegion
    
    wstUltFila = RangoDestino.Rows.Count 'Establece el numero de la ultima fila del destino
    wsTNuevafila = wstUltFila + 1 'Establece el numero de la nueva fila a ingresar
    ID_Value = wsT.Cells(wstUltFila, 1).Value + 1 'Obtenemos el ID del nuevo registro
    wsT.Cells(wsTNuevafila, 1).Value = ID_Value 'Lo digitamos
    
    'Primera fila
    RowOp.Copy Destination:=wsT.Range("B" & wsTNuevafila) 'Pasa la fecha
    RowOp.Offset(0, 2).Copy Destination:=wsT.Range("F" & wsTNuevafila) 'Pasa el numero de operacion
    RowOp.Offset(0, 3).Copy Destination:=wsT.Range("C" & wsTNuevafila) 'Pasa la descripción
    
    'Evalua si estamos ante un cargo o un abono y lo copia a la columna correspondiente
    If RowOp.Offset(0, 4).Value <> 0 Then
        RowOp.Offset(0, 4).Copy Destination:=wsT.Range("D" & wsTNuevafila)
    Else
        RowOp.Offset(0, 5).Copy Destination:=wsT.Range("D" & wsTNuevafila)
    End If
    
    wsT.Range("H" & wsTNuevafila).Value = "CLP" 'Digita la Moneda
    wsT.Range("I" & wsTNuevafila).Value = "PERSONAL" 'Digita el centro de costo
    
    'Evalua si el registro que se ingresó es un cargo o un abono y Digita la cuenta activo en caso que sea un abono
    If RowOp.Offset(0, 4).Value = 0 Then
        wsT.Range("G" & wsTNuevafila).Value = Cta
    End If
    
    'Segunda fila
    wsTNuevafila = wsTNuevafila + 1 'Establece la nueva fila nuevamente
    
    wsT.Cells(wsTNuevafila, 1).Value = ID_Value 'Digita el ID
    RowOp.Copy Destination:=wsT.Range("B" & wsTNuevafila) 'Pasa la fecha
    RowOp.Offset(0, 2).Copy Destination:=wsT.Range("F" & wsTNuevafila) 'Pasa el numero de operacion
    RowOp.Offset(0, 3).Copy Destination:=wsT.Range("C" & wsTNuevafila) 'Pasa la descripción
    
    'Evalua si estamos ante un cargo o un abono y lo copia a la columna correspondiente
    If RowOp.Offset(0, 4).Value <> 0 Then
        RowOp.Offset(0, 4).Copy Destination:=wsT.Range("E" & wsTNuevafila)
    Else
        RowOp.Offset(0, 5).Copy Destination:=wsT.Range("E" & wsTNuevafila)
    End If
    
    wsT.Range("H" & wsTNuevafila).Value = "CLP" 'Digita la Moneda
    wsT.Range("I" & wsTNuevafila).Value = "PERSONAL" 'Digita el centro de costo
    
    'Evalua si el registro que se ingresó es un cargo o un abono y Digita la cuenta activo en caso que sea un cargo
    If RowOp.Offset(0, 4).Value <> 0 Then
        wsT.Range("G" & wsTNuevafila).Value = Cta
    End If
    
Next

End Sub
