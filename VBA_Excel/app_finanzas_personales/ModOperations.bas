Attribute VB_Name = "ModOperations"
Sub LoadForm()
Tipo_de_operacion.Show 0
End Sub
Sub Borrarult()

Dim RangoDestino As Range
Dim UltFilaNum As Integer

Set RangoDestino = ThisWorkbook.Sheets("TRANS").Range("A1").CurrentRegion
    UltFilaNum = RangoDestino.Rows.Count
    
    ThisWorkbook.Sheets("TRANS").Cells(UltFilaNum, 1).EntireRow.Delete Shift:=xlUp
    
    UltFilaNum = RangoDestino.Rows.Count
    ThisWorkbook.Sheets("TRANS").Cells(UltFilaNum, 1).EntireRow.Delete Shift:=xlUp
End Sub
Sub Show_TRANS()
Inicio:
   Contraseña = InputBox("Introducir contraseña")
   If Contraseña = "123456" Then
      Sheets("TRANS").Visible = xlSheetVisible
   Else
        Respuesta = MsgBox("  Contraseña errónea" & Chr(10) & Chr(10) & "¿ Deseas volver a intentarlo ?", vbCritical + vbYesNo)
        If Respuesta = vbYes Then GoTo Inicio
   End If
End Sub

Sub Hide_Trans()
   Sheets("TRANS").Visible = xlSheetVeryHidden
End Sub


