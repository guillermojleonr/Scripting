Private Sub Workbook_BeforeClose(Cancel As Boolean)
Application.OnKey "^+R"
End Sub

Private Sub Workbook_Open()
'Hotkeys
Application.OnKey "^+R", "LoadForm"

'Crea menu de comandos para ejecutar los procedimientos
Dim objPopUp As CommandBarPopup
Dim objBtn As CommandBarButton
   

With Application.CommandBars("Worksheet Menu Bar") 'Primera etapa para crear el menu
   On Error Resume Next
   .Controls("MyFunction").Delete
   On Error GoTo 0
   Set objPopUp = .Controls.Add( _
      Type:=msoControlPopup, _
      before:=.Controls.Count, _
      temporary:=True)
End With

objPopUp.Caption = "Procedimientos" ' Nombre del desplegable

Set objBtn = objPopUp.Controls.Add 'Boton 1
With objBtn
   .Caption = "Importar desde cartola"
   .OnAction = "BankImport"
   .Style = msoButtonCaption
End With

Set objBtn = objPopUp.Controls.Add 'Boton 2
With objBtn
   .Caption = "Formulario de registro"
   .OnAction = "LoadForm"
   .Style = msoButtonCaption
End With
   
End Sub

