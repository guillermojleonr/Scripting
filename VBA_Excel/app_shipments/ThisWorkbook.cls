Option Explicit

Sub Workbook_Beforeclientose(Cancel As Boolean) 'Asigna la hotkey CTRL + Mayus + K a "nada", para quitarle la asignaci�n que ten�a
Application.OnKey "^+K"
Application.OnKey "^+N"


With Application.CommandBars("Worksheet Menu Bar")
      On Error Resume Next
      .Controls("&MyFunction").Delete
      On Error GoTo 0
End With

End Sub

Sub Workbook_Open()
'Asigna Hotkeys para ejecutar procedimientos con atajos de teclientado
Application.OnKey "^+K", "MPW"
Application.OnKey "^+N", "OrdenaPorColores"

'Crea menu de comandos para ejecutar los procedimientos
Dim objPopUp As CommandBarPopup
Dim objBtn As CommandBarButton
   
   'Primera etapa para crear el menu
   With Application.CommandBars("Worksheet Menu Bar")
      On Error Resume Next
      .Controls("MyFunction").Delete
      On Error GoTo 0
      Set objPopUp = .Controls.Add( _
         Type:=msoControlPopup, _
         before:=.Controls.Count, _
         temporary:=True)
   End With
   
   objPopUp.Caption = "Procedimientos" ' Nombre del desplegable
   
   'Boton MPW
   Set objBtn = objPopUp.Controls.Add
   With objBtn
      .Caption = "Ajustador registro client"
      .OnAction = "MPW"
      .Style = msoButtonCaption
   End With
   
   'Boton Ordenar por colores
   Set objBtn = objPopUp.Controls.Add
   With objBtn
      .Caption = "Ordenar por colores"
      .OnAction = "OrdenaPorColores"
      .Style = msoButtonCaption
   End With
   
   'Boton ajustador recepcion client
   Set objBtn = objPopUp.Controls.Add
   With objBtn
      .Caption = "Ajustador Recepcion client"
      .OnAction = "Adjusterrouting_manifest"
      .Style = msoButtonCaption
   End With
   
   'Boton ajustador cuenta client
   Set objBtn = objPopUp.Controls.Add
   With objBtn
      .Caption = "Cuenta client (deprecated)"
      .OnAction = "AdjusterCtaclient2"
      .Style = msoButtonCaption
   End With
   
   'Boton Ajustador manifiesto routing_manifest
   Set objBtn = objPopUp.Controls.Add
   With objBtn
      .Caption = "Manifiesto routing_manifest"
      .OnAction = "Adjusterrouting_manifest"
      .Style = msoButtonCaption
   End With
   
    'Boton Filtrador-Generador de hojas en manifiesto company
   Set objBtn = objPopUp.Controls.Add
   With objBtn
      .Caption = "Manifiesto company"
      .OnAction = "FiltradorManifiestocompany"
      .Style = msoButtonCaption
   End With
   
    'Boton Cuenta client para enviar
   Set objBtn = objPopUp.Controls.Add
   With objBtn
      .Caption = "Cuenta client"
      .OnAction = "clientAccountFormatting"
      .Style = msoButtonCaption
   End With

   'Boton test
   Set objBtn = objPopUp.Controls.Add
   With objBtn
      .Caption = "test"
      .OnAction = "test"
      .Style = msoButtonCaption
   End With
End Sub





