VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Tipo_de_Operacion 
   Caption         =   "Tipo_de_Operacion"
   ClientHeight    =   2145
   ClientLeft      =   120
   ClientTop       =   456
   ClientWidth     =   3672
   OleObjectBlob   =   "Tipo_de_operacion.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Tipo_de_operacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButtonContinuar_Click()

    If OptionButtonEgreso.Value = True Then
    Load Egreso
    'El numero 0 indica que el formulario es no modal (se puede seguir usando la aplicacion con el formulario abierto)
    Egreso.Show 0
    Unload Me
    
    ElseIf OptionButtonTransferencia.Value = True Then
    Load Transferencia
    Transferencia.Show 0
    Unload Me
    
    ElseIf OptionButtonIngreso.Value = True Then
    Load Ingreso
    Ingreso.Show 0
    Unload Me
    
    ElseIf OptionButtonCXC.Value = True Then
    Load Cuenta_por_cobrar
    Cuenta_por_cobrar.Show 0
    Unload Me
    
    ElseIf OptionButtonCXP.Value = True Then
    Load Cuenta_por_pagar
    Cuenta_por_pagar.Show 0
    Unload Me

Else
MsgBox "Debe seleccionar una opción"
    
End If

End Sub

Private Sub CommandButtonCancelar_Click()
Unload Me

End Sub

