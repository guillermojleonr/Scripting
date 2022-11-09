VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormSelectType 
   Caption         =   "Tipo de registro"
   clientientHeight    =   2028
   clientientLeft      =   105
   clientientTop       =   450
   clientientWidth     =   2505
   OleObjectBlob   =   "UserFormSelectType.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormSelectType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclientaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButtonContinuar_clientick()
    If Button_type1.Value = True Then
        Call Seteo_Cuadre_1
        Unload Me
    End If
    
    If Button_Notype1.Value = True Then
        Call Seteo_Cuadre_clientNOtype1
        Unload Me
    End If
    
End Sub

Private Sub CommandButtonSalir_clientick()
  Unload Me
End Sub

