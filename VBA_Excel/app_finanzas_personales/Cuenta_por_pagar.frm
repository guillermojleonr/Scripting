VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Cuenta_por_pagar 
   Caption         =   "Cuenta por pagar"
   ClientHeight    =   4104
   ClientLeft      =   120
   ClientTop       =   456
   ClientWidth     =   7836
   OleObjectBlob   =   "Cuenta_por_pagar.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Cuenta_por_pagar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
'EVENTO INITIALIZE: carga datos al momento de inicializar el formulario
'Estos datos se cargan pero se pueden editar posteriormente, para que sean fijos debe usarse el evento
'Change y/o usar la propiedad Value del TextBox para precargar un dato fijo

    'Declaro variables
    Dim rango, celda As Range
    Dim lastIDRendicion As Long
    Dim ws As Worksheet
    
    'Carga ultima ID Rendicion
    Set ws = Sheets("TRANS")
    lastIDRendicion = ws.ListObjects("TRANS").Range.Columns(11).Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Value
    TextBoxIDRendicion.Value = lastIDRendicion
    
    'Carga Fecha
    TextBoxFecha.Value = Format(Date, "yyyy/mm/dd")
    
    'Carga ComboBoxes usando rangos dinamicos
    'Cuenta Debe
    Set rango = Worksheets("CUENTAS_2").Range("CUENTAS2")
    For Each celda In rango
        ComboBoxCuentaDebe.AddItem celda.Value
    Next celda
    'Cuenta Haber
    Set rango = Worksheets("CUENTAS_2").Range("CUENTAS2")
    For Each celda In rango
        ComboBoxCuentaHaber.AddItem celda.Value
    Next celda
    'Moneda
    Set rango = Worksheets("LISTAS").Range("MONEDA")
    For Each celda In rango
        ComboBoxMoneda.AddItem celda.Value
    Next celda
    'Centro de costo
    Set rango = Worksheets("LISTAS").Range("CENTRO_DE_COSTO")
    For Each celda In rango
        ComboBoxCentrodecosto.AddItem celda.Value
    Next celda
 
End Sub

'TextBox Limitados
'ID
Private Sub TextBoxID_Change()
    Dim RangoDestino As Range
    Dim Nuevafila As Integer
    Set RangoDestino = ThisWorkbook.Sheets("TRANS").Range("A1").CurrentRegion
        Nuevafila = RangoDestino.Rows.Count
    
    TextBoxID.Text = ThisWorkbook.Sheets("TRANS").Cells(Nuevafila, 1).Value + 1
End Sub

'EVENTO CLICK: Establece las acciones a ejecutar al presionar los CommandButton

Private Sub CommandButtonGuardar_Click()
    Dim RangoDestino As Range
    Dim Nuevafila As Integer
    
    Set RangoDestino = ThisWorkbook.Sheets("TRANS").Range("A1").CurrentRegion
        'Establece el numero de la suma de la cuenta del numero de filas del Rangodestino + 1 fila mas
        Nuevafila = RangoDestino.Rows.Count + 1
    
    'VALIDACIONES DE DATOS
    'Fecha
    If Not IsDate(TextBoxFecha) Then
        MsgBox ("Ingrese una fecha v�lida (yyyy/mm/dd)")
        TextBoxFecha.SetFocus
        Exit Sub
    End If
    
    'Carga de datos de la primera fila
    With ThisWorkbook.Sheets("TRANS")
        .Cells(Nuevafila, 1).Value = Me.TextBoxID.Value
        .Cells(Nuevafila, 2).Value = Me.TextBoxFecha.Value
        .Cells(Nuevafila, 3).Value = Me.TextBoxDescripcion.Value
        .Cells(Nuevafila, 4).Value = Me.TextBoxMonto.Value
        .Cells(Nuevafila, 6).Value = Me.TextBoxN�Documento.Value
        .Cells(Nuevafila, 7).Value = Me.ComboBoxCuentaDebe.Value
        .Cells(Nuevafila, 8).Value = Me.ComboBoxMoneda.Value
        .Cells(Nuevafila, 9).Value = Me.ComboBoxCentrodecosto.Value
    End With
    
    'Carga de datos de la segunda fila
    
    Set RangoDestino = ThisWorkbook.Sheets("TRANS").Range("A1").CurrentRegion
        Nuevafila = RangoDestino.Rows.Count + 1
    
    With ThisWorkbook.Sheets("TRANS")
        .Cells(Nuevafila, 1).Value = Me.TextBoxID.Value
        .Cells(Nuevafila, 2).Value = Me.TextBoxFecha.Value
        .Cells(Nuevafila, 3).Value = Me.TextBoxDescripcion.Value
        .Cells(Nuevafila, 5).Value = Me.TextBoxMonto.Value
        .Cells(Nuevafila, 6).Value = Me.TextBoxN�Documento.Value
        .Cells(Nuevafila, 7).Value = Me.ComboBoxCuentaHaber.Value
        .Cells(Nuevafila, 8).Value = Me.ComboBoxMoneda.Value
        .Cells(Nuevafila, 9).Value = Me.ComboBoxCentrodecosto.Value
        .Cells(Nuevafila, 11).Value = Me.TextBoxIDRendicion.Value
        .Cells(Nuevafila, 10).Value = Me.TextBoxContraparte.Value
    End With
    
    
    MsgBox "Carga de datos completada"
    Unload Me
End Sub

Private Sub CommandButtonCancelar_Click()
    Unload Me

End Sub

Private Sub CommandButtonGuardarMySQL_Click()
    
    Dim ID As Integer
    Dim Fecha As String
    Dim Descripcion As String
    Dim Monto As Double
    Dim N�_Documento As String
    Dim Categoria As String
    Dim Cuenta As String
    Dim Moneda As String
    Dim Centro_De_Costo As String
    Dim Gasto As Double
    Dim Ingreso As Double
    Dim con As ADODB.connection
    
    ' CInt convierte el valor de Texto de un textbox en un valor tipo Integer, igual mente CDate, CDbl, etc
    
    ID = CInt(TextBoxID.Value)
    Fecha = TextBoxFecha.Value
    Descripcion = TextBoxDescripcion.Value
    Monto = TextBoxMonto.Value * -1
    N�_Documento = TextBoxN�Documento.Value
    Categoria = ComboBoxCategoria.Value
    Cuenta = ComboBoxCuenta.Value
    Moneda = ComboBoxMoneda.Value
    Centro_De_Costo = ComboBoxCentrodecosto.Value
    Gasto = 1
    Ingreso = 1

    
    Set con = New ADODB.connection
    con.Open "DRIVER={MySQL ODBC 8.0 Unicode Driver};SERVER=localhost;DATABASE=DBNAME;USER=root;PASSWORD=jgjgjg;"
    Sql = "INSERT INTO trans(ID, Fecha, Descripcion, Monto, N�_Documento, Categoria, Cuenta, Moneda, Centro_De_Costo, Gasto, Ingreso) values(" & ID & ", '" & Fecha & "', '" & Descripcion & "', " & Monto & ", '" & N�_Documento & "', '" & Categoria & "', '" & Cuenta & "', '" & Moneda & "', '" & Centro_De_Costo & "', " & Gasto & ", " & Ingreso & ") "
    con.Execute Sql, rowAffected
    
    If rowAffected = 1 Then
        MsgBox "Datos Guardados"
    Else
        MsgBox "Carga fallida"
    End If
    
    con.Close
    
    Unload Me

End Sub

