VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Egreso 
   Caption         =   "Egreso"
   ClientHeight    =   4200
   ClientLeft      =   120
   ClientTop       =   456
   ClientWidth     =   7836
   OleObjectBlob   =   "Egreso.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Egreso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub UserForm_Initialize() 'Instrucciones al iniciar el formulario
    Dim rango, celda As Range
    
    'Carga de fecha autom�tica al TextBox
    TextBoxFecha.Value = Format(Date, "yyyy/mm/dd")
    
    'Seteo el rango dinamico con los nombres de las cuentas
    Set rango = Worksheets("CUENTAS_2").Range("CUENTAS2")
    
    'Poblamos los comboboxes
    For Each celda In rango
        ComboBoxCuentaEgreso.AddItem celda.Value
        ComboBoxCuentaActivo.AddItem celda.Value
    Next celda
    
    Set rango = Worksheets("LISTAS").Range("MONEDA")
    For Each celda In rango
        ComboBoxMoneda.AddItem celda.Value
    Next celda
    
    Set rango = Worksheets("LISTAS").Range("CENTRO_DE_COSTO")
    For Each celda In rango
        ComboBoxCentrodecosto.AddItem celda.Value
    Next celda
 
End Sub

Private Sub TextBoxID_Change() 'Llena el TextBoxID con el correlativo
    Dim RangoDestino As Range
    Dim Nuevafila As Integer
    Set RangoDestino = ThisWorkbook.Sheets("TRANS").Range("A1").CurrentRegion
        Nuevafila = RangoDestino.Rows.Count
    
    TextBoxID.Text = ThisWorkbook.Sheets("TRANS").Cells(Nuevafila, 1).Value + 1
End Sub

Private Sub CommandButtonGuardar_Click() 'Carga de datos
    Dim RangoDestino As Range
    Dim Nuevafila As Integer
    
    'Validacion de dato tipo fecha
    If Not IsDate(TextBoxFecha) Then
        MsgBox ("Ingrese una fecha v�lida (yyyy/mm/dd)")
        TextBoxFecha.SetFocus
        Exit Sub
    End If
    
    
    'Obtiene el n�mero de la nueva fila a registrar
    Set RangoDestino = ThisWorkbook.Sheets("TRANS").Range("A1").CurrentRegion
        Nuevafila = RangoDestino.Rows.Count + 1
    'Carga de datos de la primera fila
    With ThisWorkbook.Sheets("TRANS")
        .Cells(Nuevafila, 1).Value = Me.TextBoxID.Value
        .Cells(Nuevafila, 2).Value = Me.TextBoxFecha.Value
        .Cells(Nuevafila, 3).Value = Me.TextBoxDescripcion.Value
        .Cells(Nuevafila, 4).Value = Me.TextBoxMonto.Value
        
        .Cells(Nuevafila, 6).Value = Me.TextBoxN�Documento.Value
        .Cells(Nuevafila, 7).Value = Me.ComboBoxCuentaEgreso.Value
        .Cells(Nuevafila, 8).Value = Me.ComboBoxMoneda.Value
        .Cells(Nuevafila, 9).Value = Me.ComboBoxCentrodecosto.Value
        .Cells(Nuevafila, 10).Value = Me.TextBoxContraparte.Value
        .Cells(Nuevafila, 11).Value = Me.TextBoxIDRendicion.Value
    End With
    
    'Obtiene el n�mero de la nueva fila a registrar
    Set RangoDestino = ThisWorkbook.Sheets("TRANS").Range("A1").CurrentRegion
        Nuevafila = RangoDestino.Rows.Count + 1
    'Carga de datos de la segunda fila
    With ThisWorkbook.Sheets("TRANS")
        .Cells(Nuevafila, 1).Value = Me.TextBoxID.Value
        .Cells(Nuevafila, 2).Value = Me.TextBoxFecha.Value
        .Cells(Nuevafila, 3).Value = Me.TextBoxDescripcion.Value
        .Cells(Nuevafila, 5).Value = Me.TextBoxMonto.Value
        
        .Cells(Nuevafila, 6).Value = Me.TextBoxN�Documento.Value
        .Cells(Nuevafila, 7).Value = Me.ComboBoxCuentaActivo.Value
        .Cells(Nuevafila, 8).Value = Me.ComboBoxMoneda.Value
        .Cells(Nuevafila, 9).Value = Me.ComboBoxCentrodecosto.Value
    End With

    MsgBox "Carga de datos completada"
    Unload Me
End Sub

Private Sub CommandButtonCancelar_Click()
    Unload Me
End Sub

Private Sub CommandButtonGuardarMySQL_Click() 'Cargar datos en una BBDD MySQL
    
    Dim ID As Integer
    Dim Fecha, Descripcion, N�_Documento, Categoria, Cuenta, Moneda, Centro_De_Costo, Sql As String
    Dim Monto, Gasto, Ingreso As Double
    Dim con As ADODB.connection
   
    'Iniciamos las variables
    ID = CInt(TextBoxID.Value) 'Cambiamos el tipo de dato a Integer
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

    'Seteamos la conexi�n
    Set con = New ADODB.connection
    'Abrimos la conexi�n, indicando versi�n del driver ODBC, Nombre del servidor, Nombre de la BBDD, usuario y password
    con.Open "DRIVER={MySQL ODBC 8.0 Unicode Driver};SERVER=localhost;DATABASE=DBNAME;USER=root;PASSWORD=jgjgjg;"
    'Construimos la query
    Sql = "INSERT INTO trans(ID, Fecha, Descripcion, Monto, N�_Documento, Categoria, Cuenta, Moneda, Centro_De_Costo, Gasto, Ingreso) values(" & ID & ", '" & Fecha & "', '" & Descripcion & "', " & Monto & ", '" & N�_Documento & "', '" & Categoria & "', '" & Cuenta & "', '" & Moneda & "', '" & Centro_De_Costo & "', " & Gasto & ", " & Ingreso & ") "
    'Ejecutamos la query
    con.Execute Sql, rowAffected
    
    'Verificamos que se cargaron los datos
    If rowAffected = 1 Then
        MsgBox "Datos Guardados"
    Else
        MsgBox "Carga fallida"
    End If
    
    'Cerramos la conexi�n
    con.Close
    
    Unload Me

End Sub

