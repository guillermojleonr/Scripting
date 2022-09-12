VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Ingreso 
   Caption         =   "Ingreso"
   ClientHeight    =   3804
   ClientLeft      =   120
   ClientTop       =   456
   ClientWidth     =   8352
   OleObjectBlob   =   "Ingreso.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Ingreso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Label6_Click()

End Sub

'Paso 1: carga de los datos de los ComboBox haciendo referencia a un rango dinamico previamente creado El evento Initialize del UserForm implica que se ejecutar� la funci�n cuando se inicialice el userform

Private Sub UserForm_Initialize()
    'Declaro variables
    Dim rango, celda As Range
    
    'Carga de fecha autom�tica al TextBox
    TextBoxFecha.Value = Format(Date, "yyyy/mm/dd")
    
    'Establezco rango = rangodinamico precreado con el nombre "CUENTA"
    Set rango = Worksheets("CUENTAS_2").Range("CUENTAS2")
    
    'Para cada celda (variable tipo Range) en rango Agregar el Valor de la misma al ComboBoxCategoria, luego continua con la siguiente celda
    For Each celda In rango
        ComboBoxCuentaIngreso.AddItem celda.Value
    Next celda
    'Se repite el codigo para el otro rango dinamico "CUENTA"
    Set rango = Worksheets("CUENTAS_2").Range("CUENTAS2")
    For Each celda In rango
        ComboBoxCuentaActivo.AddItem celda.Value
    Next celda
    
    'Se repite el codigo para el otro rango dinamico "MONEDA"
    Set rango = Worksheets("LISTAS").Range("MONEDA")
    For Each celda In rango
        ComboBoxMoneda.AddItem celda.Value
    Next celda
    
    'Se repite el codigo para el otro rango dinamico "CENTRO_DE_COSTO"
    Set rango = Worksheets("LISTAS").Range("CENTRO_DE_COSTO")
    For Each celda In rango
        ComboBoxCentrodecosto.AddItem celda.Value
    Next celda
 
End Sub

'Paso 2: Vamos a establecer las limitaciones que tienen los TextBox para introducir datos,
'como datos obligatorios, que no se puedan modificar, etc

Private Sub TextBoxID_Change()
    Dim RangoDestino As Range
    Dim Nuevafila As Integer
    Set RangoDestino = ThisWorkbook.Sheets("TRANS").Range("A1").CurrentRegion
        Nuevafila = RangoDestino.Rows.Count
    
    TextBoxID.Text = ThisWorkbook.Sheets("TRANS").Cells(Nuevafila, 1).Value + 1
End Sub

'Paso 3: Definimos el codigo para los botones de comandos
Private Sub CommandButtonGuardar_Click()
    Dim RangoDestino As Range
    Dim Nuevafila As Integer
        'The current region is a range bounded by any combination of blank rows and blank columns. Read-only.
    Set RangoDestino = ThisWorkbook.Sheets("TRANS").Range("A1").CurrentRegion
        'Establece el numero de la suma de la cuenta del numero de filas del Rangodestino + 1 fila mas
        Nuevafila = RangoDestino.Rows.Count + 1
    
    'Validacion de dato tipo fecha
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
        .Cells(Nuevafila, 7).Value = Me.ComboBoxCuentaActivo.Value
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
        .Cells(Nuevafila, 7).Value = Me.ComboBoxCuentaIngreso.Value
        .Cells(Nuevafila, 8).Value = Me.ComboBoxMoneda.Value
        .Cells(Nuevafila, 9).Value = Me.ComboBoxCentrodecosto.Value
        .Cells(Nuevafila, 10).Value = Me.TextBoxContraparte.Value
        .Cells(Nuevafila, 11).Value = Me.TextBoxIDRendicion.Value
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
    Monto = TextBoxMonto.Value
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

