VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Transferencia 
   Caption         =   "Transferencia"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   456
   ClientWidth     =   7836
   OleObjectBlob   =   "Transferencia.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Transferencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Paso 1: carga de los datos de los ComboBox haciendo referencia a un rango dinamico previamente creado
'El evento Initialize del UserForm implica que se ejecutará la función cuando se inicialice el userform
'Tambien carga los datos automáticos editables de los TextBox debido a que los datos que se cargan con el evento Initialize son editables

Private Sub UserForm_Initialize()
    Dim rango, celda As Range
    
    'Carga de fecha automática al TextBox
    TextBoxFecha.Value = Format(Date, "yyyy/mm/dd")
    
    'Establezco rango = rangodinamico precreado con el nombre "CUENTA"
    Set rango = Worksheets("CUENTAS").Range("CUENTA")
    'Para cada celda (variable tipo Range) en rango Agregar el Valor de la misma al ComboBoxCuentaRemitente, luego continua con la siguiente celda
    For Each celda In rango
        ComboBoxCuentaDebe.AddItem celda.Value
    Next celda
    'Se repite el codigo para el otro rango dinamico "MONEDA"
    Set rango = Worksheets("LISTAS").Range("MONEDA")
    For Each celda In rango
        ComboBoxMoneda.AddItem celda.Value
    Next celda
    
    'Se repite el codigo para el otro rango dinamico "CUENTA"
    Set rango = Worksheets("CUENTAS").Range("CUENTA")
    For Each celda In rango
        ComboBoxCuentaHaber.AddItem celda.Value
    Next celda
    
    'Se repite el codigo para el otro rango dinamico "CENTRO_DE_COSTO"
    Set rango = Worksheets("LISTAS").Range("CENTRO_DE_COSTO")
    For Each celda In rango
        ComboBoxCentrodecosto.AddItem celda.Value
    Next celda
    
  

End Sub

'Paso 2: Vamos a establecer las limitaciones que tienen los TextBox para introducir datos, como datos obligatorios, que no se puedan modificar, etc
'A diferencia del seteo anterior del TextBoxFecha acá usamos el evento Change para evitar la modificación por el usuario del dato

Private Sub TextBoxID_Change()
    Dim RangoDestino As Range
    Dim Nuevafila As Integer
    Set RangoDestino = ThisWorkbook.Sheets("TRANS").Range("A1").CurrentRegion
        Nuevafila = RangoDestino.Rows.Count
    
    TextBoxID.Text = ThisWorkbook.Sheets("TRANS").Cells(Nuevafila, 1).Value + 1
End Sub

Private Sub TextBoxCategoria_Change()
    'El relleno automático del TextBox se hace en el panel de propiedades del mismo, sin embargo esta función con el evento Change tiene como fin evitar la edición del TextBox
    TextBoxCategoria.Text = "TRANSFERENCIA"
End Sub

'Paso 3: Definimos el codigo para los botones de comandos
'Boton de carga de datos
Private Sub CommandButton1_Click()
    Dim RangoDestino As Range
    Dim Nuevafila As Integer
        'The current region is a range bounded by any combination of blank rows and blank columns. Read-only.
    Set RangoDestino = ThisWorkbook.Sheets("TRANS").Range("A1").CurrentRegion
        'Establece el numero de la suma de la cuenta del numero de filas del Rangodestino + 1 fila mas
        Nuevafila = RangoDestino.Rows.Count + 1
        
    'Validacion de dato tipo fecha
    If Not IsDate(TextBoxFecha) Then
        MsgBox ("Ingrese una fecha válida (mm/dd/yyyy)")
        TextBoxFecha.SetFocus
        Exit Sub
    End If
    
    'Carga de datos de la primera fila
    With ThisWorkbook.Sheets("TRANS")
        .Cells(Nuevafila, 1).Value = Me.TextBoxID.Value
        .Cells(Nuevafila, 2).Value = Me.TextBoxFecha.Value
        .Cells(Nuevafila, 3).Value = Me.TextBoxDescripcion.Value
        .Cells(Nuevafila, 4).Value = Me.TextBoxMonto.Value

        .Cells(Nuevafila, 6).Value = Me.TextBoxN°Documento.Value
        .Cells(Nuevafila, 7).Value = Me.ComboBoxCuentaDebe.Value
        .Cells(Nuevafila, 8).Value = Me.ComboBoxMoneda.Value
        .Cells(Nuevafila, 9).Value = Me.ComboBoxCentrodecosto
    End With
    
    'Carga de datos de la segunda fila
    
    Set RangoDestino = ThisWorkbook.Sheets("TRANS").Range("A1").CurrentRegion
        Nuevafila = RangoDestino.Rows.Count + 1
    
    With ThisWorkbook.Sheets("TRANS")
        .Cells(Nuevafila, 1).Value = Me.TextBoxID.Value
        .Cells(Nuevafila, 2).Value = Me.TextBoxFecha.Value
        .Cells(Nuevafila, 3).Value = Me.TextBoxDescripcion.Value
        .Cells(Nuevafila, 5).Value = Me.TextBoxMonto.Value
        
        .Cells(Nuevafila, 6).Value = Me.TextBoxN°Documento.Value
        .Cells(Nuevafila, 7).Value = Me.ComboBoxCuentaHaber.Value
        .Cells(Nuevafila, 8).Value = Me.ComboBoxMoneda.Value
        .Cells(Nuevafila, 9).Value = Me.ComboBoxCentrodecosto
    End With
    
    MsgBox "Carga Exitosa"
    Unload Me
End Sub

Private Sub CommandButton2_Click()


    Unload Me

End Sub

