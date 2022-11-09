Attribute VB_Name = "Scan"
Option Compare Database
Option Explicit

'---------------------------------------------------------------------------------------
' Module    : Scan
' Author    : Guillermo Leon
' Website   : https://savingl.cl
' Purpose   : Scan the database executing queries
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
' Procedure : scan_linebreak()
' Purpose   : Create queries to find records with line breaks, each column at a time
'----------------------------------------------------------------------------------------

Sub scan_breakline()
    
    Dim strSQL As String
    Dim fields() As Variant
    Dim field As Variant
    Dim BBDDpath As String
    Dim BBDDname As String
    Dim BBDDdestination As String
    
    'Creación de copia de seguridad antes de empezar importación
    BBDDpath = Application.CurrentProject.FullName
    BBDDname = Application.CurrentProject.Name
    BBDDdestination = "G:\Mi unidad\01_BBDD\COPIAS DE SEGURIDAD\"
    Call FileBackup(BBDDpath, BBDDdestination, BBDDname)
    
    'Definimos array con los nombres de los campos
    fields() = Array("Id", "N_INT", "NOMBRE", "DIRECCION", "Q", "TELEFONO", "VALOR_ADICIONAL", _
        "TIENDA", "CONDICION_COBRO", "VALOR", "N_PAQ", "N_REC", "ESTATUS_LOGISTICO", _
        "CAMBIO_DIRECCION", "OBSERVACION", "REGISTRADOR", "HORA_DESPACHO", "PRIORIDAD", _
        "LATITUD", "LONGITUD", "FECHA_INGRESO", "FECHA_RECEPCION", "NOTAS", "COMUNA", _
        "ZONA", "VEHICULO", "TARIFA", "PRUEBA_VALOR", "TIPO_ENVIO", "CONDUCTOR", "DIRECCION_SR", _
        "N_GUIA")
    
    'Definimos el SQL y creamos el objeto query
    For Each field In fields
        strSQL = "SELECT envios.*, IIf(InStr(1," & field & ",Chr(10))<>0,True,False) AS NewField " & _
                "FROM envios " & _
                "WHERE (((IIf(InStr(1," & field & ",Chr(10))<>0,True,False))=True));"
        Call CreateQuery(field, strSQL)
    Next field
End Sub

