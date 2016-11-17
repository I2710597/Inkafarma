Attribute VB_Name = "mdVar"
    '///Variable de cnnIONA la base de Datos.
    Public cnn As ADODB.Connection
    
    '///Variables de cnnIONES RecordSet's.
    Public prov As ADODB.Recordset
    Public medi As ADODB.Recordset
    Public emp As ADODB.Recordset
    Public cli As ADODB.Recordset
    Public fact As ADODB.Recordset
    Public bole As ADODB.Recordset
    Public detbole As ADODB.Recordset
    Public detbole1 As ADODB.Recordset
    Public detfact As ADODB.Recordset
    Public detfact1 As ADODB.Recordset
    Public guia As ADODB.Recordset
    Public detguia As ADODB.Recordset
    Public kar As ADODB.Recordset
    Public alm As ADODB.Recordset
    
    '///Variable para saber si el registro es Nuevo o Antiguo.
    Public nuevo As Boolean

    '///Variables de Busquedas.
    Public proveedor As String
    Public medicamento As String
    Public boleta As String
    Public factura As String
    Public empleado As String
    Public cliente As String
    Public guiar As String

    '///Variables de traspaso entre formularios.
    Public pasar As Boolean
    Public cantidad As String

    '///Variable de tipo de Moneda.
    Public dolar As Boolean
    Public user As String
    Public cambio As Double
