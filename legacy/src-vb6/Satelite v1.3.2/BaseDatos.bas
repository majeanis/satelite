Attribute VB_Name = "BaseDatos"
Option Explicit

Global cnn_Consulta             As ADODB.Connection
Global Cnn_Satelite             As ADODB.Connection
Global gsGlsConexionSatelite    As String

Global grsBaseDatos             As ADODB.Recordset
Global grsUsuarioReal           As ADODB.Recordset
Global grsTipoDatos             As ADODB.Recordset

Function OpenRecordSet(CnAdo As ADODB.Connection, Rs As ADODB.Recordset, sSql As String) As Boolean
    On Error GoTo ErrOpenRecordSet
    
    GrabaLog "Ejecutando " & sSql
    Set Rs = New ADODB.Recordset
    Rs.CursorLocation = adUseClient
    Rs.Open sSql, CnAdo, adOpenForwardOnly, adLockReadOnly
    Set Rs.ActiveConnection = Nothing

    GrabaLog "Ejecutado OK"
    OpenRecordSet = True
    Exit Function
    
ErrOpenRecordSet:
    Screen.MousePointer = vbNormal
    MsgBox Error, vbCritical, App.Title
    OpenRecordSet = False
    Set Rs = Nothing
    Exit Function
End Function
Public Function ConectaBaseDatos(ByVal nNumBaseDatos As Integer) As Boolean
    Dim sNomDB  As String
    
    On Error GoTo ErrConectaBaseDatos
    
    gsFormatoFechaDB = "dd/mm/yyyy"
    
    ' Filtra la base de datos
    grsBaseDatos.Filter = ""
    grsBaseDatos.Filter = "num_basedatos=" & nNumBaseDatos
    If grsBaseDatos.EOF Then
        Screen.MousePointer = vbNormal
        MsgBox "Base de datos número " & nNumBaseDatos & " no existe.", vbInformation, App.Title
        ConectaBaseDatos = False
        Exit Function
    End If
    
    sNomDB = "" & grsBaseDatos!nom_basedatos
    If OpenDataBase(cnn_Consulta, sDecodifClave("" & grsBaseDatos!gls_coneccion, -1, gsCodigoStrCon), , sNomDB) Then
        gsFormatoFechaDB = grsBaseDatos!gls_formato_fecha
        ConectaBaseDatos = True
    Else
        ConectaBaseDatos = False
    End If
    
    Exit Function
    
ErrConectaBaseDatos:
    Screen.MousePointer = vbNormal
    MsgBox "Error en conexión : " & Err.Description, vbCritical, App.Title
    ConectaBaseDatos = False
    Exit Function
End Function
Function OpenDataBase(CnAdo As ADODB.Connection, sGlsConexion As String, Optional nTimeOut As Long = 0, Optional sNomDB As String) As Boolean
    On Error GoTo ErrOpenDataBase
    
    Set CnAdo = New ADODB.Connection
    If nTimeOut > 0 Then
        CnAdo.ConnectionTimeout = nTimeOut
    Else
        CnAdo.ConnectionTimeout = 0
    End If
    CnAdo.Open sGlsConexion
    
    CnAdo.CommandTimeout = 0
    OpenDataBase = True
    
    Exit Function
    
ErrOpenDataBase:
    Screen.MousePointer = vbNormal
    If sNomDB = "" Then
        MsgBox "Error al abrir base de datos (" & sGlsConexion & ") - " & Err.Description, vbCritical, App.Title
    Else
        MsgBox "Error al abrir base de datos """ & sNomDB & """. " & Err.Description, vbCritical, App.Title
    End If
    OpenDataBase = False
End Function

Function OpenMyDataBase() As Boolean
    
    On Error GoTo ErrOpenMyDataBase
    GrabaLog "Abriendo BD Satelite..."
    OpenMyDataBase = OpenDataBase(Cnn_Satelite, gsGlsConexionSatelite)
    If OpenMyDataBase Then
        GrabaLog "BD Satelite OK"
    End If
    Exit Function
    
ErrOpenMyDataBase:
    Screen.MousePointer = vbNormal
    MsgBox "Error al abrir base de datos Satélite : " & Err.Description, vbCritical, App.Title
    OpenMyDataBase = False
End Function

Sub CloseMyDataBase()
    On Error Resume Next
    Cnn_Satelite.Close
    Set Cnn_Satelite = Nothing
End Sub


