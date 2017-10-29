Attribute VB_Name = "SateliteDB"
Option Explicit

Function db_EliminaBaseDatos(sNumBaseDatos As String) As Boolean
    Dim rsData  As ADODB.Recordset
    Dim sGlsSp  As String
    
    sGlsSp = "usp_EliminaBaseDatos " & sNumBaseDatos
    
    OpenMyDataBase
    
    db_EliminaBaseDatos = OpenRecordSet(Cnn_Satelite, rsData, sGlsSp)
    Set rsData = Nothing
    
    CloseMyDataBase
End Function

Function db_EliminaCarpetaUsuario(ByVal sNomUsuario As String, ByVal sGlsCarpeta As String) As Boolean
    Dim rsData  As ADODB.Recordset
    Dim sGlsSp  As String
    
    On Error GoTo ErrEliminaCarpetaUsuario
    
    sNomUsuario = Replace(sNomUsuario, "'", "''")
    sGlsCarpeta = Replace(sGlsCarpeta, "'", "''")
    
    sGlsSp = "usp_EliminaCarpetaUsuario "
    sGlsSp = sGlsSp & " '" & sNomUsuario & "'"
    sGlsSp = sGlsSp & ",'" & sGlsCarpeta & "'"
    
    OpenMyDataBase
    
    db_EliminaCarpetaUsuario = OpenRecordSet(Cnn_Satelite, rsData, sGlsSp)
    Set rsData = Nothing
    
    CloseMyDataBase
    
    Exit Function

ErrEliminaCarpetaUsuario:
    Screen.MousePointer = vbNormal
    MsgBox Error, vbInformation, App.Title
    db_EliminaCarpetaUsuario = False
    Exit Function
End Function


Function db_EliminaConsulta(sNumConsulta As String) As Boolean
    Dim rsData  As ADODB.Recordset
    Dim sGlsSp  As String
        
    sGlsSp = "usp_EliminaConsulta " & sNumConsulta
    
    OpenMyDataBase
    
    db_EliminaConsulta = OpenRecordSet(Cnn_Satelite, rsData, sGlsSp)
    Set rsData = Nothing
    
    CloseMyDataBase
End Function

Function db_EliminaConsultaEnCarpeta(ByVal sNomUsuario As String, ByVal sNumConsulta As String) As Boolean
    Dim rsData  As ADODB.Recordset
    Dim sGlsSp  As String
    
    On Error GoTo ErrEliminaConsultaEnCarpeta
    
    sNomUsuario = Replace(sNomUsuario, "'", "''")
    
    sGlsSp = "usp_EliminaConsultaEnCarpeta "
    sGlsSp = sGlsSp & " '" & sNomUsuario & "'"
    sGlsSp = sGlsSp & "," & sNumConsulta & ""
    
    OpenMyDataBase
    
    db_EliminaConsultaEnCarpeta = OpenRecordSet(Cnn_Satelite, rsData, sGlsSp)
    Set rsData = Nothing
    
    CloseMyDataBase
    
    Exit Function

ErrEliminaConsultaEnCarpeta:
    Screen.MousePointer = vbNormal
    MsgBox Error, vbInformation, App.Title
    db_EliminaConsultaEnCarpeta = False
    Exit Function
End Function


Function db_EliminaUsuario(ByVal sNomUsuario As String) As Boolean
    Dim rsData  As ADODB.Recordset
    Dim sGlsSp  As String
    
    sNomUsuario = Replace(sNomUsuario, "'", "''")
    
    sGlsSp = "usp_EliminaUsuario '" & sNomUsuario & "'"
    
    OpenMyDataBase
    
    db_EliminaUsuario = OpenRecordSet(Cnn_Satelite, rsData, sGlsSp)
    Set rsData = Nothing
    
    CloseMyDataBase
End Function

Function db_EliminaTabValores(sNumRegistro As String) As Boolean
    Dim rsData  As ADODB.Recordset
    Dim sGlsSp  As String
    
    sGlsSp = "usp_EliminaTabValores '" & sNumRegistro & "'"
    
    OpenMyDataBase
    
    db_EliminaTabValores = OpenRecordSet(Cnn_Satelite, rsData, sGlsSp)
    Set rsData = Nothing
    
    CloseMyDataBase
End Function

Function db_EliminaTipoUsuario(sCodTipoUsuario As String) As Boolean
    Dim rsData  As ADODB.Recordset
    Dim sGlsSp  As String
    
    sGlsSp = "usp_EliminaTipoUsuario '" & sCodTipoUsuario & "'"
    
    OpenMyDataBase
    
    db_EliminaTipoUsuario = OpenRecordSet(Cnn_Satelite, rsData, sGlsSp)
    Set rsData = Nothing
    
    CloseMyDataBase
End Function

Function db_EliminaAgrupacion(sNumPerfil As String) As Boolean
    Dim rsData  As ADODB.Recordset
    Dim sGlsSp  As String
    
    sGlsSp = "usp_EliminaPerfil " & sNumPerfil
    
    OpenMyDataBase
    
    db_EliminaAgrupacion = OpenRecordSet(Cnn_Satelite, rsData, sGlsSp)
    Set rsData = Nothing
    
    CloseMyDataBase
End Function

Function db_EliminaLote(sNumLote As String) As Boolean
'<V1.3.0>
    Dim rsData  As ADODB.Recordset
    Dim sGlsSp  As String
    
    sGlsSp = "usp_EliminaLote " & sNumLote
    
    OpenMyDataBase
    
    db_EliminaLote = OpenRecordSet(Cnn_Satelite, rsData, sGlsSp)
    Set rsData = Nothing
    
    CloseMyDataBase
'</V1.3.0>
End Function

Function db_EliminaUsuariosPorConsulta(sNumConsulta As String, sXmlConsUsuario As String) As Boolean
    Dim sGlsSp          As String
    Dim rsData          As ADODB.Recordset
    
    On Error GoTo ErrGrabaUsuariosPorConsulta
    
    ' Forma SP para grabar consulta
    sGlsSp = "usp_EliminaUsuariosPorConsulta "
    sGlsSp = sGlsSp & " " & sNumConsulta
    sGlsSp = sGlsSp & ",'" & sXmlConsUsuario & "'"
    
    OpenMyDataBase
    
    ' Graba consulta
    db_EliminaUsuariosPorConsulta = OpenRecordSet(Cnn_Satelite, rsData, sGlsSp)
    Set rsData = Nothing
        
    CloseMyDataBase
    
    Exit Function

ErrGrabaUsuariosPorConsulta:
    Screen.MousePointer = vbNormal
    MsgBox Error, vbInformation, App.Title
    db_EliminaUsuariosPorConsulta = False
    Exit Function
End Function

Function db_EliminaAgrupacionesPorConsulta(sNumConsulta As String, sXmlConsPerfil As String) As Boolean
    Dim sGlsSp          As String
    Dim rsData          As ADODB.Recordset
    
    On Error GoTo ErrEliminaPerfilesPorConsulta
    
    ' Forma SP para grabar consulta
    sGlsSp = "usp_EliminaPerfilesPorConsulta "
    sGlsSp = sGlsSp & " " & sNumConsulta
    sGlsSp = sGlsSp & ",'" & sXmlConsPerfil & "'"
    
    OpenMyDataBase
    
    ' Graba consulta
    db_EliminaAgrupacionesPorConsulta = OpenRecordSet(Cnn_Satelite, rsData, sGlsSp)
    Set rsData = Nothing
        
    CloseMyDataBase
    
    Exit Function

ErrEliminaPerfilesPorConsulta:
    Screen.MousePointer = vbNormal
    MsgBox Error, vbInformation, App.Title
    db_EliminaAgrupacionesPorConsulta = False
    Exit Function
End Function

Function db_EliminaConsultasPorUsuario(ByVal sNomUsuario As String, sXmlConsUsuario As String) As Boolean
    Dim sGlsSp          As String
    Dim rsData          As ADODB.Recordset
    
    On Error GoTo ErrEliminaConsultasPorUsuario
    
    sNomUsuario = Replace(sNomUsuario, "'", "''")
    
    ' Forma SP para grabar consulta
    sGlsSp = "usp_EliminaConsultasPorUsuario "
    sGlsSp = sGlsSp & " '" & sNomUsuario & "'"
    sGlsSp = sGlsSp & ",'" & sXmlConsUsuario & "'"
    
    OpenMyDataBase
    
    ' Graba consulta
    db_EliminaConsultasPorUsuario = OpenRecordSet(Cnn_Satelite, rsData, sGlsSp)
    Set rsData = Nothing
        
    CloseMyDataBase
    
    Exit Function

ErrEliminaConsultasPorUsuario:
    Screen.MousePointer = vbNormal
    MsgBox Error, vbInformation, App.Title
    db_EliminaConsultasPorUsuario = False
    Exit Function
End Function

Function db_EliminaAgrupacionesPorUsuario(ByVal sNomUsuario As String, sXmlPerfUsuario As String) As Boolean
    Dim sGlsSp          As String
    Dim rsData          As ADODB.Recordset
    
    On Error GoTo ErrEliminaPerfilesPorUsuario
    
    sNomUsuario = Replace(sNomUsuario, "'", "''")
    
    ' Forma SP para grabar consulta
    sGlsSp = "usp_EliminaPerfilesPorUsuario "
    sGlsSp = sGlsSp & " '" & sNomUsuario & "'"
    sGlsSp = sGlsSp & ",'" & sXmlPerfUsuario & "'"
    
    OpenMyDataBase
    
    ' Graba consulta
    db_EliminaAgrupacionesPorUsuario = OpenRecordSet(Cnn_Satelite, rsData, sGlsSp)
    Set rsData = Nothing
        
    CloseMyDataBase
    
    Exit Function

ErrEliminaPerfilesPorUsuario:
    Screen.MousePointer = vbNormal
    MsgBox Error, vbInformation, App.Title
    db_EliminaAgrupacionesPorUsuario = False
    Exit Function
End Function

Function db_EliminaLotesPorUsuario(ByVal sNomUsuario As String, sXmlLoteUsuario As String) As Boolean
    '<V1.3.0>
    Dim sGlsSp          As String
    Dim rsData          As ADODB.Recordset
    
    On Error GoTo ErrEliminaLotesPorUsuario
    
    sNomUsuario = Replace(sNomUsuario, "'", "''")
    
    ' Forma SP para eliminar lotes por usuario
    sGlsSp = "usp_EliminaLotesPorUsuario "
    sGlsSp = sGlsSp & " '" & sNomUsuario & "'"
    sGlsSp = sGlsSp & ",'" & sXmlLoteUsuario & "'"
    
    OpenMyDataBase
    
    ' Ejecuta query
    db_EliminaLotesPorUsuario = OpenRecordSet(Cnn_Satelite, rsData, sGlsSp)
    Set rsData = Nothing
        
    CloseMyDataBase
    
    Exit Function

ErrEliminaLotesPorUsuario:
    Screen.MousePointer = vbNormal
    MsgBox Error, vbInformation, App.Title
    db_EliminaLotesPorUsuario = False
    Exit Function
    '</V1.3.0>
End Function

Function db_EliminaConsultasPorAgrupacion(sNumPerfil As String, sXmlConsPerfil As String) As Boolean
    Dim sGlsSp          As String
    Dim rsData          As ADODB.Recordset
    
    On Error GoTo ErrEliminaConsultasPorPerfil
    
    ' Forma SP para grabar consulta
    sGlsSp = "usp_EliminaConsultasPorPerfil "
    sGlsSp = sGlsSp & " " & sNumPerfil
    sGlsSp = sGlsSp & ",'" & sXmlConsPerfil & "'"
    
    OpenMyDataBase
    
    ' Graba consulta
    db_EliminaConsultasPorAgrupacion = OpenRecordSet(Cnn_Satelite, rsData, sGlsSp)
    Set rsData = Nothing
        
    CloseMyDataBase
    
    Exit Function

ErrEliminaConsultasPorPerfil:
    Screen.MousePointer = vbNormal
    MsgBox Error, vbInformation, App.Title
    db_EliminaConsultasPorAgrupacion = False
    Exit Function
End Function

Function db_EliminaUsuariosPorAgrupacion(sNumPerfil As String, sXmlPerfUsuario As String) As Boolean
    Dim sGlsSp          As String
    Dim rsData          As ADODB.Recordset
    
    On Error GoTo ErrEliminaUsuariosPorPerfil
    
    ' Forma SP para grabar consulta
    sGlsSp = "usp_EliminaUsuariosPorPerfil "
    sGlsSp = sGlsSp & " " & sNumPerfil
    sGlsSp = sGlsSp & ",'" & sXmlPerfUsuario & "'"
    
    OpenMyDataBase
    
    ' Graba consulta
    db_EliminaUsuariosPorAgrupacion = OpenRecordSet(Cnn_Satelite, rsData, sGlsSp)
    Set rsData = Nothing
        
    CloseMyDataBase
    
    Exit Function

ErrEliminaUsuariosPorPerfil:
    Screen.MousePointer = vbNormal
    MsgBox Error, vbInformation, App.Title
    db_EliminaUsuariosPorAgrupacion = False
    Exit Function
End Function

Function db_EliminaUsuariosPorLote(sNumLote As String, sXmlLoteUsuario As String) As Boolean
    '<V1.3.0>
    Dim sGlsSp          As String
    Dim rsData          As ADODB.Recordset
    
    On Error GoTo ErrEliminaUsuariosPorLote
    
    ' Forma SP para eliminar usuarios por lote
    sGlsSp = "usp_EliminaUsuariosPorLote "
    sGlsSp = sGlsSp & " " & sNumLote
    sGlsSp = sGlsSp & ",'" & sXmlLoteUsuario & "'"
    
    OpenMyDataBase
    
    ' Ejecuta query
    db_EliminaUsuariosPorLote = OpenRecordSet(Cnn_Satelite, rsData, sGlsSp)
    Set rsData = Nothing
        
    CloseMyDataBase
    
    Exit Function

ErrEliminaUsuariosPorLote:
    Screen.MousePointer = vbNormal
    MsgBox Error, vbInformation, App.Title
    db_EliminaUsuariosPorLote = False
    Exit Function
    '<V1.3.0>
End Function

Function db_GrabaConsultasPorUsuarios(sXmlConsUsuario As String) As Boolean
    Dim sNomUser        As String
    Dim sGlsSp          As String
    Dim rsData          As ADODB.Recordset
    
    On Error GoTo ErrGrabaUsuariosPorConsulta
    
    sNomUser = Left("alabrin", 32)
    
    ' Forma SP para grabar consulta
    sGlsSp = "usp_GrabaConsultasPorUsuarios "
    sGlsSp = sGlsSp & " '" & sNomUser & "'"
    sGlsSp = sGlsSp & ",'" & sXmlConsUsuario & "'"
    
    OpenMyDataBase
    
    ' Graba consulta
    db_GrabaConsultasPorUsuarios = OpenRecordSet(Cnn_Satelite, rsData, sGlsSp)
    Set rsData = Nothing
        
    CloseMyDataBase
    
    Exit Function

ErrGrabaUsuariosPorConsulta:
    Screen.MousePointer = vbNormal
    MsgBox Error, vbInformation, App.Title
    db_GrabaConsultasPorUsuarios = False
    Exit Function
End Function

Function db_GrabaAgrupacionesPorUsuarios(sXmlPerfUsuario As String) As Boolean
    Dim sNomUser        As String
    Dim sGlsSp          As String
    Dim rsData          As ADODB.Recordset
    
    On Error GoTo ErrGrabaPerfilesPorUsuarios
    
    sNomUser = Left(Get_Username, 32)
    
    ' Forma SP para grabar consulta
    sGlsSp = "usp_GrabaPerfilesPorUsuarios "
    sGlsSp = sGlsSp & " '" & sNomUser & "'"
    sGlsSp = sGlsSp & ",'" & sXmlPerfUsuario & "'"
    
    OpenMyDataBase
    
    ' Graba consulta
    db_GrabaAgrupacionesPorUsuarios = OpenRecordSet(Cnn_Satelite, rsData, sGlsSp)
    Set rsData = Nothing
        
    CloseMyDataBase
    
    Exit Function

ErrGrabaPerfilesPorUsuarios:
    Screen.MousePointer = vbNormal
    MsgBox Error, vbInformation, App.Title
    db_GrabaAgrupacionesPorUsuarios = False
    Exit Function
End Function

Function db_GrabaUsuariosPorLote(sXmlLoteUsuario As String) As Boolean
    '<V1.3.0>
    Dim sNomUser        As String
    Dim sGlsSp          As String
    Dim rsData          As ADODB.Recordset
    
    On Error GoTo ErrGrabaLotesPorUsuarios
    
    sNomUser = Left(Get_Username, 32)
    
    ' Forma SP para grabar usuarios por lote
    sGlsSp = "usp_GrabaUsuariosPorLote "
    sGlsSp = sGlsSp & " '" & sNomUser & "'"
    sGlsSp = sGlsSp & ",'" & sXmlLoteUsuario & "'"
    
    OpenMyDataBase
    
    ' Ejecuta query
    db_GrabaUsuariosPorLote = OpenRecordSet(Cnn_Satelite, rsData, sGlsSp)
    Set rsData = Nothing
        
    CloseMyDataBase
    
    Exit Function

ErrGrabaLotesPorUsuarios:
    Screen.MousePointer = vbNormal
    MsgBox Error, vbInformation, App.Title
    db_GrabaUsuariosPorLote = False
    Exit Function
    '</V1.3.0>
End Function

Function db_GrabaConsultasPorAgrupacion(sXmlConsPerfil As String) As Boolean
    Dim sNomUser        As String
    Dim sGlsSp          As String
    Dim rsData          As ADODB.Recordset
    
    On Error GoTo ErrGrabaConsultasPorPerfiles
    
    sNomUser = Left(Get_Username, 32)
    
    ' Forma SP para grabar consulta
    sGlsSp = "usp_GrabaConsultasPorPerfiles "
    sGlsSp = sGlsSp & " '" & sNomUser & "'"
    sGlsSp = sGlsSp & ",'" & sXmlConsPerfil & "'"
    
    OpenMyDataBase
    
    ' Graba consulta
    db_GrabaConsultasPorAgrupacion = OpenRecordSet(Cnn_Satelite, rsData, sGlsSp)
    Set rsData = Nothing

    CloseMyDataBase
    
    Exit Function

ErrGrabaConsultasPorPerfiles:
    Screen.MousePointer = vbNormal
    MsgBox Error, vbInformation, App.Title
    db_GrabaConsultasPorAgrupacion = False
    Exit Function
End Function

Function db_GrabaLote(sNumLote As String, sNomLote As String, sNomSolicitante As String, sIndAsignarLote As String, sGlsConsultas As String) As Boolean
    '<V1.3.0>
    Dim sGlsSp          As String
    Dim sNomUserReal    As String
    Dim rsData          As ADODB.Recordset
    
    On Error GoTo ErrGrabaLote
    
    sNomUserReal = Left(Get_Username, 32)
    sNomSolicitante = LCase(Left(sNomSolicitante, 32))
    
    sGlsSp = "usp_GrabaLote "
    sGlsSp = sGlsSp & sNumLote
    sGlsSp = sGlsSp & ",'" & Replace(sNomLote, "'", "''") & "'"
    sGlsSp = sGlsSp & ",'" & sNomUserReal & "'"
    sGlsSp = sGlsSp & ",'" & sNomSolicitante & "'"
    sGlsSp = sGlsSp & ",'" & sIndAsignarLote & "'"
    sGlsSp = sGlsSp & ",'" & sGlsConsultas & "'"
    
    OpenMyDataBase
    
    ' Graba consulta
    db_GrabaLote = OpenRecordSet(Cnn_Satelite, rsData, sGlsSp)
    If db_GrabaLote Then
        sNumLote = "" & rsData!num_lote
    End If
    Set rsData = Nothing
        
    CloseMyDataBase

    Exit Function

ErrGrabaLote:
    Screen.MousePointer = vbNormal
    MsgBox Error, vbInformation, App.Title
    db_GrabaLote = False
    Exit Function
    '</V1.3.0>
End Function

Function db_GrabaTabValores(ByVal sNumRegistro As String, sCodTabla As String, ByVal sGlsValor As String) As Boolean
    Dim sGlsSp          As String
    Dim rsData          As ADODB.Recordset
    
    On Error GoTo ErrGrabaTabValores
    
    sGlsValor = Replace(sGlsValor, "'", "''")
    
    ' Forma SP para grabar Tab_Valores
    sGlsSp = "usp_GrabaTabValores "
    sGlsSp = sGlsSp & " '" & sNumRegistro & "'"
    sGlsSp = sGlsSp & ",'" & sCodTabla & "'"
    sGlsSp = sGlsSp & ",'" & sGlsValor & "'"
    
    OpenMyDataBase
    
    db_GrabaTabValores = OpenRecordSet(Cnn_Satelite, rsData, sGlsSp)
    Set rsData = Nothing
        
    CloseMyDataBase
    
    Exit Function

ErrGrabaTabValores:
    Screen.MousePointer = vbNormal
    MsgBox Error, vbInformation, App.Title
    db_GrabaTabValores = False
    Exit Function
End Function

Function db_GrabaTipoUsuario(sCodTipoUsuario As String, sXmlTipoUsuario As String, sTipoAccion As String) As Boolean
    Dim sGlsSp          As String
    Dim rsData          As ADODB.Recordset
    
    On Error GoTo ErrGrabaTipoUsuario
    
    ' Forma SP para grabar usuario
    sGlsSp = "usp_GrabaTipoUsuario "
    sGlsSp = sGlsSp & " '" & sCodTipoUsuario & "'"
    sGlsSp = sGlsSp & ",'" & sXmlTipoUsuario & "'"
    sGlsSp = sGlsSp & ",'" & sTipoAccion & "'"
    
    OpenMyDataBase
    
    ' Graba consulta
    db_GrabaTipoUsuario = OpenRecordSet(Cnn_Satelite, rsData, sGlsSp)
    Set rsData = Nothing
        
    CloseMyDataBase
    
    Exit Function

ErrGrabaTipoUsuario:
    Screen.MousePointer = vbNormal
    MsgBox Error, vbInformation, App.Title
    db_GrabaTipoUsuario = False
    Exit Function
End Function

Function db_GrabaUsuario(ByVal sNomUsuario As String, sCodTipoUsuario As String, sTipoAccion As String) As Boolean
    Dim sGlsSp          As String
    Dim rsData          As ADODB.Recordset
    
    On Error GoTo ErrGrabaUsuario
    
    sNomUsuario = Replace(sNomUsuario, "'", "''")
    
    ' Forma SP para grabar usuario
    sGlsSp = "usp_GrabaUsuario "
    sGlsSp = sGlsSp & " '" & sNomUsuario & "'"
    sGlsSp = sGlsSp & ",'" & sCodTipoUsuario & "'"
    sGlsSp = sGlsSp & ",'" & sTipoAccion & "'"
    
    OpenMyDataBase
    
    ' Graba consulta
    db_GrabaUsuario = OpenRecordSet(Cnn_Satelite, rsData, sGlsSp)
    Set rsData = Nothing
        
    CloseMyDataBase
    
    Exit Function

ErrGrabaUsuario:
    Screen.MousePointer = vbNormal
    MsgBox Error, vbInformation, App.Title
    db_GrabaUsuario = False
    Exit Function
End Function

Function db_LeeConsultaPorNombre(ByVal sNomConsulta As String, rsData As ADODB.Recordset) As Boolean
    Dim sGlsSp As String
    
    sNomConsulta = Replace(sNomConsulta, "'", "''")
    
    sGlsSp = "usp_LeeConsultaPorNombre '" & sNomConsulta & "'"
    
    db_LeeConsultaPorNombre = OpenRecordSet(Cnn_Satelite, rsData, sGlsSp)
End Function

Function db_LeeConsultasAgrupacionPorUsuario(sUsuario As String, rsData As ADODB.Recordset) As Boolean
    Dim sGlsSp As String
    
    sGlsSp = "usp_ConsultasPerfilPorUsuario '" & sUsuario & "'"
    
    db_LeeConsultasAgrupacionPorUsuario = OpenRecordSet(Cnn_Satelite, rsData, sGlsSp)
End Function

Function db_LeeConsultasEnCarpetas(ByVal sNomUsuario As String, rsData As ADODB.Recordset) As Boolean
    Dim sGlsSp As String
    
    sNomUsuario = Replace(sNomUsuario, "'", "''")
    
    sGlsSp = "usp_LeeConsultasEnCarpetas '" & sNomUsuario & "'"
    
'    OpenMyDataBase
    
    db_LeeConsultasEnCarpetas = OpenRecordSet(Cnn_Satelite, rsData, sGlsSp)

'    CloseMyDataBase
End Function


Function db_GrabaConsulta(sNumConsulta As String, ByVal sNomConsulta As String, ByVal nNumBaseDatos As Integer, ByVal sGlsQuery As String, ByVal sGlsParametros As String, ByVal sGlsFormatos As String, ByVal sGlsHorarios As String, _
                          ByVal sGlsArchivoSalida As String, ByVal sNomHojaSalida As String, ByVal nNumArea As String, ByVal nNumNegocio As String, sNomUser As String) As Boolean
    Dim sNomUserReal    As String
    Dim sGlsSp          As String
    Dim rsData          As ADODB.Recordset
    
    On Error GoTo ErrGrabaConsulta
    
    sGlsQuery = Replace(sGlsQuery, "'", "''")
    sGlsParametros = Replace(sGlsParametros, "'", "''")
    sNomUserReal = Left(Get_Username, 32)
    
    ' Forma SP para grabar consulta
    sGlsSp = "usp_GrabaConsulta "
    sGlsSp = sGlsSp & sNumConsulta
    sGlsSp = sGlsSp & ",'" & sNomConsulta & "'"
    sGlsSp = sGlsSp & "," & nNumBaseDatos
    sGlsSp = sGlsSp & ",'" & sGlsQuery & "'"
    sGlsSp = sGlsSp & ",'" & sGlsParametros & "'"
    sGlsSp = sGlsSp & ",'" & sGlsFormatos & "'"
    sGlsSp = sGlsSp & ",'" & sGlsHorarios & "'"
    '<V1.3.0>
    sGlsSp = sGlsSp & ",'" & Replace(sGlsArchivoSalida, "'", "''") & "'"
    sGlsSp = sGlsSp & ",'" & Replace(sNomHojaSalida, "'", "''") & "'"
    '</V1.3.0>
    '<V1.3.1>
    sGlsSp = sGlsSp & "," & nNumArea
    sGlsSp = sGlsSp & "," & nNumNegocio
    '</V1.3.1>
    sGlsSp = sGlsSp & ",'" & sNomUser & "'"
    sGlsSp = sGlsSp & ",'" & sNomUserReal & "'"
    
    OpenMyDataBase
    
    ' Graba consulta
    db_GrabaConsulta = OpenRecordSet(Cnn_Satelite, rsData, sGlsSp)
    If db_GrabaConsulta Then
        sNumConsulta = "" & rsData!num_consulta
    End If
    Set rsData = Nothing
        
    CloseMyDataBase
    
    Exit Function

ErrGrabaConsulta:
    Screen.MousePointer = vbNormal
    MsgBox Error, vbInformation, App.Title
    db_GrabaConsulta = False
    Exit Function
End Function

Function db_GrabaConsultaCarpeta(ByVal sNomUsuario As String, sNumConsulta As String, sNumCarpeta As String) As Boolean
    Dim sGlsSp          As String
    Dim rsData          As ADODB.Recordset
    
    On Error GoTo ErrGrabaConsultaCarpeta
    
    sNomUsuario = Replace(sNomUsuario, "'", "''")
    
    ' Forma SP para grabar la carpeta
    sGlsSp = "usp_GrabaConsultaCarpeta "
    sGlsSp = sGlsSp & " '" & sNomUsuario & "'"
    sGlsSp = sGlsSp & "," & sNumConsulta & ""
    sGlsSp = sGlsSp & "," & sNumCarpeta & ""
    
    OpenMyDataBase
    
    ' Graba consulta
    db_GrabaConsultaCarpeta = OpenRecordSet(Cnn_Satelite, rsData, sGlsSp)
    Set rsData = Nothing
        
    CloseMyDataBase
    
    Exit Function

ErrGrabaConsultaCarpeta:
    Screen.MousePointer = vbNormal
    MsgBox Error, vbInformation, App.Title
    db_GrabaConsultaCarpeta = False
    Exit Function
End Function


Function db_GrabaFormatosConsulta(sNumConsulta As String, ByVal sGlsFormatos As String) As Boolean
    Dim sGlsSp          As String
    Dim rsData          As ADODB.Recordset
    
    On Error GoTo ErrGrabaFormatosConsulta
    
    sGlsFormatos = Replace(sGlsFormatos, "'", "''")
    
    ' Forma SP para grabar consulta
    sGlsSp = "usp_GrabaFormatosConsulta "
    sGlsSp = sGlsSp & sNumConsulta
    sGlsSp = sGlsSp & ",'" & sGlsFormatos & "'"
    
    OpenMyDataBase
    
    ' Graba consulta
    db_GrabaFormatosConsulta = OpenRecordSet(Cnn_Satelite, rsData, sGlsSp)
    Set rsData = Nothing
        
    CloseMyDataBase
    
    Exit Function

ErrGrabaFormatosConsulta:
    Screen.MousePointer = vbNormal
    MsgBox Error, vbInformation, App.Title
    db_GrabaFormatosConsulta = False
    Exit Function
End Function

Function db_GrabaAgrupacion(sNumPerfil As String, ByVal sNomPerfil As String) As Boolean
    Dim sGlsSp          As String
    Dim rsData          As ADODB.Recordset
    
    On Error GoTo ErrGrabaConsulta
    
    sNomPerfil = Replace(sNomPerfil, "'", "''")
    
    ' Forma SP para grabar consulta
    sGlsSp = "usp_GrabaPerfil "
    sGlsSp = sGlsSp & sNumPerfil
    sGlsSp = sGlsSp & ",'" & sNomPerfil & "'"
    
    OpenMyDataBase
    
    ' Graba consulta
    db_GrabaAgrupacion = OpenRecordSet(Cnn_Satelite, rsData, sGlsSp)
    If db_GrabaAgrupacion Then
        sNumPerfil = "" & rsData!num_perfil
    End If
    Set rsData = Nothing
        
    CloseMyDataBase
    
    Exit Function

ErrGrabaConsulta:
    Screen.MousePointer = vbNormal
    MsgBox Error, vbInformation, App.Title
    db_GrabaAgrupacion = False
    Exit Function
End Function

Function db_GrabaBaseDatos(sNumBaseDatos As String, sNomBaseDatos As String, sGlsConeccion As String, sGlsFormatoFecha As String) As Boolean
    Dim sGlsSp          As String
    Dim rsData          As ADODB.Recordset
    
    On Error GoTo ErrGrabaBaseDatos
    
    sNomBaseDatos = Replace(sNomBaseDatos, "'", "''")
    sGlsConeccion = Replace(sGlsConeccion, "'", "''")
    sGlsFormatoFecha = Replace(sGlsFormatoFecha, "'", "''")
    
    ' Forma SP para grabar consulta
    sGlsSp = "usp_GrabaBaseDatos "
    sGlsSp = sGlsSp & sNumBaseDatos
    sGlsSp = sGlsSp & ",'" & sNomBaseDatos & "'"
    sGlsSp = sGlsSp & ",'" & sGlsConeccion & "'"
    sGlsSp = sGlsSp & ",'" & sGlsFormatoFecha & "'"
    
    OpenMyDataBase
    
    ' Graba consulta
    db_GrabaBaseDatos = OpenRecordSet(Cnn_Satelite, rsData, sGlsSp)
    If db_GrabaBaseDatos Then
        sNumBaseDatos = "" & rsData!num_basedatos
    End If
    Set rsData = Nothing
        
    CloseMyDataBase
    
    Exit Function

ErrGrabaBaseDatos:
    Screen.MousePointer = vbNormal
    MsgBox Error, vbInformation, App.Title
    db_GrabaBaseDatos = False
    Exit Function
End Function

Function db_GrabaCarpetaUsuario(sNumCarpeta As String, ByVal sNomUsuario As String, ByVal sGlsCarpeta As String) As Boolean
    Dim sGlsSp          As String
    Dim rsData          As ADODB.Recordset
    
    On Error GoTo ErrGrabaUsuario
    
    sNomUsuario = Replace(sNomUsuario, "'", "''")
    sGlsCarpeta = Replace(sGlsCarpeta, "'", "''")
    If sNumCarpeta = "" Then sNumCarpeta = "0"
    
    ' Forma SP para grabar la carpeta
    sGlsSp = "usp_GrabaCarpetaUsuario "
    sGlsSp = sGlsSp & " '" & sNumCarpeta & "'"
    sGlsSp = sGlsSp & ",'" & sNomUsuario & "'"
    sGlsSp = sGlsSp & ",'" & sGlsCarpeta & "'"
    
    OpenMyDataBase
    
    ' Graba consulta
    db_GrabaCarpetaUsuario = OpenRecordSet(Cnn_Satelite, rsData, sGlsSp)
    If db_GrabaCarpetaUsuario Then
        sNumCarpeta = "" & rsData!num_carpeta
    End If
    Set rsData = Nothing
        
    CloseMyDataBase
    
    Exit Function

ErrGrabaUsuario:
    Screen.MousePointer = vbNormal
    MsgBox Error, vbInformation, App.Title
    db_GrabaCarpetaUsuario = False
    Exit Function
End Function


Function db_GrabaLogEjecucionConsulta(sNumConsulta As String, sFecEjecucion As String, sHorEjecucion As String, sGlsTiempoUtilizado As String, sNumRegistros As String) As Boolean
    Dim sNomUser        As String
    Dim sGlsSp          As String
    Dim rsData          As ADODB.Recordset
    
    On Error GoTo ErrGrabaLogEjecucionConsulta
    
    sNomUser = Left(Get_Username, 32)
    
    ' Forma SP para grabar consulta
    sGlsSp = "usp_GrabaLogEjecucion "
    sGlsSp = sGlsSp & sNumConsulta
    sGlsSp = sGlsSp & ",'" & sNomUser & "'"
    sGlsSp = sGlsSp & ",'" & sFecEjecucion & "'"
    sGlsSp = sGlsSp & ",'" & sHorEjecucion & "'"
    sGlsSp = sGlsSp & ",'" & sGlsTiempoUtilizado & "'"
    sGlsSp = sGlsSp & "," & sNumRegistros
    
    OpenMyDataBase
    
    ' Graba consulta
    db_GrabaLogEjecucionConsulta = OpenRecordSet(Cnn_Satelite, rsData, sGlsSp)
    Set rsData = Nothing
        
    CloseMyDataBase
    
    Exit Function

ErrGrabaLogEjecucionConsulta:
    Screen.MousePointer = vbNormal
    MsgBox Error, vbInformation, App.Title
    db_GrabaLogEjecucionConsulta = False
    Exit Function
End Function

Function db_LeeTipoUsuario(sCodTipoUsuario As String, rsData As ADODB.Recordset)
    Dim sGlsSp As String
    
    sGlsSp = "usp_LeeTipoUsuario '" & sCodTipoUsuario & "'"
    
    OpenMyDataBase
    
    db_LeeTipoUsuario = OpenRecordSet(Cnn_Satelite, rsData, sGlsSp)

    CloseMyDataBase
End Function

Function db_LeeTodasConsultasPorLote(gsNumLote As String, rsData As ADODB.Recordset) As Boolean
    '<V1.3.0>
    Dim sGlsSp As String
    
    sGlsSp = "usp_LeeTodasConsultasPorLote " & IIf(gsNumLote = "", 0, gsNumLote)
    
    OpenMyDataBase
    
    db_LeeTodasConsultasPorLote = OpenRecordSet(Cnn_Satelite, rsData, sGlsSp)
    
    CloseMyDataBase
    '</V1.3.0>
End Function

Function db_LeeConsultasPorLote(gsNumLote As String, rsData As ADODB.Recordset) As Boolean
    '<V1.3.0>
    Dim sGlsSp As String
    
    sGlsSp = "usp_LeeConsultasPorLote " & IIf(gsNumLote = "", 0, gsNumLote)
    
    db_LeeConsultasPorLote = OpenRecordSet(Cnn_Satelite, rsData, sGlsSp)
    '</V1.3.0>
End Function

Function db_LeeUsuario(ByVal sNomUsuario As String, rsData As ADODB.Recordset) As Boolean
    Dim sGlsSp As String
    
    sNomUsuario = Replace(sNomUsuario, "'", "''")
    
    sGlsSp = "usp_DetalleUsuario '" & sNomUsuario & "'"
    
    db_LeeUsuario = OpenRecordSet(Cnn_Satelite, rsData, sGlsSp)
End Function

Function db_LeeConsulta(sNumConsulta As String, rsData As ADODB.Recordset) As Boolean
    Dim sGlsSp As String
    
    sGlsSp = "usp_DetalleConsulta " & sNumConsulta
    
    db_LeeConsulta = OpenRecordSet(Cnn_Satelite, rsData, sGlsSp)
End Function

Function db_LeeFormatos(sNumConsulta As String, rsData As ADODB.Recordset) As Boolean
    Dim sGlsSp As String
    
    sGlsSp = "usp_LeeFormatos " & sNumConsulta
    
    db_LeeFormatos = OpenRecordSet(Cnn_Satelite, rsData, sGlsSp)
End Function

Function db_LeeConsultas(rsData As ADODB.Recordset) As Boolean
    Dim sGlsSp As String
    
    sGlsSp = "usp_LeeConsultas"
    
    OpenMyDataBase
    
    db_LeeConsultas = OpenRecordSet(Cnn_Satelite, rsData, sGlsSp)
    
    CloseMyDataBase
End Function

Function db_LeeBasesDeDatos(rsData As ADODB.Recordset) As Boolean
    Dim sGlsSp As String
    
    sGlsSp = "usp_LeeBasesDeDatos"
    
    OpenMyDataBase
    
    db_LeeBasesDeDatos = OpenRecordSet(Cnn_Satelite, rsData, sGlsSp)
    
    CloseMyDataBase
End Function

Function db_LeeCarpetasUsuario(ByVal sNomUsuario As String, rsData As ADODB.Recordset) As Boolean
    Dim sGlsSp As String
    
    sNomUsuario = Replace(sNomUsuario, "'", "''")
    
    sGlsSp = "usp_LeeCarpetasUsuario '" & sNomUsuario & "'"
    
    OpenMyDataBase
    
    db_LeeCarpetasUsuario = OpenRecordSet(Cnn_Satelite, rsData, sGlsSp)

    CloseMyDataBase
End Function


Function db_LeeBaseDatos(sNumBaseDatos As String, rsData As ADODB.Recordset) As Boolean
    Dim sGlsSp As String
    
    sGlsSp = "usp_LeeBaseDatos " & sNumBaseDatos

    OpenMyDataBase
    
    db_LeeBaseDatos = OpenRecordSet(Cnn_Satelite, rsData, sGlsSp)
    
    CloseMyDataBase
End Function

Function db_LeeSysConfig(rsData As ADODB.Recordset) As Boolean
    Dim sGlsSp As String
    
    sGlsSp = "usp_LeeSysConfig"

    OpenMyDataBase
    
    db_LeeSysConfig = OpenRecordSet(Cnn_Satelite, rsData, sGlsSp)
    
    CloseMyDataBase
End Function

Function db_LeeParametro(sNumConsulta As String, sNomParametro As String, rsData As ADODB.Recordset) As Boolean
    Dim sGlsSp As String
    
    sGlsSp = "usp_DetalleParametro " & sNumConsulta & ",'" & sNomParametro & "'"
    
    db_LeeParametro = OpenRecordSet(Cnn_Satelite, rsData, sGlsSp)
End Function

Function db_LeeConsultasPorUsuario(sUsuario As String, rsData As ADODB.Recordset) As Boolean
    Dim sGlsSp As String
    
    sGlsSp = "usp_ConsultasPorUsuario '" & sUsuario & "'"
    
    db_LeeConsultasPorUsuario = OpenRecordSet(Cnn_Satelite, rsData, sGlsSp)
End Function

Function db_LeeAgrupacionesPorUsuario(sUsuario As String, rsData As ADODB.Recordset) As Boolean
    Dim sGlsSp As String
    
    sGlsSp = "usp_PerfilesPorUsuario '" & sUsuario & "'"
    
    db_LeeAgrupacionesPorUsuario = OpenRecordSet(Cnn_Satelite, rsData, sGlsSp)
End Function

Function db_LeeLotesPorUsuario(sUsuario As String, rsData As ADODB.Recordset) As Boolean
    '<V1.3.0>
    Dim sGlsSp As String
    
    sGlsSp = "usp_LotesPorUsuario '" & sUsuario & "'"
    
    db_LeeLotesPorUsuario = OpenRecordSet(Cnn_Satelite, rsData, sGlsSp)
    '</V1.3.0>
End Function

Function db_LeeConsultasPorAgrupacion(sNumPerfil As String, rsData As ADODB.Recordset) As Boolean
    Dim sGlsSp As String
    
    sGlsSp = "usp_ConsultasPorPerfil " & sNumPerfil
    
    db_LeeConsultasPorAgrupacion = OpenRecordSet(Cnn_Satelite, rsData, sGlsSp)
End Function

Function db_LeeUsuariosPorAgrupacion(sNumPerfil As String, rsData As ADODB.Recordset) As Boolean
    Dim sGlsSp As String
    
    sGlsSp = "usp_UsuariosPorPerfil " & sNumPerfil
    
    db_LeeUsuariosPorAgrupacion = OpenRecordSet(Cnn_Satelite, rsData, sGlsSp)
End Function

Function db_LeeUsuariosPorLote(sNumLote As String, rsData As ADODB.Recordset) As Boolean
    '<V1.3.0>
    Dim sGlsSp As String
    
    sGlsSp = "usp_UsuariosPorLote " & sNumLote
    
    db_LeeUsuariosPorLote = OpenRecordSet(Cnn_Satelite, rsData, sGlsSp)
    '</V1.3.0>
End Function

Public Function CargaBaseDatosSistema() As Boolean
    If Not OpenRecordSet(Cnn_Satelite, grsBaseDatos, "usp_LeeBasesDeDatos") Then
        CargaBaseDatosSistema = False
        Exit Function
    End If
        
    CargaBaseDatosSistema = True
    Exit Function
End Function

Function db_LeeUsuarios(rsData As ADODB.Recordset) As Boolean
    Dim sGlsSp As String
    
    sGlsSp = "usp_LeeUsuarios"
    
    OpenMyDataBase
        
    db_LeeUsuarios = OpenRecordSet(Cnn_Satelite, rsData, sGlsSp)

    CloseMyDataBase
End Function

Function db_LeeAgrupaciones(rsData As ADODB.Recordset) As Boolean
    Dim sGlsSp As String
    
    sGlsSp = "usp_LeePerfiles"
    
    OpenMyDataBase
    
    db_LeeAgrupaciones = OpenRecordSet(Cnn_Satelite, rsData, sGlsSp)
    
    CloseMyDataBase
End Function

Function db_LeeTabValores(sCodTabla As String, rsData As ADODB.Recordset) As Boolean
    '<V1.3.1>
    Dim sGlsSp As String
    
    sGlsSp = "usp_LeeTabValores '" & sCodTabla & "'"
    
    OpenMyDataBase
    
    db_LeeTabValores = OpenRecordSet(Cnn_Satelite, rsData, sGlsSp)
    
    CloseMyDataBase
    '</V1.3.1>
End Function

Function db_LeeLotes(rsData As ADODB.Recordset) As Boolean
    '<V1.3.0>
    Dim sGlsSp As String
    
    sGlsSp = "usp_LeeLotes"
    
    OpenMyDataBase
    
    db_LeeLotes = OpenRecordSet(Cnn_Satelite, rsData, sGlsSp)
    
    CloseMyDataBase
    '</V1.3.0>
End Function

Function db_LeeTipoDatos(rsData As ADODB.Recordset) As Boolean
    Dim sGlsSp As String
    
    sGlsSp = "usp_LeeTipoDatos"
    
    db_LeeTipoDatos = OpenRecordSet(Cnn_Satelite, rsData, sGlsSp)
End Function

Function db_LeeTiposUsuarios(rsData As ADODB.Recordset) As Boolean
    Dim sGlsSp As String
    
    sGlsSp = "usp_LeeTiposUsuarios"
    
    db_LeeTiposUsuarios = OpenRecordSet(Cnn_Satelite, rsData, sGlsSp)
End Function

Function db_LeeUsuariosPorConsulta(sNomConsulta As String, rsData As ADODB.Recordset) As Boolean
    Dim sGlsSp As String
    
    sGlsSp = "usp_UsuariosPorConsulta " & sNomConsulta
    
    db_LeeUsuariosPorConsulta = OpenRecordSet(Cnn_Satelite, rsData, sGlsSp)
End Function

Function db_LeeAgrupacionesPorConsulta(sNomConsulta As String, rsData As ADODB.Recordset) As Boolean
    Dim sGlsSp As String
    
    sGlsSp = "usp_PerfilesPorConsulta " & sNomConsulta
    
    db_LeeAgrupacionesPorConsulta = OpenRecordSet(Cnn_Satelite, rsData, sGlsSp)
End Function

Function db_LeeUsuariosSinConsulta(sNomConsulta As String, rsData As ADODB.Recordset) As Boolean
    Dim sGlsSp As String
    
    sGlsSp = "usp_UsuariosSinConsulta " & sNomConsulta
    
    db_LeeUsuariosSinConsulta = OpenRecordSet(Cnn_Satelite, rsData, sGlsSp)
End Function

Function db_LeeAgrupacionesSinConsulta(sNomConsulta As String, rsData As ADODB.Recordset) As Boolean
    Dim sGlsSp As String
    
    sGlsSp = "usp_PerfilesSinConsulta " & sNomConsulta
    
    db_LeeAgrupacionesSinConsulta = OpenRecordSet(Cnn_Satelite, rsData, sGlsSp)
End Function

Function db_LeeAgrupacionesSinLote(sNumLote As String, rsData As ADODB.Recordset) As Boolean
    '<V1.3.0>
    Dim sGlsSp As String
    
    sGlsSp = "usp_PerfilesSinLote " & sNumLote
    
    db_LeeAgrupacionesSinLote = OpenRecordSet(Cnn_Satelite, rsData, sGlsSp)
    '</V1.3.0>
End Function

Function db_LeeConsultasSinUsuario(ByVal sNomUsuario As String, rsData As ADODB.Recordset) As Boolean
    Dim sGlsSp As String
    
    sNomUsuario = Replace(sNomUsuario, "'", "''")
    
    sGlsSp = "usp_ConsultasSinUsuario '" & sNomUsuario & "'"
    
    db_LeeConsultasSinUsuario = OpenRecordSet(Cnn_Satelite, rsData, sGlsSp)
End Function

Function db_LeeAgrupacionesSinUsuario(ByVal sNomUsuario As String, rsData As ADODB.Recordset) As Boolean
    Dim sGlsSp As String
    
    sNomUsuario = Replace(sNomUsuario, "'", "''")
    
    sGlsSp = "usp_PerfilesSinUsuario '" & sNomUsuario & "'"
    
    db_LeeAgrupacionesSinUsuario = OpenRecordSet(Cnn_Satelite, rsData, sGlsSp)
End Function

Function db_LeeLotesSinUsuario(ByVal sNomUsuario As String, rsData As ADODB.Recordset) As Boolean
    '<V1.3.0>
    Dim sGlsSp As String
    
    sNomUsuario = Replace(sNomUsuario, "'", "''")
    
    sGlsSp = "usp_LotesSinUsuario '" & sNomUsuario & "'"
    
    db_LeeLotesSinUsuario = OpenRecordSet(Cnn_Satelite, rsData, sGlsSp)
    '</V1.3.0>
End Function

Function db_LeeConsultasSinAgrupacion(sNumPerfil As String, rsData As ADODB.Recordset) As Boolean
    Dim sGlsSp As String
    
    sGlsSp = "usp_ConsultasSinPerfil " & sNumPerfil
    
    db_LeeConsultasSinAgrupacion = OpenRecordSet(Cnn_Satelite, rsData, sGlsSp)
End Function

Function db_LeeLotesSinAgrupacion(sNumPerfil As String, rsData As ADODB.Recordset) As Boolean
    '<V1.3.0>
    Dim sGlsSp As String
    
    sGlsSp = "usp_LotesSinPerfil " & sNumPerfil
    
    db_LeeLotesSinAgrupacion = OpenRecordSet(Cnn_Satelite, rsData, sGlsSp)
    '</V1.3.0>
End Function

Function db_LeeUsuariosSinAgrupacion(sNumPerfil As String, rsData As ADODB.Recordset) As Boolean
    Dim sGlsSp As String
    
    sGlsSp = "usp_UsuariosSinPerfil " & sNumPerfil
    
    db_LeeUsuariosSinAgrupacion = OpenRecordSet(Cnn_Satelite, rsData, sGlsSp)
End Function

Function db_LeeUsuariosSinLote(sNumLote As String, rsData As ADODB.Recordset) As Boolean
    '<V1.3.0>
    Dim sGlsSp As String
    
    sGlsSp = "usp_UsuariosSinLote " & sNumLote
    
    db_LeeUsuariosSinLote = OpenRecordSet(Cnn_Satelite, rsData, sGlsSp)
    '</V1.3.0>
End Function

Function db_BloqueaConsulta(sNumConsulta As String, sIndBloqueada As String) As Boolean
    '<V1.3.1>
    Dim rsData  As ADODB.Recordset
    Dim sGlsSp  As String
    
    sGlsSp = "usp_BloqueaConsulta " & sNumConsulta & ",'" & sIndBloqueada & "'"
    
    OpenMyDataBase
    
    db_BloqueaConsulta = OpenRecordSet(Cnn_Satelite, rsData, sGlsSp)
    Set rsData = Nothing
    
    CloseMyDataBase
    '</V1.3.1>
End Function

