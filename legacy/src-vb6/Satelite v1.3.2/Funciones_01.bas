Attribute VB_Name = "Funciones_01"
'*** Módulo global de la aplicación de ejemplo Bloc de notas MDI.  ***
'*********************************************************************
Option Explicit

Private Declare Function CreateFile Lib "Kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function GetFileSize Lib "Kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Private Declare Function GetFileTime Lib "Kernel32" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Private Declare Function FileTimeToLocalFileTime Lib "Kernel32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Private Declare Function FileTimeToSystemTime Lib "Kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Private Declare Function CloseHandle Lib "Kernel32" (ByVal hObject As Long) As Long
Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Global gdNullDate       As Date
Global gdMinDateExcel   As Date

' Tipo definido por el usuario para almacenar información sobre los formularios secundarios
Type FormState
    Deleted As Integer
    Dirty As Integer
    Color As Long
End Type

Type recCarpetas
    sKey    As String
    sValue  As String
End Type

Type recCampos
    nLargo          As Long
    nTipoDato       As Integer
    sTitulo         As String
    nNumericScale   As Integer
End Type

'Public FState()  As FormState           ' Matriz de tipos definidos por el usuario
'Public Document() As New frmPrincipal     ' Matriz de objetos formulario secundario
Public gFindString As String            ' Almacena el texto de búsqueda.
Public gFindCase As Integer             ' Búsqueda que distingue mayúsculas de minúsculas
Public gCurPos As Integer               ' Almacena la posición del cursor.
Public gFirstTime As Integer            ' Posición inicial.
Public gToolsHidden As Boolean          ' Estado de la barra de herramientas.
Public Const ThisApp = "MDINote"        ' Constante del Registro del sistema.
Public Const ThisKey = "Recent Files"   ' Constante del Registro del sistema.

Global gsNomArchivoAL       As String
Global gbVersionLocal       As Boolean

Global gbCancelar           As Boolean
Global gnObjetoACrear       As Integer
Global gsTagNodoActual      As String
Global gsNombreNodoActual   As String
Global gsNomConsulta        As String
Global gsNumConsulta        As String
Global gsNomUsuario         As String
Global gsNumPerfil          As String
Global gsNomPerfil          As String
Global gsCodTipoUsuario     As String
Global gsNumBaseDatos       As String
Global gsNomArchivoEditar   As String
Global gsNumLote            As String
Global gsNomLote            As String
Global gsNomSolicitante     As String
'<V1.3.1>
Global gsNumRegTabValor     As String
'</V1.3.1>

Global gsNomParametro       As String
'<V1.3.2>
Global gsGlsParametro       As String
'</V1.3.2>
Global gsTipoDato           As String
Global gsTipoAyuda          As String
Global gsGlsAyuda           As String
Global gsIndOpcional        As String

Global gsFechaSeleccionada  As String
Global gnTopControlFecha    As Long
Global gnLeftControlFecha   As Long
Global gsFormatoFechaDB     As String

Global grsLookUp            As ADODB.Recordset
Global gsQueryLookUp        As String
Global gsResultLookUp       As String
Global gsCampoLookUp        As String

Global gsNomArchivoExportar As String
Global gsGlsSeparadorCampos As String
Global gsNomHojaExportar    As String
'<V1.3.1>
Global gsGlsOperadorFiltro  As String
'</V1.3.1>

Global Const wc_tipo_dato_integer = 1
Global Const wc_tipo_dato_float = 2
Global Const wc_tipo_dato_fecha = 3
Global Const wc_tipo_dato_otro = 4
'<INI SP1.2.1>
Global Const wc_tipo_dato_hora = 5
'<FIN SP1.2.1>

'<V1.3.0>
Global gbEjecutandoLote     As Boolean

' Define arreglo para recibir parametros consulta
Public Type lteRegParametros
    num_consulta    As String
    nom_consulta    As String
    arc_salida      As String
    hja_salida      As String
    par_consulta()  As rRegParametros
End Type
Public gaLteRegParametros()     As lteRegParametros
'</V1.3.0>
'<V1.3.1>
Global Const gsGlsValorBlanco = "{vacio}"
'</V1.3.1>

Global Const gsSignoMenor = "#$SignoMenor$#"
Global Const gsSignoComillas = "#$SignoComillas$#"
Global Const gsCodigoStrCon = "String de conexion de base de datos"

Private Type SHFILEOPSTRUCT
    hWnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAborted As Boolean
    hNameMaps As Long
    sProgress As String
End Type

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Private Const GENERIC_WRITE = &H40000000
Private Const GENERIC_READ = &H80000000
Private Const OPEN_EXISTING = 3
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2

Function CargaCarpetasUsuario(sCodPadreInicial As String, tvTreeView As TreeView, nNumMaxKey As Long) As Boolean
    Dim rsData          As ADODB.Recordset
    Dim intX            As Integer
    Dim arrCarpetas()   As recCarpetas
    Dim regCarpetas     As recCarpetas
    Dim nItem           As Integer
    Dim nAux            As Integer
    Dim sKey            As String
    Dim sFolder         As String
    Dim nKey            As Long
    Dim nPos            As Long
    Dim bExiste         As Boolean
    
    Dim mNode           As Node
    Dim sCodPadre       As String
    Dim sCodLlave       As String
    
    nNumMaxKey = 0
    
    ReDim arrCarpetas(0) As recCarpetas
    
    On Error GoTo ErrCargaCarpetasUsuario
    
    If Not db_LeeCarpetasUsuario(sCodPadreInicial, rsData) Then
        Exit Function
    End If
        
    ' Carga carpetas
    nItem = 0
    While Not rsData.EOF
        ' Verifica que se cargue sólo una vez
        bExiste = False
        For nAux = 1 To UBound(arrCarpetas)
            If LCase(arrCarpetas(nAux).sValue) = LCase(rsData!gls_carpeta) Then
                bExiste = True
                Exit For
            End If
        Next nAux
        
        If Not bExiste Then
            nItem = nItem + 1
            ReDim Preserve arrCarpetas(nItem) As recCarpetas
            
            arrCarpetas(nItem).sKey = rsData!num_carpeta
            arrCarpetas(nItem).sValue = rsData!gls_carpeta
        End If
        
        rsData.MoveNext
    Wend

    ' Muestra carpetas
    On Error Resume Next
    For intX = 1 To UBound(arrCarpetas)
        sKey = arrCarpetas(intX).sKey
        sFolder = arrCarpetas(intX).sValue
    
        ' Determina ultimo numero de carpeta
        nKey = 0
        nPos = InStr(sKey, "_")
        If nPos > 0 Then
            nKey = Val(Mid(sKey, nPos + 1))
        End If
        If nKey > nNumMaxKey Then
            nNumMaxKey = nKey
        End If
        
        ' Obtiene nombre de la carpeta
        nPos = fnUltimaPos(sFolder, "\")
        If nPos > 0 Then
            sCodPadre = LCase(Left(sFolder, nPos - 1))
            sFolder = Mid(sFolder, nPos + 1)
        Else
            sCodPadre = sCodPadreInicial
        End If
        
        sCodLlave = LCase(arrCarpetas(intX).sValue)
        Set mNode = tvTreeView.Nodes.Add(sCodPadre, tvwChild, sCodLlave, sFolder, "carpeta")
        mNode.Tag = "DIR_" & sKey
        mNode.Expanded = False
    Next intX

    CargaCarpetasUsuario = True
    Exit Function

ErrCargaCarpetasUsuario:
    Screen.MousePointer = vbNormal
    MsgBox Error, vbInformation, App.Title
    CargaCarpetasUsuario = True
    Exit Function
End Function

Sub CargaParametrosDefault(arrRegParametros() As rRegParametros, nCtaInput As Integer)
    Dim nFila   As Integer
    
    nCtaInput = 0
    For nFila = 1 To UBound(arrRegParametros)
        Select Case UCase(arrRegParametros(nFila).Tipo)
        Case "USERNAME"
            arrRegParametros(nFila).valor = Get_Username
        Case Else
            nCtaInput = nCtaInput + 1
        End Select
    Next nFila
End Sub

Function EsHoraDeEjecucion(ByVal sGlsHorario As String, sGlsHorariosEje As String) As Boolean
    Dim sGlsHoraIni1    As String
    Dim sGlsHoraFin1    As String
    Dim sGlsHoraIni2    As String
    Dim sGlsHoraFin2    As String
    
    Dim sHoraActual     As String
    Dim sGlsAMPM        As String
    
    Dim bOkHora1        As Boolean
    Dim bOkHora2        As Boolean
    
    On Error GoTo Err_EsHoraDeEjecucion
    
    If Trim(sGlsHorario) = "" Then
        EsHoraDeEjecucion = True
        Exit Function
    End If
    
    sGlsHoraIni1 = Trim(Mid(sGlsHorario, 1, 10))
    sGlsHoraFin1 = Trim(Mid(sGlsHorario, 11, 10))
    sGlsHoraIni2 = Trim(Mid(sGlsHorario, 21, 10))
    sGlsHoraFin2 = Trim(Mid(sGlsHorario, 31, 10))
    
    sHoraActual = Trim(Format(Time, "hh:mm am/pm"))
    If InStr(sHoraActual, "pm") > 0 Then
        sGlsAMPM = "p.m."
    Else
        sGlsAMPM = "a.m."
    End If
    
    sHoraActual = Trim(Format(Time, "hh:mm"))
    If InStr(sHoraActual, ":") < 1 Then
        sHoraActual = "0" & sHoraActual
    End If
    sHoraActual = LCase(sHoraActual & " " & sGlsAMPM)
    
    If sGlsHoraIni1 <> "" And sGlsHoraFin1 <> "" Then
        If sGlsHoraIni1 <= sGlsHoraFin1 Then
            If sHoraActual >= sGlsHoraIni1 And sHoraActual <= sGlsHoraFin1 Then
                bOkHora1 = True
            Else
                bOkHora1 = False
            End If
        Else
            If sHoraActual >= sGlsHoraIni1 Or sHoraActual <= sGlsHoraFin1 Then
                bOkHora1 = True
            Else
                bOkHora1 = False
            End If
        End If
        
        sGlsHorariosEje = "entre " & sGlsHoraIni1 & " y " & sGlsHoraFin1
    End If
    
    If sGlsHoraIni2 <> "" And sGlsHoraFin2 <> "" Then
        If sGlsHoraIni2 <= sGlsHoraFin2 Then
            If sHoraActual >= sGlsHoraIni2 And sHoraActual <= sGlsHoraFin2 Then
                bOkHora2 = True
            Else
                bOkHora2 = False
            End If
        Else
            If sHoraActual >= sGlsHoraIni2 Or sHoraActual <= sGlsHoraFin2 Then
                bOkHora2 = True
            Else
                bOkHora2 = False
            End If
        End If
    
        sGlsHorariosEje = sGlsHorariosEje & IIf(sGlsHorariosEje = "", "", " o ") & "entre " & sGlsHoraIni2 & " y " & sGlsHoraFin2
    End If
    
    EsHoraDeEjecucion = (bOkHora1 Or bOkHora2)
    Exit Function
    
Err_EsHoraDeEjecucion:
    Screen.MousePointer = vbNormal
    MsgBox "Error al validar hora de ejecución : " & Error, vbInformation, App.Title
    EsHoraDeEjecucion = False
    Exit Function
End Function

Function fnFindStr(grdiitf As fpSpread, nCol As Long, sTexto As String, Optional nFilaInicio As Long) As Long
    Dim nRowOld     As Long
    Dim nRow        As Long
    Dim nRowItem    As Long
    Dim nRowIni     As Long
    
    Screen.MousePointer = vbHourglass
    nRowItem = 0
    nRowOld = grdiitf.Row
    
    If nFilaInicio > 0 Then
        nRowIni = nFilaInicio
    Else
        nRowIni = 1
    End If
    
    For nRow = nRowIni To grdiitf.MaxRows
        If InStr(UCase(fsGetGrilla(grdiitf, nRow, nCol)), UCase(sTexto)) > 0 Then
            nRowItem = nRow
            Exit For
        End If
    Next nRow
    
    fnFindStr = nRowItem
    grdiitf.Row = nRowOld
    Screen.MousePointer = vbNormal
End Function

Function fnTipoDato(sGlsTipoDato As String) As Integer
    Select Case LCase(sGlsTipoDato)
    Case "entero"
        fnTipoDato = wc_tipo_dato_integer
    Case "decimal"
        fnTipoDato = wc_tipo_dato_float
    Case "fecha"
        fnTipoDato = wc_tipo_dato_fecha
    '<INI SP1.2.1>
    Case "hora"
        fnTipoDato = wc_tipo_dato_hora
    '<FIN SP1.2.1>
    Case "texto"
        fnTipoDato = wc_tipo_dato_otro
    End Select
End Function

Function fnTipoDatoCol(nTypeCol As Integer) As Integer
    '<V1.3.1>
    Dim nTipoSalida     As Integer

    Select Case nTypeCol
    Case 2, 3, 20 ' Integer
        nTipoSalida = wc_tipo_dato_integer
   
    Case 4, 5, 6, 14, 131 ' Float
        nTipoSalida = wc_tipo_dato_float
   
    Case 135 ' Fecha
        nTipoSalida = wc_tipo_dato_fecha

    Case Else
        nTipoSalida = wc_tipo_dato_otro
    End Select

    fnTipoDatoCol = nTipoSalida
    '</V1.3.1>
End Function

Function fnTipoDatoRecordset(nTypeCol As Integer) As Integer
    Dim nTipoSalida As Integer
    
    On Error GoTo Err_fnTipoDatoRecordset
    
    grsTipoDatos.Filter = "gls_proveedor='" & cnn_Consulta.Provider & "' AND num_tipo_columna_in=" & nTypeCol
    
    If grsTipoDatos.EOF Then
        Select Case nTypeCol
        Case 2, 3, 20 ' Integer
            nTipoSalida = wc_tipo_dato_integer
       
        Case 4, 5, 6, 14, 131 ' Float
            nTipoSalida = wc_tipo_dato_float
       
        Case 135 ' Fecha
            nTipoSalida = wc_tipo_dato_fecha
    
        Case Else
            nTipoSalida = wc_tipo_dato_otro
        End Select
    
    Else
        nTipoSalida = IIf(IsNull(grsTipoDatos!num_tipo_columna_out), wc_tipo_dato_otro, grsTipoDatos!num_tipo_columna_out)
        
        If nTipoSalida <> wc_tipo_dato_integer And _
           nTipoSalida <> wc_tipo_dato_float And _
           nTipoSalida <> wc_tipo_dato_fecha Then
            nTipoSalida = wc_tipo_dato_otro
        End If
    End If
    
    fnTipoDatoRecordset = nTipoSalida
    Exit Function
    
Err_fnTipoDatoRecordset:
    fnTipoDatoRecordset = wc_tipo_dato_otro
End Function

Function fsConvierteData(sValor As String, sTipo As String) As String
    Select Case LCase(sTipo)
    Case "fecha"
        fsConvierteData = Format(fdValorFecha(sValor), gsFormatoFechaDB)
    Case Else
        fsConvierteData = sValor
    End Select
End Function

Function fsConvierteTextoToLinea(ByVal sTexto As String) As String
    Dim sTextoOut   As String
    
    sTextoOut = sTexto
    sTextoOut = Replace(sTextoOut, Chr(10), " ")
    sTextoOut = Replace(sTextoOut, Chr(13), "")
    
    fsConvierteTextoToLinea = sTextoOut
End Function


Sub GrabaLog(sTexto As String)
    'Open App.Path & "\satelite.log" For Append As #3
    'Print #3, sTexto
    'Close #3
End Sub


Function HelpFecha(frmForm As Form, ctrControl As Control) As Boolean
    Dim bCancel     As Integer

    gsFechaSeleccionada = ""
    gnTopControlFecha = ctrControl.Top + frmForm.Top + 280
    gnLeftControlFecha = ctrControl.Left + frmForm.Left + ctrControl.Width + 4700
    frmFecha.Show vbModal

    If gsFechaSeleccionada <> "" Then
        Call SetMasked(ctrControl, gsFechaSeleccionada)
        HelpFecha = True
    Else
        HelpFecha = False
    End If
End Function

Sub SetMasked(rctrMask As Control, rsValMask As String)
    If TypeOf rctrMask Is TextBox Then
        rctrMask.SelStart = 0
        rctrMask.SelLength = Len(rctrMask.Text)
        rctrMask.SelText = rsValMask
        rctrMask.SelLength = 0
    Else
        rctrMask.SelStart = 0
        rctrMask.SelLength = rctrMask.MaxLength
        rctrMask.SelText = rsValMask
        rctrMask.SelLength = 0
    End If
End Sub

Function Hyoplus() As Boolean
    Dim rsData              As ADODB.Recordset
    Dim sTextLine           As String
    Dim sVersion            As String
    Dim sCodigo             As String
    Dim nID1, nID2, nID3    As Boolean
    Dim i                   As Integer
    Dim sPathExe            As String
    Dim sPathHlp            As String
    Dim sFileHlp            As String
    Dim sFileHlpCnt         As String
    Dim sFileAdmHlp         As String
    Dim sFileAdmHlpCnt      As String
    Dim sFileExe            As String
    Dim sFileExeLocal       As String
    
    Dim nFileSize           As Long
    Dim dFileDate           As Date
    Dim nFileSizeLocal      As Long
    Dim dFileDateLocal      As Date
    
    On Error GoTo ErrHyoplus
    
    GrabaLog "Validando..."
    sVersion = "Ver" & App.Major & "." & App.Minor & "." & App.Revision
    
    If Not db_LeeSysConfig(rsData) Then
        Screen.MousePointer = vbNormal
        Hyoplus = False
        End
    End If
    
    If rsData.EOF Then
        Set rsData = Nothing
        Screen.MousePointer = vbNormal
        MsgBox "Sistema no ha sido autorizado. Solicite su clave de autorización al distribuidor de este software", vbCritical, App.Title
        End
    End If
    
    'rsData.Filter = "id_1='" & sDecodifClave(App.CompanyName, 1, App.Title) & "'"
    'If rsData.EOF Then
    '    Set rsData = Nothing
    '    Screen.MousePointer = vbNormal
    '    MsgBox "Sistema no ha sido autorizado. Solicite su clave de autorización al distribuidor de este software", vbCritical, App.Title
    '    End
    'End If
    
    ' Verifica exitencia de hlp
    sPathExe = Trim("" & rsData!path_exe)
    sPathHlp = Trim("" & rsData!path_hlp)
    GrabaLog sPathExe
    GrabaLog sPathHlp
    
    If sPathExe <> "" Then
        If Right(sPathExe, 1) = "/" Or Right(sPathExe, 1) = "\" Then
            sFileExe = sPathExe & "Satelite.Exe"
        Else
            sFileExe = sPathExe & "\Satelite.Exe"
        End If
        FileInfo sFileExe, nFileSize, dFileDate
        
        sFileExeLocal = App.Path & "\Satelite.Exe"
        FileInfo sFileExeLocal, nFileSizeLocal, dFileDateLocal
        
        If nFileSize > 0 And (nFileSize <> nFileSizeLocal Or dFileDate <> dFileDateLocal) Then
            Screen.MousePointer = vbNormal
            Shell App.Path & "\Version.exe """ & sFileExe & """;""" & sFileExeLocal & """", vbNormalFocus
            End
        End If
    End If
    
    '<V1.3.0>
    If Right(sPathExe, 1) = "/" Or Right(sPathExe, 1) = "\" Then
        sFileExe = sPathExe & "Exportar.Exe"
    Else
        sFileExe = sPathExe & "\Exportar.Exe"
    End If
    sFileExeLocal = App.Path & "\Exportar.Exe"
    FileInfo sFileExe, nFileSize, dFileDate
    FileInfo sFileExeLocal, nFileSizeLocal, dFileDateLocal
    If nFileSize > 0 And (nFileSize <> nFileSizeLocal Or dFileDate <> dFileDateLocal) Then
        FileCopy sFileExe, sFileExeLocal
    End If
    '</V1.3.0>
    
    If sPathHlp <> "" Then
        If Right(sPathHlp, 1) = "/" Or Right(sPathHlp, 1) = "\" Then
            sFileHlp = sPathHlp & "Satelite.hlp"
            sFileHlpCnt = sPathHlp & "Satelite.cnt"
            sFileAdmHlp = sPathHlp & "AdmSatelite.hlp"
            sFileAdmHlpCnt = sPathHlp & "AdmSatelite.cnt"
        Else
            sFileHlp = sPathHlp & "\Satelite.hlp"
            sFileHlpCnt = sPathHlp & "\Satelite.cnt"
            sFileAdmHlp = sPathHlp & "\AdmSatelite.hlp"
            sFileAdmHlpCnt = sPathHlp & "\AdmSatelite.cnt"
        End If
        
        FileInfo sFileHlp, nFileSize, dFileDate
        FileInfo App.Path & "\Satelite.hlp", nFileSizeLocal, dFileDateLocal
        If nFileSize > 0 And (nFileSize <> nFileSizeLocal Or dFileDate <> dFileDateLocal) Then
            FileCopy sFileHlp, App.Path & "\Satelite.hlp"
            If Exist(sFileHlpCnt) Then
                FileCopy sFileHlpCnt, App.Path & "\Satelite.cnt"
            End If
        End If
    
        FileInfo sFileAdmHlp, nFileSize, dFileDate
        FileInfo App.Path & "\AdmSatelite.hlp", nFileSizeLocal, dFileDateLocal
        If nFileSize > 0 And (nFileSize <> nFileSizeLocal Or dFileDate <> dFileDateLocal) Then
            FileCopy sFileAdmHlp, App.Path & "\AdmSatelite.hlp"
            If Exist(sFileAdmHlpCnt) Then
                FileCopy sFileAdmHlpCnt, App.Path & "\AdmSatelite.cnt"
            End If
        End If
    End If
    
    nID1 = (App.CompanyName = sDecodifClave("" & rsData!id_1, -1, App.Title))
    nID2 = (sVersion = sDecodifClave("" & rsData!id_2, -1, App.Title))
    GrabaLog Trim(nID1)
    GrabaLog Trim(nID2)
    
    i = 1
    While i <= Len(sVersion) And i <= Len(App.CompanyName)
        sCodigo = sCodigo & Mid(sVersion, i, 1) & Mid(App.CompanyName, i, 1)
        i = i + 1
    Wend
    If i > Len(sVersion) And i <= Len(App.CompanyName) Then
        sCodigo = sCodigo & Mid(App.CompanyName, i)
    End If
    If i > Len(App.CompanyName) And i <= Len(sVersion) Then
        sCodigo = sCodigo & Mid(sVersion, i)
    End If
    
    nID3 = (sCodigo = sDecodifClave("" & rsData!id_3, -1, App.Title))
    GrabaLog Trim(nID3)
    
    'If Not nID1 Then
    '    Screen.MousePointer = vbNormal
    '    MsgBox "Esta licencia no está autorizada para su Organización", vbCritical, App.Title
    '    End
    'End If
    'If Not nID3 Then
    '    Screen.MousePointer = vbNormal
    '    MsgBox "Sistema no ha sido autorizado. Solicite su clave de autorización al distribuidor de este software", vbCritical, App.Title
    '    End
    'End If
     
    GrabaLog "Validado."
    Hyoplus = True
    Exit Function
    
ErrHyoplus:
    Screen.MousePointer = vbNormal
    MsgBox Error, vbCritical, App.Title
    End
End Function

Public Sub FileInfo(ByVal FileName As String, ByRef FileSize As Long, ByRef FileDate As Date)
On Error GoTo Er_FileInfo
    Dim fecha As Date
    Dim lngHandle As Long, SHDirOp As SHFILEOPSTRUCT, lngLong As Long
    Dim Ft1 As FILETIME, Ft2 As FILETIME, SysTime As SYSTEMTIME
    lngHandle = CreateFile(FileName, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)
    FileSize = GetFileSize(lngHandle, lngLong)
    GetFileTime lngHandle, Ft1, Ft1, Ft2
    FileTimeToLocalFileTime Ft2, Ft1
    FileTimeToSystemTime Ft1, SysTime
    If Left(Format(0, "ddddd"), 2) = "12" Then
        fecha = CDate(LTrim(Str$(SysTime.wMonth)) + "/" + LTrim(Str$(SysTime.wDay)) + "/" + LTrim(Str$(SysTime.wYear)) + " " + LTrim(Str$(SysTime.wHour)) + ":" + LTrim(Str$(SysTime.wMinute)) + ":" + LTrim(Str$(SysTime.wSecond)))
    Else
        fecha = CDate(LTrim(Str$(SysTime.wDay)) + "/" + LTrim(Str$(SysTime.wMonth)) + "/" + LTrim(Str$(SysTime.wYear)) + " " + LTrim(Str$(SysTime.wHour)) + ":" + LTrim(Str$(SysTime.wMinute)) + ":" + LTrim(Str$(SysTime.wSecond)))
    End If
    FileDate = fecha
    CloseHandle lngHandle
    Exit Sub
Er_FileInfo:
    FileSize = 0
    FileDate = CDate(0)
End Sub

Function EjecutaConsulta(ByVal sNumConsulta As String, nNumBaseDatos As Integer, sGlsQuery As String, arrRegParametros() As rRegParametros, _
                         rsConsulta As ADODB.Recordset, rsFormatos As ADODB.Recordset, nTotalRegUltQuery As Long, _
                         bResultadoEnGrilla As Boolean, grdResultado As fpSpread, txtResultado As TextBox, _
                         StatusBar1 As StatusBar, ProgressBar1 As ProgressBar) As Boolean
    
    Dim nCtaParamInput  As Integer
    
    On Error GoTo ErrEjecutaConsulta
    
    ' Carga parámetros
    If UBound(arrRegParametros) <= 0 Then
        ReDim gaRegParametros(0) As rRegParametros
    Else
        Call CargaParametrosDefault(arrRegParametros, nCtaParamInput)
        If nCtaParamInput > 0 Then
            gaRegParametros = arrRegParametros
            frmParametros.mnNumBaseDatos = nNumBaseDatos
        
            frmParametros.Show vbModal
            If gbCancelar Then
                EjecutaConsulta = False
                Exit Function
            Else
                arrRegParametros = gaRegParametros
            End If
        End If
    End If
        
    If Not EjecutaSentencia(sNumConsulta, nNumBaseDatos, sGlsQuery, arrRegParametros, rsConsulta, ProgressBar1, StatusBar1) Then
        EjecutaConsulta = False
        Exit Function
    Else
        nTotalRegUltQuery = rsConsulta.RecordCount
        StatusBar1.Panels(2).Text = ""
        StatusBar1.Panels(3).Text = Trim(CStr(nTotalRegUltQuery)) & " reg"
        
        If nTotalRegUltQuery = 0 Then
            MsgBox "No hay registros para esta consulta", vbInformation, App.Title
        Else
            If bResultadoEnGrilla Then
                Call CargarResultadoEnGrilla(rsConsulta, rsFormatos, grdResultado, txtResultado, ProgressBar1)
            Else
                Call CargarResultadoEnTexto(rsConsulta, rsFormatos, grdResultado, txtResultado, ProgressBar1)
            End If
        End If
    End If
    
    EjecutaConsulta = True
    Exit Function
    
ErrEjecutaConsulta:
    Screen.MousePointer = vbNormal
    MsgBox Error, vbCritical, App.Title
    
    EjecutaConsulta = False
    Exit Function
    Resume
End Function

Function fsExtraeNombreConsulta(sTag As String) As String
    Dim nPos    As Integer
    
    nPos = InStr(sTag, "[")
    If nPos > 0 Then
        fsExtraeNombreConsulta = Mid(Left(sTag, nPos - 1), 5)
    Else
        fsExtraeNombreConsulta = ""
    End If
End Function

Function fsExtraeRutaNombreConsulta(sTag As String) As String
    Dim nPos    As Integer
    Dim sFile   As String
    
    nPos = InStr(sTag, "[")
    sFile = Mid(sTag, nPos + 1)
    fsExtraeRutaNombreConsulta = Left(sFile, Len(sFile) - 1)
End Function

Function EjecutaSentencia(ByVal sNumConsulta As String, nNumBaseDatos As Integer, sGlsQuery As String, arrParametros() As rRegParametros, _
                          rsConsulta As ADODB.Recordset, ProgressBar1 As ProgressBar, StatusBar1 As StatusBar) As Boolean
    Dim sFecEjecucion   As String
    Dim sHorEjecucion   As String
    Dim sTotTiempo      As String
    Dim sNumRegistros   As String
    Dim dHorInicio      As Date
    Dim dHorFin         As Date
    
    Dim bOk             As Boolean
    Dim nItem           As Long
    Dim sSql            As String
    
    On Error GoTo ErrEjecutaSentencia
    
    Screen.MousePointer = vbHourglass
    
    bOk = True
    For nItem = 1 To UBound(arrParametros)
        If arrParametros(nItem).Opcional = False And Trim(arrParametros(nItem).valor) = "" Then
            bOk = False
            Exit For
        End If
    Next
    
    If Not bOk Then
        Screen.MousePointer = vbNormal
        MsgBox "Debe ingresar valores de todos los parámetros que son requeridos", vbInformation, App.Title
        EjecutaSentencia = False
        Exit Function
    End If
    
    ProgressBar1.Width = StatusBar1.Panels(2).Width - 45
    ProgressBar1.Top = StatusBar1.Top + 60
    
    ' Conecta a base de datos de la consulta
    StatusBar1.Panels(2).Text = "Conectando a Base de Datos ..."
    If Not ConectaBaseDatos(nNumBaseDatos) Then
        Screen.MousePointer = vbNormal
        EjecutaSentencia = False
        Exit Function
    End If
    
    ' Convierte los parametros a los valores de ejecucion
    sSql = sGlsQuery
    If Not CargaValores(sSql, arrParametros) Then
        Screen.MousePointer = vbNormal
        MsgBox "Error al cargar los parámetros", vbCritical, App.Title
        EjecutaSentencia = False
        Exit Function
    End If

    ' Ejecuta consulta final
    sFecEjecucion = Format(Date, "yyyymmdd")
    dHorInicio = Time
    
    StatusBar1.Panels(2).Text = "Ejecutando consulta ..."
    Set rsConsulta = New ADODB.Recordset
    rsConsulta.CursorLocation = adUseClient
    rsConsulta.Open sSql, cnn_Consulta, adOpenForwardOnly, adLockReadOnly
    
    DoEvents
    dHorFin = Time
    
    ' Graba registro de log de ejecucion
    If sNumConsulta <> "" Then
        sHorEjecucion = Format(dHorInicio, "hh:mm:ss")
        sTotTiempo = Format(dHorFin - dHorInicio, "hh:mm:ss")
        sNumRegistros = rsConsulta.RecordCount
        
        Call db_GrabaLogEjecucionConsulta(sNumConsulta, sFecEjecucion, sHorEjecucion, sTotTiempo, sNumRegistros)
    End If
    
    ' Fin
    Screen.MousePointer = vbNormal
    EjecutaSentencia = True
    Exit Function
    
ErrEjecutaSentencia:
    EjecutaSentencia = False
    ProgressBar1.Value = 0
    ProgressBar1.Visible = False
    Screen.MousePointer = vbNormal
    MsgBox "Error al ejecutar consulta : " & Err.Description, vbCritical, App.Title
    Exit Function
End Function

Public Function CargaValores(sSql As String, maRegParametros() As rRegParametros) As Boolean
    Dim sNombre As String
    Dim sValor As String
    Dim sTipo As String
    Dim bOk As Boolean
    Dim nX As Integer
    bOk = True
    
    If UBound(maRegParametros) > 0 And bOk Then
        For nX = 1 To UBound(maRegParametros)
            sNombre = maRegParametros(nX).Nombre
            sTipo = maRegParametros(nX).Tipo
            sValor = fsConvierteData(maRegParametros(nX).valor, sTipo)
            
            bOk = ReemplazaValor(sSql, sNombre, sTipo, sValor)
        Next
    End If
    CargaValores = bOk
    DoEvents
End Function

Public Function ReemplazaValor(sSql As String, sNombre As String, sTipo As String, sValor As String) As Boolean
    Dim nPosIni As Integer
    Dim nPosFin As Integer
    Dim bEncontrado As Boolean
    Dim nlongitud As Integer
    
    bEncontrado = False
    nPosIni = InStr(1, LCase(sSql), "@" & sNombre & "@")
    nlongitud = Len("@" & sNombre & "@")
    bEncontrado = IIf(nPosIni > 0, True, False)
    Do While bEncontrado
        sSql = Mid(sSql, 1, nPosIni - 1) & Formateador(sValor, sTipo) & Mid(sSql, nPosIni + nlongitud)
        nPosIni = InStr(1, LCase(sSql), "@" & sNombre & "@")
        nlongitud = Len("@" & sNombre & "@")
        bEncontrado = IIf(nPosIni > 0, True, False)
    Loop
    ReemplazaValor = True
End Function


Private Function Formateador(sNombre As String, sTipo As String) As String
    Select Case sTipo
    Case "D"
        Formateador = "'" & sNombre & "'"
    Case "N"
        Formateador = sNombre
    Case "T"
        Formateador = "'" & sNombre & "'"
    Case "S"
        Formateador = "'" & sNombre & "'"
    Case "L"
        Formateador = sNombre
    Case Else
        Formateador = sNombre
    End Select
End Function

Sub CargarResultadoEnGrilla(rsConsulta As ADODB.Recordset, rsFormatos As ADODB.Recordset, grdResultado As fpSpread, txtResultado As TextBox, ProgressBar1 As ProgressBar)
    Dim nX                  As Integer
    Dim nY                  As Integer
    Dim nCampos             As Integer
    Dim nReg                As Long
    Dim arrCampos()         As recCampos
    Dim nStatus             As Long
    Dim fCampos             As Field
    Dim nTotalRegQuery      As Long
        
    Dim nTipoDatoSalida     As Integer
    Dim sIndSeparadorMiles  As String
    Dim sNumDecimales       As String
    Dim sFormatoIn          As String
    Dim sFormatoOut         As String
        
    nTotalRegQuery = rsConsulta.RecordCount
    
    If nTotalRegQuery > 0 Then
        ProgressBar1.Max = nTotalRegQuery
    Else
        ProgressBar1.Max = 1
    End If
    ProgressBar1.Value = 0
    ProgressBar1.Visible = True
    
    grdResultado.Visible = False
    txtResultado.Visible = False

    nCampos = rsConsulta.Fields.Count
    grdResultado.UnitType = UnitTypeTwips
    grdResultado.MaxCols = nCampos
    grdResultado.MaxRows = nTotalRegQuery
    ReDim Preserve arrCampos(nCampos) As recCampos
    
    ' Configura las columnas
    nX = 0
    For Each fCampos In rsConsulta.Fields
        Call ConfigColumna(grdResultado, nX, fCampos.Type, fCampos.NumericScale)
        nX = nX + 1
        arrCampos(nX).nTipoDato = fnTipoDatoRecordset(fCampos.Type)
        arrCampos(nX).nNumericScale = fCampos.NumericScale
    Next
    
    With grdResultado
        .ReDraw = False
        If nTotalRegQuery > 0 Then
            rsConsulta.MoveLast
            rsConsulta.MoveFirst
        End If
        If Not rsConsulta.EOF Then
            For nReg = 1 To nTotalRegQuery
                .Row = nReg
                ProgressBar1.Value = nReg

                For nY = 1 To nCampos
                    .Col = nY
                    nTipoDatoSalida = arrCampos(nY).nTipoDato
                    
                    ' Busca formato de salida del campo
                    sIndSeparadorMiles = "S"
                    If nTipoDatoSalida = wc_tipo_dato_integer Then
                        sNumDecimales = "0"
                    Else
                        If arrCampos(nY).nNumericScale = 0 Then
                            sNumDecimales = "0"
                        Else
                            sNumDecimales = "2"
                        End If
                    End If
                    sFormatoIn = ""
                    sFormatoOut = ""
                    
                    rsFormatos.Filter = "nom_columna='" & LCase(rsConsulta(nY - 1).Name) & "'"
                    If Not rsFormatos.EOF Then
                        nTipoDatoSalida = fnTipoDato("" & rsFormatos!cod_tipo_dato_salida)
                        sIndSeparadorMiles = "" & rsFormatos!ind_separador_miles
                        sNumDecimales = "" & rsFormatos!num_decimales
                        sFormatoIn = "" & rsFormatos!gls_formato_entrada
                        sFormatoOut = "" & rsFormatos!gls_formato_salida
                        
                        sFormatoIn = Replace(sFormatoIn, gsSignoMenor, "<")
                        sFormatoIn = Replace(sFormatoIn, gsSignoComillas, """")
                        sFormatoOut = Replace(sFormatoOut, gsSignoMenor, "<")
                        sFormatoOut = Replace(sFormatoOut, gsSignoComillas, """")
                    End If
                    
                    ' Carga valor formateado
                    Select Case nTipoDatoSalida
                    Case wc_tipo_dato_float, wc_tipo_dato_integer
                        grdResultado.Text = fsFormatoValorNumerico("" & rsConsulta(nY - 1), sIndSeparadorMiles, sNumDecimales)
                    '<INI SP1.2.1>
                    Case wc_tipo_dato_fecha, wc_tipo_dato_hora
                    '<FIN SP1.2.1>
                        grdResultado.Text = fsFormatoValorFecha("" & rsConsulta(nY - 1), sFormatoIn, sFormatoOut)
                    Case Else
                        grdResultado.Text = "" & rsConsulta(nY - 1)
                    End Select
                    
                    ' Calcula ancho de la columna
                    frmMdiPadre.Label1.Caption = grdResultado.Text
                    If frmMdiPadre.Label1.Width > arrCampos(nY).nLargo Then arrCampos(nY).nLargo = frmMdiPadre.Label1.Width
                Next
                rsConsulta.MoveNext
            Next
        End If
        
        If nTotalRegQuery > 0 Then
            rsConsulta.MoveLast
            rsConsulta.MoveFirst
        End If
        .Row = 0
        .ColWidth(0) = 5
        For nX = 0 To nCampos - 1
            .Col = nX + 1
            .Text = FormatoTitulo(rsConsulta(nX).Name)
            frmMdiPadre.Label1.Caption = .Text
            If frmMdiPadre.Label1.Width > arrCampos(nX + 1).nLargo Then arrCampos(nX + 1).nLargo = frmMdiPadre.Label1.Width
            .ColWidth(nX + 1) = arrCampos(nX + 1).nLargo + 125
            
        Next
        .ReDraw = True
    End With
    
    grdResultado.Row = 1
    grdResultado.Row2 = grdResultado.MaxRows
    grdResultado.Col = 1
    grdResultado.Col2 = grdResultado.MaxCols
    grdResultado.BlockMode = True
    grdResultado.Lock = True
    grdResultado.BlockMode = False
    
    grdResultado.Visible = True
    ProgressBar1.Value = 0
    ProgressBar1.Visible = False
End Sub


Sub fnPutGrilla(grdGrilla As fpSpread, nFila As Variant, nColumna As Variant, sString As Variant)
    ' No se permite Fuera de Rango de Columna
    If nColumna > grdGrilla.MaxCols Then
        Exit Sub
    End If

    ' Cheque fuera de rango sobre la fila
    If nFila > grdGrilla.MaxRows Then
        grdGrilla.MaxRows = grdGrilla.MaxRows + 1
    End If
    grdGrilla.Row = nFila
    grdGrilla.Col = nColumna
    grdGrilla.Text = sString
End Sub

Function fdValorFecha(sFecha As String) As Date
    Dim dFecha  As Date

    On Error GoTo ErrValorFecha
    dFecha = CVDate(ConvFormFecha(sFecha))
    fdValorFecha = dFecha
    Exit Function

ErrValorFecha:
    fdValorFecha = gdNullDate
End Function

Function EsEntero(sNumero As String) As Boolean
    Dim nVal    As Double
    
    On Error GoTo ErrEsEntero
    
    nVal = CLng(sNumero)
    EsEntero = True
    Exit Function

ErrEsEntero:
    EsEntero = False
End Function
Function EsDecimal(sNumero As String) As Boolean
    Dim nVal    As Double
    
    On Error GoTo ErrEsEntero
    
    nVal = CDbl(sNumero)
    EsDecimal = True
    Exit Function

ErrEsEntero:
    EsDecimal = False
End Function

Function ConvFormFecha(sFecha As String) As String
    Dim vtFecha         As Variant
    Dim sTextMsg        As String
    Dim nPrimerPalo     As Integer
    Dim nSegundoPalo    As Integer
    
    Dim sFechaAux       As String
    Dim nPosDia         As Integer
    Dim nPosMes         As Integer
    Dim nPosAño         As Integer
    Dim nNumDigMes      As Integer
    Dim nNumDigAño      As Integer

    sFechaAux = DateValue("31/12/2110")
    
    nPosDia = InStr(sFechaAux, "31")
    
    nPosMes = InStr(sFechaAux, "12")
    nNumDigMes = 2
    If nPosMes = 0 Then
        nPosMes = InStr(sFechaAux, "7")
        nNumDigMes = 1
    End If
    
    nPosAño = InStr(sFechaAux, "2110")
    nNumDigAño = 4
    If nPosAño = 0 Then
        nPosAño = InStr(sFechaAux, "10")
        nNumDigAño = 2
    End If
    
    Mid(sFechaAux, nPosDia, 2) = Mid(sFecha, 1, 2)
    Mid(sFechaAux, nPosMes, nNumDigMes) = Mid(sFecha, 4, 2)
    Mid(sFechaAux, nPosAño, nNumDigAño) = Mid(sFecha, 7 + (4 - nNumDigAño), nNumDigAño)
    
    '<INI SP1.2.1>
    nPosDia = InStr(sFecha, " ")
    If nPosDia > 0 Then
        sFechaAux = sFechaAux & Mid(sFecha, nPosDia)
    End If
    '<FIN SP1.2.1>
    ConvFormFecha = sFechaAux
End Function



Sub ConfigColumna(grdResultado As fpSpread, nCol As Integer, nTypeCol As Integer, nNumericScale As Integer)
    grdResultado.Col = nCol '+ 1
    grdResultado.Col2 = nCol '+ 1
    grdResultado.Row = 1
    grdResultado.Row2 = grdResultado.MaxRows
    grdResultado.BlockMode = True
    
    Select Case fnTipoDatoRecordset(nTypeCol)
    
    Case wc_tipo_dato_integer
        'grdResultado.CellType = 3
        'grdResultado.TypeFloatDecimalPlaces = 0
        'grdResultado.TypeFloatMoney = False
        'grdResultado.TypeFloatSeparator = True
        'grdResultado.TypeHAlign = 1
        'grdResultado.TypeIntegerMin = -2147483648#
        'grdResultado.TypeIntegerMax = 2147483647
        
        'If gbComaDecimal() Then
        '    grdResultado.FloatDefSepChar = Asc(",")
        'Else
        '    grdResultado.FloatDefSepChar = Asc(".")
        'End If
    
        grdResultado.CellType = CellTypeStaticText ' = 5
        grdResultado.TypeHAlign = TypeHAlignRight ' = 1
    
    Case wc_tipo_dato_float
        'grdResultado.CellType = 2
        'grdResultado.TypeFloatDecimalPlaces = IIf(Environ("Satelite$FloatDecimalPlaces") = "", 2, Val(Environ("Satelite$FloatDecimalPlaces")))
        'grdResultado.TypeFloatMoney = False
        'grdResultado.TypeFloatSeparator = True
        'grdResultado.TypeHAlign = 1
        'grdResultado.TypeFloatMin = -99999999999.9999
        'grdResultado.TypeFloatMax = 99999999999.9999
        'If Not gbComaDecimal() Then
        '    grdResultado.FloatDefSepChar = Asc(",")
        '    grdResultado.FloatDefDecimalChar = Asc(".")
        'Else
        '    grdResultado.FloatDefSepChar = Asc(".")
        '    grdResultado.FloatDefDecimalChar = Asc(",")
        'End If
    
        grdResultado.CellType = CellTypeStaticText ' = 5
        grdResultado.TypeHAlign = TypeHAlignRight ' = 1
    
    Case wc_tipo_dato_fecha
        grdResultado.CellType = CellTypeStaticText ' = 5
        grdResultado.TypeHAlign = TypeHAlignRight ' = 1
        
    Case Else
        grdResultado.CellType = CellTypeStaticText ' = 5
        grdResultado.TypeHAlign = TypeHAlignLeft
    End Select

    grdResultado.BlockMode = False
End Sub

Sub CargarResultadoEnTexto(rsConsulta As ADODB.Recordset, rsFormatos As ADODB.Recordset, grdResultado As fpSpread, txtResultado As TextBox, ProgressBar1 As ProgressBar)
    Dim sLinea              As String
    Dim nCampos             As Long
    Dim fCampos             As Field
    Dim nReg                As Long
    Dim nX                  As Long
    Dim sValor              As String
    
    Dim nTipoDatoSalida     As Integer
    Dim sIndSeparadorMiles  As String
    Dim sNumDecimales       As String
    Dim sFormatoIn          As String
    Dim sFormatoOut         As String
    Dim nTotalRegQuery      As Long
    Dim arrCampos()         As recCampos
        
    nTotalRegQuery = rsConsulta.RecordCount
        
    ProgressBar1.Max = nTotalRegQuery
    ProgressBar1.Value = 0
    ProgressBar1.Visible = True
    grdResultado.Visible = False
    txtResultado.Visible = False
    txtResultado.Text = ""
    
    nCampos = rsConsulta.Fields.Count
    ReDim Preserve arrCampos(nCampos) As recCampos
    nX = 0
    For Each fCampos In rsConsulta.Fields
        nX = nX + 1
        
        arrCampos(nX).nTipoDato = fnTipoDatoRecordset(fCampos.Type)
        arrCampos(nX).sTitulo = FormatoTitulo(fCampos.Name)
        arrCampos(nX).nNumericScale = fCampos.NumericScale
                
        Select Case arrCampos(nX).nTipoDato
        Case wc_tipo_dato_integer
            arrCampos(nX).nLargo = IIf(Len(arrCampos(nX).sTitulo) > 11, Len(arrCampos(nX).sTitulo), 11)
        Case wc_tipo_dato_float
            arrCampos(nX).nLargo = IIf(Len(arrCampos(nX).sTitulo) > 50, Len(arrCampos(nX).sTitulo), 50)
        Case wc_tipo_dato_fecha
            arrCampos(nX).nLargo = IIf(Len(arrCampos(nX).sTitulo) > 35, Len(arrCampos(nX).sTitulo), 35)
        Case Else
            arrCampos(nX).nLargo = IIf(Len(arrCampos(nX).sTitulo) > fCampos.DefinedSize, Len(arrCampos(nX).sTitulo), fCampos.DefinedSize)
        End Select
    Next

    rsConsulta.MoveLast
    rsConsulta.MoveFirst
    If Not rsConsulta.EOF Then
        ' Imprime titulos de campos
        sLinea = ""
        For nX = 1 To nCampos
            sValor = Left(arrCampos(nX).sTitulo & Space(255), arrCampos(nX).nLargo)
            sLinea = sLinea & sValor & " "
        Next
        txtResultado = txtResultado & sLinea & Chr(13) & Chr(10)
    
        sLinea = ""
        For nX = 1 To nCampos
            sValor = Left("------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------", arrCampos(nX).nLargo)
            sLinea = sLinea & sValor & " "
        Next
        txtResultado = txtResultado & sLinea & Chr(13) & Chr(10)
    

        rsConsulta.MoveLast
        rsConsulta.MoveFirst
        For nReg = 1 To nTotalRegQuery
            ProgressBar1.Value = nReg

            sLinea = ""
            For nX = 1 To nCampos
                nTipoDatoSalida = arrCampos(nX).nTipoDato
                
                ' Busca formato de salida del campo
                sIndSeparadorMiles = "S"
                If nTipoDatoSalida = wc_tipo_dato_integer Then
                    sNumDecimales = "0"
                Else
                    If arrCampos(nX).nNumericScale = 0 Then
                        sNumDecimales = "0"
                    Else
                        sNumDecimales = "2"
                    End If
                End If
                sFormatoIn = ""
                sFormatoOut = ""
                
                rsFormatos.Filter = "nom_columna='" & LCase(rsConsulta(nX - 1).Name) & "'"
                If Not rsFormatos.EOF Then
                    nTipoDatoSalida = fnTipoDato("" & rsFormatos!cod_tipo_dato_salida)
                    sIndSeparadorMiles = "" & rsFormatos!ind_separador_miles
                    sNumDecimales = "" & rsFormatos!num_decimales
                    sFormatoIn = "" & rsFormatos!gls_formato_entrada
                    sFormatoOut = "" & rsFormatos!gls_formato_salida
                    
                    sFormatoIn = Replace(sFormatoIn, gsSignoMenor, "<")
                    sFormatoIn = Replace(sFormatoIn, gsSignoComillas, """")
                    sFormatoOut = Replace(sFormatoOut, gsSignoMenor, "<")
                    sFormatoOut = Replace(sFormatoOut, gsSignoComillas, """")
                End If
                
                ' Carga valor formateado
                Select Case nTipoDatoSalida
                Case wc_tipo_dato_float, wc_tipo_dato_integer
                    sValor = Right(Space(255) & fsFormatoValorNumerico("" & rsConsulta(nX - 1), sIndSeparadorMiles, sNumDecimales), arrCampos(nX).nLargo)
                    
                '<INI SP1.2.1>
                Case wc_tipo_dato_fecha, wc_tipo_dato_hora
                '<FIN SP1.2.1>
                    sValor = Left(fsFormatoValorFecha("" & rsConsulta(nX - 1), sFormatoIn, sFormatoOut) & Space(255), arrCampos(nX).nLargo)
                    
                Case Else
                    sValor = Left("" & rsConsulta(nX - 1) & Space(255), arrCampos(nX).nLargo)
                    
                End Select
                
                sLinea = sLinea & sValor & " "
            Next
            
            txtResultado = txtResultado & sLinea & Chr(13) & Chr(10)
            rsConsulta.MoveNext
        Next
    End If
    
    txtResultado.Visible = True
    ProgressBar1.Value = 0
    ProgressBar1.Visible = False
End Sub


Function fnBuscaClave(ByVal sLinea As String, ByVal sClaves As String) As Integer
    Dim sClave  As String
    Dim nPos1   As Integer
    Dim nPos    As Integer
    
    nPos = 0
    nPos1 = InStr(sClaves, "]")
    If nPos1 > 0 Then
        sClave = Left(sClaves, nPos1)
        sClaves = Mid(sClaves, nPos1 + 1)
    Else
        sClave = sClaves
    End If
    While sClave <> ""
        If InStr(sLinea, sClave) > 0 Then
            nPos = InStr(sLinea, sClave)
            sClave = ""
        Else
            nPos1 = InStr(sClaves, "]")
            If nPos1 > 0 Then
                sClave = Left(sClaves, nPos1)
                sClaves = Mid(sClaves, nPos1 + 1)
            Else
                sClave = sClaves
            End If
        End If
    Wend
    fnBuscaClave = nPos
End Function


Function sDecodifClave(palabra As String, multiplo As Integer, codigo) As String
    Dim aux     As String
    Dim i       As Integer
    Dim sApp    As String
    
    aux = ""
    sApp = codigo
    While Len(sApp) < Len(palabra)
        sApp = sApp & codigo
    Wend
    
    For i = 1 To Len(palabra)
        aux = aux + Chr(Asc(Mid(palabra, i, 1)) + (multiplo * Asc(Mid(UCase(sApp), i, 1))))
    Next
    
    sDecodifClave = aux
End Function

Sub CargaParametros(sNumConsulta As String, sGlsQuery As String, aRegParametros() As rRegParametros)
    Dim nPosIni         As Integer
    Dim nPosFin         As Integer
    Dim sParametro      As String
    Dim nIndice         As Integer
    Dim bEncontrado     As Boolean
    Dim nX              As Integer
    Dim sTipoDato       As String
    Dim sTipoAyuda      As String
    Dim sGlsAyuda       As String
    Dim sIndOpcional    As String
    Dim nTotParametros  As Integer
    Dim sGlsParametro   As String
    Dim rsData          As ADODB.Recordset
    
    On Error GoTo ErrCargaParametros
    
    ReDim aRegParametros(0) As rRegParametros
    nTotParametros = 0
    
    ' Busca todos los parámetros que estén actualmente en el query (@parametro@)
    nPosIni = InStr(1, sGlsQuery, "@")
    nIndice = nPosIni
    Do While nPosIni > 0
        nPosFin = InStr(nPosIni + 1, sGlsQuery, "@")
        If nPosIni < nPosFin Then
            sParametro = LCase(Mid(sGlsQuery, nPosIni + 1, nPosFin - nPosIni - 1))
            bEncontrado = False
            If nTotParametros > 0 Then
                For nX = 1 To nTotParametros
                    If aRegParametros(nX).Nombre = sParametro Then
                        bEncontrado = True
                    End If
                Next
            End If
            
            If Not bEncontrado Then
                sTipoDato = "Texto"
                sGlsParametro = FormatoTitulo(sParametro)
                sTipoAyuda = ""
                sGlsAyuda = ""
                sIndOpcional = "N"
                
                If sNumConsulta <> "" Then
                    If db_LeeParametro(sNumConsulta, sParametro, rsData) Then
                        If Not rsData.EOF Then
                            sGlsParametro = "" & rsData!gls_parametro
                            sGlsParametro = Replace(sGlsParametro, "#$SignoMenor$#", "<")
                            sTipoDato = "" & rsData!cod_tipo_dato
                            If "" & rsData!gls_ayuda_valores <> "" Then
                                sTipoAyuda = "" & rsData!cod_tipo_ayuda
                                sGlsAyuda = "" & rsData!gls_ayuda_valores
                                sGlsAyuda = Replace(sGlsAyuda, gsSignoMenor, "<")
                            End If
                            sIndOpcional = IIf("" & rsData!ind_opcional = "", "N", "" & rsData!ind_opcional)
                        End If
                    End If
                End If

                nTotParametros = nTotParametros + 1
                ReDim Preserve aRegParametros(nTotParametros) As rRegParametros
                aRegParametros(nTotParametros).Nombre = sParametro
                aRegParametros(nTotParametros).Descripcion = sGlsParametro
                aRegParametros(nTotParametros).Opcional = IIf(sIndOpcional = "S", True, False)
                aRegParametros(nTotParametros).Tipo = sTipoDato
                aRegParametros(nTotParametros).TipoAyuda = sTipoAyuda
                aRegParametros(nTotParametros).Ayuda = sGlsAyuda
            End If
        End If
        
        nIndice = nPosFin + 1
        nPosIni = InStr(nIndice, sGlsQuery, "@")
    Loop
    
    Set rsData = Nothing
    Exit Sub
    
ErrCargaParametros:
    Screen.MousePointer = vbNormal
    MsgBox Error, vbInformation, App.Title
    Exit Sub
End Sub

Function Exist(ByVal FileName As String) As Integer
    Dim nn As Long

    On Error GoTo noExistFile
    If Trim(FileName) <> "" Then
        nn = FileLen(FileName)
        Exist = True
    Else
        Exist = False
    End If
Exit Function

noExistFile:
    Exist = False
    Exit Function

End Function

Public Function FormatoTitulo(sTitulo As String) As String
    Dim nPosBlanco As Integer, nPosAnterior, sResultado As String
    If sTitulo = "" Then
        FormatoTitulo = ""
        Exit Function
    End If
    sResultado = Replace(Trim(sTitulo), "_", " ")
    sResultado = UCase(Mid(sResultado, 1, 1)) & LCase(Mid(sResultado, 2))
    nPosBlanco = InStr(1, sResultado, " ")
    Do While nPosBlanco > 0
        sResultado = Mid(sResultado, 1, nPosBlanco) & UCase(Mid(sResultado, nPosBlanco + 1, 1)) & LCase(Mid(sResultado, nPosBlanco + 2))
        nPosAnterior = nPosBlanco + 1
        nPosBlanco = InStr(nPosAnterior, sResultado, " ")
    Loop
    FormatoTitulo = sResultado
End Function


Function gbComaDecimal() As Boolean
    If CDbl("1,000") > 1 Then
        gbComaDecimal = False
    Else
        gbComaDecimal = True
    End If
End Function

Public Function fsGetGrilla(grdGrilla As fpSpread, nFila As Variant, nColumna As Variant) As String
    Dim nColOld As Long
    Dim nRowOld As Long
    
    nColOld = grdGrilla.Col
    nRowOld = grdGrilla.Row
    
    grdGrilla.Row = nFila
    grdGrilla.Col = nColumna
    fsGetGrilla = grdGrilla.Text
    
    grdGrilla.Col = nColOld
    grdGrilla.Row = nRowOld
End Function

Function FolderRename(sFileOld As String, sFileNew As String) As Boolean
    Dim fso, f, s
    
    On Error GoTo ErrFileRename
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.GetFolder(sFileOld)
    f.Name = sFileNew
    
    FolderRename = True
    Exit Function
    
ErrFileRename:
    FolderRename = False
End Function


Function FileRename(sFileOld As String, sFileNew As String) As Boolean
    Dim fso, f, s
    
    On Error GoTo ErrFileRename
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.GetFile(sFileOld)
    f.Name = sFileNew
    
    FileRename = True
    Exit Function
    
ErrFileRename:
    FileRename = False
End Function
Function fnUltimaPos(sTexto As String, sCaracter As String) As Integer
    Dim nPos    As Integer
    Dim nLast   As Integer
    
    nLast = 0
    
    For nPos = Len(sTexto) To 1 Step -1
        If Mid(sTexto, nPos, 1) = sCaracter Then
            nLast = nPos
            Exit For
        End If
    Next
    
    fnUltimaPos = nLast
End Function
Function fnValorDoble(sNumero As String) As Double
    On Error GoTo ErrValorDoble
    fnValorDoble = CDbl(sNumero)
    Exit Function

ErrValorDoble:
    fnValorDoble = 0
End Function
Function fnValorEntero(sNumero As String) As Long
    On Error GoTo ErrValorEntero
    fnValorEntero = CLng(sNumero)
    Exit Function

ErrValorEntero:
    fnValorEntero = 0
End Function

Function Get_Username() As String
     ' Dimension variables
     Dim lpBuff     As String * 25
     Dim Ret        As Long
     Dim UserName   As String

     ' Get the user name minus any trailing spaces found in the name.
     Ret = GetUserName(lpBuff, 25)
     UserName = Left(lpBuff, InStr(lpBuff, Chr(0)) - 1)

     Get_Username = LCase(UserName)
End Function

