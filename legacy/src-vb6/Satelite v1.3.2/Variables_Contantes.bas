Attribute VB_Name = "Variables_Constantes"
Declare Function OSGetPrivateProfileString% Lib "Kernel32" Alias "GetPrivateProfileStringA" (ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal ReturnString$, ByVal NumBytes As Integer, ByVal FileName$)
Declare Function GetWindowsDirectory Lib "Kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Integer) As Integer

Public gsPathConsultas      As String
Public gsPathGrupales       As String
Public gsDirWindows         As String
Public gsNomEditor          As String

Public gsPathSeleccionado   As String
Public gsPathInicial        As String

Public gsUsuarioReal        As String
Public gsNomUsuarioLocal    As String
Public pStr_Archivo         As String

Public Type rRegParametros   ' Defino un tipo para levantar el *.SQL
   Nombre       As String
   Descripcion  As String
   Tipo         As String
   Opcional     As Boolean
   valor        As String
   Ayuda        As String
   TipoAyuda    As String
End Type
Public gaRegParametros()    As rRegParametros

Public Type rRegFormatos
    nom_columna             As String
    cod_tipo_dato_salida    As String
    ind_separador_miles     As String
    num_decimales           As Integer
    gls_formato_entrada     As String
    gls_formato_salida      As String
End Type

'<V1.3.0>
' Asigno constantes a cada una de las vistas, ya que en version 1.2 estaban en duro
Global Const mnVistaUsuarios = 0
Global Const mnVistaConsultas = 1
Global Const mnVistaPerfiles = 2
Global Const mnVistaLotes = 3
Global Const mnVistaTiposUsuarios = 5
Global Const mnVistaBaseDatos = 6
Global Const mnVistaTabValores = 7
'</V1.3.0>

Public Function Limpia(sTexto As String, bUcase As Boolean) As String
    Dim sResultado As String
    Dim nX As Integer
    sResultado = ""
    For nX = 1 To Len(sTexto)
        If Asc(Mid(sTexto, nX, 1)) <> 0 Then
            sResultado = sResultado + Mid(sTexto, nX, 1)
        Else
            Exit For
        End If
    Next
    
    Limpia = sResultado
    
    If bUcase Then
        Limpia = UCase(Limpia)
    End If
End Function

Public Function NombreArchivo(sNombre As String) As String
    Dim sNombreLimpio As String
    Dim nPosicion As Integer
    nPosicion = InStr(1, LCase(sNombre), ".sql")
    sNombreLimpio = Mid(UCase(sNombre), 1, nPosicion - 1)
    NombreArchivo = sNombreLimpio
End Function

Public Sub Envia_Excel(Source As Control)
    Dim xlApp As Excel.Application
    Dim xlBook As Excel.Workbook
    Dim xExcel As Excel.Worksheet
    
    Set xlApp = New Excel.Application
    Set xlBook = xlApp.Workbooks.Add
    Set xExcel = xlBook.Worksheets.Add
    Dim Ancho_columna As String
    Dim xFil%, xCol%
    Dim Titulo$
    Titulo$ = Screen.ActiveForm.Caption
    Ancho_columna = Screen.ActiveForm.TextWidth("O")
    On Error GoTo Control_Errores
    xExcel.Application.Visible = True
    
    xExcel.Cells(2, 3).Font.Bold = True 'Del Titulo general
    xExcel.Cells(2, 3).Font.Size = 13
    xExcel.Cells(2, 3).Font.Italic = True
    xExcel.Cells(2, 3).Value = Titulo$
    xExcel.Cells(2, 3).Font.Color = RGB(0, 0, 230)
        
    xExcel.Rows(4).Font.Bold = True     'Fila de titulo por columna
    xExcel.Rows(4).Interior.Color = RGB(200, 200, 200)
        
    For xCol% = 0 To Source.Cols - 1  'Ancho de las columnas
            xExcel.Columns(xCol% + 1).ColumnWidth = Source.ColWidth(xCol%) / Ancho_columna
    Next xCol%
    
    For xFil% = 0 To Source.Rows - 1
        Source.Row = xFil%
        For xCol% = 0 To Source.Cols - 1
            Source.Col = xCol%
            xExcel.Cells(xFil% + 4, xCol% + 1).Value = Source.Text
        Next xCol%
    Next xFil%
    Exit Sub
Control_Errores:
    Select Case Err.Number
    Case 1004:
        MsgBox "No se ha grabado archivo", vbCritical, App.Title
        Resume Next
    Case Else:
        MsgBox "Error : " & Err.Description, vbCritical, App.Title
        Resume Next
    End Select
    Exit Sub
End Sub

Sub ObtieneVersion()
    Dim sPath           As String
    Dim sDriveLetter    As String
    Dim nPos            As Integer
    Dim fso, d, dc, s
    
    gbVersionLocal = False
    sPath = App.Path
    nPos = InStr(sPath, ":")
    If nPos > 0 Then
        sDriveLetter = LCase(Left(sPath, nPos - 1))
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set dc = fso.Drives
        For Each d In dc
            s = LCase(d.DriveLetter)
            If s = sDriveLetter Then
                If d.DriveType = 2 Then
                    gbVersionLocal = True
                End If
                
                Exit For
            End If
        Next
        Set fso = Nothing
    End If
End Sub


Public Function ObtieneTipoDato() As String
    Dim sTipoDato As String, sFormato As String
    Dim gTotalDato As Integer
    Dim nX As Integer
    sFormato = ""
    gTotalDato = frmPrincipal.grdResultado.Cols - 1
    frmPrincipal.grdResultado.Row = 0
    For nX = 1 To gTotalDato
        frmPrincipal.grdResultado.Col = nX
        Select Case UCase(Mid(frmPrincipal.grdResultado.Text, 1, 3))
        Case "FEC", "PER"
            sTipoDato = "D"
        Case "NUM", "VAL", "ANO", "MES", "DIA", "HOR", "MIN", "MAX"
            sTipoDato = "N"
        Case Else
            sTipoDato = "S"
        End Select
        sFormato = sFormato & sTipoDato & "¬"
    Next
    ObtieneTipoDato = sFormato
End Function
Public Function CadenaNula(sTexto As String) As Boolean
    If sTexto = "" Or sTexto = Empty Then
        CadenaNula = False
    Else
        CadenaNula = True
    End If
End Function
