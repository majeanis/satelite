VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "Msmask32.ocx"
Begin VB.Form frmParametros 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Parámetros"
   ClientHeight    =   1860
   ClientLeft      =   4845
   ClientTop       =   2745
   ClientWidth     =   6645
   Icon            =   "frmParametros.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1860
   ScaleWidth      =   6645
   Begin VB.CommandButton cmdEjecutar 
      Caption         =   "&Ejecutar"
      Height          =   375
      Left            =   4260
      TabIndex        =   5
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   5460
      TabIndex        =   6
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Frame fraParametros 
      Caption         =   "Parámetros"
      Height          =   1335
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   6495
      Begin VB.CommandButton cmdHlpFecParametro 
         Caption         =   "..."
         Height          =   315
         Index           =   0
         Left            =   4320
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   660
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CommandButton cmdHlpParametro 
         Caption         =   "..."
         Height          =   315
         Index           =   0
         Left            =   5040
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   300
         Visible         =   0   'False
         Width           =   255
      End
      Begin MSMask.MaskEdBox mskFecParametro 
         Height          =   315
         Index           =   0
         Left            =   3120
         TabIndex        =   1
         Top             =   660
         Visible         =   0   'False
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtValParametro 
         Height          =   315
         Index           =   0
         Left            =   3120
         TabIndex        =   0
         Top             =   300
         Visible         =   0   'False
         Width           =   1875
      End
      Begin VB.Label lblNomFecParametro 
         AutoSize        =   -1  'True
         Caption         =   "Label1"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label lblNomParametro 
         AutoSize        =   -1  'True
         Caption         =   "Nom Parametro : "
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Visible         =   0   'False
         Width           =   1230
      End
   End
End
Attribute VB_Name = "frmParametros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mnNumBaseDatos   As Integer

Function CargaHelpPorLista(ByVal sGlsAyuda As String, ByVal sNomCampo As String) As Boolean
    Dim nPos        As Integer
    Dim sValAyuda   As String

    On Error GoTo ErrCargaHelpPorLista

    Screen.MousePointer = vbHourglass
    
    Set grsLookUp = New ADODB.Recordset
    grsLookUp.CursorLocation = adUseClient
    grsLookUp.CursorType = adOpenStatic
    grsLookUp.ActiveConnection = Nothing
    grsLookUp.LockType = adLockBatchOptimistic
    
    Call grsLookUp.Fields.Append(LCase(sNomCampo), adChar, 100, adFldIsNullable)
    
    grsLookUp.Open
    
    nPos = InStr(sGlsAyuda, ",")
    While nPos > 0
        sValAyuda = Left(sGlsAyuda, nPos - 1)
        sGlsAyuda = Mid(sGlsAyuda, nPos + 1)
        
        grsLookUp.AddNew
        grsLookUp.Fields(0).Value = sValAyuda
        grsLookUp.MoveLast
        
        nPos = InStr(sGlsAyuda, ",")
    Wend
    
    If sGlsAyuda <> "" Then
        grsLookUp.AddNew
        grsLookUp.Fields(0).Value = sGlsAyuda
        grsLookUp.MoveFirst
    End If
    
    CargaHelpPorLista = True
    
    Screen.MousePointer = vbNormal
    Exit Function
    
ErrCargaHelpPorLista:
    Screen.MousePointer = vbNormal
    MsgBox "Error al cargar ayuda de valores para el parámetro " & sNomCampo & " : " & Error, vbCritical, App.Title
    CargaHelpPorLista = False
End Function

Sub ReemplazaParametros(sNomCampo As String, sGlsAyuda As String)
    Dim nX  As Integer
    
    For nX = 1 To UBound(gaRegParametros)
        If LCase(gaRegParametros(nX).Nombre) <> LCase(sNomCampo) Then
            sGlsAyuda = Replace(LCase(sGlsAyuda), "@" & UCase(gaRegParametros(nX).Nombre) & "@", gaRegParametros(nX).valor)
            sGlsAyuda = Replace(LCase(sGlsAyuda), "@" & LCase(gaRegParametros(nX).Nombre) & "@", gaRegParametros(nX).valor)
        End If
    Next nX
End Sub

Private Sub cmdCancelar_Click()
    gbCancelar = True
    Unload Me
End Sub

Private Sub cmdEjecutar_Click()
    If GuardaParametros Then
        Unload Me
    End If
End Sub

Function GuardaParametros() As Boolean
    Dim nRow            As Integer
    Dim nCta            As Integer
    Dim nTotal          As Integer
    Dim nTotCamposTexto As Integer
    Dim nTotCamposFecha As Integer
    
    nCta = 0
    nTotal = 0
    nTotCamposTexto = -1
    nTotCamposFecha = -1
    
    For nRow = 1 To UBound(gaRegParametros)
        Select Case UCase(gaRegParametros(nRow).Tipo)
            Case "FECHA"
                nTotCamposFecha = nTotCamposFecha + 1
                gaRegParametros(nRow).valor = mskFecParametro(nTotCamposFecha).Text
                If fdValorFecha(gaRegParametros(nRow).valor) = gdNullDate Then
                    MsgBox gaRegParametros(nRow).valor & " no es una fecha válida. Por favor ingrese nuevamente el valor para el parámetro """ & FormatoTitulo(gaRegParametros(nRow).Nombre) & """", vbCritical, App.Title
                    GuardaParametros = False
                    Exit Function
                End If
            Case "ENTERO"
                nTotCamposTexto = nTotCamposTexto + 1
                gaRegParametros(nRow).valor = txtValParametro(nTotCamposTexto).Text
                If Not EsEntero(gaRegParametros(nRow).valor) Then
                    MsgBox gaRegParametros(nRow).valor & " no es un número entero válido. Por favor ingrese nuevamente el valor para el parámetro """ & FormatoTitulo(gaRegParametros(nRow).Nombre) & """", vbCritical, App.Title
                    GuardaParametros = False
                    Exit Function
                End If
            Case "DECIMAL"
                nTotCamposTexto = nTotCamposTexto + 1
                gaRegParametros(nRow).valor = txtValParametro(nTotCamposTexto).Text
                If Not EsDecimal(gaRegParametros(nRow).valor) Then
                    MsgBox gaRegParametros(nRow).valor & " no es un número decimal válido. Por favor ingrese nuevamente el valor para el parámetro """ & FormatoTitulo(gaRegParametros(nRow).Nombre) & """", vbCritical, App.Title
                    GuardaParametros = False
                    Exit Function
                End If
            Case Else
                nTotCamposTexto = nTotCamposTexto + 1
                gaRegParametros(nRow).valor = txtValParametro(nTotCamposTexto).Text
        End Select
        
        If gaRegParametros(nRow).Opcional = False Then
            nTotal = nTotal + 1
            If gaRegParametros(nRow).valor <> "" Then
                nCta = nCta + 1
            End If
        End If
    Next
    
    If nCta < nTotal Then
        MsgBox "Debe ingresar todos los parámetros antes de ejecutar esta consulta", vbCritical, App.Title
        GuardaParametros = False
        Exit Function
    End If
    
    GuardaParametros = True
End Function


Private Sub cmdHlpFecParametro_Click(Index As Integer)
    If HelpFecha(Me, mskFecParametro(Index)) Then
        mskFecParametro(Index).SetFocus
    End If
End Sub

Private Sub cmdHlpParametro_Click(Index As Integer)
    Dim sGlosaCampos    As String
    Dim sCampos         As String
    Dim sGlosaWhere     As String
    Dim sGlosaOrder     As String
    Dim nVal            As Integer
    Dim nX              As Integer

    gsQueryLookUp = ""
    sCampos = ""
    sGlosaWhere = ""
    sGlosaOrder = ""

    nX = txtValParametro(Index).Tag
    
    ' Carga recordset con Ayuda de Valores
    If CargaHelp(gaRegParametros(nX).Nombre, gaRegParametros(nX).TipoAyuda, gaRegParametros(nX).Ayuda) Then
        ' Posiciona el formulario de ayuda
        frmListaOpciones.Top = Me.Top + cmdHlpParametro(Index).Top + 330
        frmListaOpciones.Left = Me.Left + cmdHlpParametro(Index).Left + cmdHlpParametro(Index).Width + 150
        frmListaOpciones.Caption = frmListaOpciones.Caption & lblNomParametro(Index).Caption
        
        gsCampoLookUp = gaRegParametros(nX).Nombre
        frmListaOpciones.Show vbModal
    
        grsLookUp.Close
        Set grsLookUp = Nothing
    
        If Not gbCancelar Then
            txtValParametro(Index) = gsResultLookUp
        End If
        gbCancelar = False
        txtValParametro(Index).SetFocus
    End If
End Sub


Function CargaHelpPorQuery(sGlsAyuda As String, sNomCampo As String) As Boolean
    Dim nPos        As Integer
    Dim sValAyuda   As String

    On Error GoTo ErrCargaHelpPorQuery

    Screen.MousePointer = vbHourglass
    
    If Not ConectaBaseDatos(mnNumBaseDatos) Then
        Screen.MousePointer = vbNormal
        CargaHelpPorQuery = False
        Exit Function
    Else
        Set grsLookUp = New ADODB.Recordset
        grsLookUp.CursorLocation = adUseClient
        grsLookUp.Open sGlsAyuda, cnn_Consulta, adOpenForwardOnly, adLockReadOnly
    End If
    
    CargaHelpPorQuery = True
    
    Screen.MousePointer = vbNormal
    Exit Function
    
ErrCargaHelpPorQuery:
    Screen.MousePointer = vbNormal
    MsgBox "Error al cargar ayuda de valores para el parámetro " & sNomCampo & " : " & Error, vbCritical, App.Title
    CargaHelpPorQuery = False
End Function


Function CargaHelp(ByVal sNomCampo As String, ByVal sTipoAyuda As String, ByVal sGlsAyuda As String) As Boolean
    Dim sCodTipo    As String
    Dim nPos        As Integer
    Dim sValAyuda   As String
    Dim sGlsAux     As String
    Dim nX          As Integer
    Dim sCodCampo   As String
    Dim sValCampo   As String
    
    Select Case LCase(sTipoAyuda)
    Case "query"
        If sGlsAyuda <> "" Then
            Call ReemplazaParametros(sNomCampo, sGlsAyuda)
            CargaHelp = CargaHelpPorQuery(sGlsAyuda, sNomCampo)
        Else
            CargaHelp = True
        End If
        
    Case "list"
        If sGlsAyuda <> "" Then
            CargaHelp = CargaHelpPorLista(sGlsAyuda, sNomCampo)
        Else
            CargaHelp = True
        End If
    End Select
End Function

Private Sub Form_Load()
    '<V1.3.0>
    ' Se muestra el nombre de la consulta en el caption
    Me.Caption = gsNomConsulta 'Me.Caption & " - " & App.CompanyName
    If gbEjecutandoLote Then
        Me.cmdEjecutar.Caption = "&Aceptar"
    Else
        Me.cmdEjecutar.Caption = "&Ejecutar"
    End If
    '</V1.3.0>
    gbCancelar = False
    CargaParametros
End Sub
Sub CargaParametros()
    Dim nX              As Integer
    Dim sNom            As String
    Dim nTotCamposTexto As Integer
    Dim nTotCamposFecha As Integer
    Dim nAlto           As Single
    Dim nPos            As Integer
    
    nTotCamposTexto = -1
    nTotCamposFecha = -1
    
    If UBound(gaRegParametros) > 0 Then
        For nX = 1 To UBound(gaRegParametros)
            sNom = gaRegParametros(nX).Descripcion
            nPos = InStr(sNom, "(")
            If nPos = 0 Then
                sNom = FormatoTitulo(sNom)
            Else
                sNom = FormatoTitulo(Left(sNom, nPos - 1)) & " " & Mid(sNom, nPos)
            End If
            
            Select Case UCase(gaRegParametros(nX).Tipo)
                Case "FECHA"
                    nTotCamposFecha = nTotCamposFecha + 1
                    If nTotCamposFecha > 0 Then
                        Load lblNomFecParametro(nTotCamposFecha)
                        Load mskFecParametro(nTotCamposFecha)
                        Load cmdHlpFecParametro(nTotCamposFecha)
                    End If
                    lblNomFecParametro(nTotCamposFecha).Caption = sNom
                    lblNomFecParametro(nTotCamposFecha).Top = 300 + ((nX - 1) * 360)
                    mskFecParametro(nTotCamposFecha).Top = 300 + ((nX - 1) * 360)
                    cmdHlpFecParametro(nTotCamposFecha).Top = 300 + ((nX - 1) * 360)
                    mskFecParametro(nTotCamposFecha).TabIndex = nX - 1
                    
                    lblNomFecParametro(nTotCamposFecha).Visible = True
                    mskFecParametro(nTotCamposFecha).Visible = True
                    cmdHlpFecParametro(nTotCamposFecha).Visible = True
                    
                    mskFecParametro(nTotCamposFecha).Tag = nX
                    Call SetMasked(mskFecParametro(nTotCamposFecha), gaRegParametros(nX).valor)
                
                Case "USERNAME"
                    nTotCamposTexto = nTotCamposTexto + 1
                    If nTotCamposTexto > 0 Then
                        Load lblNomParametro(nTotCamposTexto)
                        Load txtValParametro(nTotCamposTexto)
                    End If
                    lblNomParametro(nTotCamposTexto).Caption = sNom
                    lblNomParametro(nTotCamposTexto).Top = 300 + ((nX - 1) * 360)
                    txtValParametro(nTotCamposTexto).Top = 300 + ((nX - 1) * 360)
                    txtValParametro(nTotCamposTexto).TabIndex = nX - 1

                    lblNomParametro(nTotCamposTexto).Visible = True
                    txtValParametro(nTotCamposTexto).Visible = True
                    txtValParametro(nTotCamposTexto).Enabled = False

                    txtValParametro(nTotCamposTexto).Tag = nX
                    txtValParametro(nTotCamposTexto) = gaRegParametros(nX).valor
                
                Case Else
                    nTotCamposTexto = nTotCamposTexto + 1
                    If nTotCamposTexto > 0 Then
                        Load lblNomParametro(nTotCamposTexto)
                        Load txtValParametro(nTotCamposTexto)
                    End If
                    lblNomParametro(nTotCamposTexto).Caption = sNom
                    lblNomParametro(nTotCamposTexto).Top = 300 + ((nX - 1) * 360)
                    txtValParametro(nTotCamposTexto).Top = 300 + ((nX - 1) * 360)
                    txtValParametro(nTotCamposTexto).TabIndex = nX - 1

                    lblNomParametro(nTotCamposTexto).Visible = True
                    txtValParametro(nTotCamposTexto).Visible = True
                    txtValParametro(nTotCamposTexto).Enabled = True

                    txtValParametro(nTotCamposTexto).Tag = nX
                    txtValParametro(nTotCamposTexto) = gaRegParametros(nX).valor
                    
                    If gaRegParametros(nX).Ayuda <> "" Then
                        If nTotCamposTexto > 0 Then
                            Load cmdHlpParametro(nTotCamposTexto)
                        End If
                        cmdHlpParametro(nTotCamposTexto).Top = 300 + ((nX - 1) * 360)
                        cmdHlpParametro(nTotCamposTexto).Visible = True
                    End If
            
            End Select
            
        Next
    End If

    fraParametros.Height = 300 + (UBound(gaRegParametros) * 360) + 120
    Me.cmdCancelar.Top = fraParametros.Height + 60
    Me.cmdEjecutar.Top = fraParametros.Height + 60
    
    nAlto = cmdCancelar.Top + cmdCancelar.Height + 450
    Me.Height = nAlto
    Me.Left = (Screen.Width - Me.Width) \ 3
    Me.Top = (Screen.Height - nAlto) \ 3
End Sub

Private Sub mskFecParametro_Change(Index As Integer)
    Dim nX  As Integer
    
    nX = mskFecParametro(Index).Tag
    gaRegParametros(nX).valor = mskFecParametro(Index).Text
End Sub

Private Sub mskFecParametro_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 And cmdEjecutar.Enabled Then
        cmdEjecutar_Click
    End If
End Sub


Private Sub txtValParametro_Change(Index As Integer)
    Dim nX  As Integer
    
    nX = txtValParametro(Index).Tag
    gaRegParametros(nX).valor = txtValParametro(Index)
End Sub


Private Sub txtValParametro_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 And cmdEjecutar.Enabled Then
        cmdEjecutar_Click
    End If
End Sub


