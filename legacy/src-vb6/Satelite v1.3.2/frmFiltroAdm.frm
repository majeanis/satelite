VERSION 5.00
Begin VB.Form frmFiltroAdm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Filtro personalizado"
   ClientHeight    =   2775
   ClientLeft      =   3450
   ClientTop       =   2925
   ClientWidth     =   8655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   8655
   Begin VB.Frame fraColumna 
      Caption         =   "Mostrar las filas en las cuales:"
      Height          =   2295
      Left            =   60
      TabIndex        =   10
      Top             =   0
      Width           =   8535
      Begin VB.ComboBox cboGlsFiltro 
         Height          =   315
         Index           =   1
         Left            =   4920
         TabIndex        =   7
         Top             =   1200
         Width           =   3495
      End
      Begin VB.ComboBox cboGlsFiltro 
         Height          =   315
         Index           =   0
         Left            =   4920
         TabIndex        =   2
         Top             =   420
         Width           =   3495
      End
      Begin VB.ComboBox cboSigno 
         Height          =   315
         Index           =   1
         ItemData        =   "frmFiltroAdm.frx":0000
         Left            =   4140
         List            =   "frmFiltroAdm.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1200
         Width           =   735
      End
      Begin VB.ComboBox cboColumna 
         Height          =   315
         Index           =   1
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1200
         Width           =   3795
      End
      Begin VB.OptionButton optO 
         Caption         =   "&O"
         Height          =   255
         Left            =   1680
         TabIndex        =   4
         Top             =   840
         Width           =   615
      End
      Begin VB.OptionButton optY 
         Caption         =   "&Y"
         Height          =   255
         Left            =   900
         TabIndex        =   3
         Top             =   840
         Value           =   -1  'True
         Width           =   615
      End
      Begin VB.ComboBox cboSigno 
         Height          =   315
         Index           =   0
         ItemData        =   "frmFiltroAdm.frx":0004
         Left            =   4140
         List            =   "frmFiltroAdm.frx":0006
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   420
         Width           =   735
      End
      Begin VB.ComboBox cboColumna 
         Height          =   315
         Index           =   0
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   420
         Width           =   3795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Use * para representar cualquier serie de caracteres en comparación de tipo Like"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   1920
         Width           =   5730
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   6300
      TabIndex        =   8
      Top             =   2340
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   7500
      TabIndex        =   9
      Top             =   2340
      Width           =   1095
   End
End
Attribute VB_Name = "frmFiltroAdm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub ActivaControles(nIndice As Integer)
    Dim bEstado     As Boolean
    
    Select Case nIndice
    Case 0
        bEstado = (cboColumna(nIndice).ListIndex >= 0)
        
        cboColumna(1).Enabled = bEstado And (Trim(cboGlsFiltro(0).Text) <> "")
        cboSigno(1).Enabled = bEstado And (Trim(cboGlsFiltro(0).Text) <> "")
        cboGlsFiltro(1).Enabled = bEstado And (Trim(cboGlsFiltro(0).Text) <> "")
    Case 1
        If cboColumna(nIndice).ListIndex <= 0 Then
            cboGlsFiltro(nIndice).Text = ""
        End If
    End Select

    cmdAceptar.Enabled = (cboColumna(0).ListIndex >= 0 And Trim(cboGlsFiltro(0).Text) <> "")
    If (cboColumna(1).ListIndex > 0 And Trim(cboGlsFiltro(1).Text) = "") Then
        cmdAceptar.Enabled = False
    End If
End Sub

Sub GrabarFiltro()
    Dim i           As Integer
    Dim j           As Integer
    Dim sGlsFiltro  As String
    Dim sGlsValor   As String
    
    For i = 1 To frmAdministracion.lvConsultas.ColumnHeaders.Count
        frmAdministracion.lvConsultas.ColumnHeaders(i).Tag = ""
    Next i
    
    i = cboColumna(0).ListIndex
    j = cboColumna(1).ListIndex
    
    If i >= 0 Then
        sGlsValor = Replace(cboGlsFiltro(0), "'", "''")
        If sGlsValor = gsGlsValorBlanco Then sGlsValor = ""

        frmAdministracion.lvConsultas.ColumnHeaders(i + 1).Tag = cboSigno(0).Text & sGlsValor
    End If
    
    If j > 0 Then
        sGlsValor = Replace(cboGlsFiltro(1), "'", "''")
        If sGlsValor = gsGlsValorBlanco Then sGlsValor = ""
        
        sGlsFiltro = cboSigno(1).Text & sGlsValor
        If j = i + 1 Then
            frmAdministracion.lvConsultas.ColumnHeaders(j).Tag = frmAdministracion.lvConsultas.ColumnHeaders(j).Tag & Chr(9) & sGlsFiltro
        Else
            frmAdministracion.lvConsultas.ColumnHeaders(j).Tag = sGlsFiltro
        End If
    End If
    
    If optY.Value = True Then
        gsGlsOperadorFiltro = " and "
    Else
        gsGlsOperadorFiltro = " or "
    End If
    
    gbCancelar = False
    Unload Me
End Sub

Private Sub cboColumna_Click(Index As Integer)
    Call CargaValores(Index, cboColumna(Index).ListIndex)
    Call ActivaControles(Index)
End Sub

Private Sub cboGlsFiltro_Click(Index As Integer)
    Call ActivaControles(Index)
End Sub

Private Sub cmdAceptar_Click()
    GrabarFiltro
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    IniciaFormulario
End Sub

Sub IniciaFormulario()
    Dim i   As Integer

    Me.Left = (Screen.Width - Me.Width) \ 2
    Me.Top = (Screen.Height - Me.Height) \ 3

    Me.cboSigno(0).AddItem " =  "
    Me.cboSigno(0).AddItem " <> "
    Me.cboSigno(0).AddItem " <  "
    Me.cboSigno(0).AddItem " <= "
    Me.cboSigno(0).AddItem " >  "
    Me.cboSigno(0).AddItem " >= "
    Me.cboSigno(0).AddItem "Like"

    Me.cboSigno(1).AddItem " =  "
    Me.cboSigno(1).AddItem " <> "
    Me.cboSigno(1).AddItem " <  "
    Me.cboSigno(1).AddItem " <= "
    Me.cboSigno(1).AddItem " >  "
    Me.cboSigno(1).AddItem " >= "
    Me.cboSigno(1).AddItem "Like"

    cboSigno(0).ListIndex = 0
    cboSigno(1).ListIndex = 0
    
    Me.cboColumna(1).AddItem ""
    
    For i = 1 To frmAdministracion.lvConsultas.ColumnHeaders.Count
        cboColumna(0).AddItem frmAdministracion.lvConsultas.ColumnHeaders(i)
        cboColumna(1).AddItem frmAdministracion.lvConsultas.ColumnHeaders(i)
    Next i
    
    cmdAceptar.Enabled = False
    cboColumna(0).ListIndex = frmAdministracion.mnNumColumnaActiva
    cboGlsFiltro(0).Text = frmAdministracion.msValColumnaActiva
    'txtGlsFiltro(0).SelStart = 0
    'txtGlsFiltro(0).SelLength = Len(txtGlsFiltro(0).Text)
    'txtGlsFiltro(0).SelText = txtGlsFiltro(0).Text
    'cboColumna(1).Enabled = False
    'cboSigno(1).Enabled = False
    'cboGlsFiltro(1).Enabled = False
End Sub

Private Sub cboGlsFiltro_Change(Index As Integer)
    Call ActivaControles(Index)
End Sub


Private Sub cboGlsFiltro_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 And cmdAceptar.Enabled Then
        cmdAceptar_Click
    End If
End Sub

Sub CargaValores(nIndice As Integer, nNumColumna As Long)
    Dim i           As Long
    Dim sNomCampo   As String
    
    cboGlsFiltro(nIndice).Clear
    frmAdministracion.rsDataFiltro.Filter = ""
    
    sNomCampo = frmAdministracion.lvConsultas.ColumnHeaders(nNumColumna + 1).Key
    frmAdministracion.rsDataFiltro.Filter = "nom_campo='" & sNomCampo & "'"
    frmAdministracion.rsDataFiltro.Sort = "gls_valor"
    While Not frmAdministracion.rsDataFiltro.EOF
        cboGlsFiltro(nIndice).AddItem IIf("" & frmAdministracion.rsDataFiltro!gls_valor = "", gsGlsValorBlanco, "" & frmAdministracion.rsDataFiltro!gls_valor)
        frmAdministracion.rsDataFiltro.MoveNext
    Wend
End Sub
