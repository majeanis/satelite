VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmNuevaCarpeta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Crear nueva carpeta"
   ClientHeight    =   4935
   ClientLeft      =   4830
   ClientTop       =   3750
   ClientWidth     =   5175
   Icon            =   "frmNuevaCarpeta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   5175
   Begin VB.PictureBox picConsulta 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   360
      Picture         =   "frmNuevaCarpeta.frx":030A
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   4500
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picCarpeta 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   120
      Picture         =   "frmNuevaCarpeta.frx":0694
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   4500
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picElemento 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   120
      Picture         =   "frmNuevaCarpeta.frx":0BD6
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   420
      Width           =   240
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   4500
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   3960
      TabIndex        =   3
      Top             =   4500
      Width           =   1095
   End
   Begin VB.TextBox txtNombreCarpeta 
      Height          =   315
      Left            =   420
      TabIndex        =   0
      Top             =   360
      Width           =   4635
   End
   Begin ComctlLib.TreeView tvTreeView 
      Height          =   3255
      Left            =   120
      TabIndex        =   1
      Top             =   1140
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   5741
      _Version        =   327682
      HideSelection   =   0   'False
      Indentation     =   176
      LineStyle       =   1
      Style           =   7
      ImageList       =   "Iconos"
      Appearance      =   1
   End
   Begin ComctlLib.ImageList Iconos 
      Left            =   660
      Top             =   4440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   5
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmNuevaCarpeta.frx":1118
            Key             =   "carpeta"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmNuevaCarpeta.frx":166A
            Key             =   "consulta"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmNuevaCarpeta.frx":1BBC
            Key             =   "usuario"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmNuevaCarpeta.frx":210E
            Key             =   "grupal"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmNuevaCarpeta.frx":2660
            Key             =   "area"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblUbicacion 
      AutoSize        =   -1  'True
      Caption         =   "Seleccionar ubicación de la carpeta:"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   2595
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Nombre:"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   600
   End
End
Attribute VB_Name = "frmNuevaCarpeta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mNode           As Node
Dim msRutaActual    As String
Dim mnNumMaxCarpeta   As Long

Sub ActivaAceptar()
    cmdAceptar.Enabled = (txtNombreCarpeta <> "")
End Sub

Function BuscaNodo(sTag As String) As Long
    Dim nItem       As Long
    Dim nItemTag    As Long
    
    nItemTag = 1
    For nItem = 1 To tvTreeView.Nodes.Count
        If tvTreeView.Nodes(nItem).Tag = sTag Then
            nItemTag = nItem
            Exit For
        End If
    Next nItem
    
    BuscaNodo = nItemTag
End Function

Sub CrearCarpeta()
    Dim sPath  As String
    Dim sKey    As String

    If InStr(Me.txtNombreCarpeta.Text, "/") > 0 Or InStr(Me.txtNombreCarpeta.Text, "\") > 0 Then
        MsgBox "Nombre de la carpeta no puede contener el caracter \ o /", vbCritical, App.Title
        Exit Sub
    End If

    On Error GoTo ErrCrearCarpeta
    
    Screen.MousePointer = vbHourglass
    
    sPath = mNode.Key
    sPath = sPath & "\" & txtNombreCarpeta
    
    If Not db_GrabaCarpetaUsuario("", gsNomUsuarioLocal, sPath) Then
        Screen.MousePointer = vbNormal
        Exit Sub
    End If
    
    Screen.MousePointer = vbNormal
    
    gbCancelar = False
    Unload Me
    Exit Sub
    
ErrCrearCarpeta:
    Screen.MousePointer = vbNormal
    MsgBox Error, vbCritical, App.Title
End Sub

Sub MoverArchivoACarpeta()
    Dim sPath           As String
    Dim sKey            As String
    Dim nPos            As Integer
    Dim sNumCarpeta     As String
    Dim sNumConsulta    As String

    On Error GoTo ErrMoverArchivoACarpeta
    
    Screen.MousePointer = vbHourglass
    
    sPath = mNode.Key
    sNumCarpeta = Mid(mNode.Tag, 5)
    
    nPos = InStr(gsTagNodoActual, ";")
    If nPos = 0 Then
        sKey = gsTagNodoActual
    Else
        sKey = Left(gsTagNodoActual, nPos - 1)
    End If
    sNumConsulta = Mid(sKey, 5)
    
    If sNumCarpeta = "" Then
        If Not db_EliminaConsultaEnCarpeta(gsNomUsuarioLocal, sNumConsulta) Then
            Screen.MousePointer = vbNormal
            Exit Sub
        End If
    Else
        If Not db_GrabaConsultaCarpeta(gsNomUsuarioLocal, sNumConsulta, sNumCarpeta) Then
            Screen.MousePointer = vbNormal
            Exit Sub
        End If
    End If
    
    Screen.MousePointer = vbNormal
    
    gbCancelar = False
    Unload Me
    Exit Sub
        
ErrMoverArchivoACarpeta:
    Screen.MousePointer = vbNormal
    MsgBox Error, vbCritical, App.Title
End Sub

Private Sub cmdAceptar_Click()
    Select Case gnObjetoACrear
    Case 1 ' Nueva Carpeta
        CrearCarpeta
        
    Case 2 ' Mover Consulta
        MoverArchivoACarpeta
        
    End Select
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    IniciaForm
    CargaCarpetas
End Sub
Sub IniciaForm()
    Dim nX              As Integer
    
    Me.Left = (Screen.Width - Me.Width) \ 2
    Me.Top = (Screen.Height - Me.Height) \ 2

    Select Case gnObjetoACrear
    Case 1 ' Crear Carpeta
        Me.Caption = "Crear nueva carpeta"
        lblUbicacion.Caption = "Seleccionar ubicación de la carpeta:"
        cmdAceptar.Enabled = False
        picElemento.Picture = picCarpeta.Picture
            
    Case 2 ' Mover consulta a Carpeta
        Me.Caption = "Mover consulta a una carpeta"
        lblUbicacion.Caption = "Mover consulta a la carpeta:"
        cmdAceptar.Enabled = True
        
        If Left(gsTagNodoActual, 3) = "SQL" Then
            picElemento.Picture = picConsulta.Picture
        Else
            picElemento.Picture = picCarpeta.Picture
        End If
        txtNombreCarpeta = gsNombreNodoActual
        txtNombreCarpeta.Enabled = False
    
    End Select
    
    gbCancelar = True
End Sub
Sub CargaCarpetas()
    Dim nIndex  As Long
    
    On Error GoTo ErrCargaCarpetas
    
    Screen.MousePointer = vbHourglass
    
    msRutaActual = gsNombreNodoActual
    
    Set mNode = tvTreeView.Nodes.Add(, , LCase(gsNomUsuarioLocal), LCase(gsNomUsuarioLocal), "usuario")
    mNode.Tag = "USU"
    mNode.Expanded = True

    Call CargaCarpetasUsuario(LCase(gsNomUsuarioLocal), tvTreeView, mnNumMaxCarpeta)
    
    nIndex = BuscaNodo(gsTagNodoActual)
    Set mNode = tvTreeView.Nodes(nIndex)
    mNode.Selected = True
    mNode.EnsureVisible
    
    Screen.MousePointer = vbNormal
    Exit Sub
    
ErrCargaCarpetas:
    Screen.MousePointer = vbNormal
    MsgBox Error, vbCritical, App.Title
End Sub

Private Sub tvTreeView_DblClick()
    If cmdAceptar.Enabled Then
        cmdAceptar_Click
    End If
End Sub

Private Sub tvTreeView_NodeClick(ByVal Node As ComctlLib.Node)
    Set mNode = Node
End Sub


Private Sub txtNombreCarpeta_Change()
    ActivaAceptar
End Sub


Private Sub txtNombreCarpeta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And cmdAceptar.Enabled Then
        cmdAceptar_Click
    End If
End Sub


