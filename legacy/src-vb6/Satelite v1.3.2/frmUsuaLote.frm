VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmUsuaLote 
   Caption         =   "Asignación de usuarios por lote"
   ClientHeight    =   7035
   ClientLeft      =   3495
   ClientTop       =   3285
   ClientWidth     =   11055
   Icon            =   "frmUsuaLote.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7035
   ScaleWidth      =   11055
   Begin VB.Frame fraUsuarios 
      Caption         =   "Usuarios"
      Height          =   5475
      Left            =   60
      TabIndex        =   16
      Top             =   1140
      Width           =   3435
      Begin VB.ListBox lstUsuarios 
         Height          =   5010
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   17
         Top             =   300
         Width           =   3195
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Descripción del lote"
      Height          =   1035
      Left            =   60
      TabIndex        =   9
      Top             =   60
      Width           =   10935
      Begin VB.Label lblNomSolicitante 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3240
         TabIndex        =   19
         Top             =   660
         Width           =   1755
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Solicitado por :"
         Height          =   195
         Left            =   2100
         TabIndex        =   18
         Top             =   660
         Width           =   1050
      End
      Begin VB.Label lblTitNumConsulta 
         AutoSize        =   -1  'True
         Caption         =   "Id :"
         Height          =   195
         Left            =   540
         TabIndex        =   15
         Top             =   660
         Width           =   225
      End
      Begin VB.Label lblTitNomConsulta 
         AutoSize        =   -1  'True
         Caption         =   "Nombre :"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   300
         Width           =   645
      End
      Begin VB.Label lblNomLote 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   840
         TabIndex        =   13
         Top             =   300
         Width           =   9375
      End
      Begin VB.Label lblNumLote 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   840
         TabIndex        =   12
         Top             =   660
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha creación :"
         Height          =   195
         Left            =   5940
         TabIndex        =   11
         Top             =   660
         Width           =   1200
      End
      Begin VB.Label lblFecCreación 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   7200
         TabIndex        =   10
         Top             =   660
         Width           =   2175
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Usuarios asignados al lote"
      Height          =   5475
      Left            =   3660
      TabIndex        =   7
      Top             =   1140
      Width           =   7335
      Begin ComctlLib.ListView lvUsuarios 
         Height          =   5055
         Left            =   120
         TabIndex        =   8
         Top             =   300
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   8916
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         _Version        =   327682
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Salir"
      Height          =   315
      Left            =   9900
      TabIndex        =   6
      Top             =   6660
      Width           =   1095
   End
   Begin VB.CommandButton cmdAsignar 
      Caption         =   "&Asignar >>"
      Height          =   315
      Left            =   2340
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   6660
      Width           =   1095
   End
   Begin VB.CommandButton cmdTodos 
      Caption         =   "&Todos"
      Height          =   315
      Left            =   60
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   6660
      Width           =   1095
   End
   Begin VB.CommandButton cmdNinguno 
      Caption         =   "&Ninguno"
      Height          =   315
      Left            =   1200
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   6660
      Width           =   1095
   End
   Begin VB.CommandButton cmdNingunoQuitar 
      Caption         =   "N&inguno"
      Height          =   315
      Left            =   5940
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   6660
      Width           =   1095
   End
   Begin VB.CommandButton cmdTodosQuitar 
      Caption         =   "T&odos"
      Height          =   315
      Left            =   4800
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   6660
      Width           =   1095
   End
   Begin VB.CommandButton cmdQuitar 
      Caption         =   "<< &Quitar"
      Height          =   315
      Left            =   3660
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   6660
      Width           =   1095
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   1320
      Top             =   6960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   13
      ImageHeight     =   13
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   2
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmUsuaLote.frx":0442
            Key             =   "Blanco"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmUsuaLote.frx":0964
            Key             =   "Marcado"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmUsuaLote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mItem                   As ListItem

Sub ArcGrabarAsignar()
    Dim nItem           As Long
    Dim sXmlLoteUsuario As String
    Dim bOk             As Boolean
    Dim nItems          As Long
    
    On Error GoTo ErrArcGrabarAsignar
    
    nItems = 0
    
    ' Arma xml para inserción masiva
    sXmlLoteUsuario = "<ROOT>"
    For nItem = 0 To lstUsuarios.ListCount - 1
        If lstUsuarios.Selected(nItem) Then
            sXmlLoteUsuario = sXmlLoteUsuario & "<LoteUsuario"
            sXmlLoteUsuario = sXmlLoteUsuario & " num_lote=""" & gsNumLote & """"
            sXmlLoteUsuario = sXmlLoteUsuario & " nom_usuario=""" & lstUsuarios.List(nItem) & """"
            sXmlLoteUsuario = sXmlLoteUsuario & "/>"
            nItems = nItems + 1
        End If
    Next nItem
    sXmlLoteUsuario = sXmlLoteUsuario & "</ROOT>"
    
    If nItems = 0 Then
        MsgBox "Debe seleccionar al menos un usuario para asignar a este lote", vbCritical, App.Title
        Exit Sub
    End If
    
    If MsgBox("Está seguro que desea asignar los usuarios seleccionados a este lote", vbYesNoCancel + vbDefaultButton1, App.Title) <> vbYes Then
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
    ' Graba informacion
    bOk = db_GrabaUsuariosPorLote(sXmlLoteUsuario)
    
    ' Carga nuevamente el formulario con los datos actualizados
    If bOk Then
        CargaFormulario
    End If
    
    Screen.MousePointer = vbNormal
    Exit Sub
    
ErrArcGrabarAsignar:
    Screen.MousePointer = vbNormal
    MsgBox Error, vbCritical, App.Title
    Exit Sub
End Sub

Sub ArcMarcarUnoQuitar()
    If Me.lvUsuarios.ListItems.Count > 0 Then
        mItem.SmallIcon = IIf(mItem.SmallIcon = "Blanco", "Marcado", "Blanco")
        mItem.Selected = False
    End If
End Sub

Sub ArcMarcarTodosAsignar(bSelected As Boolean)
    Dim nItem       As Integer
    Dim nItemOld    As Integer
    
    nItemOld = lstUsuarios.ListIndex
    
    For nItem = 0 To lstUsuarios.ListCount - 1
        lstUsuarios.Selected(nItem) = bSelected
    Next nItem
    
    lstUsuarios.ListIndex = nItemOld
End Sub

Sub ArcMarcarTodosQuitar(bSelected As Boolean)
    Dim nItem       As Integer
    Dim nItemOld    As Integer
    
    nItemOld = mItem.Index
    
    For nItem = 1 To lvUsuarios.ListItems.Count
        Set mItem = lvUsuarios.ListItems.Item(nItem)
        mItem.SmallIcon = IIf(bSelected, "Marcado", "Blanco")
    Next nItem
    
    Set mItem = lvUsuarios.ListItems.Item(nItemOld)
    mItem.EnsureVisible
End Sub

Sub ArcGrabarQuitar()
    Dim nItem           As Long
    Dim sXmlLoteUsuario As String
    Dim bOk             As Boolean
    Dim nItems          As Long
    
    On Error GoTo ErrArcGrabarQuitar
    
    nItems = 0
    
    ' Arma xml para inserción masiva
    sXmlLoteUsuario = "<ROOT>"
    For nItem = 1 To lvUsuarios.ListItems.Count
        Set mItem = lvUsuarios.ListItems.Item(nItem)
        If mItem.SmallIcon = "Marcado" Then
            sXmlLoteUsuario = sXmlLoteUsuario & "<LoteUsuario "
            sXmlLoteUsuario = sXmlLoteUsuario & " num_lote=""" & gsNumLote & """"
            sXmlLoteUsuario = sXmlLoteUsuario & " nom_usuario=""" & mItem.Text & """"
            sXmlLoteUsuario = sXmlLoteUsuario & "/>"
            nItems = nItems + 1
        End If
    Next nItem
    sXmlLoteUsuario = sXmlLoteUsuario & "</ROOT>"
    
    If nItems = 0 Then
        MsgBox "Debe seleccionar al menos un usuario para quitar a este lote", vbCritical, App.Title
        Exit Sub
    End If
    
    If MsgBox("Está seguro que desea quitar los usuarios seleccionados a este lote", vbYesNoCancel + vbDefaultButton1, App.Title) <> vbYes Then
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
    ' Elimina informacion
    bOk = db_EliminaUsuariosPorLote(gsNumLote, sXmlLoteUsuario)
    
    ' Carga nuevamente el formulario con los datos actualizados
    If bOk Then
        CargaFormulario
    End If
    
    Screen.MousePointer = vbNormal
    Exit Sub
    
ErrArcGrabarQuitar:
    Screen.MousePointer = vbNormal
    MsgBox Error, vbCritical, App.Title
    Exit Sub
End Sub

Private Sub cmdAsignar_Click()
    ArcGrabarAsignar
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdNinguno_Click()
    Call ArcMarcarTodosAsignar(False)
End Sub

Private Sub cmdNingunoQuitar_Click()
    Call ArcMarcarTodosQuitar(False)
End Sub

Private Sub cmdQuitar_Click()
    ArcGrabarQuitar
End Sub

Private Sub cmdTodos_Click()
    Call ArcMarcarTodosAsignar(True)
End Sub

Private Sub cmdTodosQuitar_Click()
    Call ArcMarcarTodosQuitar(True)
End Sub

Private Sub Form_Load()
    IniciaForm
    IniciaLista
    CargaFormulario
End Sub

Sub IniciaForm()
    Me.Left = (Screen.Width - Me.Width) \ 2
    Me.Top = (Screen.Height - Me.Height) \ 2
End Sub

Sub CargaFormulario()
    Dim rsData  As ADODB.Recordset
    
    On Error GoTo ErrCargaFormulario
    
    Screen.MousePointer = vbHourglass
    
    lvUsuarios.ListItems.Clear
    lstUsuarios.Clear
    
    OpenMyDataBase
    
    ' Carga usuarios que ya tienen el lote asignado
    If db_LeeUsuariosPorLote(gsNumLote, rsData) Then
        While Not rsData.EOF
            Set mItem = lvUsuarios.ListItems.Add(, , "" & rsData!nom_usuario, , "Blanco")
            mItem.SubItems(1) = "" & rsData!fec_creacion
            
            rsData.MoveNext
        Wend
    End If

    ' Carga usuarios que NO tienen la consulta asignada
    If db_LeeUsuariosSinLote(gsNumLote, rsData) Then
        While Not rsData.EOF
            lstUsuarios.AddItem "" & rsData!nom_usuario
            
            rsData.MoveNext
        Wend
    End If

    CloseMyDataBase
    
    Set rsData = Nothing
    Screen.MousePointer = vbNormal

    If lstUsuarios.ListCount <= 0 Then
        cmdNinguno.Enabled = False
        cmdTodos.Enabled = False
        cmdAsignar.Enabled = False
    Else
        cmdNinguno.Enabled = True
        cmdTodos.Enabled = True
        cmdAsignar.Enabled = True
        lstUsuarios.ListIndex = 0
    End If

    If Me.lvUsuarios.ListItems.Count <= 0 Then
        cmdNingunoQuitar.Enabled = False
        cmdTodosQuitar.Enabled = False
        cmdQuitar.Enabled = False
    Else
        cmdNingunoQuitar.Enabled = True
        cmdTodosQuitar.Enabled = True
        cmdQuitar.Enabled = True
        
        Set mItem = lvUsuarios.ListItems.Item(1)
        mItem.EnsureVisible
        mItem.Selected = True
    End If

    On Error Resume Next
    lstUsuarios.SetFocus
    Exit Sub

ErrCargaFormulario:
    Screen.MousePointer = vbNormal
    MsgBox Error, vbCritical, App.Title
    Exit Sub
End Sub

Sub IniciaLista()
    Me.lvUsuarios.ColumnHeaders.Add , , "Usuario", 2000
    Me.lvUsuarios.ColumnHeaders.Add , , "Fec.Asginación", 2000
End Sub

Private Sub lvUsuarios_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
    lvUsuarios.SortKey = ColumnHeader.Index - 1
    If lvUsuarios.SortOrder = lvwAscending Then
        lvUsuarios.SortOrder = lvwDescending
    Else
        lvUsuarios.SortOrder = lvwAscending
    End If
    ' Establece Verdadero en Sorted para ordenar la lista.
    lvUsuarios.Sorted = True
End Sub

Private Sub lvUsuarios_DblClick()
    Call ArcMarcarUnoQuitar
End Sub

Private Sub lvUsuarios_ItemClick(ByVal Item As ComctlLib.ListItem)
    Set mItem = Item
End Sub

Private Sub lvUsuarios_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case 32 ' Marcar o quitar
        Call ArcMarcarUnoQuitar
    End Select
End Sub

