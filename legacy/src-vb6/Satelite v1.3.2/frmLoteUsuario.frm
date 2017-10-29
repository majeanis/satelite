VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmLoteUsuario 
   Caption         =   "Asignación de lotes por usuario"
   ClientHeight    =   7020
   ClientLeft      =   2460
   ClientTop       =   1755
   ClientWidth     =   11055
   Icon            =   "frmLoteUsuario.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7020
   ScaleWidth      =   11055
   Begin VB.Frame fraConsultas 
      Caption         =   "Lotes"
      Height          =   5475
      Left            =   60
      TabIndex        =   14
      Top             =   1140
      Width           =   3435
      Begin VB.ListBox lstLotes 
         Height          =   5010
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   15
         Top             =   300
         Width           =   3195
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Descripción del usuario"
      Height          =   1035
      Left            =   60
      TabIndex        =   9
      Top             =   60
      Width           =   10935
      Begin VB.Label lblTitNomUsuario 
         AutoSize        =   -1  'True
         Caption         =   "Nombre :"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   300
         Width           =   645
      End
      Begin VB.Label lblNomUsuario 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1200
         TabIndex        =   12
         Top             =   300
         Width           =   9015
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Usuario :"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   660
         Width           =   990
      End
      Begin VB.Label lblCodTipoUsuario 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1200
         TabIndex        =   10
         Top             =   660
         Width           =   1395
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Lotes asignados al usuario"
      Height          =   5475
      Left            =   3660
      TabIndex        =   7
      Top             =   1140
      Width           =   7335
      Begin ComctlLib.ListView lvLotes 
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
      Left            =   9840
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
            Picture         =   "frmLoteUsuario.frx":038A
            Key             =   "Blanco"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmLoteUsuario.frx":08AC
            Key             =   "Marcado"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmLoteUsuario"
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
    For nItem = 0 To lstLotes.ListCount - 1
        If lstLotes.Selected(nItem) Then
            sXmlLoteUsuario = sXmlLoteUsuario & "<LoteUsuario "
            sXmlLoteUsuario = sXmlLoteUsuario & " num_lote=""" & lstLotes.ItemData(nItem) & """"
            sXmlLoteUsuario = sXmlLoteUsuario & " nom_usuario=""" & gsNomUsuario & """"
            sXmlLoteUsuario = sXmlLoteUsuario & "/>"
            nItems = nItems + 1
        End If
    Next nItem
    sXmlLoteUsuario = sXmlLoteUsuario & "</ROOT>"
    
    If nItems = 0 Then
        MsgBox "Debe seleccionar al menos un lote para asignar a esta usuario", vbCritical, App.Title
        Exit Sub
    End If
    
    If MsgBox("Está seguro que desea asignar los lotes seleccionados a este usuario", vbYesNoCancel + vbDefaultButton1, App.Title) <> vbYes Then
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
    If Me.lvLotes.ListItems.Count > 0 Then
        mItem.SmallIcon = IIf(mItem.SmallIcon = "Blanco", "Marcado", "Blanco")
        mItem.Selected = False
    End If
End Sub

Sub ArcMarcarTodosAsignar(bSelected As Boolean)
    Dim nItem       As Integer
    Dim nItemOld    As Integer
    
    nItemOld = lstLotes.ListIndex
    
    For nItem = 0 To lstLotes.ListCount - 1
        lstLotes.Selected(nItem) = bSelected
    Next nItem
    
    lstLotes.ListIndex = nItemOld
End Sub

Sub ArcMarcarTodosQuitar(bSelected As Boolean)
    Dim nItem       As Integer
    Dim nItemOld    As Integer
    
    nItemOld = mItem.Index
    
    For nItem = 1 To lvLotes.ListItems.Count
        Set mItem = lvLotes.ListItems.Item(nItem)
        mItem.SmallIcon = IIf(bSelected, "Marcado", "Blanco")
    Next nItem
    
    Set mItem = lvLotes.ListItems.Item(nItemOld)
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
    For nItem = 1 To lvLotes.ListItems.Count
        Set mItem = lvLotes.ListItems.Item(nItem)
        If mItem.SmallIcon = "Marcado" Then
            sXmlLoteUsuario = sXmlLoteUsuario & "<LoteUsuario "
            sXmlLoteUsuario = sXmlLoteUsuario & " num_lote=""" & mItem.Tag & """"
            sXmlLoteUsuario = sXmlLoteUsuario & " nom_usuario=""" & gsNomUsuario & """"
            sXmlLoteUsuario = sXmlLoteUsuario & "/>"
            nItems = nItems + 1
        End If
    Next nItem
    sXmlLoteUsuario = sXmlLoteUsuario & "</ROOT>"
    
    If nItems = 0 Then
        MsgBox "Debe seleccionar al menos un lote para quitar a este usuario", vbCritical, App.Title
        Exit Sub
    End If
    
    If MsgBox("Está seguro que desea quitar los lotes seleccionados a este usuario", vbYesNoCancel + vbDefaultButton1, App.Title) <> vbYes Then
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
    ' Elimina informacion
    bOk = db_EliminaLotesPorUsuario(gsNomUsuario, sXmlLoteUsuario)
    
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
    
    lvLotes.ListItems.Clear
    lstLotes.Clear
    
    OpenMyDataBase
    
    ' Carga lotes que ya tiene el usuario
    If db_LeeLotesPorUsuario(gsNomUsuario, rsData) Then
        While Not rsData.EOF
            Set mItem = lvLotes.ListItems.Add(, , "" & rsData!nom_lote, , "Blanco")
            mItem.SubItems(1) = "" & rsData!fec_creacion
            mItem.Tag = rsData!num_lote
            
            rsData.MoveNext
        Wend
    End If

    ' Carga lotes que NO tiene el usuario
    If db_LeeLotesSinUsuario(gsNomUsuario, rsData) Then
        While Not rsData.EOF
            lstLotes.AddItem "" & rsData!nom_lote & " (Id. " & rsData!num_lote & ")"
            lstLotes.ItemData(lstLotes.ListCount - 1) = rsData!num_lote
            
            rsData.MoveNext
        Wend
    End If

    CloseMyDataBase
    
    Set rsData = Nothing
    Screen.MousePointer = vbNormal

    If lstLotes.ListCount <= 0 Then
        cmdNinguno.Enabled = False
        cmdTodos.Enabled = False
        cmdAsignar.Enabled = False
    Else
        cmdNinguno.Enabled = True
        cmdTodos.Enabled = True
        cmdAsignar.Enabled = True
        lstLotes.ListIndex = 0
    End If

    If Me.lvLotes.ListItems.Count <= 0 Then
        cmdNingunoQuitar.Enabled = False
        cmdTodosQuitar.Enabled = False
        cmdQuitar.Enabled = False
    Else
        cmdNingunoQuitar.Enabled = True
        cmdTodosQuitar.Enabled = True
        cmdQuitar.Enabled = True
        
        Set mItem = lvLotes.ListItems.Item(1)
        mItem.EnsureVisible
        mItem.Selected = True
    End If

    On Error Resume Next
    lstLotes.SetFocus
    Exit Sub

ErrCargaFormulario:
    Screen.MousePointer = vbNormal
    MsgBox Error, vbCritical, App.Title
    Exit Sub
End Sub
Sub IniciaLista()
    Me.lvLotes.ColumnHeaders.Add , , "Lote", 2500
    Me.lvLotes.ColumnHeaders.Add , , "Fec.Asginación", 2000
End Sub

Private Sub lvLotes_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
    lvLotes.SortKey = ColumnHeader.Index - 1
    If lvLotes.SortOrder = lvwAscending Then
        lvLotes.SortOrder = lvwDescending
    Else
        lvLotes.SortOrder = lvwAscending
    End If
    ' Establece Verdadero en Sorted para ordenar la lista.
    lvLotes.Sorted = True
End Sub

Private Sub lvLotes_DblClick()
    Call ArcMarcarUnoQuitar
End Sub

Private Sub lvLotes_ItemClick(ByVal Item As ComctlLib.ListItem)
    Set mItem = Item
End Sub

Private Sub lvLotes_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case 32 ' Marcar o quitar
        Call ArcMarcarUnoQuitar
    End Select
End Sub

