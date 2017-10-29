VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmConsPerfil 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asignación de consultas por agrupación"
   ClientHeight    =   7095
   ClientLeft      =   1620
   ClientTop       =   765
   ClientWidth     =   11025
   Icon            =   "frmConsPerfil.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   11025
   Begin VB.CommandButton cmdQuitar 
      Caption         =   "<< &Quitar"
      Height          =   315
      Left            =   3660
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   6660
      Width           =   1095
   End
   Begin VB.CommandButton cmdTodosQuitar 
      Caption         =   "T&odos"
      Height          =   315
      Left            =   4800
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   6660
      Width           =   1095
   End
   Begin VB.CommandButton cmdNingunoQuitar 
      Caption         =   "N&inguno"
      Height          =   315
      Left            =   5940
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   6660
      Width           =   1095
   End
   Begin VB.CommandButton cmdNinguno 
      Caption         =   "&Ninguno"
      Height          =   315
      Left            =   1200
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   6660
      Width           =   1095
   End
   Begin VB.CommandButton cmdTodos 
      Caption         =   "&Todos"
      Height          =   315
      Left            =   60
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   6660
      Width           =   1095
   End
   Begin VB.CommandButton cmdAsignar 
      Caption         =   "&Asignar >>"
      Height          =   315
      Left            =   2340
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   6660
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Salir"
      Height          =   315
      Left            =   9840
      TabIndex        =   2
      Top             =   6660
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Caption         =   "Consultas asignadas a la agrupación"
      Height          =   5475
      Left            =   3660
      TabIndex        =   7
      Top             =   1140
      Width           =   7335
      Begin ComctlLib.ListView lvConsultas 
         Height          =   5055
         Left            =   120
         TabIndex        =   1
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
   Begin VB.Frame Frame1 
      Caption         =   "Descripción de la agrupación"
      Height          =   1035
      Left            =   60
      TabIndex        =   4
      Top             =   60
      Width           =   10935
      Begin VB.Label lblTitNumConsulta 
         AutoSize        =   -1  'True
         Caption         =   "Id :"
         Height          =   195
         Left            =   540
         TabIndex        =   17
         Top             =   660
         Width           =   225
      End
      Begin VB.Label lblNumPerfil 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   840
         TabIndex        =   16
         Top             =   660
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha creación :"
         Height          =   195
         Left            =   2160
         TabIndex        =   15
         Top             =   660
         Width           =   1200
      End
      Begin VB.Label lblFecCreacion 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3420
         TabIndex        =   14
         Top             =   660
         Width           =   2175
      End
      Begin VB.Label lblNomPerfil 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   840
         TabIndex        =   6
         Top             =   300
         Width           =   9375
      End
      Begin VB.Label lblTitNomUsuario 
         AutoSize        =   -1  'True
         Caption         =   "Nombre :"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   300
         Width           =   645
      End
   End
   Begin VB.Frame fraConsultas 
      Caption         =   "Consultas"
      Height          =   5475
      Left            =   60
      TabIndex        =   3
      Top             =   1140
      Width           =   3435
      Begin VB.ListBox lstConsultas 
         Height          =   5010
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   0
         Top             =   300
         Width           =   3195
      End
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
            Picture         =   "frmConsPerfil.frx":038A
            Key             =   "Blanco"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmConsPerfil.frx":08AC
            Key             =   "Marcado"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmConsPerfil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mItem                   As ListItem

Sub ArcGrabarAsignar()
    Dim nItem           As Long
    Dim sXmlConsPerfil  As String
    Dim bOk             As Boolean
    Dim nItems          As Long
    
    On Error GoTo ErrArcGrabarAsignar
    
    nItems = 0
    
    ' Arma xml para inserción masiva
    sXmlConsPerfil = "<ROOT>"
    For nItem = 0 To lstConsultas.ListCount - 1
        If lstConsultas.Selected(nItem) Then
            sXmlConsPerfil = sXmlConsPerfil & "<ConsPerfil "
            sXmlConsPerfil = sXmlConsPerfil & " num_perfil=""" & gsNumPerfil & """"
            sXmlConsPerfil = sXmlConsPerfil & " num_consulta=""" & lstConsultas.ItemData(nItem) & """"
            sXmlConsPerfil = sXmlConsPerfil & "/>"
            nItems = nItems + 1
        End If
    Next nItem
    sXmlConsPerfil = sXmlConsPerfil & "</ROOT>"
    
    If nItems = 0 Then
        MsgBox "Debe seleccionar al menos una consulta para asignar a esta agrupación", vbCritical, App.Title
        Exit Sub
    End If
    
    If MsgBox("Está seguro que desea asignar las consultas seleccionadas a esta agrupación", vbYesNoCancel + vbDefaultButton1, App.Title) <> vbYes Then
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
    ' Graba informacion
    bOk = db_GrabaConsultasPorAgrupacion(sXmlConsPerfil)
    
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
    If Me.lvConsultas.ListItems.Count > 0 Then
        mItem.SmallIcon = IIf(mItem.SmallIcon = "Blanco", "Marcado", "Blanco")
        mItem.Selected = False
    End If
End Sub

Sub ArcMarcarTodosAsignar(bSelected As Boolean)
    Dim nItem       As Integer
    Dim nItemOld    As Integer
    
    nItemOld = lstConsultas.ListIndex
    
    For nItem = 0 To lstConsultas.ListCount - 1
        lstConsultas.Selected(nItem) = bSelected
    Next nItem
    
    lstConsultas.ListIndex = nItemOld
End Sub

Sub ArcMarcarTodosQuitar(bSelected As Boolean)
    Dim nItem       As Integer
    Dim nItemOld    As Integer
    
    nItemOld = mItem.Index
    
    For nItem = 1 To lvConsultas.ListItems.Count
        Set mItem = lvConsultas.ListItems.Item(nItem)
        mItem.SmallIcon = IIf(bSelected, "Marcado", "Blanco")
    Next nItem
    
    Set mItem = lvConsultas.ListItems.Item(nItemOld)
    mItem.EnsureVisible
End Sub

Sub ArcGrabarQuitar()
    Dim nItem           As Long
    Dim sXmlConsPerfil  As String
    Dim bOk             As Boolean
    Dim nItems          As Long
    
    On Error GoTo ErrArcGrabarQuitar
    
    nItems = 0
    
    ' Arma xml para inserción masiva
    sXmlConsPerfil = "<ROOT>"
    For nItem = 1 To lvConsultas.ListItems.Count
        Set mItem = lvConsultas.ListItems.Item(nItem)
        If mItem.SmallIcon = "Marcado" Then
            sXmlConsPerfil = sXmlConsPerfil & "<ConsPerfil "
            sXmlConsPerfil = sXmlConsPerfil & " num_perfil=""" & gsNumPerfil & """"
            sXmlConsPerfil = sXmlConsPerfil & " num_consulta=""" & mItem.Tag & """"
            sXmlConsPerfil = sXmlConsPerfil & "/>"
            nItems = nItems + 1
        End If
    Next nItem
    sXmlConsPerfil = sXmlConsPerfil & "</ROOT>"
    
    If nItems = 0 Then
        MsgBox "Debe seleccionar al menos una consulta para quitar a esta agrupación", vbCritical, App.Title
        Exit Sub
    End If
    
    If MsgBox("Está seguro que desea quitar las consultas seleccionadas a esta agrupación", vbYesNoCancel + vbDefaultButton1, App.Title) <> vbYes Then
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
    ' Elimina informacion
    bOk = db_EliminaConsultasPorAgrupacion(gsNumPerfil, sXmlConsPerfil)
    
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
    
    lvConsultas.ListItems.Clear
    lstConsultas.Clear
    
    OpenMyDataBase
    
    ' Carga consultas que ya tienen el pefil
    If db_LeeConsultasPorAgrupacion(gsNumPerfil, rsData) Then
        While Not rsData.EOF
            Set mItem = lvConsultas.ListItems.Add(, , "" & rsData!nom_consulta, , "Blanco")
            mItem.SubItems(1) = "" & rsData!fec_creacion
            mItem.Tag = rsData!num_consulta
            
            rsData.MoveNext
        Wend
    End If

    ' Carga Consultas que no estan asociadas al pefil
    If db_LeeConsultasSinAgrupacion(gsNumPerfil, rsData) Then
        While Not rsData.EOF
            lstConsultas.AddItem "" & rsData!nom_consulta & " (Id. " & rsData!num_consulta & ")"
            lstConsultas.ItemData(lstConsultas.ListCount - 1) = rsData!num_consulta
            
            rsData.MoveNext
        Wend
    End If

    CloseMyDataBase
    
    Set rsData = Nothing
    Screen.MousePointer = vbNormal

    If lstConsultas.ListCount <= 0 Then
        cmdNinguno.Enabled = False
        cmdTodos.Enabled = False
        cmdAsignar.Enabled = False
    Else
        cmdNinguno.Enabled = True
        cmdTodos.Enabled = True
        cmdAsignar.Enabled = True
        lstConsultas.ListIndex = 0
    End If

    If Me.lvConsultas.ListItems.Count <= 0 Then
        cmdNingunoQuitar.Enabled = False
        cmdTodosQuitar.Enabled = False
        cmdQuitar.Enabled = False
    Else
        cmdNingunoQuitar.Enabled = True
        cmdTodosQuitar.Enabled = True
        cmdQuitar.Enabled = True
        
        Set mItem = lvConsultas.ListItems.Item(1)
        mItem.EnsureVisible
        mItem.Selected = True
    End If

    On Error Resume Next
    lstConsultas.SetFocus
    Exit Sub

ErrCargaFormulario:
    Screen.MousePointer = vbNormal
    MsgBox Error, vbCritical, App.Title
    Exit Sub
End Sub
Sub IniciaLista()
    Me.lvConsultas.ColumnHeaders.Add , , "Consulta", 2500
    Me.lvConsultas.ColumnHeaders.Add , , "Fec.Asginación", 2000
End Sub

Private Sub lvConsultas_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
    lvConsultas.SortKey = ColumnHeader.Index - 1
    If lvConsultas.SortOrder = lvwAscending Then
        lvConsultas.SortOrder = lvwDescending
    Else
        lvConsultas.SortOrder = lvwAscending
    End If
    ' Establece Verdadero en Sorted para ordenar la lista.
    lvConsultas.Sorted = True
End Sub

Private Sub lvConsultas_DblClick()
    Call ArcMarcarUnoQuitar
End Sub

Private Sub lvConsultas_ItemClick(ByVal Item As ComctlLib.ListItem)
    Set mItem = Item
End Sub

Private Sub lvConsultas_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case 32 ' Marcar o quitar
        Call ArcMarcarUnoQuitar
    End Select
End Sub



