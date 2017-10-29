VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmAdmConsultas 
   Caption         =   "Administración de Consultas"
   ClientHeight    =   4875
   ClientLeft      =   5985
   ClientTop       =   3150
   ClientWidth     =   4935
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4875
   ScaleWidth      =   4935
   Begin VB.Frame fraConsultas 
      Caption         =   "Consultas"
      Height          =   3315
      Left            =   0
      TabIndex        =   1
      Top             =   660
      Width           =   4875
      Begin ComctlLib.ListView lvConsultas 
         Height          =   2475
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   4035
         _ExtentX        =   7117
         _ExtentY        =   4366
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         _Version        =   327682
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   240
      Top             =   4140
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   327682
   End
   Begin VB.Menu mnuConsultas 
      Caption         =   "&Consultas"
      Begin VB.Menu mnuConNuevaCons 
         Caption         =   "&Nueva consulta"
      End
      Begin VB.Menu mnuConEditarCons 
         Caption         =   "&Editar consulta"
      End
      Begin VB.Menu mnuConEliminarCons 
         Caption         =   "E&liminar consulta"
      End
      Begin VB.Menu mnuConNulo1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuConSalir 
         Caption         =   "&Salir"
      End
   End
   Begin VB.Menu mnuAsignaciones 
      Caption         =   "&Asignaciones"
      Begin VB.Menu mnuAsigUsuarios 
         Caption         =   "&Usuarios"
      End
      Begin VB.Menu mnuAsigPerfiles 
         Caption         =   "&Perfiles"
      End
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuPopNuevaCons 
         Caption         =   "&Nueva consulta"
      End
      Begin VB.Menu mnuPopEditarCons 
         Caption         =   "&Editar consulta"
      End
      Begin VB.Menu mnuPopEliminarCons 
         Caption         =   "E&liminar consulta"
      End
      Begin VB.Menu mnuPopNulo1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopAsigUsuarios 
         Caption         =   "&Usuarios"
      End
      Begin VB.Menu mnuPopAsigPerfiles 
         Caption         =   "&Perfiles"
      End
   End
End
Attribute VB_Name = "frmAdmConsultas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mItem                   As ListItem
Dim msNumUltConsultas       As String
Sub ActivarBotones()
    Dim bOpcConsultas   As Boolean
    
    bOpcConsultas = (lvConsultas.ListItems.Count > 0)
    
    Me.mnuConEditarCons.Enabled = bOpcConsultas
    Me.mnuConEliminarCons.Enabled = bOpcConsultas
    Me.mnuAsigUsuarios.Enabled = bOpcConsultas
    Me.mnuAsigPerfiles.Enabled = bOpcConsultas

    Me.mnuPopEditarCons.Enabled = bOpcConsultas
    Me.mnuPopEliminarCons.Enabled = bOpcConsultas
    Me.mnuPopAsigUsuarios.Enabled = bOpcConsultas
    Me.mnuPopAsigPerfiles.Enabled = bOpcConsultas
End Sub

Sub ArcEditarConsulta()
    On Error GoTo ErrEditarConsulta
    
    If lvConsultas.ListItems.Count > 0 Then
        gsNumConsulta = mItem
        gsNomConsulta = mItem.SubItems(1)
        frmEditarConsulta.Show vbModal
        If Not gbCancelar Then
            msNumUltConsultas = gsNumConsulta
            CargaConsultas
        End If
    End If
    
    Exit Sub
    
ErrEditarConsulta:
    MsgBox Error, vbCritical, App.Title
End Sub

Sub ArcEliminarConsulta()
    Dim sNumConsulta    As String
    Dim sNomConsulta    As String
    Dim bOk             As Boolean
    
    On Error GoTo ErrArcEliminarConsulta
    
    sNumConsulta = mItem
    sNomConsulta = mItem.SubItems(1)
    
    If MsgBox("Está seguro que desea eliminar la consulta """ & sNomConsulta & """ (Id " & sNumConsulta & ")", vbYesNoCancel + vbDefaultButton2, App.Title) <> vbYes Then
        Exit Sub
    End If
    
    Screen.MousePointer = 11
    
    ' Elimina consulta
    OpenMyDataBase
    bOk = db_EliminaConsulta(sNumConsulta)
    CloseMyDataBase
    
    Screen.MousePointer = 0
    
    If bOk Then
        MsgBox "Consulta fue eliminada", vbInformation, App.Title
        msNumUltConsultas = ""
        CargaConsultas
    End If
    
    Exit Sub
    
ErrArcEliminarConsulta:
    Screen.MousePointer = 0
    MsgBox Error, vbCritical, App.Title
    Exit Sub
End Sub

Sub ArcSalir()
    Unload Me
End Sub

Sub AsigPerfiles()

End Sub

Sub AsigUsuarios()
    gsNumConsulta = mItem
    
    frmUsuaConsulta.lblNumConsulta = gsNumConsulta
    frmUsuaConsulta.lblNomConsulta = mItem.SubItems(1)
    frmUsuaConsulta.lblNomCreador = mItem.SubItems(3)
    frmUsuaConsulta.lblFecCreación = mItem.SubItems(4)
    
    frmUsuaConsulta.Show vbModal
End Sub

Sub FormResize()
    fraConsultas.Top = Toolbar1.Top + Toolbar1.Height + 30
    fraConsultas.Left = 60
    fraConsultas.Width = Me.Width - 240
    fraConsultas.Height = Me.Height - fraConsultas.Top - 800
    
    lvConsultas.Top = 240
    lvConsultas.Left = 60
    lvConsultas.Width = fraConsultas.Width - 120
    lvConsultas.Height = fraConsultas.Height - 360
End Sub

Private Sub Form_Load()
    IniciaForm
    IniciaLista
    CargaConsultas
End Sub
Sub IniciaForm()
    Me.WindowState = vbMaximized
    
    msNumUltConsultas = ""
End Sub
Sub IniciaLista()
    Me.lvConsultas.ColumnHeaders.Add , , "Id", 500
    Me.lvConsultas.ColumnHeaders.Add , , "Consulta", 3000
    Me.lvConsultas.ColumnHeaders.Add , , "Base", 1000
    Me.lvConsultas.ColumnHeaders.Add , , "Creado por", 1000
    Me.lvConsultas.ColumnHeaders.Add , , "Fec.Creación", 2000
    Me.lvConsultas.ColumnHeaders.Add , , "Ult.Modificación", 2000
End Sub

Sub CargaConsultas()
    Dim rsData      As ADODB.Recordset
    Dim nItem       As Integer
        
    On Error GoTo ErrCargaConsultas
            
    Screen.MousePointer = 11
    
    lvConsultas.ListItems.Clear
    nItem = 1
    
    ' Abre base datos
    OpenMyDataBase
    
    ' Lee consulta
    If db_LeeConsultas(rsData) Then
        While Not rsData.EOF
            Set mItem = Me.lvConsultas.ListItems.Add(, , rsData!num_consulta)
            mItem.SubItems(1) = "" & rsData!nom_consulta
            mItem.SubItems(2) = "" & rsData!nom_basedatos
            mItem.SubItems(3) = "" & rsData!nom_creador
            mItem.SubItems(4) = "" & rsData!fec_creacion
            mItem.SubItems(5) = "" & rsData!fec_ult_actualizacion
            
            If msNumUltConsultas = rsData!num_consulta Then
                nItem = Me.lvConsultas.ListItems.Count
            End If
            
            rsData.MoveNext
        Wend
    End If
        
    ' Cierra base datos
    CloseMyDataBase
    
    If Me.lvConsultas.ListItems.Count > 0 Then
        Set mItem = lvConsultas.ListItems.Item(nItem)
        mItem.EnsureVisible
        mItem.Selected = True
        On Error Resume Next
        lvConsultas.SetFocus
    End If
    
    ActivarBotones
    
    Screen.MousePointer = 0
    Exit Sub
    
ErrCargaConsultas:
    Screen.MousePointer = 0
    MsgBox Error, vbCritical, App.Title
    Exit Sub
    Resume
End Sub

Private Sub Form_Resize()
    FormResize
End Sub


Private Sub lvConsultas_DblClick()
    If lvConsultas.ListItems.Count > 0 Then
        ArcEditarConsulta
    End If
End Sub

Private Sub lvConsultas_ItemClick(ByVal Item As ComctlLib.ListItem)
    Set mItem = Item
End Sub


Private Sub lvConsultas_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button And vbRightButton Then
        PopupMenu mnuPopup
    End If
End Sub


Private Sub mnuAsigPerfiles_Click()
    AsigPerfiles
End Sub

Private Sub mnuAsigUsuarios_Click()
    AsigUsuarios
End Sub

Private Sub mnuConEditarCons_Click()
    ArcEditarConsulta
End Sub

Sub ArcNuevaConsulta()
    gsNumConsulta = ""
    gsNomConsulta = ""
    frmEditarConsulta.Show vbModal
    If Not gbCancelar Then
        msNumUltConsultas = gsNumConsulta
        CargaConsultas
    End If
End Sub


Private Sub mnuConEliminarCons_Click()
    ArcEliminarConsulta
End Sub

Private Sub mnuConNuevaCons_Click()
    ArcNuevaConsulta
End Sub


Private Sub mnuConSalir_Click()
    ArcSalir
End Sub

Private Sub mnuPopAsigPerfiles_Click()
    AsigPerfiles
End Sub

Private Sub mnuPopAsigUsuarios_Click()
    AsigUsuarios
End Sub

Private Sub mnuPopEditarCons_Click()
    ArcEditarConsulta
End Sub

Private Sub mnuPopEliminarCons_Click()
    ArcEliminarConsulta
End Sub

Private Sub mnuPopNuevaCons_Click()
    ArcNuevaConsulta
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
    Select Case Button
    Case "Nueva"
        ArcNuevaConsulta
    Case "Editar"
        ArcEditarConsulta
    Case "Eliminar"
        ArcEliminarConsulta
    Case "Usuarios"
        AsigUsuarios
    Case "Perfiles"
        AsigPerfiles
    Case "Salir"
        ArcSalir
    End Select
End Sub


