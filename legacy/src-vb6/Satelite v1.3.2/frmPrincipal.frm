VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmPrincipal 
   Caption         =   "Consulta"
   ClientHeight    =   6195
   ClientLeft      =   2955
   ClientTop       =   1935
   ClientWidth     =   10350
   Icon            =   "frmPrincipal.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6195
   ScaleWidth      =   10350
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picSplitter 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      FillColor       =   &H00808080&
      Height          =   4800
      Left            =   3480
      ScaleHeight     =   2090.126
      ScaleMode       =   0  'User
      ScaleWidth      =   780
      TabIndex        =   0
      Top             =   180
      Visible         =   0   'False
      Width           =   72
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   1500
      TabIndex        =   7
      Top             =   5520
      Visible         =   0   'False
      Width           =   5475
      _ExtentX        =   9657
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   6
      Top             =   5880
      Width           =   10350
      _ExtentX        =   18256
      _ExtentY        =   556
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   5
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   7461
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "11:14 a.m."
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "14/04/2015"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraConsultas 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5190
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   3120
      Begin ComctlLib.TreeView tvTreeView 
         Height          =   3255
         Left            =   120
         TabIndex        =   5
         Top             =   180
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   5741
         _Version        =   327682
         HideSelection   =   0   'False
         Indentation     =   176
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         ImageList       =   "Iconos"
         Appearance      =   1
      End
   End
   Begin VB.Frame fraGrillas 
      Height          =   3375
      Left            =   3840
      TabIndex        =   1
      Top             =   0
      Width           =   5235
      Begin FPSpreadADO.fpSpread grdResultado 
         Height          =   1335
         Left            =   120
         TabIndex        =   2
         Top             =   180
         Width           =   1995
         _Version        =   524288
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SpreadDesigner  =   "frmPrincipal.frx":038A
      End
      Begin VB.TextBox txtResultado 
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   420
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   8
         Text            =   "frmPrincipal.frx":0774
         Top             =   1920
         Width           =   3555
      End
   End
   Begin ComctlLib.ImageList Iconos 
      Left            =   4320
      Top             =   3840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   8
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmPrincipal.frx":077A
            Key             =   "carpeta"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmPrincipal.frx":0CCC
            Key             =   "consulta1"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmPrincipal.frx":121E
            Key             =   "usuario"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmPrincipal.frx":1770
            Key             =   "grupal"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmPrincipal.frx":1CC2
            Key             =   "area"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmPrincipal.frx":2214
            Key             =   "consulta"
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmPrincipal.frx":2766
            Key             =   "lote"
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmPrincipal.frx":2CB8
            Key             =   "cons_bloqueada"
         EndProperty
      EndProperty
   End
   Begin VB.OLE OLE1 
      Class           =   "Excel.Sheet.8"
      Height          =   390
      Left            =   5820
      OleObjectBlob   =   "frmPrincipal.frx":320A
      TabIndex        =   4
      Top             =   3720
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Image imgSplitter 
      Height          =   4785
      Left            =   3240
      MousePointer    =   9  'Size W E
      Top             =   60
      Width           =   150
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "&Archivo"
      Begin VB.Menu mnuArcNueva 
         Caption         =   "&Nueva ventana"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuArcActualizar 
         Caption         =   "Actualizar ventana"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuArcNulo1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuArcNuevaCarpeta 
         Caption         =   "Nueva carpeta"
      End
      Begin VB.Menu mnuArcEliminarCarpeta 
         Caption         =   "Eliminar carpeta"
      End
      Begin VB.Menu mnuArcNulo2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuArcNuevaConsulta 
         Caption         =   "&Crear consulta"
      End
      Begin VB.Menu mnuArcEditarConsulta 
         Caption         =   "&Editar consulta"
      End
      Begin VB.Menu mnuArcMoverConsulta 
         Caption         =   "&Mover consulta"
      End
      Begin VB.Menu mnuArcNulo3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuArcAdministracion 
         Caption         =   "Módulo de &Administración"
      End
      Begin VB.Menu mnuArcNulo4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuArcCerrarVentana 
         Caption         =   "&Cerrar Ventana"
      End
      Begin VB.Menu mnuArcSalir 
         Caption         =   "&Salir"
      End
   End
   Begin VB.Menu mnuEditar 
      Caption         =   "&Editar"
      Begin VB.Menu mnuEdiBuscarTexto 
         Caption         =   "&Buscar texto"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuEdiBuscarSgte 
         Caption         =   "Buscar &siguiente"
         Shortcut        =   {F3}
      End
   End
   Begin VB.Menu mnuConsulta 
      Caption         =   "&Consulta"
      Begin VB.Menu mnuArcEjecutar 
         Caption         =   "&Ejecutar consulta"
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnuConNulo1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuConResEnGrila 
         Caption         =   "Resultado en &Grilla"
         Checked         =   -1  'True
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuConResEnTexto 
         Caption         =   "Resultado en &Texto"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuConNulo2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuArcExportar 
         Caption         =   "Exportar a E&xcel"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuArcExpArchivo 
         Caption         =   "Exportar a &Archivo..."
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "Ve&ntana"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWindowCascade 
         Caption         =   "&Cascada"
      End
      Begin VB.Menu mnuWindowVert 
         Caption         =   "&Vertical"
      End
      Begin VB.Menu mnuWindowHort 
         Caption         =   "&Horizontal"
      End
   End
   Begin VB.Menu mnuPopupRes 
      Caption         =   "Popup Resultado"
      Visible         =   0   'False
      Begin VB.Menu mnuPopResExcel 
         Caption         =   "Exportar a &Excel"
      End
      Begin VB.Menu mnuPopResTexto 
         Caption         =   "Exportar a &Archivo"
      End
   End
   Begin VB.Menu mnuPopConsultas 
      Caption         =   "Popup Consultas"
      Visible         =   0   'False
      Begin VB.Menu mnuPopEdiEjecutar 
         Caption         =   "&Ejecutar consulta"
      End
      Begin VB.Menu mnuPopEdiEditar 
         Caption         =   "E&ditar consulta"
      End
      Begin VB.Menu mnuPopConNuevaConsulta 
         Caption         =   "&Crear consulta"
      End
      Begin VB.Menu mnuPopConNulo1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopMoverConsulta 
         Caption         =   "&Mover consulta"
      End
   End
   Begin VB.Menu mnuPopCarpetas 
      Caption         =   "Popup Carpetas"
      Visible         =   0   'False
      Begin VB.Menu mnuPopNuevaCarpeta 
         Caption         =   "&Nueva carpeta"
      End
      Begin VB.Menu mnuPopEliminarCarpeta 
         Caption         =   "&Eliminar carpeta"
      End
      Begin VB.Menu mnuPopNulo1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopCarNuevaConsulta 
         Caption         =   "&Crear consulta"
      End
   End
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public msNomUsuarioLocal    As String
Public msGlsQuery           As String
Public msNumConsulta        As String
Public mnNumBaseDatos       As Integer
Public msGlsHorario         As String

Dim maRegParametros()       As rRegParametros
Dim mrsUsuarioLocal         As ADODB.Recordset

Dim nPosHor1                As Integer
Dim bRedimensiona           As Boolean
Dim grsConsulta             As ADODB.Recordset
Dim grsFormatos             As ADODB.Recordset
Dim nLimiteVert             As Integer
Dim nLimiteHort             As Integer
Dim mbMoving                As Boolean
Dim mNode                   As Node
Dim msTextoDeBusqueda       As String
Dim mnTotalRegUltQuery      As Long
Dim msLastTag               As String
Sub ActivarBotones()
    Dim nCtaInput   As Integer
    '<V1.3.1>
    Dim bBloqueada  As Boolean
    '</V1.3.1>
    
    On Error GoTo ErrActivarBotones
        
    Select Case Left(mNode.Tag, 3)
    Case "USU", "COM", "PER", "LCO"
        frmMdiPadre.Toolbar1(1).Buttons(1).Enabled = False
        frmMdiPadre.Toolbar1(1).Buttons(2).Enabled = False
        frmMdiPadre.Toolbar1(1).Buttons(3).Enabled = False
        
        mnuArcNuevaConsulta.Enabled = ((Left(mNode.Tag, 3) = "USU") And ("" & mrsUsuarioLocal!ind_crear_consultas = "S"))
        mnuPopCarNuevaConsulta.Enabled = ((Left(mNode.Tag, 3) = "USU") And ("" & mrsUsuarioLocal!ind_crear_consultas = "S"))
        mnuPopConNuevaConsulta.Enabled = ((Left(mNode.Tag, 3) = "USU") And ("" & mrsUsuarioLocal!ind_crear_consultas = "S"))
        mnuArcEditarConsulta.Enabled = False
        mnuPopEdiEditar.Enabled = False
        mnuArcMoverConsulta.Enabled = False
        mnuPopMoverConsulta.Enabled = False
        
        mnuArcNuevaCarpeta.Enabled = (msNomUsuarioLocal = msNomUsuarioLocal) And (Left(mNode.Tag, 3) = "USU")
        mnuPopNuevaCarpeta.Enabled = (msNomUsuarioLocal = msNomUsuarioLocal) And (Left(mNode.Tag, 3) = "USU")
        mnuArcEliminarCarpeta.Enabled = False
        mnuPopEliminarCarpeta.Enabled = False
        
        mnuEdiBuscarTexto.Enabled = False
        mnuEdiBuscarSgte.Enabled = False
        mnuArcEjecutar.Enabled = False
        mnuPopEdiEjecutar.Enabled = False
        mnuConResEnGrila.Enabled = False
        mnuConResEnTexto.Enabled = False
        mnuArcExportar.Enabled = False
        mnuArcExpArchivo.Enabled = False
    
    Case "DIR"
        frmMdiPadre.Toolbar1(1).Buttons(1).Enabled = False
        frmMdiPadre.Toolbar1(1).Buttons(2).Enabled = False
        frmMdiPadre.Toolbar1(1).Buttons(3).Enabled = False
        
        mnuArcNuevaConsulta.Enabled = ("" & mrsUsuarioLocal!ind_crear_consultas = "S")
        mnuPopCarNuevaConsulta.Enabled = ("" & mrsUsuarioLocal!ind_crear_consultas = "S")
        mnuPopConNuevaConsulta.Enabled = ("" & mrsUsuarioLocal!ind_crear_consultas = "S")
        mnuArcEditarConsulta.Enabled = False
        mnuPopEdiEditar.Enabled = False
        mnuArcMoverConsulta.Enabled = False
        mnuPopMoverConsulta.Enabled = False
        
        mnuArcNuevaCarpeta.Enabled = (msNomUsuarioLocal = msNomUsuarioLocal)
        mnuPopNuevaCarpeta.Enabled = (msNomUsuarioLocal = msNomUsuarioLocal)
        mnuArcEliminarCarpeta.Enabled = (msNomUsuarioLocal = msNomUsuarioLocal)
        mnuPopEliminarCarpeta.Enabled = (msNomUsuarioLocal = msNomUsuarioLocal)
        
        mnuEdiBuscarTexto.Enabled = False
        mnuEdiBuscarSgte.Enabled = False
        mnuArcEjecutar.Enabled = False
        mnuPopEdiEjecutar.Enabled = False
        mnuConResEnGrila.Enabled = False
        mnuConResEnTexto.Enabled = False
        mnuArcExportar.Enabled = False
        mnuArcExpArchivo.Enabled = False
    
    '<V1.3.0>
    Case "LOT"
        frmMdiPadre.Toolbar1(1).Buttons(1).Enabled = False
        frmMdiPadre.Toolbar1(1).Buttons(2).Enabled = False
        frmMdiPadre.Toolbar1(1).Buttons(3).Enabled = False
        
        mnuArcNuevaConsulta.Enabled = False
        mnuPopCarNuevaConsulta.Enabled = False
        mnuPopConNuevaConsulta.Enabled = False
        mnuArcEditarConsulta.Enabled = False
        mnuPopEdiEditar.Enabled = False
        mnuArcMoverConsulta.Enabled = False
        mnuPopMoverConsulta.Enabled = False
        
        mnuArcNuevaCarpeta.Enabled = False
        mnuPopNuevaCarpeta.Enabled = False
        mnuArcEliminarCarpeta.Enabled = False
        mnuPopEliminarCarpeta.Enabled = False
        
        mnuEdiBuscarTexto.Enabled = False
        mnuEdiBuscarSgte.Enabled = False
        mnuArcEjecutar.Enabled = True
        mnuPopEdiEjecutar.Enabled = False
        mnuConResEnGrila.Enabled = False
        mnuConResEnTexto.Enabled = False
        mnuArcExportar.Enabled = False
        mnuArcExpArchivo.Enabled = False
        frmMdiPadre.Toolbar1(1).Buttons(1).Enabled = mnuArcEjecutar.Enabled
    '</V1.3.0>
    
    Case Else
        '<V1.3.1>
        bBloqueada = (mNode.Image = "cons_bloqueada")
        '</V1.3.1>
        
        If Left(mNode.Parent.Tag, 3) = "PER" Then
            mnuArcNuevaConsulta.Enabled = False
            mnuPopCarNuevaConsulta.Enabled = False
            mnuPopConNuevaConsulta.Enabled = False
            mnuArcEditarConsulta.Enabled = False
            mnuPopEdiEditar.Enabled = False
            mnuArcMoverConsulta.Enabled = False
            mnuPopMoverConsulta.Enabled = False
        Else
            '<V1.3.1>
            mnuArcNuevaConsulta.Enabled = ("" & mrsUsuarioLocal!ind_crear_consultas = "S")
            mnuPopCarNuevaConsulta.Enabled = ("" & mrsUsuarioLocal!ind_crear_consultas = "S")
            mnuPopConNuevaConsulta.Enabled = ("" & mrsUsuarioLocal!ind_crear_consultas = "S")
            mnuArcEditarConsulta.Enabled = ("" & mrsUsuarioLocal!ind_modificar_consultas = "S") And Not bBloqueada
            mnuPopEdiEditar.Enabled = ("" & mrsUsuarioLocal!ind_modificar_consultas = "S") And Not bBloqueada
            mnuArcMoverConsulta.Enabled = (msNomUsuarioLocal = msNomUsuarioLocal) And Not bBloqueada
            mnuPopMoverConsulta.Enabled = (msNomUsuarioLocal = msNomUsuarioLocal) And Not bBloqueada
            '</V1.3.1>
        End If
    
        mnuEdiBuscarTexto.Enabled = (mnTotalRegUltQuery > 0)
        mnuEdiBuscarSgte.Enabled = False
        mnuArcEjecutar.Enabled = ("" & mrsUsuarioLocal!ind_ejecutar_consultas = "S") And Not bBloqueada
        
        mnuArcNuevaCarpeta.Enabled = False
        mnuPopNuevaCarpeta.Enabled = False
        mnuArcEliminarCarpeta.Enabled = False
        mnuPopEliminarCarpeta.Enabled = False
        
        mnuPopEdiEjecutar.Enabled = ("" & mrsUsuarioLocal!ind_ejecutar_consultas = "S") And Not bBloqueada
        mnuConResEnGrila.Enabled = True
        mnuConResEnTexto.Enabled = True
    
        mnuArcExportar.Enabled = (mnuConResEnGrila.Checked And mnTotalRegUltQuery > 0)
        mnuArcExpArchivo.Enabled = (mnTotalRegUltQuery > 0)
    
        frmMdiPadre.Toolbar1(1).Buttons(1).Enabled = mnuArcEjecutar.Enabled
        frmMdiPadre.Toolbar1(1).Buttons(2).Enabled = mnuArcExportar.Enabled
        
        '<V1.3.0>
        Call CargaParametrosDefault(maRegParametros(), nCtaInput)
        frmMdiPadre.Toolbar1(1).Buttons(3).Enabled = mnuArcEjecutar.Enabled And (nCtaInput > 0)
        '</V1.3.0>
    
    End Select
    
    mnuArcAdministracion.Enabled = ("" & mrsUsuarioLocal!ind_administrador = "S")
    
    Exit Sub
    
ErrActivarBotones:
    frmMdiPadre.Toolbar1(1).Buttons(1).Enabled = False
    frmMdiPadre.Toolbar1(1).Buttons(2).Enabled = False
    frmMdiPadre.Toolbar1(1).Buttons(3).Enabled = False
End Sub

Sub ArcAdministracion()
    frmAdministracion.Show
End Sub

Sub ArcCerrarVentana()
    Unload Me
End Sub

Sub ArcEliminarCarpeta()
    Dim sMensaje    As String
    Dim sFolderIni  As String
    
    If Left(mNode.Tag, 3) = "DIR" Then
        sFolderIni = mNode.Key
    
        sMensaje = "Al eliminar esta carpeta, se eliminarán todos las carpetas internas, y las consultas que pertenecen a ellas "
        sMensaje = sMensaje & " serán trasladadas al nodo raíz (" & msNomUsuarioLocal & ")" & Chr(13) & Chr(10)
        sMensaje = sMensaje & "Está seguro que desea eliminar la carpeta """ & mNode & """"
        
        If MsgBox(sMensaje, vbYesNoCancel + vbDefaultButton2, App.Title) <> vbYes Then
            Exit Sub
        End If
    
        Screen.MousePointer = vbHourglass
        If db_EliminaCarpetaUsuario(msNomUsuarioLocal, sFolderIni) Then
            ArcActualizarVentana
        End If
        Screen.MousePointer = vbNormal
    End If
End Sub

Sub ArcMoverConsulta()
    gsTagNodoActual = mNode.Tag
    gsNombreNodoActual = mNode.Text
    
    gnObjetoACrear = 2
    frmNuevaCarpeta.Show vbModal
    If Not gbCancelar Then
        ArcActualizarVentana
    End If
End Sub

Sub ArcNuevaConsulta()
    gsNumConsulta = ""
    gsNomConsulta = ""
    gsNomUsuarioLocal = msNomUsuarioLocal
    frmEditarConsulta.Show vbModal
    If Not gbCancelar Then
        ArcActualizarVentana
    End If
End Sub

Sub ArcSalir()
    End
End Sub

Sub BuscarSiguiente()
    Dim sTexto      As String
    Dim nFila       As Long
    Dim nColumna    As Long
    
    sTexto = msTextoDeBusqueda
    
    Screen.MousePointer = vbHourglass
    For nFila = grdResultado.Row + 1 To Me.grdResultado.MaxRows
        For nColumna = 1 To Me.grdResultado.MaxCols
            If InStr(LCase(fsGetGrilla(grdResultado, nFila, nColumna)), LCase(sTexto)) > 0 Then
                grdResultado.Row = nFila
                grdResultado.Col = nColumna
                grdResultado.Action = 0
                Screen.MousePointer = vbNormal
                grdResultado.SetFocus
                msTextoDeBusqueda = sTexto
                Me.mnuEdiBuscarSgte.Enabled = True
                Exit Sub
            End If
        Next nColumna
    Next nFila
    Screen.MousePointer = vbNormal
    
    MsgBox "No hay mas registros con el texto ingresado", vbInformation, App.Title
    Me.mnuEdiBuscarSgte.Enabled = False
    msTextoDeBusqueda = ""
End Sub

Sub BuscarTexto()
    Dim sTexto      As String
    Dim nFila       As Long
    Dim nColumna    As Long
    
    sTexto = InputBox("Ingrese texto que desea buscar")
    If Trim(sTexto) = "" Then
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    For nFila = 1 To Me.grdResultado.MaxRows
        For nColumna = 1 To Me.grdResultado.MaxCols
            If InStr(LCase(fsGetGrilla(grdResultado, nFila, nColumna)), LCase(sTexto)) > 0 Then
                grdResultado.Row = nFila
                grdResultado.Col = nColumna
                grdResultado.Action = 0
                Screen.MousePointer = vbNormal
                grdResultado.SetFocus
                msTextoDeBusqueda = sTexto
                Me.mnuEdiBuscarSgte.Enabled = True
                Exit Sub
            End If
        Next nColumna
    Next nFila
    Screen.MousePointer = vbNormal
    
    MsgBox "No se encontró el texto ingresado", vbInformation, App.Title
    Me.mnuEdiBuscarSgte.Enabled = False
    msTextoDeBusqueda = ""
End Sub

Sub CargaConsultas(sNomUsuario As String)
    Dim sNumConsulta    As String
    Dim sNomConsulta    As String
    Dim sNomDueño       As String
    Dim rsData          As ADODB.Recordset
    Dim rsDataCarpeta   As ADODB.Recordset
    Dim sKey            As String
    Dim sTag            As String
    Dim sKeyPadre       As String
    '<V1.3.1>
    Dim sIndBloqueada   As String
    '</V1.3.1>
    
    On Error GoTo ErrCargaConsultas
    
    ' Carga Consultas en carpetas del usuario
    If Not db_LeeConsultasEnCarpetas(msNomUsuarioLocal, rsDataCarpeta) Then
        Exit Sub
    End If
    
    ' Carga Consultas personales por usuario
    If db_LeeConsultasPorUsuario(msNomUsuarioLocal, rsData) Then
        While Not rsData.EOF
            sNumConsulta = "" & rsData!num_consulta
            sNomConsulta = "" & rsData!nom_consulta
            sNomDueño = msNomUsuarioLocal
            '<V1.3.1>
            sIndBloqueada = "" & rsData!ind_bloqueada
            '</V1.3.1>
            sKey = "SQL" & "_" & sNumConsulta
            sTag = sKey & ";" & sNomDueño
            
            sKeyPadre = sNomUsuario
            rsDataCarpeta.Filter = "num_consulta=" & sNumConsulta
            If Not rsDataCarpeta.EOF Then
                sKeyPadre = LCase("" & rsDataCarpeta!gls_carpeta)
            End If
            
            '<V1.3.1>
            Call CargaHijo(sNomUsuario, sKeyPadre, "", sTag, sNomConsulta, sIndBloqueada)
            '</V1.3.1>
            
            rsData.MoveNext
        Wend
    End If
    
    Set rsData = Nothing
    Exit Sub

ErrCargaConsultas:
    Screen.MousePointer = vbNormal
    MsgBox Error, vbInformation, App.Title
    Exit Sub
End Sub

Sub CargaFormulario()
    Dim nMaxKey As Long
    
    IniciaForms
    bRedimensiona = False
    
    Screen.MousePointer = vbHourglass
    Me.StatusBar1.Panels(2).Text = "Cargando consultas ..."
    
    Set mNode = tvTreeView.Nodes.Add(, , msNomUsuarioLocal, msNomUsuarioLocal, "usuario")
    mNode.Tag = "USU"
    mNode.Expanded = True
    
    Set mNode = tvTreeView.Nodes.Add(, , "COM", "Consultas Grupales", "grupal")
    mNode.Tag = "COM"
    mNode.Expanded = True
    
    '<V1.3.0>
    Set mNode = tvTreeView.Nodes.Add(, , "LCO", "Consultas por Lote", "lote")
    mNode.Tag = "LCO"
    mNode.Expanded = True
    '</V1.3.0>
    
    Call CargaCarpetasUsuario(msNomUsuarioLocal, tvTreeView, nMaxKey)
    
    OpenMyDataBase
    
    Call CargaConsultas(msNomUsuarioLocal)
    '<V1.3.1>
    Call CargaAgrupaciones(msNomUsuarioLocal, "COM")
    '</V1.3.1>
    '<V1.3.0>
    Call CargaLotes("LCO")
    '</V1.3.0>
    
    CloseMyDataBase
    
    Screen.MousePointer = vbNormal
    
    bRedimensiona = True
    grdResultado.Top = 150
    grdResultado.Left = 60
    grdResultado.Visible = False
    Me.mnuEdiBuscarSgte.Enabled = False
    Me.StatusBar1.Panels(2).Text = ""
    
    Set mNode = tvTreeView.Nodes.Item(1)
    Me.Caption = "Consultas " & mNode
    msLastTag = mNode.Tag
End Sub

Public Sub CargaAgrupaciones(sNomUsuario As String, sCodPadre As String)
    Dim sNumPerfil      As String
    Dim sNomPerfil      As String
    Dim sNumConsulta    As String
    Dim sNomConsulta    As String
    Dim sCodLlave       As String
    Dim rsData          As ADODB.Recordset
    Dim sNumPerfilOld   As String
    '<V1.3.1>
    Dim sKeyPadre       As String
    Dim sGlsTag         As String
    Dim sIndBloqueada   As String
    '</V1.3.1>
    
    On Error GoTo ErrCargaAgrupaciones
    
    sNumPerfilOld = ""
    
    ' Carga Consultas asignadas por Perfiles
    If db_LeeConsultasAgrupacionPorUsuario(msNomUsuarioLocal, rsData) Then
        While Not rsData.EOF
            sNumPerfil = "" & rsData!num_perfil
            sNomPerfil = "" & rsData!nom_perfil
            sNumConsulta = "" & rsData!num_consulta
            sNomConsulta = "" & rsData!nom_consulta
            
            ' Carga nombre del perfil
            If sNumPerfil <> sNumPerfilOld Then
                ' Carga Consulta
                sCodLlave = sCodPadre & "_" & sNumPerfil
                Set mNode = tvTreeView.Nodes.Add(sCodPadre, tvwChild, sCodLlave, sNomPerfil, "carpeta")
                mNode.Tag = "PER" & "_" & sNumPerfil
                mNode.Expanded = True
            
                sNumPerfilOld = sNumPerfil
            End If
            
            ' Carga Consulta
            sCodLlave = sCodPadre & "_" & sNumPerfil & "_" & sNumConsulta
            '<V1.3.1>
            sKeyPadre = sCodPadre & "_" & sNumPerfil
            sGlsTag = "SQL" & "_" & sNumConsulta
            sIndBloqueada = "" & rsData!ind_bloqueada
            
            Call CargaHijo(sNomUsuario, sKeyPadre, sCodLlave, sGlsTag, sNomConsulta, sIndBloqueada)
            '</V1.3.1>
            
            rsData.MoveNext
        Wend
    End If
    
    Set rsData = Nothing
    Exit Sub
    
ErrCargaAgrupaciones:
    Screen.MousePointer = vbNormal
    MsgBox Error, vbInformation, App.Title
    Exit Sub
End Sub

Public Sub CargaLotes(sCodPadre As String)
    '<V1.3.0>
    Dim sNumLote      As String
    Dim sNomLote      As String
    Dim sCodLlave     As String
    Dim rsData        As ADODB.Recordset
    
    On Error GoTo ErrCargaLotes
    
    ' Carga lotes del usuario
    If db_LeeLotesPorUsuario(msNomUsuarioLocal, rsData) Then
        While Not rsData.EOF
            sNumLote = "" & rsData!num_lote
            sNomLote = "" & rsData!nom_lote
            
            ' Carga Lote
            sCodLlave = sCodPadre & "_" & sNumLote
            Set mNode = tvTreeView.Nodes.Add(sCodPadre, tvwChild, sCodLlave, sNomLote, "lote")
            mNode.Tag = "LOT" & "_" & sNumLote
            
            rsData.MoveNext
        Wend
    End If
    
    Set rsData = Nothing
    Exit Sub
    
ErrCargaLotes:
    Screen.MousePointer = vbNormal
    MsgBox Error, vbInformation, App.Title
    Exit Sub
    '</V1.3.0>
End Sub

Sub CargaHijo(sNomUsuario As String, sKeyPadre As String, sKeyNodo As String, sTagNodo As String, sNombre As String, sIndBloqueada As String)
    Dim nNode       As Node
    '<V1.3.1>
    Dim sGlsIcono   As String
    '</V1.3.1>
    
    On Error GoTo SinNodoPadre
    
    '<V1.3.1>
    If sIndBloqueada = "S" Then
        sGlsIcono = "cons_bloqueada"
    Else
        sGlsIcono = "consulta"
    End If
    Set nNode = tvTreeView.Nodes.Add(sKeyPadre, tvwChild, sKeyNodo, sNombre, sGlsIcono)
    '</V1.3.1>
    
    nNode.Tag = sTagNodo
    
    Exit Sub
    
SinNodoPadre:
    On Error Resume Next
    
    '<V1.3.1>
    Set nNode = tvTreeView.Nodes.Add(sNomUsuario, tvwChild, sKeyNodo, sNombre, sGlsIcono)
    '</V1.3.1>
    nNode.Tag = sTagNodo
    
    Exit Sub
End Sub

Sub EditarArchivo()
    Dim sFile   As String
    
    On Error GoTo ErrEditarArchivo
    
    sFile = fsExtraeRutaNombreConsulta(mNode.Tag)
    If gsNomEditor = "" Then
        MsgBox "Archivo Notepad.exe y Wordpad.exe no fueron encontrados en forma automática. " & _
               "Por favor configure su editor de texto y luego intente editar este archivo.", vbInformation, App.Title
        Exit Sub
    End If
    
    If Not Exist(gsNomEditor) Then
        MsgBox "Archivo " & gsNomEditor & " no se encuentra disponible. " & _
               "Por favor configure su editor de texto y luego intente editar este archivo.", vbInformation, App.Title
        Exit Sub
    End If
    
    Call Shell(gsNomEditor & " """ & sFile & """", vbMaximizedFocus)
    Exit Sub
    
ErrEditarArchivo:
    MsgBox Error, vbCritical, App.Title
End Sub

Sub ArcEditarConsulta()
    Dim sNumConsulta    As String
    Dim sNomDueño       As String
    Dim nPos            As Integer
    
    On Error GoTo ErrEditarArchivo
    
    If Left(mNode.Tag, 3) = "SQL" Then
        
        sNumConsulta = Mid(mNode.Tag, 5)
        nPos = InStr(sNumConsulta, ";")
        If nPos = 0 Then
            sNomDueño = ""
        Else
            sNomDueño = Mid(sNumConsulta, nPos + 1)
            sNumConsulta = Left(sNumConsulta, nPos - 1)
        End If
        
        If LCase(msNomUsuarioLocal) <> LCase(sNomDueño) Then
            MsgBox "Usted no es el dueño de esta consulta, no puede modificarla", vbCritical, App.Title
            Exit Sub
        End If
        
        gsNumConsulta = sNumConsulta
        gsNomConsulta = mNode.Text
        gsNomUsuarioLocal = msNomUsuarioLocal
        frmEditarConsulta.Show vbModal
        If Not gbCancelar Then
            ArcActualizarVentana
        End If
    End If
    
    Exit Sub
    
ErrEditarArchivo:
    MsgBox Error, vbCritical, App.Title
End Sub

Public Sub EjecutarConsulta()
    '<V1.3.0>
    ' Se separa la validación y ejecución de la rutina LeeDetalleConsulta
    Call LeeDetalleConsulta
    
    If ValidaConsulta Then
        If EjecutaConsulta(msNumConsulta, mnNumBaseDatos, msGlsQuery, maRegParametros, _
                           grsConsulta, grsFormatos, mnTotalRegUltQuery, mnuConResEnGrila.Checked, grdResultado, txtResultado, _
                           StatusBar1, ProgressBar1) Then
            ActivarBotones
        End If
    End If
    '</V1.3.0>
End Sub

Sub EjecutarLote()
    '<V1.3.0>
    Dim sNumLote    As String
    Dim sNomLote    As String
    Dim nPos        As Integer
    
    ' Determina numero del lote y confirma ejecución
    sNumLote = Mid(mNode.Tag, 5)
    sNomLote = mNode.Text
    
    If MsgBox("Está seguro que desea ejecutar el lote " & sNumLote & " (" & sNomLote & ")", vbYesNoCancel + vbDefaultButton2, App.Title) <> vbYes Then
        Exit Sub
    End If
    
    ' Ejecuta lote
    Call EjecutaConsultasPorLote(sNumLote)
    '</V1.3.0>
End Sub

Sub EjecutarNodo()
    '<V1.3.0>
    ' Ejecuta el nodo (Consulta o Lote de Consultas)
    Select Case Left(mNode.Tag, 3)
    Case "SQL"
        Call EjecutarConsulta
    Case "LOT"
        Call EjecutarLote
    End Select
    '</V1.3.0>
End Sub

Sub Exportar_A_Archivo()
    Dim sLinea              As String
    Dim nCampos             As Long
    Dim fCampos             As Field
    Dim nLargo()            As Long
    Dim nTipo()             As Integer
    Dim sTitulo()           As String
    Dim nReg                As Long
    Dim nX                  As Long
    Dim sValor              As String
    Dim nTotalRegQuery      As Long
    Dim bArchivoAbierto     As Boolean
                
                
    On Error GoTo ErrExportar_A_Archivo
    
    Screen.MousePointer = vbHourglass
    
    bArchivoAbierto = False
    nTotalRegQuery = grsConsulta.RecordCount
        
    ProgressBar1.Max = nTotalRegQuery
    ProgressBar1.Value = 0
    ProgressBar1.Visible = True
    
    Open gsNomArchivoExportar For Output As #1
    bArchivoAbierto = True
    
    nCampos = 0
    nX = 0
    For Each fCampos In grsConsulta.Fields
        nX = nX + 1
        nCampos = nCampos + 1
        ReDim Preserve nLargo(nX) As Long
        ReDim Preserve nTipo(nX) As Integer
        ReDim Preserve sTitulo(nX) As String
        
        nTipo(nX) = fCampos.Type
        sTitulo(nX) = FormatoTitulo(fCampos.Name)
        
        Select Case fnTipoDatoRecordset(nTipo(nX))
        Case wc_tipo_dato_integer
            nLargo(nX) = IIf(Len(sTitulo(nX)) > 11, Len(sTitulo(nX)), 11)
        Case wc_tipo_dato_float
            nLargo(nX) = IIf(Len(sTitulo(nX)) > 50, Len(sTitulo(nX)), 50)
        Case wc_tipo_dato_fecha
            nLargo(nX) = IIf(Len(sTitulo(nX)) > 24, Len(sTitulo(nX)), 24)
        Case Else
            nLargo(nX) = IIf(Len(sTitulo(nX)) > fCampos.DefinedSize, Len(sTitulo(nX)), fCampos.DefinedSize)
        End Select
    Next

    grsConsulta.MoveLast
    grsConsulta.MoveFirst
    If Not grsConsulta.EOF Then
        If gsGlsSeparadorCampos = "" Then
            If nCampos = 1 Then
                For nReg = 1 To nTotalRegQuery
                    ProgressBar1.Value = nReg
                    sLinea = "" & grsConsulta(nX - 1)
                    
                    Print #1, sLinea
                    grsConsulta.MoveNext
                Next
            Else
                For nReg = 1 To nTotalRegQuery
                    ProgressBar1.Value = nReg
        
                    sLinea = ""
                    For nX = 1 To nCampos - 1
                        sValor = Left("" & grsConsulta(nX - 1) & Space(255), nLargo(nX))
                        sLinea = sLinea & sValor & " "
                    Next
                    sValor = Left("" & grsConsulta(nX - 1) & Space(255), nLargo(nX))
                    sLinea = sLinea & sValor
                    
                    Print #1, sLinea
                    grsConsulta.MoveNext
                Next
            End If
        Else
            For nReg = 1 To nTotalRegQuery
                ProgressBar1.Value = nReg
    
                sLinea = ""
                For nX = 1 To nCampos - 1
                    sValor = "" & grsConsulta(nX - 1)
                    sLinea = sLinea & sValor & gsGlsSeparadorCampos
                Next
                sValor = "" & grsConsulta(nX - 1)
                sLinea = sLinea & sValor
                
                Print #1, sLinea
                grsConsulta.MoveNext
            Next
        End If
    End If
    
    ProgressBar1.Value = 0
    ProgressBar1.Visible = False
    
    Close #1
    bArchivoAbierto = False
    
    Screen.MousePointer = vbNormal
    MsgBox "Archivo exportado correctamente", vbInformation, App.Title
    
    Exit Sub
    
ErrExportar_A_Archivo:
    Screen.MousePointer = vbNormal
    MsgBox Error, vbCritical, App.Title
    If bArchivoAbierto Then
        Close #1
    End If
End Sub

Public Sub ExportarExcel()
    Dim sGlsError       As String
    
    If Not grdResultado.Visible Then
        MsgBox "Debe ejecutar la consulta primero. Asegúrese que Resultado en Grilla esté activado en el menú Consulta", vbCritical, App.Title
        Exit Sub
    End If
    If grdResultado.MaxRows < 1 Then
        MsgBox "No existen registros a procesar", vbCritical, App.Title
        Exit Sub
    End If
    
    frmExpExcel.msNomConsulta = mNode
    frmExpExcel.Show vbModal
    If Not gbCancelar Then
        '<V1.3.0>
        ' Cambio de Sub a Function para mostrar el error afuera de la rutina
        If Not ExportarToFile(Me.Caption, grdResultado, grsConsulta, grsFormatos, gaRegParametros, ProgressBar1, StatusBar1, False, sGlsError) Then
            MsgBox sGlsError, vbCritical, App.Title
        End If
        '</V1.3.0>
    End If
End Sub

Sub FormResize(y As Long)
    Dim nAlto   As Long
    Dim nAncho  As Long
    Dim nTab    As Integer
    
    AutoRedraw = False
    
    ' Ajusta altura de las consultas
    nAncho = y - 15
    If nAncho < 0 Then nAncho = 0
    fraConsultas.Width = nAncho
    tvTreeView.Width = nAncho - 240
    
    nAlto = Me.Height - fraGrillas.Top - StatusBar1.Height - 415
    If nAlto < 0 Then nAlto = 0
    fraConsultas.Height = nAlto
    fraGrillas.Height = fraConsultas.Height
    tvTreeView.Height = fraConsultas.Height - 360
    imgSplitter.Left = y
    
    ' Ajusta posicion del la grilla
    fraGrillas.Left = y + 90
    fraGrillas.Width = Me.Width - fraGrillas.Left - 140
      
    ' Ajusta altura de las grillas
    nAlto = fraGrillas.Height - Me.grdResultado.Top - 70
    If nAlto < 0 Then nAlto = 0
    grdResultado.Height = nAlto
    nAlto = nAlto - 60
    If nAlto < 0 Then nAlto = 0
    txtResultado.Height = nAlto
    
    ' Ajusta anchos
    'nAncho = fraConsultas.Width - File1.Left - 60
    If nAncho < 0 Then nAncho = 0
    'File1.Width = nAncho
    
    nAncho = fraGrillas.Width - grdResultado.Left - 70
    If nAncho < 0 Then nAncho = 0
    grdResultado.Width = nAncho
    nAncho = nAncho - 60
    If nAncho < 0 Then nAncho = 0
    txtResultado.Width = nAncho

    imgSplitter.Height = fraConsultas.Height
    picSplitter.Height = fraConsultas.Height

    AutoRedraw = True
End Sub

Sub IniciaFormulario()
    frmMdiPadre.Toolbar1(0).Visible = False
    frmMdiPadre.Toolbar1(1).Visible = True
    frmMdiPadre.Toolbar1(2).Visible = False
    
    txtResultado.Top = grdResultado.Top
    txtResultado.Left = grdResultado.Left
    
    txtResultado.Visible = False
    grdResultado.Visible = False
    grdResultado.MaxRows = 0
    txtResultado.Text = ""
    
    Me.Show
    msNomUsuarioLocal = gsNomUsuarioLocal
    Me.StatusBar1.Panels(1).Text = UCase(msNomUsuarioLocal)
    
    Screen.MousePointer = vbHourglass
    Me.StatusBar1.Panels(2).Text = "Obteniendo informacion del usuario ..."
    
    OpenMyDataBase
    Call db_LeeUsuario(msNomUsuarioLocal, mrsUsuarioLocal)
    CloseMyDataBase
    
    Screen.MousePointer = vbNormal
    ActivarBotones
End Sub

Sub ArcCerrar()
    Unload Me
End Sub

Function ValidaConsulta() As Boolean
    Dim sGlsHorarios    As String
    
    '<V1.3.0>
    ' Se separa la validación en una rutina para ser llamada desde ejecución Simple o por Lote
    If mnNumBaseDatos <= 0 Then
        Screen.MousePointer = vbNormal
        MsgBox "Consulta no tiene definida la base de datos a utilizar", vbInformation, App.Title
        ValidaConsulta = False
        Exit Function
    End If
    
    If msGlsQuery = "" Then
        Screen.MousePointer = vbNormal
        MsgBox "Consulta no tiene definido el query a ejecutar", vbCritical, App.Title
        ValidaConsulta = False
        Exit Function
    End If
    
    If Not EsHoraDeEjecucion(msGlsHorario, sGlsHorarios) Then
        Screen.MousePointer = vbNormal
        MsgBox "Horario definido para esta consulta no permite su ejecución en este momento (" & sGlsHorarios & ")", vbCritical, App.Title
        ValidaConsulta = False
        Exit Function
    End If
    
    ValidaConsulta = True
    '</V1.3.0>
End Function

Public Sub VerParametros()
    gaRegParametros = maRegParametros
    frmParametros.mnNumBaseDatos = mnNumBaseDatos
    frmParametros.Show vbModal
    
    If Not gbCancelar Then
        maRegParametros = gaRegParametros
        
        If EjecutaSentencia(msNumConsulta, mnNumBaseDatos, msGlsQuery, maRegParametros, grsConsulta, ProgressBar1, StatusBar1) Then
            mnTotalRegUltQuery = grsConsulta.RecordCount
            StatusBar1.Panels(2).Text = ""
            StatusBar1.Panels(3).Text = Trim(CStr(mnTotalRegUltQuery)) & " reg"
            
            If mnTotalRegUltQuery = 0 Then
                MsgBox "No hay registros para esta consulta", vbInformation, App.Title
            Else
                If mnuConResEnGrila.Checked Then
                    Call CargarResultadoEnGrilla(grsConsulta, grsFormatos, grdResultado, txtResultado, ProgressBar1)
                Else
                    Call CargarResultadoEnTexto(grsConsulta, grsFormatos, grdResultado, txtResultado, ProgressBar1)
                End If
            End If
        End If
    End If
End Sub

Private Sub Form_Activate()
    App.HelpFile = App.Path & "\Satelite.hlp"
    
    frmMdiPadre.Toolbar1(0).Visible = False
    frmMdiPadre.Toolbar1(1).Visible = True
    frmMdiPadre.Toolbar1(2).Visible = False
    
    ActivarBotones
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMdiPadre.Toolbar1(1).Visible = False
    frmMdiPadre.Toolbar1(0).Visible = True
End Sub

Private Sub grdResultado_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button And vbRightButton And grdResultado.MaxRows > 0 Then PopupMenu mnuPopupRes
End Sub

Private Sub imgSplitter_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    With imgSplitter
        picSplitter.Move .Left, .Top, .Width \ 2, .Height - 20
    End With
    picSplitter.Visible = True
    mbMoving = True
End Sub

Private Sub imgSplitter_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim sglPos As Single
    
    If mbMoving Then
        sglPos = x + imgSplitter.Left
        If sglPos < 500 Then
            picSplitter.Left = 500
        ElseIf sglPos > Me.Width - 500 Then
            picSplitter.Left = Me.Width - 500
        Else
            picSplitter.Left = sglPos
        End If
    End If
End Sub

Private Sub imgSplitter_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    FormResize picSplitter.Left
    picSplitter.Visible = False
    SaveSetting App.Title, "Settings", "SplitterVer", imgSplitter.Left
    mbMoving = False
End Sub

Private Sub Form_Load()
    IniciaFormulario
    CargaFormulario
End Sub

Sub LeeDetalleConsulta()
    Dim nPos            As Integer
    Dim sNomDueño       As String
    
    mnTotalRegUltQuery = 0
    grdResultado.MaxRows = 0
    txtResultado.Text = ""
    grdResultado.Visible = False
    txtResultado.Visible = False
    msGlsQuery = ""
        
    msNumConsulta = Mid(mNode.Tag, 5)
    nPos = InStr(msNumConsulta, ";")
    If nPos = 0 Then
        sNomDueño = ""
    Else
        sNomDueño = Mid(msNumConsulta, nPos + 1)
        msNumConsulta = Left(msNumConsulta, nPos - 1)
    End If
    
    '<V1.3.0>
    gsNomConsulta = mNode
    '</V1.3.0>
    
    Call CargaConsulta(msNumConsulta, mnNumBaseDatos, msGlsQuery, grsFormatos, msGlsHorario)

    '<V1.3.0>
    ' Se eliminó la funcionalidad de bEjecutar la consulta de acuerdo al valor de este parámetro
    'If bEjecutar Then
    '...
    'End If
    '</V1.3.0>
End Sub

Sub CargaConsulta(sNumConsulta As String, nNumBaseDatos As Integer, sGlsQuery As String, rsFormatos As ADODB.Recordset, sGlsHorario As String)
    Dim rsData          As ADODB.Recordset
    
    On Error GoTo ErrCargaConsulta
        
    ' Abre base datos
    OpenMyDataBase
    
    ' Lee consulta
    If db_LeeConsulta(sNumConsulta, rsData) Then
        If Not rsData.EOF Then
            nNumBaseDatos = Val(rsData!num_basedatos)
            sGlsQuery = "" & rsData!gls_query
            sGlsHorario = "" & rsData!gls_horario_ejecucion
        
            Call CargaParametros(sNumConsulta, sGlsQuery, maRegParametros())
            Call db_LeeFormatos(sNumConsulta, rsFormatos)
        End If
    End If
        
    ' Cierra base datos
    CloseMyDataBase
    
    Set rsData = Nothing
    Exit Sub

ErrCargaConsulta:
    Screen.MousePointer = vbNormal
    MsgBox Error, vbInformation, App.Title
    Exit Sub
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState <> 1 And Me.WindowState <> 1 Then
        If Me.Width < 500 Then Me.Width = 500
        FormResize imgSplitter.Left
    End If
End Sub

Private Sub mnuArcActualizar_Click()
    ArcActualizarVentana
End Sub

Private Sub mnuArcAdministracion_Click()
    ArcAdministracion
End Sub

Private Sub mnuArcCerrarVentana_Click()
    ArcCerrarVentana
End Sub

Private Sub mnuArcEjecutar_Click()
    '<V1.3.0>
    ' Modificado para ejecutar Consulta o Lote
    EjecutarNodo
    '</V1.3.0>
End Sub

Private Sub mnuArcEliminarCarpeta_Click()
    ArcEliminarCarpeta
End Sub

Private Sub mnuArcExpArchivo_Click()
    ExportarArchivo
End Sub

Private Sub mnuArcExportar_Click()
    ExportarExcel
End Sub

Private Sub mnuArcMoverConsulta_Click()
    ArcMoverConsulta
End Sub

Private Sub mnuArcNueva_Click()
    ArcNuevaVentana
End Sub

Private Sub mnuArcNuevaCarpeta_Click()
    ArcNuevaCarpeta
End Sub

Private Sub mnuArcNuevaConsulta_Click()
    ArcNuevaConsulta
End Sub

Private Sub mnuArcSalir_Click()
    End
End Sub

Private Sub mnuConResEnGrila_Click()
    If Not mnuConResEnGrila.Checked Then
        mnuConResEnGrila.Checked = True
        mnuConResEnTexto.Checked = False
        If mnTotalRegUltQuery > 0 Then
            If grdResultado.MaxRows = 0 Then
                Call CargarResultadoEnGrilla(grsConsulta, grsFormatos, grdResultado, txtResultado, ProgressBar1)
            Else
                txtResultado.Visible = False
                grdResultado.Visible = True
            End If
        End If
        ActivarBotones
    End If
End Sub

Private Sub mnuConResEnTexto_Click()
    If Not mnuConResEnTexto.Checked Then
        mnuConResEnGrila.Checked = False
        mnuConResEnTexto.Checked = True
        If mnTotalRegUltQuery > 0 Then
            If txtResultado = "" Then
                Call CargarResultadoEnTexto(grsConsulta, grsFormatos, grdResultado, txtResultado, ProgressBar1)
            Else
                grdResultado.Visible = False
                txtResultado.Visible = True
            End If
        End If
        ActivarBotones
    End If
End Sub

Private Sub mnuEdiBuscarSgte_Click()
    Call BuscarSiguiente
End Sub

Private Sub mnuEdiBuscarTexto_Click()
    Call BuscarTexto
End Sub

Private Sub mnuPopConNuevaConsulta_Click()
    ArcNuevaConsulta
End Sub

Private Sub mnuPopEdiEditar_Click()
    ArcEditarConsulta
End Sub

Private Sub mnuPopResExcel_Click()
    ExportarExcel
End Sub

Private Sub mnuArcEditarConsulta_Click()
    ArcEditarConsulta
End Sub

Private Sub mnuPopEdiEjecutar_Click()
    '<V1.3.0>
    ' Modificado para ejecutar Consulta o Lote
    EjecutarNodo
    '</V1.3.0>
End Sub

Private Sub mnuPopEliminarCarpeta_Click()
    ArcEliminarCarpeta
End Sub

Private Sub mnuPopMoverConsulta_Click()
    ArcMoverConsulta
End Sub

Private Sub mnuPopNuevaCarpeta_Click()
    ArcNuevaCarpeta
End Sub

Private Sub mnuPopCarNuevaConsulta_Click()
    ArcNuevaConsulta
End Sub

Private Sub mnuPopResTexto_Click()
    ExportarArchivo
End Sub

Private Sub mnuWindowCascade_Click()
    ' Organiza los formularios secundarios en cascada.
    frmMdiPadre.Arrange vbCascade
End Sub

Private Sub mnuWindowVert_Click()
    ' Organiza los formularios secundarios en mosaico.
    frmMdiPadre.Arrange vbTileVertical
End Sub

Private Sub mnuWindowHort_Click()
    ' Organiza los formularios secundarios en mosaico.
    frmMdiPadre.Arrange vbTileHorizontal
End Sub

Private Sub tvTreeView_Collapse(ByVal Node As ComctlLib.Node)
    SaveSetting App.Title, "Expanded", Node.Key, "0"
End Sub

Private Sub tvTreeView_DblClick()
    '<V1.3.0>
    ' Modificado para agregar ejecución del lote
    If Me.mnuArcEjecutar.Enabled Then
        EjecutarNodo
    End If
    '<V1.3.0>
End Sub

Private Sub tvTreeView_Expand(ByVal Node As ComctlLib.Node)
    SaveSetting App.Title, "Expanded", Node.Key, "1"
End Sub

Private Sub tvTreeView_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button And vbRightButton Then
        Select Case Left(mNode.Tag, 3)
        Case "USU", "DIR"
            PopupMenu mnuPopCarpetas
        Case "SQL"
            PopupMenu mnuPopConsultas
        End Select
    End If
End Sub

Private Sub tvTreeView_NodeClick(ByVal Node As Node)
    Set mNode = Node
    If mNode.Tag = "USU" Then
        Me.Caption = "Consultas " & mNode
    ElseIf mNode.Tag = "COM" Then
        Me.Caption = mNode & " " & msNomUsuarioLocal
    Else
        Me.Caption = mNode
    End If
    
    If mNode.Tag <> msLastTag Then
        msLastTag = mNode.Tag
        If Left(mNode.Tag, 3) = "SQL" Then
            Call LeeDetalleConsulta
        End If
        ActivarBotones
    End If
End Sub

Sub IniciaForms()
    imgSplitter.Left = GetSetting(App.Title, "Settings", "SplitterVer", 3000)
    
    FormResize imgSplitter.Left
    
    tvTreeView.Nodes.Clear
End Sub

Sub ExportarArchivo()
    frmExpArchivo.msNomConsulta = mNode
    frmExpArchivo.Show vbModal
    If Not gbCancelar Then
        Call Exportar_A_Archivo
    End If
End Sub

Sub ArcNuevaVentana()
    Dim x As New frmPrincipal
    x.Show
End Sub

Sub ArcActualizarVentana()
    CargaFormulario
    ActivarBotones
End Sub

Sub ArcNuevaCarpeta()
    gsTagNodoActual = mNode.Tag
    gsNombreNodoActual = mNode.Text
    
    gnObjetoACrear = 1
    frmNuevaCarpeta.Show vbModal
    If Not gbCancelar Then
        ArcActualizarVentana
    End If
End Sub

Sub EjecutaConsultasPorLote(sNumLote As String)
    '<V1.3.0>
    Dim sNumConsulta        As String
    Dim sGlsHorarios        As String
    Dim bResultadoEnGrilla  As String
    Dim nTotConsultas       As Integer
    Dim nIndex              As Integer
    Dim nCtaParamInput      As Integer
    Dim nTotalRegUltQuery   As Integer
    Dim sGlsError           As String
    Dim rsData              As ADODB.Recordset
    
    On Error GoTo ErrConsultasDelLote

    nTotConsultas = 0
    ReDim gaLteRegParametros(0) As lteRegParametros
        
    ' Abre base datos
    OpenMyDataBase

    ' Lee consulta
    If db_LeeConsultasPorLote(sNumLote, rsData) Then
        While Not rsData.EOF
            sNumConsulta = rsData!num_consulta
            gsNomConsulta = "" & rsData!nom_consulta
            
            ' Carga parametros de la consulta
            ReDim Preserve gaLteRegParametros(nTotConsultas) As lteRegParametros
            gaLteRegParametros(nTotConsultas).num_consulta = sNumConsulta
            gaLteRegParametros(nTotConsultas).nom_consulta = gsNomConsulta
            gaLteRegParametros(nTotConsultas).arc_salida = "" & rsData!gls_archivo_salida
            gaLteRegParametros(nTotConsultas).hja_salida = "" & rsData!nom_hoja_salida

            ' Carga datos de la consulta
            Call CargaConsulta(sNumConsulta, mnNumBaseDatos, msGlsQuery, grsFormatos, msGlsHorario)
            
            ' Valida condiciones para la ejecuión de la consulta
            If Not ValidaConsulta Then
                CloseMyDataBase
                Set rsData = Nothing
                Exit Sub
            End If

            'Se cargan los paramétros de la consulta.
            If UBound(maRegParametros) <= 0 Then
                ReDim gaRegParametros(0) As rRegParametros
                ReDim gaLteRegParametros(nTotConsultas).par_consulta(0) As rRegParametros
            Else
                Call CargaParametrosDefault(maRegParametros, nCtaParamInput)
                If nCtaParamInput > 0 Then
                    gaRegParametros = maRegParametros
                    frmParametros.mnNumBaseDatos = mnNumBaseDatos

                    gbEjecutandoLote = True
                    frmParametros.Show vbModal
                    gbEjecutandoLote = False
                    
                    If gbCancelar Then
                        CloseMyDataBase
                        Set rsData = Nothing
                        Screen.MousePointer = vbNormal
                        MsgBox "Ejecución de Lote ha sido cancelado", vbCritical, App.Title
                        Exit Sub
                    Else
                        maRegParametros = gaRegParametros
                        gaLteRegParametros(nTotConsultas).par_consulta = maRegParametros
                    End If
                End If
            End If

            ' Si la consulta no tiene nombre de archivo, se pide su ingreso
            If gaLteRegParametros(nTotConsultas).arc_salida = "" Then
                gbEjecutandoLote = True
                frmExpExcel.msNomConsulta = gsNomConsulta
                frmExpExcel.Show vbModal
                gbEjecutandoLote = False
                
                If gbCancelar Then
                    CloseMyDataBase
                    Set rsData = Nothing
                    Screen.MousePointer = vbNormal
                    MsgBox "Ejecución de Lote ha sido cancelado", vbCritical, App.Title
                    Exit Sub
                End If
                
                gaLteRegParametros(nTotConsultas).arc_salida = gsNomArchivoExportar
                gaLteRegParametros(nIndex).hja_salida = gsNomHojaExportar
            End If
            
            nTotConsultas = nTotConsultas + 1
            rsData.MoveNext
        Wend
    End If

    ' Lee consulta
    bResultadoEnGrilla = True
    nIndex = 0

    If nTotConsultas > 0 Then
        For nIndex = 0 To nTotConsultas - 1
            sNumConsulta = gaLteRegParametros(nIndex).num_consulta
            gsNomConsulta = gaLteRegParametros(nIndex).nom_consulta
            GrabaLog "Ejecutando consulta " & sNumConsulta
            Call CargaConsulta(sNumConsulta, mnNumBaseDatos, msGlsQuery, grsFormatos, msGlsHorario)

            ' Trasapasa los parámetros cargados en el ciclo anterior
            maRegParametros = gaLteRegParametros(nIndex).par_consulta

            If EjecutaSentencia(sNumConsulta, mnNumBaseDatos, msGlsQuery, maRegParametros, grsConsulta, ProgressBar1, StatusBar1) Then
                nTotalRegUltQuery = grsConsulta.RecordCount
                StatusBar1.Panels(2).Text = ""
                StatusBar1.Panels(3).Text = Trim(CStr(nTotalRegUltQuery)) & "reg"

                GrabaLog "Cargando resultado en grilla"
                If bResultadoEnGrilla Then
                    Call CargarResultadoEnGrilla(grsConsulta, grsFormatos, grdResultado, txtResultado, ProgressBar1)
                Else
                    Call CargarResultadoEnTexto(grsConsulta, grsFormatos, grdResultado, txtResultado, ProgressBar1)
                End If
                    
                ' Se exporta automaticamente a Excel
                GrabaLog "Exportando data"
                gsNomArchivoExportar = gaLteRegParametros(nIndex).arc_salida
                gsNomHojaExportar = gaLteRegParametros(nIndex).hja_salida
                If gsNomArchivoExportar <> "" Then
                    If Not ExportarToFile(gsNomConsulta, grdResultado, grsConsulta, grsFormatos, maRegParametros, ProgressBar1, StatusBar1, True, sGlsError) Then
                        CloseMyDataBase
                        Set rsData = Nothing
                        Screen.MousePointer = vbNormal
                        MsgBox sGlsError, vbCritical, App.Title
                        Exit Sub
                    End If
                End If
                
                ' Limpia la ventana de resultados
                mnTotalRegUltQuery = 0
                grdResultado.MaxRows = 0
                txtResultado.Text = ""
                grdResultado.Visible = False
                txtResultado.Visible = False
                msGlsQuery = ""
            End If
        Next
    End If

    ' Cierra base datos
    CloseMyDataBase

    Screen.MousePointer = vbNormal
    MsgBox "Lote ha sido ejecutado sin errores", vbInformation, App.Title
    
    Set rsData = Nothing
    Exit Sub

ErrConsultasDelLote:
    Screen.MousePointer = vbNormal
    MsgBox Error, vbInformation, App.Title
    Exit Sub
    Resume
    '<V1.3.0>
End Sub

