VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmAdministracion 
   Caption         =   "Módulo de Administración"
   ClientHeight    =   3975
   ClientLeft      =   2040
   ClientTop       =   2310
   ClientWidth     =   9345
   Icon            =   "frmAdministracion.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3975
   ScaleWidth      =   9345
   Begin VB.Frame fraVista 
      Caption         =   "Consultas"
      Height          =   3135
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9015
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
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   60
      Top             =   3180
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   2
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmAdministracion.frx":038A
            Key             =   "ConsBloq_N"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmAdministracion.frx":08DC
            Key             =   "ConsBloq_S"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "&Archivo"
      Begin VB.Menu mnuArcNuevo 
         Caption         =   "&Nuevo"
      End
      Begin VB.Menu mnuArcEditar 
         Caption         =   "&Editar"
      End
      Begin VB.Menu mnuArcEliminar 
         Caption         =   "E&liminar"
      End
      Begin VB.Menu mnuArcNulo1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuArcAsigUsuarios 
         Caption         =   "Asignar &Usuarios"
      End
      Begin VB.Menu mnuArcAsigConsultas 
         Caption         =   "Asignar &Consultas"
      End
      Begin VB.Menu mnuArcAsigAgrupacion 
         Caption         =   "Asignar &Perfiles"
      End
      Begin VB.Menu mnuArcAsigLotes 
         Caption         =   "Asignar &Lotes"
      End
      Begin VB.Menu mnuArcNulo2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuArcAbrirCons 
         Caption         =   "Abrir consultas a su nombre"
      End
      Begin VB.Menu mnuArcBloquear 
         Caption         =   "&Bloquear consulta"
      End
      Begin VB.Menu mnuArcNulo3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuArcCerrarVentana 
         Caption         =   "Cerrar ventana"
      End
      Begin VB.Menu mnuArcSalir 
         Caption         =   "&Salir"
      End
   End
   Begin VB.Menu mnuVer 
      Caption         =   "&Ver"
      Begin VB.Menu mnuVerVista 
         Caption         =   "Detalle de &Usuarios"
         Index           =   0
      End
      Begin VB.Menu mnuVerVista 
         Caption         =   "Detalle de &Consultas"
         Index           =   1
      End
      Begin VB.Menu mnuVerVista 
         Caption         =   "Detalle de &Perfiles"
         Index           =   2
      End
      Begin VB.Menu mnuVerVista 
         Caption         =   "Detalle de &Lotes"
         Index           =   3
      End
      Begin VB.Menu mnuVerVista 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuVerVista 
         Caption         =   "&Tipos de usuarios"
         Index           =   5
      End
      Begin VB.Menu mnuVerVista 
         Caption         =   "&Bases de datos"
         Index           =   6
      End
      Begin VB.Menu mnuVerVista 
         Caption         =   "Tabla de &Valores"
         Index           =   7
      End
      Begin VB.Menu mnuVerNulo1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVerFiltro 
         Caption         =   "&Aplicar filtro"
      End
   End
   Begin VB.Menu mnWindow 
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
   Begin VB.Menu mnuPopColumna 
      Caption         =   "Columna"
      Visible         =   0   'False
      Begin VB.Menu mnuPopColOrdenar 
         Caption         =   "Ordenar"
      End
      Begin VB.Menu mnuPopColNulo1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopColFiltrar 
         Caption         =   "Filtrar"
         Begin VB.Menu mnuPopColFiltrarPor 
            Caption         =   "Filtrar por ..."
            Index           =   0
         End
      End
   End
End
Attribute VB_Name = "frmAdministracion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mItem                   As ListItem
Dim msNumUltConsultas       As String
Dim msNomUltUsuario         As String
Dim msNumUltPerfil          As String
'<V1.3.0>
Dim msNumUltLote            As String
'</V1.3.0>
'<V1.3.1>
Dim msNumUltRegTabValor     As String
'</V1.3.1>
Dim msCodUltTipoUsuario     As String
Dim msNumUltBaseDatos       As String
Dim mnVistaActual           As Integer

'<V1.3.1>
Dim rsData                  As ADODB.Recordset
Public rsDataFiltro         As ADODB.Recordset

Public mnNumColumnaActiva   As Long
Public msValColumnaActiva   As String
'</V1.3.1>

Sub ActivarBotones()
    Dim bOpcConsultas   As Boolean
    Dim bOpcAsigUsu     As Boolean
    Dim bOpcAsigCon     As Boolean
    Dim bOpcAsigPer     As Boolean
    Dim bOpcAsigLot     As Boolean
    
    bOpcConsultas = (lvConsultas.ListItems.Count > 0)
    
    Me.mnuArcEditar.Enabled = bOpcConsultas
    Me.mnuArcEliminar.Enabled = bOpcConsultas
    
    Select Case mnVistaActual
    Case mnVistaUsuarios
        bOpcAsigUsu = False
        bOpcAsigCon = bOpcConsultas
        bOpcAsigPer = bOpcConsultas
        bOpcAsigLot = bOpcConsultas
    Case mnVistaConsultas
        bOpcAsigUsu = bOpcConsultas
        bOpcAsigCon = False
        bOpcAsigPer = bOpcConsultas
        bOpcAsigLot = False
    Case mnVistaPerfiles
        bOpcAsigUsu = bOpcConsultas
        bOpcAsigCon = bOpcConsultas
        bOpcAsigPer = False
        bOpcAsigLot = False
    '<V1.3.0>
    Case mnVistaLotes
        bOpcAsigUsu = bOpcConsultas
        bOpcAsigCon = False
        bOpcAsigPer = False
        bOpcAsigLot = False
    '</V1.3.0>
    Case mnVistaTiposUsuarios, mnVistaBaseDatos
        bOpcAsigUsu = False
        bOpcAsigCon = False
        bOpcAsigPer = False
        bOpcAsigLot = False
    '<V1.3.1>
    Case mnVistaTabValores
        bOpcAsigUsu = False
        bOpcAsigCon = False
        bOpcAsigPer = False
        bOpcAsigLot = False
    '</V1.3.1>
    End Select
    
    Me.mnuArcAsigUsuarios.Enabled = bOpcAsigUsu
    Me.mnuArcAsigConsultas.Enabled = bOpcAsigCon
    Me.mnuArcAsigAgrupacion.Enabled = bOpcAsigPer
    Me.mnuArcAbrirCons.Enabled = (bOpcConsultas And mnVistaActual = mnVistaUsuarios)
    '<V1.3.1>
    Me.mnuArcBloquear.Enabled = (bOpcConsultas And mnVistaActual = mnVistaConsultas)
    '</V1.3.1>

    frmMdiPadre.Toolbar1(2).Buttons(1).Enabled = True
    frmMdiPadre.Toolbar1(2).Buttons(2).Enabled = bOpcConsultas
    frmMdiPadre.Toolbar1(2).Buttons(3).Enabled = bOpcConsultas
    
    '<V1.3.1>
    ' Se incrementa en 2 el numero del boton
    frmMdiPadre.Toolbar1(2).Buttons(14).Enabled = bOpcAsigUsu
    frmMdiPadre.Toolbar1(2).Buttons(15).Enabled = bOpcAsigCon
    frmMdiPadre.Toolbar1(2).Buttons(16).Enabled = bOpcAsigPer
    
    ' Se configura la opcion para Lotes
    Me.mnuArcAsigLotes.Enabled = bOpcAsigLot
    frmMdiPadre.Toolbar1(2).Buttons(17).Enabled = bOpcAsigLot
    '</V1.3.1>
End Sub

Sub ArcAbrirModConsultas()
    Dim x As New frmPrincipal
    
    gsNomUsuarioLocal = mItem
    x.Show
End Sub

Sub ArcEditarBaseDato()
    On Error GoTo ErrArcEditarBaseDato
    
    If lvConsultas.ListItems.Count > 0 Then
        gsNumBaseDatos = mItem
        frmEditarBaseDatos.Show vbModal
        If Not gbCancelar Then
            msNumUltBaseDatos = gsNumBaseDatos
            Call CargaBaseDatos(True)
        End If
    End If
    
    Exit Sub
    
ErrArcEditarBaseDato:
    MsgBox Error, vbCritical, App.Title
End Sub

Sub ArcEditarConsulta()
    On Error GoTo ErrEditarConsulta
    
    If lvConsultas.ListItems.Count > 0 Then
        gsNumConsulta = mItem
        gsNomConsulta = mItem.SubItems(1)
        gsNomUsuarioLocal = gsUsuarioReal
        frmEditarConsulta.Show vbModal
        If Not gbCancelar Then
            msNumUltConsultas = gsNumConsulta
            Call CargaConsultas(True)
        End If
    End If
    
    Exit Sub
    
ErrEditarConsulta:
    MsgBox Error, vbCritical, App.Title
End Sub

Sub ArcEditarElemento()
    Select Case mnVistaActual
    Case mnVistaUsuarios
        ArcEditarUsuario
    Case mnVistaConsultas
        ArcEditarConsulta
    Case mnVistaPerfiles
        ArcEditarAgrupacion
    '<V1.3.0>
    Case mnVistaLotes
        ArcEditarLote
    '</V1.3.0>
    Case mnVistaTiposUsuarios
        ArcEditarTipoUsuario
    Case mnVistaBaseDatos
        ArcEditarBaseDato
    '<V1.3.1>
    Case mnVistaTabValores
        ArcEditarTabValores
    '</V1.3.1>
    End Select
End Sub

Sub ArcEditarAgrupacion()
    On Error GoTo ErrEditarAgrupacion
    
    If lvConsultas.ListItems.Count > 0 Then
        gsNumPerfil = mItem
        gsNomPerfil = mItem.SubItems(1)
        frmEditarPerfil.Show vbModal
        If Not gbCancelar Then
            msNumUltPerfil = gsNumPerfil
            Call CargaAgrupaciones(True)
        End If
    End If
    
    Exit Sub
    
ErrEditarAgrupacion:
    MsgBox Error, vbCritical, App.Title
End Sub

Sub ArcEditarLote()
'<V1.3.0>
    On Error GoTo ErrEditarLote
    
    If lvConsultas.ListItems.Count > 0 Then
        gsNumLote = mItem
        gsNomLote = mItem.SubItems(1)
        gsNomSolicitante = mItem.SubItems(2)
        gsNomUsuarioLocal = gsUsuarioReal
        frmEditarLote.Show vbModal
        If Not gbCancelar Then
            msNumUltLote = gsNumLote
            Call CargaLotes(True)
        End If
    End If
    
    Exit Sub
    
ErrEditarLote:
    MsgBox Error, vbCritical, App.Title
'</V1.3.0>
End Sub

Sub ArcEditarTabValores()
    '<V1.3.1>
    On Error GoTo ErrEditarTabValores
    
    If lvConsultas.ListItems.Count > 0 Then
        gsNumRegTabValor = mItem
        frmEditarTabValores.Show vbModal
        If Not gbCancelar Then
            msNumUltRegTabValor = gsNumRegTabValor
            Call CargaTabValores(True)
        End If
    End If
    
    Exit Sub
    
ErrEditarTabValores:
    MsgBox Error, vbCritical, App.Title
    '</V1.3.1>
End Sub


Sub ArcEditarTipoUsuario()
    On Error GoTo ErrEditarTipoUsuario
    
    If lvConsultas.ListItems.Count > 0 Then
        gsCodTipoUsuario = mItem
        frmEditarTipoUsuario.Show vbModal
        If Not gbCancelar Then
            msCodUltTipoUsuario = gsCodTipoUsuario
            Call CargaTiposUsuarios(True)
        End If
    End If
    
    Exit Sub
    
ErrEditarTipoUsuario:
    MsgBox Error, vbCritical, App.Title
End Sub

Sub ArcEditarUsuario()
    On Error GoTo ErrEditarUsuario
    
    If lvConsultas.ListItems.Count > 0 Then
        gsNomUsuario = mItem
        frmEditarUsuario.Show vbModal
        If Not gbCancelar Then
            msNomUltUsuario = gsNomUsuario
            Call CargaUsuarios(True)
        End If
    End If
    
    Exit Sub
    
ErrEditarUsuario:
    MsgBox Error, vbCritical, App.Title
End Sub

Sub ArcEliminarBaseDatos()
    Dim sNumBaseDatos   As String
    Dim sNomBaseDatos   As String
    Dim bOk             As Boolean
    
    On Error GoTo ErrArcEliminarBaseDatos
    
    sNumBaseDatos = mItem
    sNomBaseDatos = mItem.SubItems(1)
    
    If MsgBox("Está seguro que desea eliminar la base de datos """ & sNomBaseDatos & """", vbYesNoCancel + vbDefaultButton2, App.Title) <> vbYes Then
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
    ' Elimina consulta
    bOk = db_EliminaBaseDatos(sNumBaseDatos)
    
    Screen.MousePointer = vbNormal
    
    If bOk Then
        MsgBox "Base de datos fue eliminada", vbInformation, App.Title
        msNumUltBaseDatos = ""
        Call CargaBaseDatos(True)
    End If
    
    Exit Sub
    
ErrArcEliminarBaseDatos:
    Screen.MousePointer = vbNormal
    MsgBox Error, vbCritical, App.Title
    Exit Sub
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
    
    Screen.MousePointer = vbHourglass
    
    ' Elimina consulta
    bOk = db_EliminaConsulta(sNumConsulta)
    
    Screen.MousePointer = vbNormal
    
    If bOk Then
        MsgBox "Consulta fue eliminada", vbInformation, App.Title
        msNumUltConsultas = ""
        Call CargaConsultas(True)
    End If
    
    Exit Sub
    
ErrArcEliminarConsulta:
    Screen.MousePointer = vbNormal
    MsgBox Error, vbCritical, App.Title
    Exit Sub
End Sub

Sub ArcEliminarElemento()
    Select Case mnVistaActual
    Case mnVistaUsuarios
        ArcEliminarUsuario
    Case mnVistaConsultas
        ArcEliminarConsulta
    Case mnVistaPerfiles
        ArcEliminarAgrupacion
    '<V1.3.0>
    Case mnVistaLotes
        ArcEliminarLotes
    '</V1.3.0>
    Case mnVistaTiposUsuarios
        ArcEliminarTipoUsuario
    Case mnVistaBaseDatos
        ArcEliminarBaseDatos
    '<V1.3.1>
    Case mnVistaTabValores
        ArcEliminarTabValores
    '</V1.3.1>
    End Select
End Sub

Sub ArcEliminarAgrupacion()
    Dim sNumPerfil  As String
    Dim sNomPerfil  As String
    Dim bOk         As Boolean
    
    On Error GoTo ErrArcEliminarAgrupacion
    
    sNumPerfil = mItem
    sNomPerfil = mItem.SubItems(1)
    
    If MsgBox("Está seguro que desea eliminar esta agrupación """ & sNomPerfil & """", vbYesNoCancel + vbDefaultButton2, App.Title) <> vbYes Then
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
    ' Elimina consulta
    bOk = db_EliminaAgrupacion(sNumPerfil)
    
    Screen.MousePointer = vbNormal
    
    If bOk Then
        MsgBox "Agrupación fue eliminada", vbInformation, App.Title
        msNumUltPerfil = ""
        Call CargaAgrupaciones(True)
    End If
    
    Exit Sub
    
ErrArcEliminarAgrupacion:
    Screen.MousePointer = vbNormal
    MsgBox Error, vbCritical, App.Title
    Exit Sub
End Sub

Sub ArcEliminarLotes()
'<V1.3.0>
    Dim sNumLote  As String
    Dim sNomLote  As String
    Dim bOk         As Boolean
    
    On Error GoTo ErrArcEliminarLote
    
    sNumLote = mItem
    sNomLote = mItem.SubItems(1)
    
    If MsgBox("Está seguro que desea eliminar este a lote """ & sNomLote & """", vbYesNoCancel + vbDefaultButton2, App.Title) <> vbYes Then
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
    ' Elimina lote
    bOk = db_EliminaLote(sNumLote)
    
    Screen.MousePointer = vbNormal
    
    If bOk Then
        MsgBox "Lote fue eliminado", vbInformation, App.Title
        msNumUltLote = ""
        Call CargaLotes(True)
    End If
    
    Exit Sub
    
ErrArcEliminarLote:
    Screen.MousePointer = vbNormal
    MsgBox Error, vbCritical, App.Title
    Exit Sub
'</V1.3.0>
End Sub

Sub ArcEliminarTabValores()
    Dim sNumRegistro    As String
    Dim bOk             As Boolean
    
    On Error GoTo ErrArcEliminarTabValores
    
    sNumRegistro = mItem
    
    If MsgBox("Está seguro que desea eliminar el registro """ & sNumRegistro & """", vbYesNoCancel + vbDefaultButton2, App.Title) <> vbYes Then
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
    ' Elimina tab_valores
    bOk = db_EliminaTabValores(sNumRegistro)
    
    Screen.MousePointer = vbNormal
    
    If bOk Then
        MsgBox "Registro fue eliminado", vbInformation, App.Title
        msNumUltRegTabValor = ""
        Call CargaTabValores(True)
    End If
    
    Exit Sub
    
ErrArcEliminarTabValores:
    Screen.MousePointer = vbNormal
    MsgBox Error, vbCritical, App.Title
    Exit Sub
End Sub

Sub ArcEliminarTipoUsuario()
    Dim sCodTipoUsuario As String
    Dim bOk             As Boolean
    
    On Error GoTo ErrArcEliminarTipoUsuario
    
    sCodTipoUsuario = mItem
    
    If MsgBox("Está seguro que desea eliminar el tipo de usuario """ & sCodTipoUsuario & """", vbYesNoCancel + vbDefaultButton2, App.Title) <> vbYes Then
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
    ' Elimina consulta
    bOk = db_EliminaTipoUsuario(sCodTipoUsuario)
    
    Screen.MousePointer = vbNormal
    
    If bOk Then
        MsgBox "Tipo de usuario fue eliminado", vbInformation, App.Title
        msCodUltTipoUsuario = ""
        Call CargaTiposUsuarios(True)
    End If
    
    Exit Sub
    
ErrArcEliminarTipoUsuario:
    Screen.MousePointer = vbNormal
    MsgBox Error, vbCritical, App.Title
    Exit Sub
End Sub

Sub ArcEliminarUsuario()
    Dim sNomUsuario    As String
    Dim bOk             As Boolean
    
    On Error GoTo ErrArcEliminarUsuario
    
    sNomUsuario = mItem
    
    If MsgBox("Está seguro que desea eliminar el usuario """ & sNomUsuario & """", vbYesNoCancel + vbDefaultButton2, App.Title) <> vbYes Then
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
    ' Elimina consulta
    bOk = db_EliminaUsuario(sNomUsuario)
    
    Screen.MousePointer = vbNormal
    
    If bOk Then
        MsgBox "Usuario fue eliminado", vbInformation, App.Title
        msNomUltUsuario = ""
        Call CargaUsuarios(True)
    End If
    
    Exit Sub
    
ErrArcEliminarUsuario:
    Screen.MousePointer = vbNormal
    MsgBox Error, vbCritical, App.Title
    Exit Sub
End Sub

Sub ArcNuevoBaseDatos()
    On Error GoTo ErrArcNuevoBaseDatos
    
    gsNumBaseDatos = ""
    frmEditarBaseDatos.Show vbModal
    If Not gbCancelar Then
        msNumUltBaseDatos = gsNumBaseDatos
        Call CargaBaseDatos(True)
    End If
    
    Exit Sub
    
ErrArcNuevoBaseDatos:
    MsgBox Error, vbCritical, App.Title
End Sub

Sub ArcNuevoElemento()
    Select Case mnVistaActual
    Case mnVistaUsuarios
        ArcNuevoUsuario
    Case mnVistaConsultas
        ArcNuevaConsulta
    Case mnVistaPerfiles
        ArcNuevaAgrupacion
    '<V1.3.0>
    Case mnVistaLotes
        ArcNuevoLote
    '</V1.3.0>
    Case mnVistaTiposUsuarios
        ArcNuevoTipoUsuario
    Case mnVistaBaseDatos
        ArcNuevoBaseDatos
    '<V1.3.1>
    Case mnVistaTabValores
        ArcNuevoTabValores
    '</V1.3.1>
    End Select
End Sub

Sub ArcNuevaAgrupacion()
    gsNumPerfil = ""
    gsNomPerfil = ""
    frmEditarPerfil.Show vbModal
    If Not gbCancelar Then
        msNumUltPerfil = gsNumPerfil
        Call CargaAgrupaciones(True)
    End If
End Sub

Sub ArcNuevoLote()
'<V1.3.0>
    gsNumLote = ""
    gsNomLote = ""
    gsNomSolicitante = ""
    gsNomUsuarioLocal = gsUsuarioReal
    frmEditarLote.Show vbModal
    If Not gbCancelar Then
        msNumUltLote = gsNumLote
        Call CargaLotes(True)
    End If
'</V1.3.0>
End Sub

Sub ArcNuevoTabValores()
    '<V1.3.1>
    On Error GoTo ErrNuevoTabValores
    
    gsNumRegTabValor = ""
    frmEditarTabValores.Show vbModal
    If Not gbCancelar Then
        msNumUltRegTabValor = ""
        Call CargaTabValores(True)
    End If
    
    Exit Sub
    
ErrNuevoTabValores:
    MsgBox Error, vbCritical, App.Title
    '</V1.3.1>
End Sub

Sub ArcNuevoTipoUsuario()
    On Error GoTo ErrEditarTipoUsuario
    
    gsCodTipoUsuario = ""
    frmEditarTipoUsuario.Show vbModal
    If Not gbCancelar Then
        msCodUltTipoUsuario = gsCodTipoUsuario
        Call CargaTiposUsuarios(True)
    End If
    
    Exit Sub
    
ErrEditarTipoUsuario:
    MsgBox Error, vbCritical, App.Title
End Sub

Sub ArcNuevoUsuario()
    gsNomUsuario = ""
    frmEditarUsuario.Show vbModal
    If Not gbCancelar Then
        msNomUltUsuario = gsNomUsuario
        Call CargaUsuarios(True)
    End If
End Sub

Sub ArcCerrarVentana()
    Unload Me
End Sub

Sub AsigConsultas()
    Select Case mnVistaActual
    Case mnVistaUsuarios ' Usuarios
        AsigConsultasPorUsuario
    Case mnVistaPerfiles ' Perfiles
        AsigConsultasPorAgrupacion
    End Select
End Sub

Sub AsigConsultasPorAgrupacion()
    gsNumPerfil = mItem
    
    frmConsPerfil.lblNumPerfil = gsNumPerfil
    frmConsPerfil.lblNomPerfil = mItem.SubItems(1)
    frmConsPerfil.lblFecCreacion = mItem.SubItems(2)
    
    frmConsPerfil.Show vbModal
End Sub

Sub AsigConsultasPorUsuario()
    gsNomUsuario = mItem
    
    frmConsUsuario.lblNomUsuario = gsNomUsuario
    frmConsUsuario.lblCodTipoUsuario = mItem.SubItems(1)
    
    frmConsUsuario.Show vbModal
End Sub

Sub AsigAgrupacion()
    Select Case mnVistaActual
    Case mnVistaUsuarios
        AsigAgrupacionPorUsuario
    Case mnVistaConsultas
        AsigAgrupacionPorConsulta
    End Select
End Sub

Sub AsigAgrupacionPorConsulta()
    gsNumConsulta = mItem
    
    frmPerfConsulta.lblNumConsulta = gsNumConsulta
    frmPerfConsulta.lblNomConsulta = mItem.SubItems(1)
    frmPerfConsulta.lblNomCreador = mItem.SubItems(3)
    frmPerfConsulta.lblFecCreación = mItem.SubItems(4)
        
    frmPerfConsulta.Show vbModal
End Sub

Sub AsigAgrupacionPorUsuario()
    gsNomUsuario = mItem
    
    frmPerfUsuario.lblNomUsuario = gsNomUsuario
    frmPerfUsuario.lblCodTipoUsuario = mItem.SubItems(1)
    
    frmPerfUsuario.Show vbModal
End Sub

Sub AsigLotes()
'<V1.3.0>
    Select Case mnVistaActual
    Case mnVistaUsuarios
        AsigLotesPorUsuario
    End Select
'</V1.3.0>
End Sub

Sub AsigLotesPorUsuario()
'<V1.3.0>
    gsNomUsuario = mItem
    
    frmLoteUsuario.lblNomUsuario = gsNomUsuario
    frmLoteUsuario.lblCodTipoUsuario = mItem.SubItems(1)
    
    frmLoteUsuario.Show vbModal
'</V1.3.0>
End Sub

Sub CargaTabValores(bIndLeerData As Boolean)
    '<V1.3.1>
    Dim nItem       As Integer
        
    On Error GoTo ErrCargaTabValores
            
    Screen.MousePointer = vbHourglass
    
    lvConsultas.Visible = False
    IniciaListaTabValores
    nItem = 1
    
    '<V1.3.1>
    lvConsultas.ColumnHeaders(1).Key = "num_registro"
    lvConsultas.ColumnHeaders(2).Key = "cod_tabla"
    lvConsultas.ColumnHeaders(3).Key = "gls_valor"
    
    If bIndLeerData Then
        IniciaRecordSetFiltro
        Call db_LeeTabValores("", rsData)
    End If
    
    If Not (rsData Is Nothing) Then
        rsData.Filter = "cod_tabla <> 'ADMIN'"
        While Not rsData.EOF
            Set mItem = Me.lvConsultas.ListItems.Add(, , "" & rsData!num_registro)
            mItem.SubItems(1) = "" & rsData!cod_tabla
            mItem.SubItems(2) = "" & rsData!gls_valor

            If bIndLeerData Then
                Call GuardaDistinct(1, mItem.SubItems(1))
                Call GuardaDistinct(2, mItem.SubItems(2))
            End If

            If msNumUltRegTabValor = rsData!num_registro Then
                nItem = Me.lvConsultas.ListItems.Count
            End If
            
            rsData.MoveNext
        Wend
    End If
    
    If Me.lvConsultas.ListItems.Count > 0 Then
        Set mItem = lvConsultas.ListItems.Item(nItem)
        mItem.EnsureVisible
        mItem.Selected = True
        On Error Resume Next
        lvConsultas.SetFocus
    End If
    
    ActivarBotones
    
    lvConsultas.Visible = True
    Screen.MousePointer = vbNormal
    
    Exit Sub
    
ErrCargaTabValores:
    Screen.MousePointer = vbNormal
    MsgBox Error, vbCritical, App.Title
    lvConsultas.Visible = True
    Exit Sub
    '</V1.3.1>
End Sub

Sub FiltroAsignar(nColHeader As Long, ByVal sGlsValor As String)
    '<V1.3.1>
    Dim i           As Integer
    Dim j           As Integer
    Dim sGlsFiltro  As String
    
    For i = 1 To frmAdministracion.lvConsultas.ColumnHeaders.Count
        frmAdministracion.lvConsultas.ColumnHeaders(i).Tag = ""
    Next i
        
    If sGlsValor = gsGlsValorBlanco Then sGlsValor = ""
    lvConsultas.ColumnHeaders(nColHeader + 1).Tag = " =  " & Replace(sGlsValor, "'", "''")
    '</V1.3.1>
End Sub

Sub AsigUsuarios()
    Select Case mnVistaActual
    Case mnVistaConsultas
        AsigUsuariosPorConsulta
    Case mnVistaPerfiles
        AsigUsuariosPorAgrupacion
    Case mnVistaLotes
        AsigUsuariosPorLote
    End Select
End Sub

Sub AsigUsuariosPorConsulta()
    gsNumConsulta = mItem
    
    frmUsuaConsulta.lblNumConsulta = gsNumConsulta
    frmUsuaConsulta.lblNomConsulta = mItem.SubItems(1)
    frmUsuaConsulta.lblNomCreador = mItem.SubItems(3)
    frmUsuaConsulta.lblFecCreación = mItem.SubItems(4)
    
    frmUsuaConsulta.Show vbModal
End Sub

Sub AsigUsuariosPorAgrupacion()
    gsNumPerfil = mItem
    
    frmUsuaPerfil.lblNumPerfil = gsNumPerfil
    frmUsuaPerfil.lblNomPerfil = mItem.SubItems(1)
    frmUsuaPerfil.lblFecCreación = mItem.SubItems(2)
    
    frmUsuaPerfil.Show vbModal
End Sub

Sub AsigUsuariosPorLote()
'<V1.3.0>
    gsNumLote = mItem
    gsNomLote = mItem.SubItems(1)
    
    frmUsuaLote.lblNomLote = gsNomLote
    frmUsuaLote.lblNumLote = gsNumLote
    frmUsuaLote.lblNomSolicitante = mItem.SubItems(2)
    frmUsuaLote.lblFecCreación = mItem.SubItems(3)
    
    frmUsuaLote.Show vbModal
'</V1.3.0>
End Sub

Sub CambiaIcono(sIndBloqueada As String)
    '<V1.3.1>
    mItem.SubItems(7) = sIndBloqueada
    If sIndBloqueada = "S" Then
        mItem.SmallIcon = "ConsBloq_S"
    Else
        mItem.SmallIcon = "ConsBloq_N"
    End If
    '</V1.3.1>
End Sub

Sub CargaBaseDatos(bIndLeerData As Boolean)
    Dim nItem       As Integer
        
    On Error GoTo ErrCargaBaseDatos
            
    Screen.MousePointer = vbHourglass
    
    lvConsultas.Visible = False
    IniciaListaBaseDatos
    nItem = 1
    
    '<V1.3.1>
    lvConsultas.ColumnHeaders(1).Key = "num_basedatos"
    lvConsultas.ColumnHeaders(2).Key = "nom_basedatos"
    lvConsultas.ColumnHeaders(3).Key = "gls_coneccion"
    lvConsultas.ColumnHeaders(4).Key = "gls_formato_fecha"
    
    If bIndLeerData Then
        IniciaRecordSetFiltro
        Call db_LeeBasesDeDatos(rsData)
    End If
    
    If Not (rsData Is Nothing) Then
    '</V1.3.1>
        While Not rsData.EOF
            Set mItem = Me.lvConsultas.ListItems.Add(, , rsData!num_basedatos)
            mItem.SubItems(1) = "" & rsData!nom_basedatos
            mItem.SubItems(2) = "" & rsData!gls_coneccion
            mItem.SubItems(3) = "" & rsData!gls_formato_fecha

            '<V1.3.1>
            If bIndLeerData Then
                Call GuardaDistinct(1, mItem.SubItems(1))
                Call GuardaDistinct(2, mItem.SubItems(2))
                Call GuardaDistinct(3, mItem.SubItems(3))
            End If
            '</V1.3.1>

            If msNumUltBaseDatos = rsData!num_basedatos Then
                nItem = Me.lvConsultas.ListItems.Count
            End If
            
            rsData.MoveNext
        Wend
    End If
        
    If Me.lvConsultas.ListItems.Count > 0 Then
        Set mItem = lvConsultas.ListItems.Item(nItem)
        mItem.EnsureVisible
        mItem.Selected = True
        On Error Resume Next
        lvConsultas.SetFocus
    End If
    
    ActivarBotones
    
    lvConsultas.Visible = True
    Screen.MousePointer = vbNormal
    
    Exit Sub
    
ErrCargaBaseDatos:
    Screen.MousePointer = vbNormal
    MsgBox Error, vbCritical, App.Title
    lvConsultas.Visible = True
    Exit Sub
    Resume
End Sub

Sub CargaAgrupaciones(bIndLeerData As Boolean)
    Dim nItem       As Integer
        
    On Error GoTo ErrCargaAgrupaciones
            
    Screen.MousePointer = vbHourglass
    
    lvConsultas.Visible = False
    IniciaListaAgrupaciones
    nItem = 1
    
    '<V1.3.1>
    lvConsultas.ColumnHeaders(1).Key = "num_perfil"
    lvConsultas.ColumnHeaders(2).Key = "nom_perfil"
    lvConsultas.ColumnHeaders(3).Key = "fec_creacion"
    
    If bIndLeerData Then
        IniciaRecordSetFiltro
        Call db_LeeAgrupaciones(rsData)
    End If
    
    If Not (rsData Is Nothing) Then
    '</V1.3.1>
        While Not rsData.EOF
            Set mItem = Me.lvConsultas.ListItems.Add(, , "" & rsData!num_perfil)
            mItem.SubItems(1) = "" & rsData!nom_perfil
            mItem.SubItems(2) = "" & rsData!fec_creacion

            '<V1.3.1>
            If bIndLeerData Then
                Call GuardaDistinct(1, mItem.SubItems(1))
                Call GuardaDistinct(2, mItem.SubItems(2))
            End If
            '</V1.3.1>
            
            If msNumUltPerfil = rsData!num_perfil Then
                nItem = Me.lvConsultas.ListItems.Count
            End If
            
            rsData.MoveNext
        Wend
    End If
        
    If Me.lvConsultas.ListItems.Count > 0 Then
        Set mItem = lvConsultas.ListItems.Item(nItem)
        mItem.EnsureVisible
        mItem.Selected = True
        On Error Resume Next
        lvConsultas.SetFocus
    End If
    
    ActivarBotones
    
    lvConsultas.Visible = True
    Screen.MousePointer = vbNormal
    
    Exit Sub
    
ErrCargaAgrupaciones:
    Screen.MousePointer = vbNormal
    MsgBox Error, vbCritical, App.Title
    lvConsultas.Visible = True
    Exit Sub
End Sub

Sub CargaLotes(bIndLeerData As Boolean)
    '<V1.3.0>
    Dim nItem       As Integer
        
    On Error GoTo ErrCargaLotes
            
    Screen.MousePointer = vbHourglass
    
    lvConsultas.Visible = False
    IniciaListaLotes
    nItem = 1
    
    '<V1.3.1>
    lvConsultas.ColumnHeaders(1).Key = "num_lote"
    lvConsultas.ColumnHeaders(2).Key = "nom_lote"
    lvConsultas.ColumnHeaders(3).Key = "nom_solicitante"
    lvConsultas.ColumnHeaders(4).Key = "fec_creacion"
    
    If bIndLeerData Then
        IniciaRecordSetFiltro
        Call db_LeeLotes(rsData)
    End If
    
    If Not (rsData Is Nothing) Then
    '</V1.3.1>
        While Not rsData.EOF
            Set mItem = Me.lvConsultas.ListItems.Add(, , "" & rsData!num_lote)
            mItem.SubItems(1) = "" & rsData!nom_lote
            mItem.SubItems(2) = "" & rsData!nom_solicitante
            mItem.SubItems(3) = "" & rsData!fec_creacion

            '<V1.3.1>
            If bIndLeerData Then
                Call GuardaDistinct(1, mItem.SubItems(1))
                Call GuardaDistinct(2, mItem.SubItems(2))
                Call GuardaDistinct(3, mItem.SubItems(3))
            End If
            '</V1.3.1>

            If msNumUltLote = rsData!num_lote Then
                nItem = Me.lvConsultas.ListItems.Count
            End If
            
            rsData.MoveNext
        Wend
    End If
    
    If Me.lvConsultas.ListItems.Count > 0 Then
        Set mItem = lvConsultas.ListItems.Item(nItem)
        mItem.EnsureVisible
        mItem.Selected = True
        On Error Resume Next
        lvConsultas.SetFocus
    End If
    
    ActivarBotones
    
    lvConsultas.Visible = True
    Screen.MousePointer = vbNormal
    
    Exit Sub
    
ErrCargaLotes:
    Screen.MousePointer = vbNormal
    MsgBox Error, vbCritical, App.Title
    lvConsultas.Visible = True
    Exit Sub
    '</V1.3.0>
End Sub

Sub CargaTiposUsuarios(bIndLeerData As Boolean)
    Dim nItem       As Integer
        
    On Error GoTo ErrCargaTiposUsuarios
            
    Screen.MousePointer = vbHourglass
    
    lvConsultas.Visible = False
    IniciaListaTiposUsuarios
    nItem = 1
    
    '<V1.3.1>
    lvConsultas.ColumnHeaders(1).Key = "cod_tipo_usuario"
    lvConsultas.ColumnHeaders(2).Key = "ind_administrador"
    lvConsultas.ColumnHeaders(3).Key = "ind_crear_consultas"
    lvConsultas.ColumnHeaders(4).Key = "ind_autoasignar_consultas"
    lvConsultas.ColumnHeaders(5).Key = "ind_modificar_consultas"
    lvConsultas.ColumnHeaders(6).Key = "ind_eliminar_consultas"
    lvConsultas.ColumnHeaders(7).Key = "ind_ejecutar_consultas"
    
    If bIndLeerData Then
        IniciaRecordSetFiltro
        OpenMyDataBase
        Call db_LeeTiposUsuarios(rsData)
        CloseMyDataBase
    End If
    
    If Not (rsData Is Nothing) Then
    '</V1.3.1>
        While Not rsData.EOF
            Set mItem = Me.lvConsultas.ListItems.Add(, , "" & rsData!cod_tipo_usuario)
            mItem.SubItems(1) = "" & rsData!ind_administrador
            mItem.SubItems(2) = "" & rsData!ind_crear_consultas
            mItem.SubItems(3) = "" & rsData!ind_autoasignar_consultas
            mItem.SubItems(4) = "" & rsData!ind_modificar_consultas
            mItem.SubItems(5) = "" & rsData!ind_eliminar_consultas
            mItem.SubItems(6) = "" & rsData!ind_ejecutar_consultas

            '<V1.3.1>
            If bIndLeerData Then
                Call GuardaDistinct(1, mItem.SubItems(1))
                Call GuardaDistinct(2, mItem.SubItems(2))
                Call GuardaDistinct(3, mItem.SubItems(3))
                Call GuardaDistinct(4, mItem.SubItems(4))
                Call GuardaDistinct(5, mItem.SubItems(5))
                Call GuardaDistinct(6, mItem.SubItems(6))
            End If
            '</V1.3.1>

            If msCodUltTipoUsuario = "" & rsData!cod_tipo_usuario Then
                nItem = Me.lvConsultas.ListItems.Count
            End If
            
            rsData.MoveNext
        Wend
    End If
    
    If Me.lvConsultas.ListItems.Count > 0 Then
        Set mItem = lvConsultas.ListItems.Item(nItem)
        mItem.EnsureVisible
        mItem.Selected = True
        On Error Resume Next
        lvConsultas.SetFocus
    End If
    
    ActivarBotones
    
    lvConsultas.Visible = True
    Screen.MousePointer = vbNormal
    
    Exit Sub
    
ErrCargaTiposUsuarios:
    Screen.MousePointer = vbNormal
    MsgBox Error, vbCritical, App.Title
    lvConsultas.Visible = True
    Exit Sub
End Sub

Sub CargaUsuarios(bIndLeerData As Boolean)
    Dim nItem       As Integer
        
    On Error GoTo ErrCargaUsuarios
            
    Screen.MousePointer = vbHourglass
    
    lvConsultas.Visible = False
    IniciaListaUsuarios
    nItem = 1
            
    '<V1.3.1>
    lvConsultas.ColumnHeaders(1).Key = "nom_usuario"
    lvConsultas.ColumnHeaders(2).Key = "cod_tipo_usuario"
    lvConsultas.ColumnHeaders(3).Key = "ind_administrador"
    lvConsultas.ColumnHeaders(4).Key = "ind_crear_consultas"
    lvConsultas.ColumnHeaders(5).Key = "ind_modificar_consultas"
    lvConsultas.ColumnHeaders(6).Key = "ind_eliminar_consultas"
    lvConsultas.ColumnHeaders(7).Key = "ind_ejecutar_consultas"
    
    If bIndLeerData Then
        IniciaRecordSetFiltro
        Call db_LeeUsuarios(rsData)
    End If
    
    If Not (rsData Is Nothing) Then
    '</V1.3.1>
        While Not rsData.EOF
            Set mItem = Me.lvConsultas.ListItems.Add(, , "" & rsData!nom_usuario)
            mItem.SubItems(1) = "" & rsData!cod_tipo_usuario
            mItem.SubItems(2) = "" & rsData!ind_administrador
            mItem.SubItems(3) = "" & rsData!ind_crear_consultas
            mItem.SubItems(4) = "" & rsData!ind_modificar_consultas
            mItem.SubItems(5) = "" & rsData!ind_eliminar_consultas
            mItem.SubItems(6) = "" & rsData!ind_ejecutar_consultas

            '<V1.3.1>
            If bIndLeerData Then
                Call GuardaDistinct(1, mItem.SubItems(1))
                Call GuardaDistinct(2, mItem.SubItems(2))
                Call GuardaDistinct(3, mItem.SubItems(3))
                Call GuardaDistinct(4, mItem.SubItems(4))
                Call GuardaDistinct(5, mItem.SubItems(5))
                Call GuardaDistinct(6, mItem.SubItems(6))
            End If
            '</V1.3.1>

            If msNomUltUsuario = rsData!nom_usuario Then
                nItem = Me.lvConsultas.ListItems.Count
            End If
            
            rsData.MoveNext
        Wend
    End If
    
    If Me.lvConsultas.ListItems.Count > 0 Then
        Set mItem = lvConsultas.ListItems.Item(nItem)
        mItem.EnsureVisible
        mItem.Selected = True
        On Error Resume Next
        lvConsultas.SetFocus
    End If
    
    ActivarBotones
    
    lvConsultas.Visible = True
    Screen.MousePointer = vbNormal
    
    Exit Sub
    
ErrCargaUsuarios:
    Screen.MousePointer = vbNormal
    MsgBox Error, vbCritical, App.Title
    lvConsultas.Visible = True
    Exit Sub
End Sub

Sub FormResize()
    fraVista.Width = Me.Width - 120
    fraVista.Height = Me.Height - fraVista.Top - 480
    
    lvConsultas.Top = 240
    lvConsultas.Left = 60
    lvConsultas.Width = fraVista.Width - 120
    lvConsultas.Height = fraVista.Height - 360
End Sub

Sub GuardaDistinct(nItem As Long, sSubItem As String)
    '<V1.3.1>
    Dim sNomCampo   As String
    
    sNomCampo = Me.lvConsultas.ColumnHeaders(nItem + 1).Key
    rsDataFiltro.Filter = "nom_campo='" & sNomCampo & "' and gls_valor='" & Replace(sSubItem, "'", "''") & "'"
    If rsDataFiltro.EOF Then
        rsDataFiltro.AddNew Array("nom_campo", "gls_valor"), Array(sNomCampo, sSubItem)
        rsDataFiltro.Update
    End If
    '</V1.3.1>
End Sub

Sub IniciaListaBaseDatos()
    lvConsultas.ColumnHeaders.Clear
    lvConsultas.ColumnHeaders.Add , , "Id", 500
    lvConsultas.ColumnHeaders.Add , , "Nombre", 2000
    lvConsultas.ColumnHeaders.Add , , "String de Conexión", 8000
    lvConsultas.ColumnHeaders.Add , , "Formato Fechas", 2000
    lvConsultas.ListItems.Clear
    
    fraVista.Caption = "Vista de todas las bases de datos"
    mnuArcNuevo.Caption = "&Nueva base de datos"
    mnuArcEditar.Caption = "&Editar base de datos"
    mnuArcEliminar.Caption = "E&liminar base de datos"
    
    frmMdiPadre.Toolbar1(2).Buttons(1).ToolTipText = "Nueva base de datos"
    frmMdiPadre.Toolbar1(2).Buttons(2).ToolTipText = "Editar base de datos"
    frmMdiPadre.Toolbar1(2).Buttons(3).ToolTipText = "Eliminar base de datos"
End Sub

Sub IniciaListaAgrupaciones()
    lvConsultas.ColumnHeaders.Clear
    lvConsultas.ColumnHeaders.Add , , "Id", 1000
    lvConsultas.ColumnHeaders.Add , , "Nombre", 2000
    lvConsultas.ColumnHeaders.Add , , "Fec.Creacion", 2000
    lvConsultas.ListItems.Clear
    
    fraVista.Caption = "Vista de todas las agrupaciones"
    mnuArcNuevo.Caption = "&Nueva agrupación"
    mnuArcEditar.Caption = "&Editar agrupación"
    mnuArcEliminar.Caption = "E&liminar agrupación"
    
    frmMdiPadre.Toolbar1(2).Buttons(1).ToolTipText = "Nueva agrupación"
    frmMdiPadre.Toolbar1(2).Buttons(2).ToolTipText = "Editar agrupación"
    frmMdiPadre.Toolbar1(2).Buttons(3).ToolTipText = "Eliminar agrupación"
End Sub

Sub IniciaListaLotes()
'<V1.3.0>
    lvConsultas.ColumnHeaders.Clear
    lvConsultas.ColumnHeaders.Add , , "Id", 1000
    lvConsultas.ColumnHeaders.Add , , "Nombre", 2000
    lvConsultas.ColumnHeaders.Add , , "Solicitante", 2000
    lvConsultas.ColumnHeaders.Add , , "Fec.Creacion", 2200
    lvConsultas.ListItems.Clear
    
    fraVista.Caption = "Vista de todos los lotes"
    mnuArcNuevo.Caption = "&Nuevo lote"
    mnuArcEditar.Caption = "&Editar lote"
    mnuArcEliminar.Caption = "E&liminar lote"
    
    frmMdiPadre.Toolbar1(2).Buttons(1).ToolTipText = "Nuevo lote"
    frmMdiPadre.Toolbar1(2).Buttons(2).ToolTipText = "Editar lote"
    frmMdiPadre.Toolbar1(2).Buttons(3).ToolTipText = "Eliminar lote"
'</V1.3.0>
End Sub

Sub IniciaListaTabValores()
    '<V1.3.1>
    lvConsultas.ColumnHeaders.Clear
    lvConsultas.ColumnHeaders.Add , , "Id", 1000
    lvConsultas.ColumnHeaders.Add , , "Tabla", 2000
    lvConsultas.ColumnHeaders.Add , , "Valor", 5000
    lvConsultas.ListItems.Clear
    
    fraVista.Caption = "Vista de tabla de valores"
    mnuArcNuevo.Caption = "&Nuevo valor"
    mnuArcEditar.Caption = "&Editar valor"
    mnuArcEliminar.Caption = "E&liminar valor"
    
    frmMdiPadre.Toolbar1(2).Buttons(1).ToolTipText = "Nuevo valor"
    frmMdiPadre.Toolbar1(2).Buttons(2).ToolTipText = "Editar valor"
    frmMdiPadre.Toolbar1(2).Buttons(3).ToolTipText = "Eliminar valor"
    '</V1.3.1>
End Sub

Sub IniciaListaTiposUsuarios()
    lvConsultas.ColumnHeaders.Clear
    lvConsultas.ColumnHeaders.Add , , "Tipo", 800
    lvConsultas.ColumnHeaders.Add , , "Administrador", 1200
    lvConsultas.ColumnHeaders.Add , , "Crea Consultas", 1500
    lvConsultas.ColumnHeaders.Add , , "Asigna Consultas", 1500
    lvConsultas.ColumnHeaders.Add , , "Modifica Consultas", 1500
    lvConsultas.ColumnHeaders.Add , , "Elimina Consultas", 1500
    lvConsultas.ColumnHeaders.Add , , "Ejecuta Consultas", 1500
    lvConsultas.ListItems.Clear
    
    fraVista.Caption = "Vista de todos los tipos de usuarios"
    mnuArcNuevo.Caption = "&Nuevo tipo de usuario"
    mnuArcEditar.Caption = "&Editar tipo de usuario"
    mnuArcEliminar.Caption = "E&liminar tipo de usuario"
    
    frmMdiPadre.Toolbar1(2).Buttons(1).ToolTipText = "Nuevo tipo de usuario"
    frmMdiPadre.Toolbar1(2).Buttons(2).ToolTipText = "Editar tipo de usuario"
    frmMdiPadre.Toolbar1(2).Buttons(3).ToolTipText = "Eliminar tipo de usuario"
End Sub

Sub IniciaListaUsuarios()
    lvConsultas.ColumnHeaders.Clear
    lvConsultas.ColumnHeaders.Add , , "Usuario", 2000
    lvConsultas.ColumnHeaders.Add , , "Tipo", 800
    lvConsultas.ColumnHeaders.Add , , "Administrador", 1200
    lvConsultas.ColumnHeaders.Add , , "Crea Consultas", 1500
    lvConsultas.ColumnHeaders.Add , , "Modifica Consultas", 1500
    lvConsultas.ColumnHeaders.Add , , "Elimina Consultas", 1500
    lvConsultas.ColumnHeaders.Add , , "Ejecuta Consultas", 1500
    lvConsultas.ListItems.Clear
    
    fraVista.Caption = "Vista de todos los usuarios"
    mnuArcNuevo.Caption = "&Nuevo usuario"
    mnuArcEditar.Caption = "&Editar usuario"
    mnuArcEliminar.Caption = "E&liminar usuario"
    
    frmMdiPadre.Toolbar1(2).Buttons(1).ToolTipText = "Nuevo usuario"
    frmMdiPadre.Toolbar1(2).Buttons(2).ToolTipText = "Editar usuario"
    frmMdiPadre.Toolbar1(2).Buttons(3).ToolTipText = "Eliminar usuario"
End Sub

Sub IniciaRecordSetFiltro()
    Set rsDataFiltro = Nothing
    Set rsDataFiltro = New ADODB.Recordset
    With rsDataFiltro.Fields
        .Append "nom_campo", adVarChar, 32
        .Append "gls_valor", adVarChar, 255
    End With
    rsDataFiltro.Open , , adOpenStatic, adLockOptimistic
End Sub

Sub MuestraVista(Index As Integer, bIndLeeData As Boolean)
    '<V1.3.1>
    Select Case Index
    Case mnVistaUsuarios
        Call CargaUsuarios(bIndLeeData)
    Case mnVistaConsultas
        Call CargaConsultas(bIndLeeData)
    Case mnVistaPerfiles
        Call CargaAgrupaciones(bIndLeeData)
    '<V1.3.0>
    Case mnVistaLotes
        Call CargaLotes(bIndLeeData)
    '</V1.3.0>
    Case mnVistaTiposUsuarios
        Call CargaTiposUsuarios(bIndLeeData)
    Case mnVistaBaseDatos
        Call CargaBaseDatos(bIndLeeData)
    '<V1.3.0>
    Case mnVistaTabValores
        Call CargaTabValores(bIndLeeData)
    '</V1.3.0>
    End Select
    '</V1.3.1>
    
    SaveSetting App.Title, "Settings", "CurrentView", Index
End Sub

Sub OrdenaPorColumna()
    '<V1.3.1>
    lvConsultas.SortKey = mnNumColumnaActiva
    If lvConsultas.SortOrder = lvwAscending Then
        lvConsultas.SortOrder = lvwDescending
    Else
        lvConsultas.SortOrder = lvwAscending
    End If
    
    ' Establece Verdadero en Sorted para ordenar la lista.
    lvConsultas.Sorted = True
    '</V1.3.1>
End Sub

Sub PreparaVista(Index As Integer)
    Dim nItem   As Integer
    
    If Index = mnVistaActual Then
        frmMdiPadre.Toolbar1(2).Buttons(5 + Index).Value = tbrPressed
    Else
        For nItem = 0 To mnuVerVista.UBound
            mnuVerVista(nItem).Checked = (nItem = Index)
            frmMdiPadre.Toolbar1(2).Buttons(5 + nItem).Value = IIf(nItem = Index, tbrPressed, tbrUnpressed)
        Next nItem
        
        mnVistaActual = Index
        Call MuestraVista(mnVistaActual, True)
    End If
End Sub

Sub MuestraItem()
    '<V1.3.1>
    Dim sIndBloqueada   As String
    
    Select Case mnVistaActual
    Case mnVistaConsultas
        sIndBloqueada = mItem.SubItems(7)
        If sIndBloqueada = "S" Then
            Me.mnuArcBloquear.Caption = "Des&bloquear consulta"
        Else
            Me.mnuArcBloquear.Caption = "&Bloquear consulta"
        End If
    End Select
    '</V1.3.1>
End Sub

Private Sub Form_Activate()
    App.HelpFile = App.Path & "\AdmSatelite.hlp"
    
    frmMdiPadre.Toolbar1(0).Visible = False
    frmMdiPadre.Toolbar1(1).Visible = False
    frmMdiPadre.Toolbar1(2).Visible = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMdiPadre.Toolbar1(2).Visible = False
    frmMdiPadre.Toolbar1(0).Visible = True
End Sub


Private Sub mnuArcAbrirCons_Click()
    ArcAbrirModConsultas
End Sub

Private Sub mnuArcAsigLotes_Click()
    AsigLotes
End Sub

Private Sub mnuArcBloquear_Click()
    '<V1.3.1>
    ArcBloquearConsulta
    '</V1.3.1>
End Sub

Private Sub mnuArcCerrarVentana_Click()
    ArcCerrarVentana
End Sub

Private Sub mnuArcSalir_Click()
    End
End Sub



Private Sub mnuPopColFiltrarPor_Click(Index As Integer)
    '<V1.3.1>
    Select Case Index
    Case 0
        Call FiltroSeleccionar
    Case Else
        Call FiltroAsignar(mnNumColumnaActiva, mnuPopColFiltrarPor(Index).Caption)
        Call FiltroAplicar
    End Select
    '</V1.3.1>
End Sub

Private Sub mnuPopColOrdenar_Click()
    '<V1.3.1>
    OrdenaPorColumna
    '</V1.3.1>
End Sub

Private Sub mnuPopColumna_Click()
    '<V1.3.1>
    Dim i           As Integer
    Dim j           As Integer
    Dim sNomCampo   As String
    
    On Error GoTo ErrPopup
    
    Me.mnuPopColFiltrarPor(0).Caption = "Filtrar por ..."
    
    sNomCampo = Me.lvConsultas.ColumnHeaders(mnNumColumnaActiva + 1).Key
    rsDataFiltro.Filter = "nom_campo='" & sNomCampo & "'"
    
    If rsDataFiltro.RecordCount > 10 Then
        ' Crear solo el valor de la columna actual
        Load mnuPopColFiltrarPor(1)
        Me.mnuPopColFiltrarPor(1).Caption = IIf(msValColumnaActiva = "", gsGlsValorBlanco, msValColumnaActiva)
        For i = 2 To mnuPopColFiltrarPor.Count - 1
            Unload mnuPopColFiltrarPor(i)
        Next i
    Else
        ' Crea todos los valores disitintos encontrados
        rsDataFiltro.Sort = "gls_valor"
        i = 0
        While Not rsDataFiltro.EOF
            i = i + 1
            Load mnuPopColFiltrarPor(i)
            Me.mnuPopColFiltrarPor(i).Caption = IIf("" & rsDataFiltro!gls_valor = "", gsGlsValorBlanco, "" & rsDataFiltro!gls_valor)
            rsDataFiltro.MoveNext
        Wend
        
        For j = i + 1 To mnuPopColFiltrarPor.Count - 1
            Unload mnuPopColFiltrarPor(j)
        Next j
    End If
    
    Exit Sub
    
ErrPopup:
    If Err = 360 Then
        Resume Next
    End If
    '</V1.3.1>
End Sub

Private Sub mnuVerFiltro_Click()
    '<V1.3.1>
    If mnuVerFiltro.Caption = "&Aplicar filtro" Then
        Call FiltroSeleccionar
    Else
        Call FiltroQuitar
    End If
    '</V1.3.1>
End Sub

Private Sub mnuWindowVert_Click()
    ' Organiza los formularios secundarios en mosaico.
    frmMdiPadre.Arrange vbTileVertical
End Sub

Private Sub mnuWindowHort_Click()
    ' Organiza los formularios secundarios en mosaico.
    frmMdiPadre.Arrange vbTileHorizontal
End Sub

Private Sub Form_Load()
    IniciaForm
    Call MuestraVista(mnVistaActual, True)
End Sub

Sub IniciaForm()
    frmMdiPadre.Toolbar1(0).Visible = False
    frmMdiPadre.Toolbar1(1).Visible = False
    frmMdiPadre.Toolbar1(2).Visible = True
    
    Me.WindowState = vbMaximized
    
    msNumUltConsultas = ""
    
    frmMdiPadre.Toolbar1(2).Buttons(5).Value = tbrUnpressed
    frmMdiPadre.Toolbar1(2).Buttons(6).Value = tbrUnpressed
    frmMdiPadre.Toolbar1(2).Buttons(7).Value = tbrUnpressed
    frmMdiPadre.Toolbar1(2).Buttons(9).Value = tbrUnpressed
    frmMdiPadre.Toolbar1(2).Buttons(10).Value = tbrUnpressed
    
    mnVistaActual = GetSetting(App.Title, "Settings", "CurrentView", 0)
    mnuVerVista(mnVistaActual).Checked = True
    frmMdiPadre.Toolbar1(2).Buttons(5 + mnVistaActual).Value = tbrPressed
End Sub

Sub IniciaListaConsultas()
    lvConsultas.ColumnHeaders.Clear
    lvConsultas.ColumnHeaders.Add , , "Id", 500
    lvConsultas.ColumnHeaders.Add , , "Consulta", 3000
    lvConsultas.ColumnHeaders.Add , , "Base", 1000
    lvConsultas.ColumnHeaders.Add , , "Area", 1500
    lvConsultas.ColumnHeaders.Add , , "Negocio", 1500
    lvConsultas.ColumnHeaders.Add , , "Dueño", 1000
    lvConsultas.ColumnHeaders.Add , , "Creado por", 1000
    lvConsultas.ColumnHeaders.Add , , "Fec.Creación", 2000
    lvConsultas.ColumnHeaders.Add , , "Ult.Modificación", 2000
    lvConsultas.ColumnHeaders.Add , , "Bloqueada", 800
    lvConsultas.ListItems.Clear
    
    fraVista.Caption = "Vista de todas las consultas"
    mnuArcNuevo.Caption = "&Nueva consulta"
    mnuArcEditar.Caption = "&Editar consulta"
    mnuArcEliminar.Caption = "E&liminar consulta"
    
    frmMdiPadre.Toolbar1(2).Buttons(1).ToolTipText = "Nueva consulta"
    frmMdiPadre.Toolbar1(2).Buttons(2).ToolTipText = "Editar consulta"
    frmMdiPadre.Toolbar1(2).Buttons(3).ToolTipText = "Eliminar consulta"
End Sub

Sub CargaConsultas(bIndLeerData As Boolean)
    Dim nItem       As Integer
    '<V1.3.1>
    Dim sIndBloqueada   As String
    '</V1.3.1>
        
    On Error GoTo ErrCargaConsultas
            
    Screen.MousePointer = vbHourglass
    
    lvConsultas.Visible = False
    IniciaListaConsultas
    nItem = 1

    '<V1.3.1>
    lvConsultas.ColumnHeaders(1).Key = "num_consulta"
    lvConsultas.ColumnHeaders(2).Key = "nom_consulta"
    lvConsultas.ColumnHeaders(3).Key = "nom_basedatos"
    lvConsultas.ColumnHeaders(4).Key = "gls_area"
    lvConsultas.ColumnHeaders(5).Key = "gls_negocio"
    lvConsultas.ColumnHeaders(6).Key = "nom_dueno"
    lvConsultas.ColumnHeaders(7).Key = "nom_creador"
    lvConsultas.ColumnHeaders(8).Key = "fec_creacion"
    lvConsultas.ColumnHeaders(9).Key = "fec_ult_actualizacion"
    lvConsultas.ColumnHeaders(10).Key = "ind_bloqueada"
    
    If bIndLeerData Then
        IniciaRecordSetFiltro
        Call db_LeeConsultas(rsData)
    End If
    
    If Not (rsData Is Nothing) Then
    '</V1.3.1>
        While Not rsData.EOF
            '<V1.3.1>
            sIndBloqueada = "" & rsData!ind_bloqueada
            If sIndBloqueada = "S" Then
                Set mItem = Me.lvConsultas.ListItems.Add(, , rsData!num_consulta, , "ConsBloq_S")
            Else
                Set mItem = Me.lvConsultas.ListItems.Add(, , rsData!num_consulta)
            End If
            
            mItem.SubItems(1) = "" & rsData!nom_consulta
            mItem.SubItems(2) = "" & rsData!nom_basedatos
            mItem.SubItems(3) = "" & rsData!gls_area
            mItem.SubItems(4) = "" & rsData!gls_negocio
            mItem.SubItems(5) = "" & rsData!nom_dueno
            mItem.SubItems(6) = "" & rsData!nom_creador
            mItem.SubItems(7) = "" & rsData!fec_creacion
            mItem.SubItems(8) = "" & rsData!fec_ult_actualizacion
            mItem.SubItems(9) = "" & rsData!ind_bloqueada
            
            If bIndLeerData Then
                Call GuardaDistinct(1, mItem.SubItems(1))
                Call GuardaDistinct(2, mItem.SubItems(2))
                Call GuardaDistinct(3, mItem.SubItems(3))
                Call GuardaDistinct(4, mItem.SubItems(4))
                Call GuardaDistinct(5, mItem.SubItems(5))
                Call GuardaDistinct(6, mItem.SubItems(6))
                Call GuardaDistinct(7, mItem.SubItems(7))
                Call GuardaDistinct(8, mItem.SubItems(8))
                Call GuardaDistinct(9, mItem.SubItems(9))
            End If
            '</V1.3.1>
            
            If msNumUltConsultas = rsData!num_consulta Then
                nItem = Me.lvConsultas.ListItems.Count
            End If
            
            rsData.MoveNext
        Wend
    End If
        
    If Me.lvConsultas.ListItems.Count > 0 Then
        Set mItem = lvConsultas.ListItems.Item(nItem)
        mItem.EnsureVisible
        mItem.Selected = True
        On Error Resume Next
        lvConsultas.SetFocus
    End If
    
    ActivarBotones
    lvConsultas.Visible = True
    Screen.MousePointer = vbNormal
    
    Exit Sub
    
ErrCargaConsultas:
    Screen.MousePointer = vbNormal
    MsgBox Error, vbCritical, App.Title
    lvConsultas.Visible = True
    Exit Sub
End Sub

Private Sub Form_Resize()
    FormResize
End Sub

Private Sub lvConsultas_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
    '<V1.3.1>
    mnNumColumnaActiva = ColumnHeader.Index - 1
    If lvConsultas.ListItems.Count > 0 Then
        If mnNumColumnaActiva = 0 Then
            msValColumnaActiva = ""
        Else
            msValColumnaActiva = mItem.SubItems(mnNumColumnaActiva)
        End If
    End If
    PopupMenu mnuPopColumna
    '</V1.3.1>
End Sub

Private Sub lvConsultas_DblClick()
    If lvConsultas.ListItems.Count > 0 Then
        ArcEditarElemento
    End If
End Sub

Private Sub lvConsultas_ItemClick(ByVal Item As ComctlLib.ListItem)
    Set mItem = Item
    '<V1.3.1>
    MuestraItem
    '</V1.3.1>
End Sub

Private Sub lvConsultas_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button And vbRightButton Then
        PopupMenu mnuArchivo
    End If
End Sub

Private Sub mnuArcEditar_Click()
    ArcEditarElemento
End Sub

Sub ArcNuevaConsulta()
    gsNumConsulta = ""
    gsNomConsulta = ""
    gsNomUsuarioLocal = gsUsuarioReal
    frmEditarConsulta.Show vbModal
    If Not gbCancelar Then
        msNumUltConsultas = gsNumConsulta
        Call CargaConsultas(True)
    End If
End Sub

Private Sub mnuArcEliminar_Click()
    ArcEliminarElemento
End Sub

Private Sub mnuArcNuevo_Click()
    ArcNuevoElemento
End Sub

Private Sub mnuArcAsigConsultas_Click()
    AsigConsultas
End Sub

Private Sub mnuArcAsigAgrupacion_Click()
    AsigAgrupacion
End Sub

Private Sub mnuArcAsigUsuarios_Click()
    AsigUsuarios
End Sub

Private Sub mnuVerVista_Click(Index As Integer)
    Call PreparaVista(Index)
End Sub

Private Sub mnuWindowCascade_Click()
    ' Organiza los formularios secundarios en cascada.
    frmMdiPadre.Arrange vbCascade
End Sub

Sub ArcBloquearConsulta()
    '<V1.3.1>
    Dim sNumConsulta        As String
    Dim sNomConsulta        As String
    Dim sIndBloqueada       As String
    Dim sIndBloqueadaFin    As String
    Dim sGlsPre             As String
    Dim bOk                 As Boolean
    
    On Error GoTo ErrArcBloquearConsulta
    
    sNumConsulta = mItem
    sNomConsulta = mItem.SubItems(1)
    sIndBloqueada = mItem.SubItems(7)
    sGlsPre = IIf(sIndBloqueada = "S", "des", "")
    
    If MsgBox("Está seguro que desea " & sGlsPre & "bloquear la consulta """ & sNomConsulta & """ (Id " & sNumConsulta & ")", vbYesNoCancel + vbDefaultButton2, App.Title) <> vbYes Then
        Exit Sub
    End If
    
    sIndBloqueadaFin = IIf(sIndBloqueada = "S", "N", "S")
    
    ' Bloquea consulta
    Screen.MousePointer = vbHourglass
    
    bOk = db_BloqueaConsulta(sNumConsulta, sIndBloqueadaFin)
    
    Screen.MousePointer = vbNormal
    
    If bOk Then
        Call CambiaIcono(sIndBloqueadaFin)
        MsgBox "Consulta fue " & sGlsPre & "bloqueada", vbInformation, App.Title
    End If
    
    Exit Sub
    
ErrArcBloquearConsulta:
    Screen.MousePointer = vbNormal
    MsgBox Error, vbCritical, App.Title
    Exit Sub
    '</V1.3.1>
End Sub

Sub FiltroAplicar()
    '<V1.3.1>
    Dim nCol            As Integer
    Dim nColData        As Integer
    Dim sNomCampo       As String
    Dim nTipoDato       As Integer
    Dim sSigno          As String
    Dim sFiltro         As String
    
    Dim sFiltroTotal    As String
    Dim nPos            As Integer
    Dim sLimitador      As String
    
    On Error GoTo ErrFiltroAplicar
    
    Screen.MousePointer = vbHourglass
    
    sFiltroTotal = ""
    For nCol = 1 To Me.lvConsultas.ColumnHeaders.Count
        If lvConsultas.ColumnHeaders(nCol).Tag <> "" Then
            sNomCampo = lvConsultas.ColumnHeaders(nCol).Key
            nColData = fnIndiceCampoRecordset(sNomCampo)
            nTipoDato = fnTipoDatoCol(rsData.Fields(nColData).Type)
            
            If nTipoDato = wc_tipo_dato_fecha Or nTipoDato = wc_tipo_dato_otro Then
                sLimitador = "'"
            Else
                sLimitador = ""
            End If
            
            nPos = InStr(lvConsultas.ColumnHeaders(nCol).Tag, Chr(9))
            If nPos > 0 Then
                sSigno = " " & Trim(Left(Left(lvConsultas.ColumnHeaders(nCol).Tag, nPos - 1), 4)) & " "
                sFiltro = Mid(Left(lvConsultas.ColumnHeaders(nCol).Tag, nPos - 1), 5)
                sFiltroTotal = sFiltroTotal & sNomCampo & sSigno & sLimitador & sFiltro & sLimitador
                
                sFiltroTotal = sFiltroTotal & gsGlsOperadorFiltro
                
                sSigno = " " & Trim(Left(Mid(lvConsultas.ColumnHeaders(nCol).Tag, nPos + 1), 4)) & " "
                sFiltro = Mid(Mid(lvConsultas.ColumnHeaders(nCol).Tag, nPos + 1), 5)
                sFiltroTotal = sFiltroTotal & sNomCampo & sSigno & sLimitador & sFiltro & sLimitador
            Else
                If sFiltroTotal <> "" Then
                    sFiltroTotal = sFiltroTotal & gsGlsOperadorFiltro
                End If
                
                sSigno = " " & Trim(Left(lvConsultas.ColumnHeaders(nCol).Tag, 4)) & " "
                sFiltro = Mid(lvConsultas.ColumnHeaders(nCol).Tag, 5)
                sFiltroTotal = sFiltroTotal & sNomCampo & sSigno & sLimitador & sFiltro & sLimitador
            End If
        End If
    Next nCol

    rsData.Filter = sFiltroTotal
    
    lvConsultas.ListItems.Clear
    Call MuestraVista(mnVistaActual, False)
    
    mnuVerFiltro.Caption = "&Quitar filtro"
    'mnuVerFiltro.Checked = True
    'frmMdiPadre.Toolbar1(2).Buttons(10).Value = tbrPressed
    
    Screen.MousePointer = vbNormal
    Exit Sub
    
ErrFiltroAplicar:
    Screen.MousePointer = vbNormal
    MsgBox Error, vbCritical, App.Title
    lvConsultas.Visible = True
    mnuVerFiltro.Checked = False
    frmMdiPadre.Toolbar1(2).Buttons(10).Value = tbrUnpressed
    Exit Sub
    '</V1.3.1>
End Sub

Function fnIndiceCampoRecordset(sNomCampo As String) As Integer
    '<V1.3.1>
    Dim nCol    As Integer
    
    For nCol = 0 To rsData.Fields.Count - 1
        If LCase(sNomCampo) = LCase(rsData.Fields(nCol).Name) Then
            Exit For
        End If
    Next nCol
    
    fnIndiceCampoRecordset = nCol
    '</V1.3.1>
End Function


Public Sub FiltroSeleccionar()
    '<V1.3.1>
    gbCancelar = True
    'frmMdiPadre.Toolbar1(2).Buttons(10).Value = tbrPressed
    frmFiltroAdm.Show vbModal
    If gbCancelar Then
        'frmMdiPadre.Toolbar1(2).Buttons(10).Value = tbrUnpressed
    Else
        FiltroAplicar
    End If
    '</V1.3.1>
End Sub

Sub FiltroQuitar()
    '<V1.3.1>
    Screen.MousePointer = vbHourglass
    
    rsData.Filter = ""
    
    lvConsultas.ListItems.Clear
    Call MuestraVista(mnVistaActual, False)
    
    Screen.MousePointer = vbNormal
    
    mnuVerFiltro.Caption = "&Aplicar filtro"
    'mnuVerFiltro.Checked = False
    'frmMdiPadre.Toolbar1(2).Buttons(10).Value = tbrUnpressed
    '</V1.3.1>
End Sub

