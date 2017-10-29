VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "Comctl32.ocx"
Begin VB.MDIForm frmMdiPadre 
   BackColor       =   &H8000000C&
   Caption         =   "Satélite - Sistema de Consultas "
   ClientHeight    =   3960
   ClientLeft      =   3705
   ClientTop       =   2970
   ClientWidth     =   9810
   Icon            =   "frmMdiPadre.frx":0000
   LinkTopic       =   "MDIForm1"
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9810
      _ExtentX        =   17304
      _ExtentY        =   1058
      ButtonWidth     =   926
      ButtonHeight    =   900
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   4
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Módulo de consultas"
            Object.Tag             =   ""
            ImageKey        =   "ModConsultas"
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Módulo de Administración"
            Object.Tag             =   ""
            ImageKey        =   "ModAdministracion"
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Salir de la aplicación"
            Object.Tag             =   ""
            ImageKey        =   "ModSalir"
         EndProperty
      EndProperty
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H000000FF&
         Caption         =   "Label1"
         Height          =   195
         Left            =   3720
         TabIndex        =   1
         Top             =   240
         Width           =   480
      End
   End
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Index           =   1
      Left            =   0
      TabIndex        =   3
      Top             =   600
      Width           =   9810
      _ExtentX        =   17304
      _ExtentY        =   1058
      ButtonWidth     =   926
      ButtonHeight    =   900
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   6
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   "Ejecutar consulta"
            ImageKey        =   "Ejecutar"
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   "Exportar consulta a Excel"
            ImageKey        =   "Excel"
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   "Parámetros"
            ImageKey        =   "Parametros"
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Cerrar ventana actual"
            Object.Tag             =   ""
            ImageKey        =   "Cerrar"
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   "Salir de la aplicacion"
            ImageKey        =   "Salir"
         EndProperty
      EndProperty
   End
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Index           =   2
      Left            =   0
      TabIndex        =   2
      Top             =   1200
      Width           =   9810
      _ExtentX        =   17304
      _ExtentY        =   1058
      ButtonWidth     =   926
      ButtonHeight    =   900
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   20
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Nuevo usuario"
            Object.Tag             =   ""
            ImageKey        =   "Nuevo"
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Editar usuario"
            Object.Tag             =   ""
            ImageKey        =   "Editar"
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Eliminar usuario"
            Object.Tag             =   ""
            ImageKey        =   "Eliminar"
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Vista de usuarios"
            Object.Tag             =   ""
            ImageKey        =   "VisUsuarios"
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Vista de consultas"
            Object.Tag             =   ""
            ImageKey        =   "VisConsultas"
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Vista de agrupaciones"
            Object.Tag             =   ""
            ImageKey        =   "VisPerfiles"
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Vista de lotes"
            Object.Tag             =   ""
            ImageKey        =   "VisLotes"
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Ver tipos de usuarios"
            Object.Tag             =   ""
            ImageKey        =   "VisTipoUsuarios"
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Ver bases de datos"
            Object.Tag             =   ""
            ImageKey        =   "VisBaseDatos"
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Ver tabla de valores"
            Object.Tag             =   ""
            ImageKey        =   "VisTabValores"
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button14 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Usuarios asignados"
            Object.Tag             =   ""
            ImageKey        =   "AsigUsuarios"
         EndProperty
         BeginProperty Button15 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Consultas asignadas"
            Object.Tag             =   ""
            ImageKey        =   "AsigConsultas"
         EndProperty
         BeginProperty Button16 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Agrupaciones asignadas"
            Object.Tag             =   ""
            ImageKey        =   "AsigGrupos"
         EndProperty
         BeginProperty Button17 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Lotes asignadas"
            Object.Tag             =   ""
            ImageKey        =   "AsigLotes"
         EndProperty
         BeginProperty Button18 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button19 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Cerrar ventana actual"
            Object.Tag             =   ""
            ImageKey        =   "Cerrar"
         EndProperty
         BeginProperty Button20 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Salir"
            Object.Tag             =   ""
            ImageKey        =   "Salir"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2220
      Top             =   2220
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   660
      Top             =   2160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   28
      ImageHeight     =   28
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   27
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMdiPadre.frx":030A
            Key             =   "Ejecutar"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMdiPadre.frx":0C8C
            Key             =   "Excel"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMdiPadre.frx":160E
            Key             =   "Parametros"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMdiPadre.frx":1F90
            Key             =   "Cerrar"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMdiPadre.frx":2912
            Key             =   "Salir"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMdiPadre.frx":3294
            Key             =   "Nuevo"
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMdiPadre.frx":3C16
            Key             =   "Editar"
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMdiPadre.frx":4598
            Key             =   "Eliminar"
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMdiPadre.frx":4F1A
            Key             =   "kk_usuarios"
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMdiPadre.frx":589C
            Key             =   "VisConsultas"
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMdiPadre.frx":621E
            Key             =   "kk_perfiles"
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMdiPadre.frx":6BA0
            Key             =   "VisTipoUsuarios"
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMdiPadre.frx":7522
            Key             =   "VisBaseDatos"
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMdiPadre.frx":7EA4
            Key             =   "AsigUsuarios"
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMdiPadre.frx":8826
            Key             =   "AsigConsultas"
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMdiPadre.frx":91A8
            Key             =   "kk_AsigPerfiles"
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMdiPadre.frx":9B2A
            Key             =   "ModConsultas"
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMdiPadre.frx":A4AC
            Key             =   "ModAdministracion"
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMdiPadre.frx":AE2E
            Key             =   "ModSalir"
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMdiPadre.frx":B7B0
            Key             =   "kk_AsigLotes"
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMdiPadre.frx":C132
            Key             =   "kk_grupos"
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMdiPadre.frx":CAB4
            Key             =   "VisLotes"
         EndProperty
         BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMdiPadre.frx":D436
            Key             =   "VisUsuarios"
         EndProperty
         BeginProperty ListImage24 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMdiPadre.frx":DDB8
            Key             =   "VisPerfiles"
         EndProperty
         BeginProperty ListImage25 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMdiPadre.frx":E73A
            Key             =   "AsigLotes"
         EndProperty
         BeginProperty ListImage26 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMdiPadre.frx":F0BC
            Key             =   "AsigGrupos"
         EndProperty
         BeginProperty ListImage27 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMdiPadre.frx":FA3E
            Key             =   "VisTabValores"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "&Archivo"
      Begin VB.Menu mnuArcModConsultas 
         Caption         =   "Módulo de &Consultas"
      End
      Begin VB.Menu mnuArcModAdministracion 
         Caption         =   "Módulo de &Administración"
      End
      Begin VB.Menu mnuArcNulo1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSalida 
         Caption         =   "&Salir"
      End
   End
End
Attribute VB_Name = "frmMdiPadre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** Formulario MDI principal de la aplicación de ejemplo       ***
'*** Bloc de notas MDI.                                         ***
'******************************************************************
Option Explicit
Sub IniciaForm()
    Me.WindowState = vbMaximized

    Me.Caption = Me.Caption & " - " & App.CompanyName
    
    frmMdiPadre.Toolbar1(0).Visible = False
    frmMdiPadre.Toolbar1(1).Visible = False
    frmMdiPadre.Toolbar1(2).Visible = False
    'Show
    
    gsNomUsuarioLocal = gsUsuarioReal
    Me.mnuArcModAdministracion.Enabled = ("" & grsUsuarioReal!ind_administrador = "S")
End Sub


Private Sub MDIForm_Load()
    IniciaForm
    
    ' Si es Administrador, queda en el menu principal
    If "" & grsUsuarioReal!ind_administrador = "S" Then
        frmMdiPadre.Toolbar1(0).Visible = True

    ' Sino, muestra la pantalla de consultas
    Else
        Me.mnuArcModAdministracion.Enabled = False
        frmMdiPadre.Toolbar1(0).Buttons(2).Enabled = False
        frmPrincipal.Show
    End If
End Sub

Private Sub mnuOptions_Click()
    ' Alterna la visibilidad de las barras de herramientas.
'    mnuOptionsToolbar.Checked = frmMdiPadre.picToolbar.Visible
End Sub

Private Sub mnuArcModAdministracion_Click()
    frmAdministracion.Show
End Sub

Private Sub mnuArcModConsultas_Click()
    frmPrincipal.Show
End Sub

Private Sub mnuSalida_Click()
    End
End Sub

Private Sub Toolbar1_ButtonClick(Index As Integer, ByVal Button As ComctlLib.Button)
    Dim x   As String
    On Error GoTo FinApp
    
    Select Case Index
    Case 0
        Select Case Button.Image
        Case "ModConsultas"
            gsNomUsuarioLocal = gsUsuarioReal
            frmPrincipal.Show
        Case "ModAdministracion"
            frmAdministracion.Show
        Case "ModSalir"
            End
        End Select
    
    Case 1
        Select Case Button.Image
        Case "Ejecutar"
            Call frmMdiPadre.ActiveForm.EjecutarNodo
        Case "Excel"
            Call frmMdiPadre.ActiveForm.ExportarExcel
        Case "Parametros"
            Call frmMdiPadre.ActiveForm.VerParametros
        Case "Cerrar"
            Call frmMdiPadre.ActiveForm.ArcCerrarVentana
        Case "Salir"
            End
        End Select

    Case 2
        Select Case Button.Image
        Case "Nuevo"
            Call frmMdiPadre.ActiveForm.ArcNuevoElemento
        Case "Editar"
            Call frmMdiPadre.ActiveForm.ArcEditarElemento
        Case "Eliminar"
            Call frmMdiPadre.ActiveForm.ArcEliminarElemento
        Case "VisUsuarios"
            Call frmMdiPadre.ActiveForm.PreparaVista(mnVistaUsuarios)
        Case "VisConsultas"
            Call frmMdiPadre.ActiveForm.PreparaVista(mnVistaConsultas)
        Case "VisPerfiles"
            Call frmMdiPadre.ActiveForm.PreparaVista(mnVistaPerfiles)
        '<V1.3.0>
        Case "VisLotes"
            Call frmMdiPadre.ActiveForm.PreparaVista(mnVistaLotes)
        '</V1.3.0>
        Case "VisTipoUsuarios"
            Call frmMdiPadre.ActiveForm.PreparaVista(mnVistaTiposUsuarios)
        Case "VisBaseDatos"
            Call frmMdiPadre.ActiveForm.PreparaVista(mnVistaBaseDatos)
        '<V1.3.0>
        Case "VisTabValores"
            Call frmMdiPadre.ActiveForm.PreparaVista(mnVistaTabValores)
        '</V1.3.1>
            
        Case "AsigUsuarios"
            Call frmMdiPadre.ActiveForm.AsigUsuarios
        Case "AsigConsultas"
            Call frmMdiPadre.ActiveForm.AsigConsultas
        Case "AsigGrupos"
            Call frmMdiPadre.ActiveForm.AsigAgrupacion
        '<V1.3.0>
        Case "AsigLotes"
            Call frmMdiPadre.ActiveForm.AsigLotes
        '</V1.3.0>
            
        Case "Cerrar"
            Call frmMdiPadre.ActiveForm.ArcCerrarVentana
        Case "Salir"
            End
        End Select
    End Select
    
    Exit Sub

FinApp:
    End
End Sub


