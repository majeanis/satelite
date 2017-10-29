VERSION 5.00
Begin VB.Form frmConfigurar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configurar sistema"
   ClientHeight    =   2625
   ClientLeft      =   3615
   ClientTop       =   2925
   ClientWidth     =   9630
   Icon            =   "frmConfigurar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   9630
   Begin VB.Frame Frame1 
      Caption         =   "Editor de archivo de query"
      Height          =   675
      Left            =   60
      TabIndex        =   10
      Top             =   1380
      Width           =   9495
      Begin VB.TextBox txtNomEditor 
         Height          =   315
         Left            =   2580
         TabIndex        =   2
         Top             =   240
         Width           =   6435
      End
      Begin VB.CommandButton cmdHlpEditor 
         Caption         =   "..."
         Height          =   315
         Left            =   9060
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   240
         Width           =   315
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Utilitario de edición de Archivos : "
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   300
         Width           =   2355
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   7260
      TabIndex        =   3
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   8460
      TabIndex        =   4
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Frame fraConsultas 
      Caption         =   "Carpetas de Consultas"
      Height          =   1335
      Left            =   60
      TabIndex        =   5
      Top             =   0
      Width           =   9495
      Begin VB.CommandButton cmdHelpCarpGru 
         Caption         =   "..."
         Height          =   315
         Left            =   9060
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   720
         Width           =   315
      End
      Begin VB.TextBox txtCarpetaGrupal 
         Height          =   315
         Left            =   1920
         TabIndex        =   1
         Top             =   720
         Width           =   7095
      End
      Begin VB.CommandButton cmdHelpCarpPer 
         Caption         =   "..."
         Height          =   315
         Left            =   9060
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   360
         Width           =   315
      End
      Begin VB.TextBox txtCarpetaPersonal 
         Height          =   315
         Left            =   1920
         TabIndex        =   0
         Top             =   360
         Width           =   7095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Consultas grupales : "
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   780
         Width           =   1470
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Consultas personales : "
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   420
         Width           =   1635
      End
   End
End
Attribute VB_Name = "frmConfigurar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Sub VerAceptar()
    cmdAceptar.Enabled = (txtCarpetaPersonal <> "")
End Sub

Private Sub cmdAceptar_Click()
    If txtCarpetaPersonal <> gsPathConsultas Then
        If Not Exist(txtCarpetaPersonal) Then
            MsgBox "Carpeta de consultas personales no existe.", vbCritical, App.Title
            Exit Sub
        End If
    End If
    
    If txtCarpetaGrupal <> "" And txtCarpetaGrupal <> gsPathGrupales Then
        If Not Exist(txtCarpetaGrupal) Then
            MsgBox "Carpeta de consultas grupales no existe.", vbCritical, App.Title
            Exit Sub
        End If
    End If
    
    If txtNomEditor <> "" And txtNomEditor.Text = gsNomEditor Then
        If Not Exist(txtNomEditor) Then
            MsgBox "Archivo " & txtNomEditor & " no existe.", vbCritical, App.Title
            Exit Sub
        End If
    End If
        
    If txtCarpetaPersonal <> gsPathConsultas Then
        gsPathConsultas = txtCarpetaPersonal
        SaveSetting App.Title, "Path", "User", gsPathConsultas
        gbCancelar = False
    End If
    
    If txtCarpetaGrupal <> gsPathGrupales Then
        gsPathGrupales = txtCarpetaGrupal
        SaveSetting App.Title, "Path", "Group", gsPathGrupales
        gbCancelar = False
    End If
        
    If txtNomEditor <> gsNomEditor Then
        gsNomEditor = txtNomEditor
        SaveSetting App.Title, "Edit", "Editor", gsNomEditor
    End If
    
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdHelpCarpPer_Click()
    gsPathInicial = txtCarpetaPersonal.Text
    gsPathSeleccionado = ""
    frmSeleccionaCarpeta.lblDescripcion.Caption = "Seleccione la carpeta que utilizará el sistema para buscar sus consultas personales"
    frmSeleccionaCarpeta.Show vbModal
    If gsPathSeleccionado <> "" Then
        txtCarpetaPersonal.Text = gsPathSeleccionado
    End If
    txtCarpetaPersonal.SetFocus
End Sub

Private Sub cmdHelpCarpGru_Click()
    gsPathInicial = txtCarpetaGrupal.Text
    gsPathSeleccionado = ""
    frmSeleccionaCarpeta.lblDescripcion.Caption = "Seleccione la carpeta que utilizará el sistema para buscar sus consultas grupales"
    frmSeleccionaCarpeta.Show vbModal
    If gsPathSeleccionado <> "" Then
        txtCarpetaGrupal.Text = gsPathSeleccionado
    End If
    txtCarpetaGrupal.SetFocus
End Sub

Private Sub cmdHlpEditor_Click()
    On Error GoTo ErrEditor
    
    frmMdiPadre.CommonDialog1.CancelError = True
    frmMdiPadre.CommonDialog1.Filter = "*.exe"
    frmMdiPadre.CommonDialog1.FileName = "*.exe"
    frmMdiPadre.CommonDialog1.DialogTitle = "Editor de Archivos"
    frmMdiPadre.CommonDialog1.InitDir = gsDirWindows
    frmMdiPadre.CommonDialog1.Action = 1
    
    If frmMdiPadre.CommonDialog1.FileName <> "" Then
        Me.txtNomEditor = frmMdiPadre.CommonDialog1.FileName
    End If
    
ErrEditor:
    Exit Sub
End Sub

Private Sub Form_Load()
    IniciaForm
End Sub
Sub IniciaForm()
    Me.Left = (Screen.Width - Me.Width) \ 2
    Me.Top = (Screen.Height - Me.Height) \ 3
    
    Me.txtCarpetaPersonal.Text = gsPathConsultas
    Me.txtCarpetaGrupal.Text = gsPathGrupales
    Me.txtNomEditor.Text = gsNomEditor
    
    gbCancelar = True
End Sub

Private Sub txtCarpetaPersonal_Change()
    VerAceptar
End Sub

