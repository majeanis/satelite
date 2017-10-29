VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Begin VB.Form frmClave 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Clave de Autorización"
   ClientHeight    =   2085
   ClientLeft      =   5625
   ClientTop       =   3120
   ClientWidth     =   6270
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmClave.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   6270
   Begin Threed.SSFrame ssfFondo 
      Height          =   1935
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   6135
      _Version        =   65536
      _ExtentX        =   10821
      _ExtentY        =   3413
      _StockProps     =   14
      Caption         =   "Satélite"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Font3D          =   3
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Salir"
         Height          =   315
         Left            =   2760
         TabIndex        =   1
         Top             =   1440
         Width           =   1035
      End
      Begin VB.CommandButton cmdCopiar 
         Caption         =   "&Copiar"
         Height          =   315
         Left            =   4380
         TabIndex        =   0
         Top             =   1020
         Width           =   1035
      End
      Begin VB.Label lblMensaje 
         Caption         =   "Label1"
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   300
         Width           =   5835
      End
      Begin VB.Label lblTitClave 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Ingrese Clave de Autorización :"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   4065
      End
   End
End
Attribute VB_Name = "frmClave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim msCodInstalacion As String

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdCopiar_Click()
    Clipboard.Clear
    Clipboard.SetText msCodInstalacion
End Sub

Private Sub Form_Load()
    msCodInstalacion = gsCodInstalacion
    
    lblMensaje.Caption = "Contáctese con su distribuidor y solicite su Clave de Autorización para ingresar al Sistema."
    Me.lblTitClave = "Código de instalación : " & UCase(msCodInstalacion)
End Sub

