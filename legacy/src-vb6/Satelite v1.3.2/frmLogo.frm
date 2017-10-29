VERSION 5.00
Begin VB.Form frmLogo 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H0000C0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   3495
   ClientLeft      =   6975
   ClientTop       =   1965
   ClientWidth     =   6915
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLogo.frx":0000
   ScaleHeight     =   3495
   ScaleWidth      =   6915
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   6300
      Top             =   2580
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Height          =   2325
      Left            =   120
      Picture         =   "frmLogo.frx":1E48
      ScaleHeight     =   2265
      ScaleWidth      =   2265
      TabIndex        =   1
      Top             =   1020
      Width           =   2325
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © 2006 Todos los derechos reservados"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2580
      TabIndex        =   5
      Top             =   3060
      Width           =   3510
   End
   Begin VB.Label lblVersion 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   300
      Left            =   2640
      TabIndex        =   4
      Top             =   2460
      Width           =   930
   End
   Begin VB.Label lblNomCompañia 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Compañía"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   780
      Left            =   2580
      TabIndex        =   3
      Top             =   1620
      Width           =   4125
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00808080&
      X1              =   2580
      X2              =   6720
      Y1              =   1500
      Y2              =   1500
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Se autoriza este producto a :"
      ForeColor       =   &H00404000&
      Height          =   195
      Left            =   2580
      TabIndex        =   2
      Top             =   1260
      Width           =   2040
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00C00000&
      BorderWidth     =   3
      X1              =   6900
      X2              =   6900
      Y1              =   0
      Y2              =   3720
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00C00000&
      BorderWidth     =   3
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   3720
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      X1              =   0
      X2              =   6900
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C00000&
      BorderWidth     =   3
      X1              =   0
      X2              =   6900
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Sistema de Consultas Satélite"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   675
      Left            =   180
      TabIndex        =   0
      Top             =   240
      Width           =   6495
   End
End
Attribute VB_Name = "frmLogo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mnTimeIni   As Date
Dim mnTimeFin   As Date

Sub CargaBaseSistema()
    Dim sFile       As String
    Dim sVal        As String * 255
    Dim sVer        As String

    On Error GoTo ErrCargaBaseSistema

    sFile = App.Path & "\Mae\Satelite.cfg"
    If Not Exist(sFile) Then
        Screen.MousePointer = vbNormal
        MsgBox "Archivo " & sFile & " no fue encontrado. Informe al administrador del sistema", vbCritical, App.Title
        End
    End If
    
    Call OSGetPrivateProfileString("DataBase", "Connection", "", sVal, 255, sFile)
    gsGlsConexionSatelite = Limpia(sVal, False)
    GrabaLog gsGlsConexionSatelite
    
    sVal = ""
    Call OSGetPrivateProfileString("DataBase", "Version", "", sVal, 255, sFile)
    sVer = Limpia(sVal, False)
    If UCase(sDecodifClave(sVer, -1, App.Title)) = "DEMO" Then
        gsGlsConexionSatelite = sDecodifClave(gsGlsConexionSatelite, -1, App.Title)
    End If
    
    If Not OpenMyDataBase Then
        End
    End If
    
    CargaBaseDatosSistema
    Call db_LeeTipoDatos(grsTipoDatos)
    Call db_LeeUsuario(gsUsuarioReal, grsUsuarioReal)
    CloseMyDataBase
    
    If grsUsuarioReal.EOF Then
        Screen.MousePointer = vbNormal
        MsgBox "Usuario """ & gsUsuarioReal & """ no ha sido creado. Informe al administrador del sistema", vbCritical, App.Title
        End
    End If
    
    Exit Sub
    
ErrCargaBaseSistema:
    Screen.MousePointer = vbNormal
    MsgBox "CargaBaseSistema: " & Error, vbCritical, App.Title
    End
End Sub

Sub IniciaSistema()
    Dim nTimeUtilizado  As Long
    
    mnTimeIni = Time()
    
    lblNomCompañia = App.CompanyName
    lblVersion = "Versión " & App.Major & "." & App.Minor & "." & App.Revision
    Me.Refresh
    gdMinDateExcel = fdValorFecha("01/01/1900")
    
    gsUsuarioReal = Get_Username
    GrabaLog "Usuario:" & gsUsuarioReal
    CargaBaseSistema
    
    If Not Hyoplus Then
        MsgBox "Sistema no autorizado. Solicite su clave de instalación", vbCritical, App.Title
        End
    End If
    
    mnTimeFin = Time()
    nTimeUtilizado = DateDiff("s", mnTimeIni, mnTimeFin) * 1000
    
    If nTimeUtilizado < 4000 Then
        Timer1.Interval = 4000 - nTimeUtilizado
        Timer1.Enabled = True
    Else
        FinalizaForm
    End If
End Sub

Private Sub Form_Activate()
    IniciaSistema
End Sub

Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    IniciaForm
End Sub
Sub IniciaForm()
    Me.Left = (Screen.Width - Me.Width) \ 2
    Me.Top = (Screen.Height - Me.Height) \ 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Screen.MousePointer = vbNormal
End Sub

Private Sub Timer1_Timer()
    FinalizaForm
End Sub
Sub FinalizaForm()
    GrabaLog "FinalizaForm"
    frmMdiPadre.Show
    Unload Me
End Sub

