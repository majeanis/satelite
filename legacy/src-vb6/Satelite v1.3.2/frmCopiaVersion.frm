VERSION 5.00
Begin VB.Form frmCopiaVersion 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H0000C0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   3495
   ClientLeft      =   5475
   ClientTop       =   3195
   ClientWidth     =   6915
   Icon            =   "frmCopiaVersion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCopiaVersion.frx":08CA
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
      Picture         =   "frmCopiaVersion.frx":2712
      ScaleHeight     =   2265
      ScaleWidth      =   2265
      TabIndex        =   1
      Top             =   1020
      Width           =   2325
   End
   Begin VB.Label lblFileNew 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "..."
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2640
      TabIndex        =   4
      Top             =   1740
      Width           =   135
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Copiando nueva version ..."
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2640
      TabIndex        =   3
      Top             =   1380
      Width           =   1905
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
      TabIndex        =   2
      Top             =   3060
      Width           =   3510
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
Attribute VB_Name = "frmCopiaVersion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub CopiaArchivo(sFileNew As String, sFileOld As String)
    sFileNew = Mid(sFileNew, 2)
    sFileNew = Left(sFileNew, Len(sFileNew) - 1)

    sFileOld = Mid(sFileOld, 2)
    sFileOld = Left(sFileOld, Len(sFileOld) - 1)
    
    lblFileNew.Caption = sFileNew
    lblFileNew.Refresh
    
    On Error GoTo ErrCopiaArchivo
    
    FileCopy sFileNew, sFileOld
    Exit Sub

ErrCopiaArchivo:
    DoEvents
    ' Se intenta copiar el archivo hasta que el timer timer
    Resume
End Sub

Sub CopiaVersion()
    Dim sComando    As String
    Dim sFileOld    As String
    Dim sFileNew    As String
    Dim nPos        As Integer
    
    Screen.MousePointer = 11

    sComando = Command()
    If sComando = "" Then
        Screen.MousePointer = 0
        End
    End If

    nPos = InStr(sComando, ";")
    If nPos = 0 Then
        Screen.MousePointer = 0
        End
    End If
    
    sFileNew = Left(sComando, nPos - 1)
    sFileOld = Mid(sComando, nPos + 1)

    If sFileOld = "" Or sFileNew = "" Then
        Screen.MousePointer = 0
        End
    End If
    
    Me.Timer1.Interval = 50000 ' 5 minutos
    Me.Timer1.Enabled = True
    
    CopiaArchivo sFileNew, sFileOld
    
    Screen.MousePointer = 0
    Shell sFileOld, vbNormalFocus
    End
End Sub

Private Sub Form_Load()
    IniciaForm
    CopiaVersion
End Sub
Sub IniciaForm()
    Me.Left = (Screen.Width - Me.Width) \ 2
    Me.Top = (Screen.Height - Me.Height) \ 2
    
    Me.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Screen.MousePointer = 0
End Sub

Private Sub Timer1_Timer()
    Screen.MousePointer = 0
    MsgBox "No se pudo realizar la actualización de su versión. Informe al adminstrador de este sistema", vbCritical, "Satelite"
    End
End Sub


