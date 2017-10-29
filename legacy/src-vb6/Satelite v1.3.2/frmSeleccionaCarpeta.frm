VERSION 5.00
Begin VB.Form frmSeleccionaCarpeta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Selección de Carpeta de Consultas"
   ClientHeight    =   5415
   ClientLeft      =   5415
   ClientTop       =   2055
   ClientWidth     =   4185
   Icon            =   "frmSeleccionaCarpeta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   4185
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   4980
      Width           =   1095
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   4980
      Width           =   1095
   End
   Begin VB.DirListBox dirCarpeta 
      Height          =   3690
      Left            =   60
      TabIndex        =   1
      Top             =   1200
      Width           =   4035
   End
   Begin VB.DriveListBox drvCarpeta 
      Height          =   315
      Left            =   60
      TabIndex        =   0
      Top             =   840
      Width           =   4035
   End
   Begin VB.Label lblDescripcion 
      Caption         =   "Accion"
      Height          =   615
      Left            =   60
      TabIndex        =   4
      Top             =   120
      Width           =   4035
   End
End
Attribute VB_Name = "frmSeleccionaCarpeta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAceptar_Click()
    gsPathSeleccionado = dirCarpeta.Path
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub dirCarpeta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And cmdAceptar.Enabled Then
        cmdAceptar_Click
    End If
End Sub

Private Sub drvCarpeta_Change()
    On Error GoTo ErrDrive
    dirCarpeta.Path = drvCarpeta.Drive
    dirCarpeta.Visible = True
    cmdAceptar.Enabled = True
    Exit Sub
    
ErrDrive:
    MsgBox Error, vbCritical, App.Title
    dirCarpeta.Visible = False
    cmdAceptar.Enabled = False
End Sub

Private Sub drvCarpeta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And cmdAceptar.Enabled Then
        cmdAceptar_Click
    End If
End Sub

Private Sub Form_Load()
    IniciaForm
End Sub
Sub IniciaForm()
    Dim nPos    As Integer
    Dim sDrive  As String
    Dim sPath   As String
    
    On Error GoTo ErrIniciaForm
    
    Me.Left = (Screen.Width - Me.Width) \ 2
    Me.Top = (Screen.Height - Me.Height) \ 2
    
    nPos = InStr(gsPathInicial, ":")
    If nPos > 0 Then
        sDrive = Left(gsPathInicial, nPos + 1)
        drvCarpeta.Drive = sDrive
        If Exist(gsPathInicial) Then
            dirCarpeta.Path = gsPathInicial
        Else
            dirCarpeta.Path = sDrive
        End If
    End If
    Exit Sub
    
ErrIniciaForm:
    If Err = 68 Then
        drvCarpeta.Drive = "C:\"
        dirCarpeta.Path = "C:\"
    Else
        MsgBox Error, vbCritical, App.Title
    End If
    Exit Sub
End Sub
