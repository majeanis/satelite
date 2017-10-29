VERSION 5.00
Begin VB.Form frmExpArchivo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exportar a archivo"
   ClientHeight    =   2685
   ClientLeft      =   5235
   ClientTop       =   5445
   ClientWidth     =   7905
   Icon            =   "frmExpArchivo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   7905
   Begin VB.Frame Frame2 
      Caption         =   "Exportar a"
      Height          =   795
      Left            =   60
      TabIndex        =   11
      Top             =   1440
      Width           =   7755
      Begin VB.CommandButton cmdHelpCarpPer 
         Caption         =   "..."
         Height          =   315
         Left            =   7320
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   300
         Width           =   315
      End
      Begin VB.TextBox txtNomArchivo 
         Height          =   315
         Left            =   900
         TabIndex        =   0
         Top             =   300
         Width           =   6375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Archivo : "
         Height          =   195
         Left            =   180
         TabIndex        =   12
         Top             =   360
         Width           =   675
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Separadores"
      Height          =   1395
      Left            =   60
      TabIndex        =   10
      Top             =   0
      Width           =   6015
      Begin VB.CheckBox chkTipo 
         Caption         =   "&Sin separador (segun ancho de columnas)"
         Height          =   255
         Index           =   5
         Left            =   2520
         TabIndex        =   9
         Top             =   960
         Value           =   1  'Checked
         Width           =   3435
      End
      Begin VB.TextBox txtOtro 
         Height          =   315
         Left            =   3300
         TabIndex        =   8
         Top             =   600
         Width           =   495
      End
      Begin VB.CheckBox chkTipo 
         Caption         =   "&Otro : "
         Height          =   255
         Index           =   4
         Left            =   2520
         TabIndex        =   7
         Top             =   660
         Width           =   735
      End
      Begin VB.CheckBox chkTipo 
         Caption         =   "&Espacio"
         Height          =   255
         Index           =   3
         Left            =   2520
         TabIndex        =   6
         Top             =   360
         Width           =   1335
      End
      Begin VB.CheckBox chkTipo 
         Caption         =   "&Coma"
         Height          =   255
         Index           =   2
         Left            =   180
         TabIndex        =   5
         Top             =   960
         Width           =   1335
      End
      Begin VB.CheckBox chkTipo 
         Caption         =   "&Punto y coma"
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   4
         Top             =   660
         Width           =   1335
      End
      Begin VB.CheckBox chkTipo 
         Caption         =   "&Tabulación"
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   3
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   6720
      TabIndex        =   2
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   5520
      TabIndex        =   1
      Top             =   2280
      Width           =   1095
   End
End
Attribute VB_Name = "frmExpArchivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public msNomConsulta    As String
Sub FijaCheck(Index As Integer)
    Dim nIndice As Integer
    
    For nIndice = 0 To Me.chkTipo.UBound
        If Index <> nIndice Then
            chkTipo(nIndice).Value = 0
        End If
    Next nIndice
End Sub

Sub HelpNombreArchivo()
    On Error GoTo ErrHelpNombreArchivo
    
    frmMdiPadre.CommonDialog1.CancelError = True
    frmMdiPadre.CommonDialog1.Filter = "Todos los Archivos (*.*)|*.*"
    frmMdiPadre.CommonDialog1.DialogTitle = "Exportar consulta a archivo"
    frmMdiPadre.CommonDialog1.Action = 2
    
    If frmMdiPadre.CommonDialog1.FileName <> "" Then
        Me.txtNomArchivo = frmMdiPadre.CommonDialog1.FileName
    End If
    
    Exit Sub
    
ErrHelpNombreArchivo:
    Exit Sub
End Sub

Sub IniciaForm()
    Me.HelpContextID = 3
    Me.Left = (Screen.Width - Me.Width) \ 2
    Me.Top = (Screen.Height - Me.Height) \ 2
    Me.cmdAceptar.Enabled = False
        
    Me.txtNomArchivo = GetSetting(App.Title, "FileRecent", msNomConsulta, "")
    gbCancelar = True

    Me.chkTipo(0).Tag = Chr(9)
    Me.chkTipo(1).Tag = ";"
    Me.chkTipo(2).Tag = ","
    Me.chkTipo(3).Tag = " "
End Sub

Private Sub chkTipo_Click(Index As Integer)
    If Me.chkTipo(Index).Value = 1 Then
        Call FijaCheck(Index)
    End If
End Sub

Private Sub chkTipo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 And Me.cmdAceptar.Enabled Then
        cmdAceptar_Click
    End If
End Sub


Private Sub cmdAceptar_Click()
    If ExportarArchivo Then
        gbCancelar = False
        Unload Me
    End If
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdHelpCarpPer_Click()
    HelpNombreArchivo
End Sub

Private Sub Form_Load()
    IniciaForm
End Sub
Function ExportarArchivo() As Boolean
    Dim sFile   As String
    Dim nFila   As Long
    Dim sCons   As String
    Dim nPos    As Integer
    Dim nCheck  As Integer
        
    On Error GoTo ErrExportarArchivo
    
    nCheck = -1
    For nFila = 0 To Me.chkTipo.UBound
        If Me.chkTipo(nFila).Value = 1 Then
            nCheck = nFila
            Exit For
        End If
    Next nFila
    
    If nCheck = -1 Then
        MsgBox "Debe especificar el tipo de separador de campos", vbCritical, App.Title
        ExportarArchivo = False
        Exit Function
    End If
    
    If chkTipo(4).Value = 1 And Me.txtOtro = "" Then
        MsgBox "Debe especificar el caracter de separacion de campos, en tipo Otro", vbCritical, App.Title
        ExportarArchivo = False
        Exit Function
    End If
    
    sFile = Me.txtNomArchivo
    
    If Exist(sFile) Then
        If MsgBox("Archivo " & sFile & " ya existe. Desea reemplazarlo", vbQuestion + vbYesNoCancel + vbDefaultButton2, App.Title) <> vbYes Then
            ExportarArchivo = False
            Exit Function
        End If
    End If
    
    If sFile <> "" Then
        SaveSetting App.Title, "FileRecent", msNomConsulta, sFile
        gsNomArchivoExportar = sFile
        If nCheck = 4 Then
            gsGlsSeparadorCampos = Me.txtOtro
        Else
            gsGlsSeparadorCampos = Me.chkTipo(nCheck).Tag
        End If
    End If
    
    ExportarArchivo = True
    Exit Function
    
ErrExportarArchivo:
    Screen.MousePointer = vbNormal
    MsgBox Error, vbCritical, App.Title
    
    ExportarArchivo = False
End Function

Private Sub txtNomArchivo_Change()
    Me.cmdAceptar.Enabled = (Trim(txtNomArchivo.Text) <> "")
End Sub


Private Sub txtNomArchivo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Me.cmdAceptar.Enabled Then
        cmdAceptar_Click
    End If
End Sub


Private Sub txtOtro_Change()
    Me.chkTipo(4).Value = 1
End Sub


