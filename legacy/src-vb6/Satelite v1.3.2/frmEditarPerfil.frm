VERSION 5.00
Begin VB.Form frmEditarPerfil 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edición de agrupación"
   ClientHeight    =   1215
   ClientLeft      =   5985
   ClientTop       =   4785
   ClientWidth     =   7215
   Icon            =   "frmEditarPerfil.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1215
   ScaleWidth      =   7215
   Begin VB.Frame Frame1 
      Caption         =   "Descripción de la agrupación"
      Height          =   735
      Left            =   60
      TabIndex        =   3
      Top             =   0
      Width           =   7095
      Begin VB.TextBox txtNomPerfil 
         Height          =   315
         Left            =   1440
         TabIndex        =   0
         Top             =   300
         Width           =   4095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "(máx. 32 caracteres)"
         Height          =   195
         Left            =   5580
         TabIndex        =   5
         Top             =   360
         Width           =   1440
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nombre :"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   645
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   6060
      TabIndex        =   2
      Top             =   780
      Width           =   1095
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   4860
      TabIndex        =   1
      Top             =   780
      Width           =   1095
   End
End
Attribute VB_Name = "frmEditarPerfil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim msNumPerfil           As String

Function GrabarPerfil()
    Dim sNumPerfil  As String
    Dim sNomPerfil  As String
    Dim bOk         As Boolean
    
    On Error GoTo ErrGrabarPerfil
        
    sNumPerfil = IIf(msNumPerfil = "", "0", msNumPerfil)
    
    ' Valida consistencia de informacion
    If Trim(txtNomPerfil) = "" Then
        MsgBox "No ha ingresado un nombre para esta agrupación", vbCritical, App.Title
        GrabarPerfil = False
        Exit Function
    End If
    
    If Len(Trim(txtNomPerfil)) > 32 Then
        If MsgBox("Nombre de la agrupación excede los 32 caracteres. El nombre se ajustará a los primeros 32 caracteres. Desea continuar", vbQuestion + vbYesNoCancel + vbDefaultButton1, App.Title) <> vbYes Then
            GrabarPerfil = False
            Exit Function
        End If
    End If
    
    sNomPerfil = Left(Trim(txtNomPerfil), 32)
    
    Screen.MousePointer = vbHourglass
    
    ' Graba informacion
    bOk = db_GrabaAgrupacion(sNumPerfil, sNomPerfil)
    
    Screen.MousePointer = vbNormal
    
    If Not bOk Then
        GrabarPerfil = False
    Else
        If sNumPerfil <> msNumPerfil Then
            MsgBox "Agrupación fue creada con el número " & sNumPerfil, vbInformation, App.Title
            gsNumPerfil = sNumPerfil
        Else
            MsgBox "Agrupación fue grabada correctamente", vbInformation, App.Title
        End If
        
        cmdAceptar.Enabled = False
        GrabarPerfil = True
        gbCancelar = False
    End If
    
    Exit Function
    
ErrGrabarPerfil:
    Screen.MousePointer = vbNormal
    MsgBox Err.Description, vbCritical, App.Title
    GrabarPerfil = False
End Function

Private Sub cmdAceptar_Click()
    If GrabarPerfil() Then
        Unload Me
    End If
End Sub

Private Sub cmdCancelar_Click()
    gbCancelar = True
    cmdAceptar.Enabled = False
    Unload Me
End Sub

Private Sub Form_Load()
    IniciaForm
    cmdAceptar.Enabled = False
End Sub
Sub IniciaForm()
    Dim nX  As Integer
    
    Me.HelpContextID = 28
    Me.Left = (Screen.Width - Me.Width) \ 2
    Me.Top = (Screen.Height - Me.Height) \ 3
    
    msNumPerfil = gsNumPerfil
    txtNomPerfil = gsNomPerfil
    
    gbCancelar = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim nResp   As Integer
    
    If cmdAceptar.Enabled Then
        nResp = MsgBox("Desea grabar los cambios realizados sobre este usuario", vbYesNoCancel + vbDefaultButton1, App.Title)
        If nResp = vbYes Then
            If Not GrabarPerfil() Then
                Cancel = True
            End If
        ElseIf nResp = vbCancel Then
            Cancel = True
        Else
            gbCancelar = True
        End If
    End If
End Sub



Private Sub txtNomPerfil_Change()
    VerAceptar
End Sub
Sub VerAceptar()
    cmdAceptar.Enabled = (Trim(txtNomPerfil.Text) <> "")
End Sub

