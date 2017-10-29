VERSION 5.00
Begin VB.Form frmEditarUsuario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edición de usuario"
   ClientHeight    =   1575
   ClientLeft      =   3765
   ClientTop       =   5145
   ClientWidth     =   7215
   Icon            =   "frmEditarUsuario.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   7215
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   4860
      TabIndex        =   2
      Top             =   1140
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   6060
      TabIndex        =   3
      Top             =   1140
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Descripción del usuario"
      Height          =   1095
      Left            =   60
      TabIndex        =   4
      Top             =   0
      Width           =   7095
      Begin VB.TextBox txtNomUsuario 
         Height          =   315
         Left            =   1440
         TabIndex        =   0
         Top             =   300
         Width           =   4035
      End
      Begin VB.ComboBox cboTipoUsuario 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   660
         Width           =   2715
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "(máx. 32 caracteres)"
         Height          =   195
         Left            =   5520
         TabIndex        =   7
         Top             =   360
         Width           =   1440
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nombre :"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   645
      End
      Begin VB.Label lblBaseDatos 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Usuario :"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmEditarUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim msNomUsuario           As String

Sub CargaTiposUsuarios()
    Dim rsData          As ADODB.Recordset
        
    On Error GoTo ErrCargaTiposUsuarios
            
    Screen.MousePointer = vbHourglass
    
    ' Abre base datos
    OpenMyDataBase
    
    ' Carga Tipos de Usuarios
    If db_LeeTiposUsuarios(rsData) Then
        While Not rsData.EOF
            cboTipoUsuario.AddItem "" & rsData!cod_tipo_usuario 'ind_crear_consultas ind_autoasignar_consultas ind_modificar_consultas ind_eliminar_consultas ind_ejecutar_consultas
            
            rsData.MoveNext
        Wend
    End If
            
    ' Cierra base datos
    CloseMyDataBase
    
    Screen.MousePointer = vbNormal
    Exit Sub
    
ErrCargaTiposUsuarios:
    Screen.MousePointer = vbNormal
    MsgBox Error, vbInformation, App.Title
    Exit Sub
End Sub

Function GrabarUsuario() As Boolean
    Dim sNomUsuario     As String
    Dim sCodTipoUsuario As String
    Dim sCodTipoAccion  As String
    Dim bOk             As Boolean
    
    On Error GoTo ErrGrabarUsuario
        
    ' Valida consistencia de informacion
    If Trim(txtNomUsuario) = "" Then
        MsgBox "No ha ingresado un nombre para este usuario", vbCritical, App.Title
        GrabarUsuario = False
        Exit Function
    End If
    
    If cboTipoUsuario.ListIndex < 0 Then
        MsgBox "No ha seleccionado tipo de usuario", vbCritical, App.Title
        GrabarUsuario = False
        Exit Function
    End If
    
    If Len(Trim(txtNomUsuario)) > 32 Then
        If MsgBox("Nombre de usuario excede los 32 caracteres. El nombre se ajustará a los primeros 32 caracteres. Desea continuar", vbQuestion + vbYesNoCancel + vbDefaultButton1, App.Title) <> vbYes Then
            GrabarUsuario = False
            Exit Function
        End If
    End If
    
    sCodTipoUsuario = cboTipoUsuario.Text
    sNomUsuario = Left(LCase(Trim(Me.txtNomUsuario)), 32)
    
    If msNomUsuario = "" Then
        sCodTipoAccion = "INS"
    Else
        sCodTipoAccion = "UPD"
    End If
    
    Screen.MousePointer = vbHourglass
    
    ' Graba informacion
    bOk = db_GrabaUsuario(sNomUsuario, sCodTipoUsuario, sCodTipoAccion)
    
    Screen.MousePointer = vbNormal
    
    If Not bOk Then
        GrabarUsuario = False
    Else
        MsgBox "Usuario fue grabada correctamente", vbInformation, App.Title
        gsNomUsuario = sNomUsuario
        
        cmdAceptar.Enabled = False
        GrabarUsuario = True
        gbCancelar = False
    End If
    
    Exit Function
    
ErrGrabarUsuario:
    Screen.MousePointer = vbNormal
    MsgBox Err.Description, vbCritical, App.Title
    GrabarUsuario = False
End Function

Private Sub cboTipoUsuario_Click()
    VerAceptar
End Sub


Private Sub cmdAceptar_Click()
    If GrabarUsuario() Then
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
    If msNomUsuario <> "" Then
        CargaUsuario
    Else
        CargaTiposUsuarios
    End If
    cmdAceptar.Enabled = False
End Sub
Sub IniciaForm()
    Dim nX  As Integer
    
    Me.HelpContextID = 12
    Me.Left = (Screen.Width - Me.Width) \ 2
    Me.Top = (Screen.Height - Me.Height) \ 3
    
    msNomUsuario = gsNomUsuario
    txtNomUsuario = gsNomUsuario
    If msNomUsuario = "" Then
        txtNomUsuario.Enabled = True
    Else
        txtNomUsuario.Enabled = False
    End If
    
    gbCancelar = True
End Sub

Sub CargaUsuario()
    Dim rsData          As ADODB.Recordset
    Dim sTipoUsuario    As String
    Dim i               As Integer
        
    On Error GoTo ErrCargaUsuario
            
    Screen.MousePointer = vbHourglass
    i = -1
    
    ' Abre base datos
    OpenMyDataBase
    
    ' Lee usuario
    If db_LeeUsuario(msNomUsuario, rsData) Then
        If Not rsData.EOF Then
            Me.txtNomUsuario = msNomUsuario
            sTipoUsuario = "" & rsData!cod_tipo_usuario
        End If
    End If
    
    ' Carga Tipos de Usuarios
    If db_LeeTiposUsuarios(rsData) Then
        While Not rsData.EOF
            cboTipoUsuario.AddItem "" & rsData!cod_tipo_usuario 'ind_crear_consultas ind_autoasignar_consultas ind_modificar_consultas ind_eliminar_consultas ind_ejecutar_consultas
            If "" & rsData!cod_tipo_usuario = sTipoUsuario Then
                i = cboTipoUsuario.ListCount - 1
            End If
                        
            rsData.MoveNext
        Wend
    End If
    
    ' Cierra base datos
    CloseMyDataBase
    
    If i >= 0 Then
        cboTipoUsuario.ListIndex = i
    End If
    
    Screen.MousePointer = vbNormal
    Exit Sub
    
ErrCargaUsuario:
    Screen.MousePointer = vbNormal
    MsgBox Error, vbInformation, App.Title
    Exit Sub
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Dim nResp   As Integer
    
    If cmdAceptar.Enabled Then
        nResp = MsgBox("Desea grabar los cambios realizados sobre este usuario", vbYesNoCancel + vbDefaultButton1, App.Title)
        If nResp = vbYes Then
            If Not GrabarUsuario() Then
                Cancel = True
            End If
        ElseIf nResp = vbCancel Then
            Cancel = True
        Else
            gbCancelar = True
        End If
    End If
End Sub


Private Sub txtNomUsuario_Change()
    VerAceptar
End Sub

Sub VerAceptar()
    cmdAceptar.Enabled = (Trim(txtNomUsuario.Text) <> "" And Me.cboTipoUsuario.ListIndex >= 0)
End Sub


Private Sub txtNomUsuario_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(LCase(Chr(KeyAscii)))
End Sub


