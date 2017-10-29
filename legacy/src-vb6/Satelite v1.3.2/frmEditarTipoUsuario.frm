VERSION 5.00
Begin VB.Form frmEditarTipoUsuario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edición de tipo de usuario"
   ClientHeight    =   2715
   ClientLeft      =   5265
   ClientTop       =   3735
   ClientWidth     =   6735
   Icon            =   "frmEditarTipoUsuario.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   6735
   Begin VB.Frame Frame1 
      Caption         =   "Descripción del tipo de usuario"
      Height          =   2235
      Left            =   60
      TabIndex        =   9
      Top             =   0
      Width           =   6615
      Begin VB.CheckBox chkIndEjecutarConsultas 
         Alignment       =   1  'Right Justify
         Caption         =   "Ejecutar consultas :"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1860
         Width           =   1875
      End
      Begin VB.CheckBox chkIndEliminarConsultas 
         Alignment       =   1  'Right Justify
         Caption         =   "Eliminar consultas : "
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1560
         Width           =   1875
      End
      Begin VB.CheckBox chkIndModificarConsultas 
         Alignment       =   1  'Right Justify
         Caption         =   "Modificar consultas : "
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1260
         Width           =   1875
      End
      Begin VB.CheckBox chkIndAutoasignarConsultas 
         Alignment       =   1  'Right Justify
         Caption         =   "Asignar automáticamente las consultas creadas : "
         Height          =   255
         Left            =   2400
         TabIndex        =   3
         Top             =   960
         Width           =   3915
      End
      Begin VB.CheckBox chkIndCrearConsultas 
         Alignment       =   1  'Right Justify
         Caption         =   "Crear consultas :"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   960
         Width           =   1875
      End
      Begin VB.CheckBox chkIndAdministrador 
         Alignment       =   1  'Right Justify
         Caption         =   "Administrador :"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   660
         Width           =   1875
      End
      Begin VB.TextBox txtCodTipoUsuario 
         Height          =   315
         Left            =   1800
         TabIndex        =   0
         Top             =   300
         Width           =   1995
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "(máx. 12 caracteres)"
         Height          =   195
         Left            =   3840
         TabIndex        =   11
         Top             =   360
         Width           =   1440
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Usuario :"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   990
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   5580
      TabIndex        =   8
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   4380
      TabIndex        =   7
      Top             =   2280
      Width           =   1095
   End
End
Attribute VB_Name = "frmEditarTipoUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim msCodTipoUsuario    As String

Function GrabarTipoUsuario()
    Dim sCodTipoUsuario As String
    Dim sXmlTipoUsuario As String
    Dim sCodTipoAccion  As String
    Dim bOk             As Boolean
    
    On Error GoTo ErrGrabarTipoUsuario
        
    ' Valida consistencia de informacion
    If Trim(txtCodTipoUsuario) = "" Then
        MsgBox "No ha ingresado un nombre para este tipo de usuario", vbCritical, App.Title
        GrabarTipoUsuario = False
        Exit Function
    End If
    
    If Len(Trim(txtCodTipoUsuario)) > 32 Then
        If MsgBox("Nombre del tipo de usuario excede los 12 caracteres. El nombre se ajustará a los primeros 12 caracteres. Desea continuar", vbQuestion + vbYesNoCancel + vbDefaultButton1, App.Title) <> vbYes Then
            GrabarTipoUsuario = False
            Exit Function
        End If
    End If
    
    sCodTipoUsuario = Left(UCase(txtCodTipoUsuario), 12)
    
    sXmlTipoUsuario = "<ROOT>"
    sXmlTipoUsuario = sXmlTipoUsuario & "<TipoUsuario "
    sXmlTipoUsuario = sXmlTipoUsuario & " ind_administrador=""" & IIf(Me.chkIndAdministrador.Value = 1, "S", "N") & """"
    sXmlTipoUsuario = sXmlTipoUsuario & " ind_crear_consultas=""" & IIf(Me.chkIndCrearConsultas.Value = 1, "S", "N") & """"
    sXmlTipoUsuario = sXmlTipoUsuario & " ind_autoasignar_consultas=""" & IIf(Me.chkIndAutoasignarConsultas.Value = 1, "S", "N") & """"
    sXmlTipoUsuario = sXmlTipoUsuario & " ind_modificar_consultas=""" & IIf(Me.chkIndModificarConsultas.Value = 1, "S", "N") & """"
    sXmlTipoUsuario = sXmlTipoUsuario & " ind_eliminar_consultas=""" & IIf(Me.chkIndEliminarConsultas.Value = 1, "S", "N") & """"
    sXmlTipoUsuario = sXmlTipoUsuario & " ind_ejecutar_consultas=""" & IIf(Me.chkIndEjecutarConsultas.Value = 1, "S", "N") & """"
    sXmlTipoUsuario = sXmlTipoUsuario & "/>"
    sXmlTipoUsuario = sXmlTipoUsuario & "</ROOT>"
    
    If msCodTipoUsuario = "" Then
        sCodTipoAccion = "INS"
    Else
        sCodTipoAccion = "UPD"
    End If
    
    Screen.MousePointer = vbHourglass
    
    ' Graba informacion
    bOk = db_GrabaTipoUsuario(sCodTipoUsuario, sXmlTipoUsuario, sCodTipoAccion)
    
    Screen.MousePointer = vbNormal
    
    If Not bOk Then
        GrabarTipoUsuario = False
    Else
        MsgBox "Tipo de usuario fue grabado correctamente", vbInformation, App.Title
        gsCodTipoUsuario = sCodTipoUsuario
        
        cmdAceptar.Enabled = False
        GrabarTipoUsuario = True
        gbCancelar = False
    End If
    
    Exit Function
    
ErrGrabarTipoUsuario:
    Screen.MousePointer = vbNormal
    MsgBox Err.Description, vbCritical, App.Title
    GrabarTipoUsuario = False
End Function

Private Sub chkIndAdministrador_Click()
    VerAceptar
End Sub

Private Sub chkIndAutoasignarConsultas_Click()
    VerAceptar
End Sub

Private Sub chkIndCrearConsultas_Click()
    VerAceptar
End Sub


Private Sub chkIndEjecutarConsultas_Click()
    VerAceptar
End Sub

Private Sub chkIndEliminarConsultas_Click()
    VerAceptar
End Sub

Private Sub chkIndModificarConsultas_Click()
    VerAceptar
End Sub

Private Sub cmdAceptar_Click()
    If GrabarTipoUsuario() Then
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
    If msCodTipoUsuario <> "" Then
        CargaTipoUsuario
    End If
    cmdAceptar.Enabled = False
End Sub
Sub VerAceptar()
    cmdAceptar.Enabled = (Trim(txtCodTipoUsuario.Text) <> "")
End Sub

Sub IniciaForm()
    Dim nX  As Integer
    
    Me.HelpContextID = 8
    Me.Left = (Screen.Width - Me.Width) \ 2
    Me.Top = (Screen.Height - Me.Height) \ 3
    
    msCodTipoUsuario = gsCodTipoUsuario
    txtCodTipoUsuario = msCodTipoUsuario
    
    If msCodTipoUsuario = "" Then
        txtCodTipoUsuario.Enabled = True
    Else
        txtCodTipoUsuario.Enabled = False
    End If

    gbCancelar = True
End Sub

Sub CargaTipoUsuario()
    Dim rsData          As ADODB.Recordset
    Dim sTipoUsuario    As String
        
    On Error GoTo ErrCargaTipoUsuario
            
    Screen.MousePointer = vbHourglass
    
    ' Lee tipo de usuario
    If db_LeeTipoUsuario(msCodTipoUsuario, rsData) Then
        If Not rsData.EOF Then
            Me.chkIndAdministrador.Value = IIf("" & rsData!ind_administrador = "S", 1, 0)
            Me.chkIndCrearConsultas.Value = IIf("" & rsData!ind_crear_consultas = "S", 1, 0)
            Me.chkIndAutoasignarConsultas.Value = IIf("" & rsData!ind_autoasignar_consultas = "S", 1, 0)
            Me.chkIndModificarConsultas.Value = IIf("" & rsData!ind_modificar_consultas = "S", 1, 0)
            Me.chkIndEliminarConsultas.Value = IIf("" & rsData!ind_eliminar_consultas = "S", 1, 0)
            Me.chkIndEjecutarConsultas.Value = IIf("" & rsData!ind_ejecutar_consultas = "S", 1, 0)
        End If
    End If
        
    Screen.MousePointer = vbNormal
    Exit Sub
    
ErrCargaTipoUsuario:
    Screen.MousePointer = vbNormal
    MsgBox Error, vbInformation, App.Title
    Exit Sub
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Dim nResp   As Integer
    
    If cmdAceptar.Enabled Then
        nResp = MsgBox("Desea grabar los cambios realizados sobre este tipo de usuario", vbYesNoCancel + vbDefaultButton1, App.Title)
        If nResp = vbYes Then
            If Not GrabarTipoUsuario() Then
                Cancel = True
            End If
        ElseIf nResp = vbCancel Then
            Cancel = True
        Else
            gbCancelar = True
        End If
    End If
End Sub



Private Sub txtCodTipoUsuario_Change()
    VerAceptar
End Sub

Private Sub txtCodTipoUsuario_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


