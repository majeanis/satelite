VERSION 5.00
Begin VB.Form frmEditarTabValores 
   Caption         =   "Edición de Tabla de Valores"
   ClientHeight    =   1680
   ClientLeft      =   2970
   ClientTop       =   2595
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   ScaleHeight     =   1680
   ScaleWidth      =   7335
   Begin VB.Frame Frame1 
      Caption         =   "Edición de Tabla de Valores"
      Height          =   1095
      Left            =   60
      TabIndex        =   4
      Top             =   60
      Width           =   7215
      Begin VB.ComboBox cboNomTabla 
         Height          =   315
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   300
         Width           =   2715
      End
      Begin VB.TextBox txtGlsValor 
         Height          =   315
         Left            =   720
         TabIndex        =   1
         Top             =   660
         Width           =   6375
      End
      Begin VB.Label lblNomTabla 
         AutoSize        =   -1  'True
         Caption         =   "Tabla :"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Valor :"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   450
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   6180
      TabIndex        =   3
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   4980
      TabIndex        =   2
      Top             =   1200
      Width           =   1095
   End
End
Attribute VB_Name = "frmEditarTabValores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim msNumRegistro   As String
Sub CargaTablas()
    Dim rsData  As ADODB.Recordset
    
    ' Carga todas las tablas de valores
    If db_LeeTabValores("ADMIN", rsData) Then
        While Not rsData.EOF
            cboNomTabla.AddItem "" & rsData!gls_valor
            rsData.MoveNext
        Wend
    End If
    
    If cboNomTabla.ListCount = 1 Then
        cboNomTabla.ListIndex = 0
    End If
End Sub

Sub CargaRegistro()
    Dim rsData  As ADODB.Recordset
    
    ' Carga el registro Tab_Valores
    If db_LeeTabValores("", rsData) Then
        rsData.Filter = "num_registro=" & msNumRegistro
        If Not rsData.EOF Then
            cboNomTabla.AddItem "" & rsData!cod_tabla
            txtGlsValor = "" & rsData!gls_valor
        End If
    End If

    If cboNomTabla.ListCount = 1 Then
        cboNomTabla.ListIndex = 0
    End If
End Sub

Function GrabarCodigoValor() As Boolean
    Dim sCodTabla       As String
    Dim sGlsValor       As String
    Dim sNumRegistro    As String
    Dim bOk             As Boolean
    
    On Error GoTo ErrGrabarCodigoValor
    
    sCodTabla = cboNomTabla.Text
    sGlsValor = txtGlsValor.Text
    sNumRegistro = msNumRegistro
    
    Screen.MousePointer = vbHourglass
    
    ' Graba informacion
    bOk = db_GrabaTabValores(sNumRegistro, sCodTabla, sGlsValor)
    
    Screen.MousePointer = vbNormal
    
    If Not bOk Then
        GrabarCodigoValor = False
    Else
        MsgBox "Registro fue grabado correctamente", vbInformation, App.Title
        
        cmdAceptar.Enabled = False
        GrabarCodigoValor = True
        gbCancelar = False
    End If
    
    Exit Function
    
ErrGrabarCodigoValor:
    Screen.MousePointer = vbNormal
    MsgBox Err.Description, vbCritical, App.Title
    GrabarCodigoValor = False
End Function

Sub VerAceptar()
    cmdAceptar.Enabled = (cboNomTabla.ListIndex >= 0 And txtGlsValor.Text <> "")
End Sub

Private Sub cboNomTabla_Change()
    VerAceptar
End Sub

Private Sub cmdAceptar_Click()
    If GrabarCodigoValor Then
        If msNumRegistro = "" Then
            txtGlsValor.Text = ""
            txtGlsValor.SetFocus
        Else
            Unload Me
        End If
    End If
End Sub

Private Sub cmdCancelar_Click()
    cmdAceptar.Enabled = False
    Unload Me
End Sub

Private Sub Form_Load()
    IniciaForm
    
    msNumRegistro = gsNumRegTabValor
    If msNumRegistro = "" Then
        ' Carga todas las tablas en caso que venga por la opcion Nuevo Valor
        CargaTablas
    Else
        ' Carga el registro del nuevo valor cuando viene por la opcion Editar Valor
        CargaRegistro
    End If
    cmdAceptar.Enabled = False
End Sub

Sub IniciaForm()
    Me.Left = (Screen.Width - Me.Width) \ 2
    Me.Top = (Screen.Height - Me.Height) \ 3
    
    gbCancelar = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim nResp   As Integer
    
    If cmdAceptar.Enabled Then
        nResp = MsgBox("Desea grabar los cambios realizados sobre este valor", vbYesNoCancel + vbDefaultButton1, App.Title)
        If nResp = vbYes Then
            If Not GrabarCodigoValor() Then
                Cancel = True
            End If
        ElseIf nResp = vbCancel Then
            Cancel = True
        End If
    End If
End Sub

Private Sub txtGlsValor_Change()
    VerAceptar
End Sub

