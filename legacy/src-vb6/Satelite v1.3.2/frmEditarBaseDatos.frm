VERSION 5.00
Begin VB.Form frmEditarBaseDatos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edición de Base de datos"
   ClientHeight    =   2235
   ClientLeft      =   2985
   ClientTop       =   4245
   ClientWidth     =   10335
   Icon            =   "frmEditarBaseDatos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   10335
   Begin VB.Frame Frame1 
      Caption         =   "Descripción de la base de datos"
      Height          =   1755
      Left            =   60
      TabIndex        =   5
      Top             =   0
      Width           =   10215
      Begin VB.CommandButton cmdProbar 
         Caption         =   "&Probar"
         Height          =   315
         Left            =   8940
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1020
         Width           =   1095
      End
      Begin VB.TextBox txtGlsFormatoFecha 
         Height          =   315
         Left            =   1980
         TabIndex        =   2
         Top             =   1260
         Width           =   2235
      End
      Begin VB.TextBox txtGlsConeccion 
         Height          =   315
         Left            =   960
         TabIndex        =   1
         Top             =   660
         Width           =   9075
      End
      Begin VB.TextBox txtNomBaseDatos 
         Height          =   315
         Left            =   960
         TabIndex        =   0
         Top             =   300
         Width           =   5715
      End
      Begin VB.Label lblMensaje 
         Alignment       =   1  'Right Justify
         Caption         =   "Label4"
         Height          =   255
         Left            =   7560
         TabIndex        =   10
         Top             =   1380
         Visible         =   0   'False
         Width           =   2475
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Formato filtro en Fechas :"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   1320
         Width           =   1785
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Conexión :"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   750
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
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   9180
      TabIndex        =   4
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   7980
      TabIndex        =   3
      Top             =   1800
      Width           =   1095
   End
End
Attribute VB_Name = "frmEditarBaseDatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim msNumBaseDatos    As String

Function GrabarBaseDatos()
    Dim sNumBaseDatos       As String
    Dim sNomBaseDatos       As String
    Dim sGlsConeccion       As String
    Dim sGlsFormatoFecha    As String
    Dim bOk                 As Boolean
    
    On Error GoTo ErrGrabarBaseDatos
        
    sNumBaseDatos = IIf(msNumBaseDatos = "", "0", msNumBaseDatos)
        
    ' Valida consistencia de informacion
    If Trim(txtNomBaseDatos) = "" Then
        MsgBox "No ha ingresado un nombre para esta base de datos", vbCritical, App.Title
        GrabarBaseDatos = False
        Exit Function
    End If
    
    If Trim(txtGlsConeccion) = "" Then
        MsgBox "No ha ingresado el string de conexión de la base de datos", vbCritical, App.Title
        GrabarBaseDatos = False
        Exit Function
    End If
    
    If Len(Trim(txtNomBaseDatos)) > 80 Then
        If MsgBox("Nombre de la base de datos excede los 80 caracteres. El nombre se ajustará a los primeros 80 caracteres. Desea continuar", vbQuestion + vbYesNoCancel + vbDefaultButton1, App.Title) <> vbYes Then
            GrabarBaseDatos = False
            Exit Function
        End If
    End If
    
    If Len(Trim(txtGlsConeccion)) > 500 Then
        If MsgBox("Conexión de la base de datos excede los 500 caracteres. La conexión se ajustará a los primeros 500 caracteres. Desea continuar", vbQuestion + vbYesNoCancel + vbDefaultButton1, App.Title) <> vbYes Then
            GrabarBaseDatos = False
            Exit Function
        End If
    End If
    
    sNomBaseDatos = Left(Trim(txtNomBaseDatos), 80)
    sGlsConeccion = sDecodifClave(Left(Trim(txtGlsConeccion), 500), 1, gsCodigoStrCon)
    sGlsFormatoFecha = Left(Trim(txtGlsFormatoFecha), 32)
    
    Screen.MousePointer = vbHourglass
    
    ' Graba informacion
    bOk = db_GrabaBaseDatos(sNumBaseDatos, sNomBaseDatos, sGlsConeccion, sGlsFormatoFecha)
    
    Screen.MousePointer = vbNormal
    
    If Not bOk Then
        GrabarBaseDatos = False
    Else
        If sNumBaseDatos <> msNumBaseDatos Then
            MsgBox "Base de datos fue creada con el número " & sNumBaseDatos, vbInformation, App.Title
            gsNumBaseDatos = sNumBaseDatos
        Else
            MsgBox "Base de datos fue grabada correctamente", vbInformation, App.Title
        End If
        
        cmdAceptar.Enabled = False
        GrabarBaseDatos = True
        gbCancelar = False
    End If
    
    Exit Function
    
ErrGrabarBaseDatos:
    Screen.MousePointer = vbNormal
    MsgBox Err.Description, vbCritical, App.Title
    GrabarBaseDatos = False
End Function

Sub VerAceptar()
    cmdAceptar.Enabled = (Trim(txtNomBaseDatos.Text) <> "" And Trim(txtGlsConeccion.Text) <> "")
End Sub

Private Sub cmdAceptar_Click()
    If GrabarBaseDatos() Then
        OpenMyDataBase
        CargaBaseDatosSistema
        CloseMyDataBase
        Unload Me
    End If
End Sub

Private Sub cmdCancelar_Click()
    gbCancelar = True
    cmdAceptar.Enabled = False
    Unload Me
End Sub

Private Sub cmdProbar_Click()
    ProbarConexion
End Sub

Private Sub Form_Load()
    IniciaForm
    If msNumBaseDatos <> "" Then
        CargaBaseDatos
        cmdAceptar.Enabled = False
    End If
End Sub
Sub IniciaForm()
    Dim nX  As Integer
    
    Me.HelpContextID = 4
    Me.Left = (Screen.Width - Me.Width) \ 2
    Me.Top = (Screen.Height - Me.Height) \ 3
    
    msNumBaseDatos = gsNumBaseDatos
    
    gbCancelar = True
End Sub

Sub CargaBaseDatos()
    Dim rsData          As ADODB.Recordset
        
    On Error GoTo ErrCargaBaseDatos
            
    Screen.MousePointer = vbHourglass
    
    ' Lee base de datos
    If db_LeeBaseDatos(msNumBaseDatos, rsData) Then
        If Not rsData.EOF Then
            Me.txtNomBaseDatos = "" & rsData!nom_basedatos
            Me.txtGlsConeccion = sDecodifClave("" & rsData!gls_coneccion, -1, gsCodigoStrCon)
            Me.txtGlsFormatoFecha = "" & rsData!gls_formato_fecha
        End If
    End If
        
    Screen.MousePointer = vbNormal
    Exit Sub
    
ErrCargaBaseDatos:
    Screen.MousePointer = vbNormal
    MsgBox Error, vbInformation, App.Title
    Exit Sub
    Resume
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Dim nResp   As Integer
    
    If cmdAceptar.Enabled Then
        nResp = MsgBox("Desea grabar los cambios realizados sobre esta base de datos", vbYesNoCancel + vbDefaultButton1, App.Title)
        If nResp = vbYes Then
            If Not GrabarBaseDatos() Then
                Cancel = True
            End If
        ElseIf nResp = vbCancel Then
            Cancel = True
        Else
            gbCancelar = True
        End If
    End If
End Sub

Sub ProbarConexion()
    lblMensaje.Caption = "Probando conexión ..."
    lblMensaje.Visible = True
    lblMensaje.Refresh
    
    Screen.MousePointer = vbHourglass
    If OpenDataBase(cnn_Consulta, txtGlsConeccion, 30) Then
        Screen.MousePointer = vbNormal
        lblMensaje.Caption = ""
        MsgBox "Conexión exitosa", vbInformation, App.Title
    End If
    Screen.MousePointer = vbNormal

    lblMensaje.Caption = ""
    lblMensaje.Visible = False
End Sub

Private Sub txtGlsConeccion_Change()
    cmdProbar.Enabled = (Trim(txtGlsConeccion.Text) <> "")
    VerAceptar
End Sub


Private Sub txtGlsFormatoFecha_Change()
    VerAceptar
End Sub


Private Sub txtNomBaseDatos_Change()
    VerAceptar
End Sub


