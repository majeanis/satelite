VERSION 5.00
Begin VB.Form frmDefParametros 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Definición de Parámetros"
   ClientHeight    =   6930
   ClientLeft      =   4530
   ClientTop       =   3630
   ClientWidth     =   7545
   Icon            =   "frmDefParametros.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   7545
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   5160
      TabIndex        =   13
      Top             =   6480
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   6360
      TabIndex        =   12
      Top             =   6480
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Caption         =   "Descripción del parámetro"
      Height          =   1455
      Left            =   60
      TabIndex        =   9
      Top             =   60
      Width           =   7395
      Begin VB.CheckBox chkOpcional 
         Caption         =   "Opcional"
         Height          =   195
         Left            =   5400
         TabIndex        =   3
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox txtGlsParametro 
         Height          =   315
         Left            =   1740
         TabIndex        =   1
         Top             =   660
         Width           =   5535
      End
      Begin VB.TextBox txtNomParametro 
         Height          =   315
         Left            =   1740
         TabIndex        =   0
         Top             =   300
         Width           =   5535
      End
      Begin VB.ComboBox cboTipoDato 
         Height          =   315
         Left            =   1740
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1020
         Width           =   2295
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Título Parámetro : "
         Height          =   195
         Left            =   180
         TabIndex        =   14
         Top             =   720
         Width           =   1320
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nombre Parámetro : "
         Height          =   195
         Left            =   180
         TabIndex        =   11
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lblBaseDatos 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Dato : "
         Height          =   195
         Left            =   180
         TabIndex        =   10
         Top             =   1080
         Width           =   1065
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ayuda de Valores"
      Height          =   4875
      Left            =   60
      TabIndex        =   8
      Top             =   1560
      Width           =   7395
      Begin VB.TextBox txtValAyuda 
         Height          =   3495
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   1260
         Width           =   7155
      End
      Begin VB.OptionButton optTipoAyuda 
         Caption         =   "Ayuda según query"
         Height          =   195
         Index           =   1
         Left            =   1200
         TabIndex        =   6
         Top             =   960
         Width           =   2295
      End
      Begin VB.OptionButton optTipoAyuda 
         Caption         =   "Lista de valores"
         Height          =   195
         Index           =   0
         Left            =   1200
         TabIndex        =   5
         Top             =   720
         Width           =   2295
      End
      Begin VB.CheckBox chkAyudaValores 
         Alignment       =   1  'Right Justify
         Caption         =   "Parámetro con ayuda de valores"
         Height          =   255
         Left            =   180
         TabIndex        =   4
         Top             =   360
         Width           =   2655
      End
   End
End
Attribute VB_Name = "frmDefParametros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public msNomParametro   As String
Public msGlsParametro   As String
Public msTipoDato       As String
Public msTipoAyuda      As String
Public msGlsAyuda       As String
Public msIndOpcional    As String

Sub AceptarParametro()
    gsNomParametro = txtNomParametro
    gsGlsParametro = txtGlsParametro
    gsTipoDato = cboTipoDato
    gsIndOpcional = IIf(chkOpcional.Value, "S", "N")
    
    If Me.chkAyudaValores.Value = 0 Then
        gsGlsAyuda = ""
    Else
        If optTipoAyuda(0).Value = True Then
            gsTipoAyuda = "List"
            gsGlsAyuda = Replace(txtValAyuda, Chr(13) & Chr(10), ",")
            
            While Right(gsGlsAyuda, 1) = ","
                gsGlsAyuda = Left(gsGlsAyuda, Len(gsGlsAyuda) - 1)
            Wend
            If gsGlsAyuda = "" Then
                MsgBox "Lista de ayuda valores no debe quedar en blanco", vbInformation, App.Title
                Exit Sub
            End If
        Else
            gsTipoAyuda = "Query"
            gsGlsAyuda = txtValAyuda
            If gsGlsAyuda = "" Then
                MsgBox "Query de ayuda de valores no debe quedar en blanco", vbInformation, App.Title
                Exit Sub
            End If
        End If
    End If
    
    Unload Me
End Sub

Function FormatoLista(ByVal sGlsAyuda As String) As String
    Dim sFormatoLista   As String
    Dim nPos            As Integer
    
    sFormatoLista = ""
    nPos = InStr(sGlsAyuda, ",")
    While nPos > 0
        sFormatoLista = sFormatoLista & Left(sGlsAyuda, nPos - 1) & Chr(13) & Chr(10)
        sGlsAyuda = Mid(sGlsAyuda, nPos + 1)
        
        nPos = InStr(sGlsAyuda, ",")
    Wend
    sFormatoLista = sFormatoLista & sGlsAyuda & Chr(13) & Chr(10)
    
    FormatoLista = sFormatoLista
End Function

Private Sub chkAyudaValores_Click()
    If chkAyudaValores.Value = 0 Then
        optTipoAyuda(0).Enabled = False
        optTipoAyuda(1).Enabled = False
        txtValAyuda.Enabled = False
    Else
        optTipoAyuda(0).Enabled = True
        optTipoAyuda(1).Enabled = True
        txtValAyuda.Enabled = True
    End If
End Sub

Private Sub cmdAceptar_Click()
    AceptarParametro
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    IniciaForm
End Sub
Sub IniciaForm()
    Dim nPos    As Integer
    Dim nItem   As Integer
    
    Me.HelpContextID = 20
    Me.Left = (Screen.Width - Me.Width) \ 2
    Me.Top = (Screen.Height - Me.Height) \ 3
    
    gsTipoDato = ""
    gsGlsAyuda = ""
    
    nItem = 0
    
    optTipoAyuda(0).Enabled = False
    optTipoAyuda(1).Enabled = False
    txtValAyuda.Enabled = False
    
    Me.cboTipoDato.AddItem "Texto"
    cboTipoDato.ListIndex = cboTipoDato.ListCount - 1
    If LCase(cboTipoDato) = LCase(msTipoDato) Then
        nItem = cboTipoDato.ListCount - 1
    End If
    
    Me.cboTipoDato.AddItem "Entero"
    cboTipoDato.ListIndex = cboTipoDato.ListCount - 1
    If LCase(cboTipoDato) = LCase(msTipoDato) Then
        nItem = cboTipoDato.ListCount - 1
    End If
    
    Me.cboTipoDato.AddItem "Decimal"
    cboTipoDato.ListIndex = cboTipoDato.ListCount - 1
    If LCase(cboTipoDato) = LCase(msTipoDato) Then
        nItem = cboTipoDato.ListCount - 1
    End If
    
    Me.cboTipoDato.AddItem "Fecha"
    cboTipoDato.ListIndex = cboTipoDato.ListCount - 1
    If LCase(cboTipoDato) = LCase(msTipoDato) Then
        nItem = cboTipoDato.ListCount - 1
    End If
    
    Me.cboTipoDato.AddItem "Username"
    cboTipoDato.ListIndex = cboTipoDato.ListCount - 1
    If LCase(cboTipoDato) = LCase(msTipoDato) Then
        nItem = cboTipoDato.ListCount - 1
    End If
    
    cboTipoDato.ListIndex = nItem
    
    If msNomParametro <> "" Then
        txtNomParametro = msNomParametro
        txtGlsParametro = msGlsParametro
        txtNomParametro.Enabled = False
    Else
        txtNomParametro = ""
        txtGlsParametro = ""
        txtNomParametro.Enabled = True
    End If
    
    chkOpcional.Value = IIf(msIndOpcional = "S", 1, 0)
    
    If msGlsAyuda = "" Then
        chkAyudaValores.Value = 0
    Else
        chkAyudaValores.Value = 1
        If Trim(msGlsAyuda) = "" Then
            chkAyudaValores.Value = 0
        Else
            Select Case LCase(msTipoAyuda)
            Case "list"
                optTipoAyuda(0).Value = True
                txtValAyuda = FormatoLista(msGlsAyuda)
            Case "query"
                optTipoAyuda(1).Value = True
                txtValAyuda = msGlsAyuda
            Case Else
                chkAyudaValores.Value = 0
                txtValAyuda = ""
            End Select
            
        End If
    End If
End Sub

