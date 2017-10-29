VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{A8B3B723-0B5A-101B-B22E-00AA0037B2FC}#1.0#0"; "Grid32.ocx"
Begin VB.Form frmFecha 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   3720
   ClientLeft      =   8055
   ClientTop       =   4785
   ClientWidth     =   4335
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmFecha.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   4335
   ShowInTaskbar   =   0   'False
   Begin Threed.SSFrame fraFondo 
      Height          =   2715
      Left            =   0
      TabIndex        =   0
      Top             =   -60
      Width           =   4215
      _Version        =   65536
      _ExtentX        =   7435
      _ExtentY        =   4789
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShadowStyle     =   1
      Begin Threed.SSCommand cmdHoy 
         Height          =   315
         Left            =   960
         TabIndex        =   1
         Top             =   2280
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   556
         _StockProps     =   78
         Caption         =   "Hoy"
      End
      Begin Threed.SSPanel pnlMes 
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   180
         Width           =   3975
         _Version        =   65536
         _ExtentX        =   7011
         _ExtentY        =   556
         _StockProps     =   15
         Caption         =   "SSPanel1"
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.CommandButton cmdMesAnterior 
            Height          =   195
            Left            =   360
            Picture         =   "frmFecha.frx":038A
            Style           =   1  'Graphical
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   60
            Width           =   255
         End
         Begin VB.CommandButton cmdAñoAnterior 
            Height          =   195
            Left            =   60
            Picture         =   "frmFecha.frx":0428
            Style           =   1  'Graphical
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   60
            Width           =   255
         End
         Begin VB.CommandButton cmdAñoSiguiente 
            Height          =   195
            Left            =   3660
            Picture         =   "frmFecha.frx":04C6
            Style           =   1  'Graphical
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   60
            Width           =   255
         End
         Begin VB.CommandButton cmdMesSiguiente 
            Height          =   195
            Left            =   3360
            Picture         =   "frmFecha.frx":0564
            Style           =   1  'Graphical
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   60
            Width           =   255
         End
      End
      Begin MSGrid.Grid grdCalendario 
         Height          =   1695
         Left            =   120
         TabIndex        =   7
         Top             =   540
         Width           =   3975
         _Version        =   65536
         _ExtentX        =   7011
         _ExtentY        =   2990
         _StockProps     =   77
         ForeColor       =   16711680
         BackColor       =   16777215
         Rows            =   7
         Cols            =   7
         FixedCols       =   0
         ScrollBars      =   0
      End
      Begin Threed.SSCommand cmdNinguna 
         Height          =   315
         Left            =   2280
         TabIndex        =   8
         Top             =   2280
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   556
         _StockProps     =   78
         Caption         =   "Ninguna"
      End
   End
End
Attribute VB_Name = "frmFecha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim msMeses(12) As String
Dim mfMesActual As Date

Dim mnNumColHoy As Integer
Dim mnNumRowHoy As Integer

Sub CargaMes()
    Dim sFecha  As String
    Dim dFecha  As Date
    
    On Error GoTo ErrCargaMes
    
    sFecha = Format(Date, "dd/mm/yyyy")
    sFecha = ConvFormFecha(sFecha)
    dFecha = DateValue(sFecha)
    
    mfMesActual = DateAdd("d", -DatePart("d", dFecha) + 1, dFecha)
    Call MuestraMes(mfMesActual)
    Exit Sub
    
ErrCargaMes:
    Exit Sub
End Sub

Sub DimensionaForms()
    fraFondo.Left = 15
    fraFondo.Top = 15
    
    Me.Width = fraFondo.Width + 30
    Me.Top = gnTopControlFecha
    Me.Left = gnLeftControlFecha - Me.Width
    Me.Height = fraFondo.Height + 30
    
    grdCalendario.Row = 0
    grdCalendario.ColAlignment(0) = 2
    grdCalendario.ColAlignment(1) = 2
    grdCalendario.ColAlignment(2) = 2
    grdCalendario.ColAlignment(3) = 2
    grdCalendario.ColAlignment(4) = 2
    grdCalendario.ColAlignment(5) = 2
    grdCalendario.ColAlignment(6) = 2
    
    grdCalendario.Col = 0: grdCalendario.Text = "  Lu"
    grdCalendario.Col = 1: grdCalendario.Text = "  Ma"
    grdCalendario.Col = 2: grdCalendario.Text = "  Mi"
    grdCalendario.Col = 3: grdCalendario.Text = "  Ju"
    grdCalendario.Col = 4: grdCalendario.Text = "  Vi"
    grdCalendario.Col = 5: grdCalendario.Text = "  Sa"
    grdCalendario.Col = 6: grdCalendario.Text = "  Do"
End Sub

Sub IniciaVariables()
msMeses(1) = "Enero"
msMeses(2) = "Febrero"
msMeses(3) = "Marzo"
msMeses(4) = "Abril"
msMeses(5) = "Mayo"
msMeses(6) = "Junio"
msMeses(7) = "Julio"
msMeses(8) = "Agosto"
msMeses(9) = "Septiembre"
msMeses(10) = "Octubre"
msMeses(11) = "Noviembre"
msMeses(12) = "Diciembre"
End Sub

Sub MuestraMes(fMesProceso As Date)
    Dim bDebeProcesar   As Integer
    Dim nFila           As Long
    Dim nFilaAux        As Integer
    Dim nColumna        As Long
    Dim nColumnaIni     As Long
    Dim fFechaHoy       As Date
    Dim fUltimoDia      As Date
    Dim nMesActual      As Long
    Dim sGlsMes         As String

    sGlsMes = msMeses(Format$(fMesProceso, "m")) & " " & Format$(fMesProceso, "yyyy")
    pnlMes.Caption = sGlsMes
    
    'IniGrilla
    mnNumColHoy = -1
    mnNumRowHoy = -1
    
    fFechaHoy = fMesProceso
    fUltimoDia = DateAdd("d", -1, DateAdd("m", 1, fFechaHoy))
    nColumnaIni = Format$(fFechaHoy, "w", vbMonday) - 1
    nMesActual = Val(Format$(fFechaHoy, "m"))

    ' Limpia columnas iniciales
    grdCalendario.Row = 1
    For nColumna = 0 To nColumnaIni - 1
        grdCalendario.Col = nColumna
        grdCalendario.Text = ""
    Next nColumna

    nFila = 1
    While nMesActual = Val(Format$(fFechaHoy, "m"))
        For nColumna = nColumnaIni To grdCalendario.Cols - 1
            grdCalendario.Row = nFila
            grdCalendario.Col = nColumna
            grdCalendario.Text = DatePart("d", fFechaHoy)
            
            If fFechaHoy = Date Then
                mnNumColHoy = nColumna
                mnNumRowHoy = nFila
            End If
            
            fFechaHoy = DateAdd("d", 1, fFechaHoy)
            If nMesActual <> Val(Format$(fFechaHoy, "m")) Then
                Exit For
            End If
        Next nColumna
        nColumnaIni = 0
        nFila = nFila + 1
    Wend
    
    nColumna = nColumna + 1
    For nFilaAux = nFila - 1 To 6
        While nColumna < 7
            grdCalendario.Row = nFilaAux
            grdCalendario.Col = nColumna
            grdCalendario.Text = ""
            nColumna = nColumna + 1
        Wend
        nColumna = 0
    Next nFilaAux
    
    If mnNumColHoy = -1 Then
        grdCalendario.Col = 0
        grdCalendario.Row = 1
        grdCalendario.SelStartCol = 1
        grdCalendario.SelEndCol = 0
        grdCalendario.SelStartRow = 1
        grdCalendario.SelEndRow = 1
        
    Else
        grdCalendario.Col = mnNumColHoy
        grdCalendario.Row = mnNumRowHoy
        grdCalendario.SelStartCol = mnNumColHoy
        grdCalendario.SelEndCol = mnNumColHoy
        grdCalendario.SelStartRow = mnNumRowHoy
        grdCalendario.SelEndRow = mnNumRowHoy
    End If
End Sub

Private Sub cmdAñoAnterior_Click()
    mfMesActual = DateAdd("m", -12, mfMesActual)
    Call MuestraMes(mfMesActual)
End Sub

Private Sub cmdAñoSiguiente_Click()
    mfMesActual = DateAdd("m", 12, mfMesActual)
    Call MuestraMes(mfMesActual)
End Sub

Private Sub cmdHoy_Click()
    gsFechaSeleccionada = Format$(Date, "dd/mm/yyyy")
    Unload Me
End Sub

Private Sub cmdMesAnterior_Click()
    mfMesActual = DateAdd("m", -1, mfMesActual)
    Call MuestraMes(mfMesActual)
End Sub

Private Sub cmdMesSiguiente_Click()
    mfMesActual = DateAdd("m", 1, mfMesActual)
    Call MuestraMes(mfMesActual)
End Sub

Private Sub cmdNinguna_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    DimensionaForms
    IniciaVariables
    CargaMes
End Sub

Private Sub grdCalendario_Click()
    Dim sFecha As String
    
    If grdCalendario.Text = "" Then
        ' Vuelve a marcar el dia de hoy si corresponde
        If mnNumColHoy = -1 Then
            grdCalendario.Col = 0
            grdCalendario.Row = 1
            grdCalendario.SelStartCol = 1
            grdCalendario.SelEndCol = 0
            grdCalendario.SelStartRow = 1
            grdCalendario.SelEndRow = 1
            
        Else
            grdCalendario.Col = mnNumColHoy
            grdCalendario.Row = mnNumRowHoy
            grdCalendario.SelStartCol = mnNumColHoy
            grdCalendario.SelEndCol = mnNumColHoy
            grdCalendario.SelStartRow = mnNumRowHoy
            grdCalendario.SelEndRow = mnNumRowHoy
        End If
        cmdNinguna.SetFocus
        Exit Sub
    End If
    
    sFecha = Format$(mfMesActual, "dd/mm/yyyy")
    gsFechaSeleccionada = Right$("0" & grdCalendario.Text, 2) + Mid$(sFecha, 3)
    Unload Me
End Sub



