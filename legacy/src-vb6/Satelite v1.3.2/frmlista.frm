VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmListaOpciones 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ayuda para "
   ClientHeight    =   1380
   ClientLeft      =   2865
   ClientTop       =   2610
   ClientWidth     =   5700
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "frmlista.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1380
   ScaleWidth      =   5700
   Begin Threed.SSCommand cmdCancelar 
      Height          =   375
      Left            =   4320
      TabIndex        =   2
      Top             =   900
      Width           =   1275
      _Version        =   65536
      _ExtentX        =   2249
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "&Cancelar"
   End
   Begin Threed.SSCommand cmdAceptar 
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   900
      Width           =   1275
      _Version        =   65536
      _ExtentX        =   2249
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "&Aceptar"
   End
   Begin Threed.SSPanel pnlFondo 
      Height          =   555
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   5595
      _Version        =   65536
      _ExtentX        =   9869
      _ExtentY        =   979
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   1
      Autosize        =   3
      Begin FPSpreadADO.fpSpread grdDatos 
         Height          =   525
         Left            =   15
         TabIndex        =   4
         Top             =   15
         Visible         =   0   'False
         Width           =   5565
         _Version        =   524288
         _ExtentX        =   9816
         _ExtentY        =   926
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SpreadDesigner  =   "frmlista.frx":038A
         UnitType        =   0
      End
   End
   Begin VB.Label lblLargo 
      AutoSize        =   -1  'True
      Caption         =   "Largo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   360
      TabIndex        =   3
      Top             =   840
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "&Archivo"
      Begin VB.Menu mnuArcSalir 
         Caption         =   "&Salir"
      End
   End
   Begin VB.Menu mnuBuscar 
      Caption         =   "&Buscar"
      Begin VB.Menu mnuBusString 
         Caption         =   "Bu&scar..."
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuBusSiguiente 
         Caption         =   "B&uscar siguiente"
         Shortcut        =   +{F3}
      End
   End
End
Attribute VB_Name = "frmListaOpciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mnAnchoGrilla   As Long
Dim mnColResult     As Integer
Dim msTextoBusqueda As String
Private Sub AjustaForm()
    Dim nAlto   As Long
    Dim nFilas  As Long
    Dim nAnchoMin   As Long

    nAnchoMin = cmdAceptar.Width + cmdCancelar.Width + 30
    If mnAnchoGrilla < nAnchoMin Then
        mnAnchoGrilla = nAnchoMin
    End If

    If mnAnchoGrilla > Screen.Width * 0.9 Then
        mnAnchoGrilla = Screen.Width * 0.9
    End If

    nFilas = grdDatos.MaxRows
    If nFilas > 16 Then
        nFilas = 16
    End If
    nAlto = (nFilas + 2) * (grdDatos.RowHeight(0) + 15) + 60
    pnlFondo.Height = nAlto

    pnlFondo.Width = mnAnchoGrilla + 110

    If nFilas <= grdDatos.MaxRows Then
        pnlFondo.Width = pnlFondo.Width + 255
    End If

    Me.Width = pnlFondo.Width + 215
    Me.Height = pnlFondo.Height + cmdAceptar.Height + 850

    cmdCancelar.Top = pnlFondo.Top + pnlFondo.Height + 60
    cmdCancelar.Left = Me.Width - cmdCancelar.Width - 150
    cmdAceptar.Top = cmdCancelar.Top
    cmdAceptar.Left = cmdCancelar.Left - cmdAceptar.Width - 15

    If frmListaOpciones.Left + frmListaOpciones.Width > (Screen.Width - 500) Then
        frmListaOpciones.Left = (Screen.Width - 500) - frmListaOpciones.Width
    End If
    If frmListaOpciones.Top + frmListaOpciones.Height > Screen.Height - 500 Then
        frmListaOpciones.Top = Screen.Height - frmListaOpciones.Height - 500
    End If
    
    pnlFondo.Visible = True

    If grdDatos.MaxRows = 2 And fsGetGrilla(grdDatos, 1, 1) = "" Then
        grdDatos.Visible = False
        cmdAceptar.Visible = False
        pnlFondo.Caption = "No hay ayuda disponible para este campo."
        cmdCancelar.Caption = "&Aceptar"
    Else
        grdDatos.Visible = True
        grdDatos.SetFocus
    End If
End Sub

Sub BuscarSiguiente()
    Dim nFila   As Long

    If msTextoBusqueda <> "" Then
        nFila = fnFindStr(grdDatos, grdDatos.ActiveCol, msTextoBusqueda, grdDatos.ActiveRow + 1)
        If nFila <= 0 Then
            MsgBox "No hay mas filas con """ & msTextoBusqueda & """ dentro de esta columna", vbInformation, App.Title
        Else
            grdDatos.Row = nFila
            grdDatos.Col = grdDatos.ActiveCol
            grdDatos.Action = 0
        End If
    End If
End Sub

Sub BuscarString()
    Dim nFila   As Long

    msTextoBusqueda = Trim(InputBox("Ingrese texto a buscar"))
    If msTextoBusqueda = "" Then
        Me.mnuBusSiguiente.Enabled = False
    Else
        nFila = fnFindStr(grdDatos, grdDatos.ActiveCol, msTextoBusqueda)
        If nFila <= 0 Then
            MsgBox "No se encontró el texto dentro de esta columna", vbInformation, App.Title
            Me.mnuBusSiguiente.Enabled = False
        Else
            grdDatos.Row = nFila
            grdDatos.Col = grdDatos.ActiveCol
            grdDatos.Action = 0
            Me.mnuBusSiguiente.Enabled = True
        End If
    End If
End Sub

Private Sub CargaLista()
    Dim sSql        As String
    Dim bOk         As Integer
    Dim nNumCampos  As Integer
    Dim nCol        As Integer
    Dim nRow        As Integer
    Dim fCampos     As Field
    Dim nAncho()    As Long

    Screen.MousePointer = vbHourglass
    
    pnlFondo.Visible = False
    pnlFondo.Refresh
        
    On Error GoTo ErrCargaLista
    
    grdDatos.UnitType = 2
    grdDatos.RowHeadersShow = False
    mnColResult = 1

    ' Cuenta los campos del query
    nNumCampos = 0
    For Each fCampos In grsLookUp.Fields
        nNumCampos = nNumCampos + 1
    Next
    ReDim nAncho(nNumCampos) As Long
    
    ' Crea titulos de las celdas
    grdDatos.MaxCols = nNumCampos
    grdDatos.MaxRows = 1
    nCol = 1
    For Each fCampos In grsLookUp.Fields
        lblLargo = FormatoTitulo(fCampos.Name)
        If lblLargo.Width > nAncho(nCol) Then
            nAncho(nCol) = lblLargo.Width
        End If
        Call fnPutGrilla(grdDatos, 0, nCol, lblLargo.Caption)
        If LCase(fCampos.Name) = LCase(gsCampoLookUp) Then
            mnColResult = nCol
        End If
        
        nCol = nCol + 1
    Next

    nRow = 0
    While Not grsLookUp.EOF
        nRow = nRow + 1
        For nCol = 1 To nNumCampos
            lblLargo = Format(grsLookUp(nCol - 1), "&")
            If lblLargo.Width > nAncho(nCol) Then
                nAncho(nCol) = lblLargo.Width
            End If
            Call fnPutGrilla(grdDatos, nRow, nCol, lblLargo.Caption)
        Next nCol
        grsLookUp.MoveNext
    Wend
    
    ' COnfigura celdas segun tipo de dato
    nCol = 1
    mnAnchoGrilla = 0
    For Each fCampos In grsLookUp.Fields
        Call ConfigColumna(grdDatos, nCol, fCampos.Type, fCampos.NumericScale)
        grdDatos.ColWidth(nCol) = nAncho(nCol) + 240
        mnAnchoGrilla = mnAnchoGrilla + grdDatos.ColWidth(nCol)
        nCol = nCol + 1
    Next
    
    Screen.MousePointer = vbNormal

    grdDatos.Row = 1
    grdDatos.Col = 1
    msTextoBusqueda = ""
    Exit Sub
    
ErrCargaLista:
    Screen.MousePointer = vbNormal
    MsgBox Error, vbCritical, App.Title
    Exit Sub
End Sub

Private Sub cmdAceptar_Click()
    Dim nFila   As Integer

    nFila = grdDatos.ActiveRow
    gsResultLookUp = fsGetGrilla(grdDatos, nFila, mnColResult)
    gbCancelar = False
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub fnBuscaFila(sLetra As String)
    Dim nRowOld As Integer
    Dim nRow    As Integer
    Dim nCol    As Integer

    nRowOld = grdDatos.Row
    nCol = grdDatos.ActiveCol
    For nRow = nRowOld + 1 To grdDatos.MaxRows
        If UCase(Left(fsGetGrilla(grdDatos, nRow, nCol), 1)) = UCase(sLetra) Then
            grdDatos.Row = nRow
            grdDatos.Col = nCol
            grdDatos.Action = 0
            Exit Sub
        End If
    Next nRow
    For nRow = 1 To nRowOld - 1
        If UCase(Left(fsGetGrilla(grdDatos, nRow, nCol), 1)) = UCase(sLetra) Then
            grdDatos.Row = nRow
            grdDatos.Col = nCol
            grdDatos.Action = 0
            Exit Sub
        End If
    Next nRow
    grdDatos.Row = nRowOld
End Sub

Private Sub Form_Activate()
    CargaLista
    AjustaForm
    gbCancelar = True
End Sub

Private Sub grdDatos_DblClick(ByVal Col As Long, ByVal Row As Long)
    Call cmdAceptar_Click
End Sub

Private Sub grdDatos_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case 13
        cmdAceptar_Click
    Case 32 To 255
        Call fnBuscaFila(Chr(KeyAscii))
    End Select
End Sub

Private Sub mnuArcSalir_Click()
    cmdCancelar_Click
End Sub


Private Sub mnuBusSiguiente_Click()
    Call BuscarSiguiente
End Sub

Private Sub mnuBusString_Click()
    Call BuscarString
End Sub


