VERSION 5.00
Begin VB.Form frmDefFormato 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Formato especial de columnas"
   ClientHeight    =   3375
   ClientLeft      =   4515
   ClientTop       =   2685
   ClientWidth     =   7935
   Icon            =   "frmDefFormato.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   7935
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   5700
      TabIndex        =   6
      Top             =   2940
      Width           =   1095
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   4500
      TabIndex        =   5
      Top             =   2940
      Width           =   1095
   End
   Begin VB.Frame fraColumna 
      Caption         =   "Columna"
      Height          =   1215
      Left            =   60
      TabIndex        =   11
      Top             =   0
      Width           =   6735
      Begin VB.CommandButton cmdAnterior 
         Caption         =   "<"
         Height          =   255
         Left            =   5940
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   840
         Width           =   315
      End
      Begin VB.CommandButton cmdSiguiente 
         Caption         =   ">"
         Height          =   255
         Left            =   6300
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   840
         Width           =   315
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Título : "
         Height          =   195
         Left            =   240
         TabIndex        =   20
         Top             =   300
         Width           =   555
      End
      Begin VB.Label lblTitulo 
         AutoSize        =   -1  'True
         Caption         =   "Título"
         Height          =   195
         Left            =   1020
         TabIndex        =   19
         Top             =   300
         Width           =   420
      End
      Begin VB.Label lblTipoDato 
         AutoSize        =   -1  'True
         Caption         =   "TipoDato"
         Height          =   195
         Left            =   1020
         TabIndex        =   15
         Top             =   900
         Width           =   660
      End
      Begin VB.Label lblNomColumna 
         AutoSize        =   -1  'True
         Caption         =   "Columna"
         Height          =   195
         Left            =   1020
         TabIndex        =   14
         Top             =   600
         Width           =   615
      End
      Begin VB.Label lblBaseDatos 
         AutoSize        =   -1  'True
         Caption         =   "Tipo : "
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   900
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nombre : "
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   600
         Width           =   690
      End
   End
   Begin VB.Frame fraFormato 
      Caption         =   "Mostrar como"
      Height          =   1635
      Left            =   60
      TabIndex        =   7
      Top             =   1260
      Width           =   7755
      Begin VB.CommandButton cmdDown 
         Appearance      =   0  'Flat
         Height          =   150
         Left            =   3960
         Picture         =   "frmDefFormato.frx":038A
         Style           =   1  'Graphical
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   780
         Width           =   255
      End
      Begin VB.CommandButton cmdUp 
         Appearance      =   0  'Flat
         Height          =   150
         Left            =   3960
         Picture         =   "frmDefFormato.frx":042C
         Style           =   1  'Graphical
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   610
         Width           =   255
      End
      Begin VB.ListBox lstTipoDato 
         Height          =   1230
         Left            =   120
         TabIndex        =   0
         Top             =   300
         Width           =   1635
      End
      Begin VB.CheckBox chkMiles 
         Alignment       =   1  'Right Justify
         Caption         =   "Separador de miles"
         Height          =   195
         Left            =   1875
         TabIndex        =   1
         Top             =   300
         Width           =   1875
      End
      Begin VB.TextBox txtNumDecimales 
         Height          =   315
         Left            =   3540
         TabIndex        =   2
         Text            =   "0"
         Top             =   610
         Width           =   375
      End
      Begin VB.TextBox txtFormatoIn 
         Height          =   255
         Left            =   5940
         TabIndex        =   3
         Text            =   "dd/mm/yyyy"
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtFormatoOut 
         Height          =   255
         Left            =   5940
         TabIndex        =   4
         Text            =   "dd/mm/yyyy"
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblMuestra 
         AutoSize        =   -1  'True
         Caption         =   "Valor"
         Height          =   195
         Left            =   2700
         TabIndex        =   22
         Top             =   1200
         Width           =   360
      End
      Begin VB.Label lblTitMuestra 
         AutoSize        =   -1  'True
         Caption         =   "Muestra :"
         Height          =   195
         Left            =   1920
         TabIndex        =   21
         Top             =   1200
         Width           =   660
      End
      Begin VB.Label lblAyuda 
         AutoSize        =   -1  'True
         Caption         =   "dd : dia"
         Height          =   195
         Left            =   7080
         TabIndex        =   18
         Top             =   660
         Width           =   525
      End
      Begin VB.Label lblNumDecimales 
         AutoSize        =   -1  'True
         Caption         =   "Número de decimales"
         Height          =   195
         Left            =   1920
         TabIndex        =   10
         Top             =   660
         Width           =   1530
      End
      Begin VB.Label lblFormatoIn 
         AutoSize        =   -1  'True
         Caption         =   "Formato de entrada"
         Height          =   195
         Left            =   4500
         TabIndex        =   9
         Top             =   300
         Width           =   1380
      End
      Begin VB.Label lblFormatoOut 
         AutoSize        =   -1  'True
         Caption         =   "Formato de salida"
         Height          =   195
         Left            =   4500
         TabIndex        =   8
         Top             =   660
         Width           =   1245
      End
   End
End
Attribute VB_Name = "frmDefFormato"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mnNumColumna     As Integer
Public msNumConsulta    As String

Dim msTipoDato          As String
Dim mvValorColumna      As Variant
Dim mbCargandoCampo     As Boolean

Public mrsFormatos      As ADODB.Recordset
Dim arrFormatos()       As rRegFormatos

Sub CargaArregloFormatos()
    Dim nTotCampos  As Integer
    Dim nCol        As Integer
    Dim sFormatoIn  As String
    Dim sFormatoOut As String
    
    nTotCampos = frmEditarConsulta.mrsData.Fields.Count
    ReDim arrFormatos(nTotCampos) As rRegFormatos

    For nCol = 1 To nTotCampos
        ' Busca campo en la base de datos
        mrsFormatos.Filter = "nom_columna='" & LCase(frmEditarConsulta.mrsData(nCol - 1).Name) & "'"
        
        ' Si lo encuentra, carga informacion desde la base de datos
        If mrsFormatos.EOF Then
            arrFormatos(nCol).nom_columna = ""
        Else
            sFormatoIn = "" & mrsFormatos!gls_formato_entrada
            sFormatoOut = "" & mrsFormatos!gls_formato_salida
            
            sFormatoIn = Replace(sFormatoIn, gsSignoMenor, "<")
            sFormatoIn = Replace(sFormatoIn, gsSignoComillas, """")
            sFormatoOut = Replace(sFormatoOut, gsSignoMenor, "<")
            sFormatoOut = Replace(sFormatoOut, gsSignoComillas, """")
            
            arrFormatos(nCol).nom_columna = "" & mrsFormatos!nom_columna
            arrFormatos(nCol).cod_tipo_dato_salida = "" & mrsFormatos!cod_tipo_dato_salida
            arrFormatos(nCol).ind_separador_miles = "" & mrsFormatos!ind_separador_miles
            arrFormatos(nCol).num_decimales = "" & mrsFormatos!num_decimales
            arrFormatos(nCol).gls_formato_entrada = sFormatoIn
            arrFormatos(nCol).gls_formato_salida = sFormatoOut
        End If
    Next nCol
End Sub

Sub ConfiguraTipoDato(nCol As Integer)
    chkMiles.Visible = False
    chkMiles.Enabled = False
    lblNumDecimales.Visible = False
    lblNumDecimales.Enabled = False
    txtNumDecimales.Visible = False
    txtNumDecimales.Enabled = False
    lblFormatoIn.Visible = False
    lblFormatoIn.Enabled = False
    txtFormatoIn.Visible = False
    txtFormatoIn.Enabled = False
    lblFormatoOut.Visible = False
    lblFormatoOut.Enabled = False
    txtFormatoOut.Visible = False
    txtFormatoOut.Enabled = False
    cmdDown.Visible = False
    cmdDown.Enabled = False
    cmdUp.Visible = False
    cmdUp.Enabled = False
    lblAyuda.Visible = False
    
    Select Case lstTipoDato.ListIndex
    Case 0 ' Numero
        If msTipoDato = wc_tipo_dato_float Then
            txtNumDecimales.Text = IIf(arrFormatos(nCol).nom_columna = "", "2", arrFormatos(nCol).num_decimales)
        Else
            txtNumDecimales.Text = IIf(arrFormatos(nCol).nom_columna = "", "0", arrFormatos(nCol).num_decimales)
        End If
        chkMiles.Visible = True
        chkMiles.Enabled = True
        chkMiles.Value = IIf(arrFormatos(nCol).nom_columna = "", 1, IIf(arrFormatos(nCol).ind_separador_miles = "S", 1, 0))
        lblNumDecimales.Visible = True
        lblNumDecimales.Enabled = True
        txtNumDecimales.Visible = True
        txtNumDecimales.Enabled = True
        cmdDown.Visible = True
        cmdDown.Enabled = True
        cmdUp.Visible = True
        cmdUp.Enabled = True
        
    Case 1, 2 ' Fecha u Hora
        If msTipoDato = wc_tipo_dato_fecha Then
            lblFormatoOut.Top = chkMiles.Top
            txtFormatoOut.Top = chkMiles.Top
        
            lblFormatoOut.Visible = True
            lblFormatoOut.Enabled = True
            txtFormatoOut.Visible = True
            txtFormatoOut.Enabled = True
        
            txtFormatoIn.Text = IIf(arrFormatos(nCol).nom_columna = "", "", arrFormatos(nCol).gls_formato_entrada)
            txtFormatoOut.Text = IIf(arrFormatos(nCol).nom_columna = "", "", arrFormatos(nCol).gls_formato_salida)
        Else
            lblFormatoIn.Top = chkMiles.Top
            txtFormatoIn.Top = chkMiles.Top
            lblFormatoOut.Top = chkMiles.Top + txtFormatoIn.Height + 30
            txtFormatoOut.Top = chkMiles.Top + txtFormatoIn.Height + 30
            
            lblFormatoIn.Visible = True
            lblFormatoIn.Enabled = True
            txtFormatoIn.Visible = True
            txtFormatoIn.Enabled = True
            
            lblFormatoOut.Visible = True
            lblFormatoOut.Enabled = True
            txtFormatoOut.Visible = True
            txtFormatoOut.Enabled = True
        
            txtFormatoIn.Text = IIf(arrFormatos(nCol).nom_columna = "", "", arrFormatos(nCol).gls_formato_entrada)
            txtFormatoOut.Text = IIf(arrFormatos(nCol).nom_columna = "", "", arrFormatos(nCol).gls_formato_salida)
        End If
        
        lblAyuda.Top = txtFormatoOut.Top
        lblAyuda.Left = txtFormatoOut.Left + txtFormatoOut.Width + 60
        If lstTipoDato.ListIndex = 1 Then ' Fecha
            lblAyuda = "(y=año; m=mes; d=día)"
            lblAyuda.Visible = True
        Else
            lblAyuda = "(h=hora; m=minuto; s=segundo)"
            lblAyuda.Visible = True
        End If
        
    End Select

    MuestraValor
End Sub

Function GrabaRecordsetFormatos() As Boolean
    Dim nCol            As Integer
    Dim rsFormatos      As ADODB.Recordset
    
    On Error GoTo ErrGrabaRecordsetFormatos
    
    Set rsFormatos = New ADODB.Recordset
    rsFormatos.CursorLocation = adUseClient
    rsFormatos.CursorType = adOpenStatic
    rsFormatos.ActiveConnection = Nothing
    rsFormatos.LockType = adLockBatchOptimistic
    
    Call rsFormatos.Fields.Append("num_consulta", adInteger, , adFldIsNullable)
    Call rsFormatos.Fields.Append("nom_columna", adVarChar, 50, adFldIsNullable)
    Call rsFormatos.Fields.Append("cod_tipo_dato_salida", adVarChar, 12, adFldIsNullable)
    Call rsFormatos.Fields.Append("ind_separador_miles", adVarChar, 1, adFldIsNullable)
    Call rsFormatos.Fields.Append("num_decimales", adInteger, , adFldIsNullable)
    Call rsFormatos.Fields.Append("gls_formato_entrada", adVarChar, 132, adFldIsNullable)
    Call rsFormatos.Fields.Append("gls_formato_salida", adVarChar, 132, adFldIsNullable)
    
    rsFormatos.Open
    

    Screen.MousePointer = vbHourglass
    
    For nCol = 1 To UBound(arrFormatos)
        If arrFormatos(nCol).nom_columna <> "" Then
            rsFormatos.AddNew
            
            rsFormatos.Fields(1).Value = LCase(arrFormatos(nCol).nom_columna)
            rsFormatos.Fields(2).Value = arrFormatos(nCol).cod_tipo_dato_salida
            rsFormatos.Fields(3).Value = arrFormatos(nCol).ind_separador_miles
            rsFormatos.Fields(4).Value = arrFormatos(nCol).num_decimales
            rsFormatos.Fields(5).Value = arrFormatos(nCol).gls_formato_entrada
            rsFormatos.Fields(6).Value = arrFormatos(nCol).gls_formato_salida
            rsFormatos.MoveLast
        End If
    Next nCol
    
    Set frmEditarConsulta.mrsFormatos = rsFormatos
    
    Screen.MousePointer = vbNormal
    
    GrabaRecordsetFormatos = True
    Exit Function
    
ErrGrabaRecordsetFormatos:
    Screen.MousePointer = vbNormal
    MsgBox Err.Description, vbCritical, App.Title
    GrabaRecordsetFormatos = False
End Function

Function GrabaFormatos() As Boolean
    Dim nCol            As Integer
    Dim sGlsFormatos    As String
    Dim sFormatoIn      As String
    Dim sFormatoOut     As String
    
    On Error GoTo ErrGrabaFormatos

    Screen.MousePointer = vbHourglass
    
    sGlsFormatos = "<ROOT>"
    For nCol = 1 To UBound(arrFormatos)
        If arrFormatos(nCol).nom_columna <> "" Then
            sFormatoIn = Replace(arrFormatos(nCol).gls_formato_entrada, "<", gsSignoMenor)
            sFormatoIn = Replace(sFormatoIn, """", gsSignoComillas)
            sFormatoOut = Replace(arrFormatos(nCol).gls_formato_salida, "<", gsSignoMenor)
            sFormatoOut = Replace(sFormatoOut, """", gsSignoComillas)
            
            sGlsFormatos = sGlsFormatos & "<Formatos"
            sGlsFormatos = sGlsFormatos & " nom_columna=""" & LCase(arrFormatos(nCol).nom_columna) & """"
            sGlsFormatos = sGlsFormatos & " cod_tipo_dato_salida=""" & arrFormatos(nCol).cod_tipo_dato_salida & """"
            sGlsFormatos = sGlsFormatos & " ind_separador_miles=""" & arrFormatos(nCol).ind_separador_miles & """"
            sGlsFormatos = sGlsFormatos & " num_decimales=""" & arrFormatos(nCol).num_decimales & """"
            sGlsFormatos = sGlsFormatos & " gls_formato_entrada=""" & sFormatoIn & """"
            sGlsFormatos = sGlsFormatos & " gls_formato_salida=""" & sFormatoOut & """"
            sGlsFormatos = sGlsFormatos & "/>"
        End If
    Next nCol
    sGlsFormatos = sGlsFormatos & "</ROOT>"
    
    ' Graba informacion
    GrabaFormatos = db_GrabaFormatosConsulta(msNumConsulta, sGlsFormatos)
    
    Screen.MousePointer = vbNormal
    Exit Function
    
ErrGrabaFormatos:
    Screen.MousePointer = vbNormal
    MsgBox Err.Description, vbCritical, App.Title
    GrabaFormatos = False
End Function


Function GuardaFormatos() As Boolean
    If msNumConsulta = "" Then
        GuardaFormatos = GrabaRecordsetFormatos
    Else
        GuardaFormatos = GrabaFormatos
    End If
End Function


Sub InicializaFormatoColumna(nCol)
    arrFormatos(nCol).nom_columna = lblNomColumna
    If lstTipoDato.ListIndex = 0 Then
        If Val(Me.txtNumDecimales) = 0 Then
            arrFormatos(nCol).cod_tipo_dato_salida = "Entero"
        Else
            arrFormatos(nCol).cod_tipo_dato_salida = "Decimal"
        End If
    Else
        arrFormatos(nCol).cod_tipo_dato_salida = lstTipoDato
    End If
End Sub

Sub MuestraColumna(nCol As Integer)
    Dim nTipoDatoSalida As Integer
    
    mbCargandoCampo = True
    
    msTipoDato = fnTipoDatoRecordset(frmEditarConsulta.mrsData(nCol).Type)
    
    lblNomColumna = frmEditarConsulta.mrsData(nCol).Name
    lblTitulo = FormatoTitulo(frmEditarConsulta.mrsData(nCol).Name)
    mvValorColumna = frmEditarConsulta.mrsData(nCol).Value
    
    If arrFormatos(nCol + 1).nom_columna = "" Then
        nTipoDatoSalida = msTipoDato
    Else
        nTipoDatoSalida = fnTipoDato(arrFormatos(nCol + 1).cod_tipo_dato_salida)
    End If
    
    Select Case msTipoDato
    Case wc_tipo_dato_integer
        lblTipoDato = "Entero"
    Case wc_tipo_dato_float
        lblTipoDato = "Decimal"
    Case wc_tipo_dato_fecha
        lblTipoDato = "Fecha"
    Case wc_tipo_dato_otro
        lblTipoDato = "Texto"
    End Select
       
    Select Case nTipoDatoSalida
    Case wc_tipo_dato_integer
        lstTipoDato.ListIndex = 0
    Case wc_tipo_dato_float
        lstTipoDato.ListIndex = 0
    Case wc_tipo_dato_fecha
        lstTipoDato.ListIndex = 1
    '<INI SP1.2.1>
    Case wc_tipo_dato_hora
        lstTipoDato.ListIndex = 2
    '<FIN SP1.2.1>
    Case wc_tipo_dato_otro
        lstTipoDato.ListIndex = 3
    End Select
       
    frmEditarConsulta.grdResultado.Col = nCol + 1
    frmEditarConsulta.grdResultado.Action = 0
    
    Me.cmdSiguiente.Enabled = (mnNumColumna < frmEditarConsulta.mrsData.Fields.Count - 1)
    Me.cmdAnterior.Enabled = (mnNumColumna > 0)
    
    Call ConfiguraTipoDato(nCol + 1)
    MuestraValor
    
    mbCargandoCampo = False
End Sub


Sub MuestraValor()

    Select Case Me.lstTipoDato.ListIndex
    Case 0 ' Numero
        lblMuestra = fsFormatoValorNumerico(mvValorColumna, IIf(chkMiles.Value = 1, "S", "N"), txtNumDecimales)
        
    Case 1, 2 ' Fecha u Hora
        lblMuestra = fsFormatoValorFecha(mvValorColumna, txtFormatoIn.Text, txtFormatoOut.Text)
    
    Case Else
        lblMuestra = "" & mvValorColumna
    End Select


End Sub


Private Sub chkMiles_Click()
    If Not mbCargandoCampo Then
        If arrFormatos(mnNumColumna + 1).nom_columna = "" Then
            Call InicializaFormatoColumna(mnNumColumna + 1)
        End If
        cmdAceptar.Enabled = True
        arrFormatos(mnNumColumna + 1).ind_separador_miles = IIf(chkMiles.Value = 1, "S", "N")
    End If
    
    MuestraValor
End Sub


Private Sub cmdAceptar_Click()
    If GuardaFormatos Then
        cmdAceptar.Enabled = False
        gbCancelar = False
        Unload Me
    End If
End Sub

Private Sub cmdAnterior_Click()
    If mnNumColumna > 0 Then
        mnNumColumna = mnNumColumna - 1
        Call MuestraColumna(mnNumColumna)
    End If
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdDown_Click()
    If Val(txtNumDecimales) <= 1 Then
        txtNumDecimales = "0"
    Else
        txtNumDecimales = Trim(Val(txtNumDecimales) - 1)
    End If
    txtNumDecimales.SetFocus
End Sub

Private Sub cmdSiguiente_Click()
    If mnNumColumna < frmEditarConsulta.mrsData.Fields.Count - 1 Then
        mnNumColumna = mnNumColumna + 1
        Call MuestraColumna(mnNumColumna)
    End If
End Sub

Private Sub cmdUp_Click()
    If Val(txtNumDecimales) = 0 Then
        txtNumDecimales = "1"
    Else
        txtNumDecimales = Trim(Val(txtNumDecimales) + 1)
    End If
    txtNumDecimales.SetFocus
End Sub

Private Sub Form_Load()
    CargaArregloFormatos
    IniciaForm
End Sub
Sub IniciaForm()
    Dim nDelta  As Long
    
    Me.HelpContextID = 21
    mbCargandoCampo = False
    
    fraFormato.Width = fraColumna.Width
    Me.Width = fraFormato.Left + fraFormato.Width + 120
    
    nDelta = txtFormatoIn.Left - lblFormatoIn.Left
    lblFormatoIn.Left = chkMiles.Left
    txtFormatoIn.Left = lblFormatoIn.Left + nDelta
    lblFormatoOut.Left = lblFormatoIn.Left
    txtFormatoOut.Left = txtFormatoIn.Left
    
    Me.Left = (Screen.Width - Me.Width) \ 2
    Me.Top = (Screen.Height - Me.Height) \ 2
    
    lstTipoDato.AddItem "Número"
    lstTipoDato.AddItem "Fecha"
    lstTipoDato.AddItem "Hora"
    lstTipoDato.AddItem "Texto"
    
    Call MuestraColumna(mnNumColumna)
    cmdAceptar.Enabled = False
    gbCancelar = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim nResp   As Integer
    
    If cmdAceptar.Enabled Then
        nResp = MsgBox("Desea grabar los cambios realizados sobre esta consulta", vbYesNoCancel + vbDefaultButton1, App.Title)
        If nResp = vbYes Then
            If Not GuardaFormatos() Then
                Cancel = True
            Else
                cmdAceptar.Enabled = False
                gbCancelar = False
            End If
        ElseIf nResp = vbCancel Then
            Cancel = True
        End If
    End If
End Sub

Private Sub lblBaseDatos_DblClick()
    On Error Resume Next
    MsgBox "Provider (" & cnn_Consulta.Provider & "), Type (" & frmEditarConsulta.mrsData(mnNumColumna).Type & ")"
End Sub


Private Sub lstTipoDato_Click()
    If Not mbCargandoCampo Then
        If arrFormatos(mnNumColumna + 1).nom_columna = "" Then
            Call InicializaFormatoColumna(mnNumColumna + 1)
        End If
        cmdAceptar.Enabled = True
        If lstTipoDato.ListIndex = 0 Then
            If Val(Me.txtNumDecimales) = 0 Then
                arrFormatos(mnNumColumna + 1).cod_tipo_dato_salida = "Entero"
            Else
                arrFormatos(mnNumColumna + 1).cod_tipo_dato_salida = "Decimal"
            End If
        Else
            arrFormatos(mnNumColumna + 1).cod_tipo_dato_salida = lstTipoDato
        End If
    End If
    
    Call ConfiguraTipoDato(mnNumColumna + 1)
End Sub

Private Sub txtFormatoIn_Change()
    If Not mbCargandoCampo Then
        If arrFormatos(mnNumColumna + 1).nom_columna = "" Then
            Call InicializaFormatoColumna(mnNumColumna + 1)
        End If
        cmdAceptar.Enabled = True
        arrFormatos(mnNumColumna + 1).gls_formato_entrada = txtFormatoIn.Text
    End If
    
    MuestraValor
End Sub

Private Sub txtFormatoOut_Change()
    If Not mbCargandoCampo Then
        If arrFormatos(mnNumColumna + 1).nom_columna = "" Then
            Call InicializaFormatoColumna(mnNumColumna + 1)
        End If
        cmdAceptar.Enabled = True
        arrFormatos(mnNumColumna + 1).gls_formato_salida = txtFormatoOut.Text
    End If
    
    MuestraValor
End Sub

Private Sub txtNumDecimales_Change()
    If Not mbCargandoCampo Then
        If arrFormatos(mnNumColumna + 1).nom_columna = "" Then
            Call InicializaFormatoColumna(mnNumColumna + 1)
        End If
        cmdAceptar.Enabled = True
        arrFormatos(mnNumColumna + 1).num_decimales = txtNumDecimales.Text
        arrFormatos(mnNumColumna + 1).ind_separador_miles = IIf(chkMiles.Value = 1, "S", "N")
    End If
    
    MuestraValor
End Sub


Private Sub txtNumDecimales_GotFocus()
    txtNumDecimales.SelStart = 0
    txtNumDecimales.SelLength = Len(txtNumDecimales)
End Sub


