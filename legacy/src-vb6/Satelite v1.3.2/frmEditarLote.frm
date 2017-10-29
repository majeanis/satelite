VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "Comctl32.ocx"
Begin VB.Form frmEditarLote 
   Caption         =   "Edición de Lotes"
   ClientHeight    =   7530
   ClientLeft      =   4365
   ClientTop       =   2250
   ClientWidth     =   8790
   Icon            =   "frmEditarLote.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7530
   ScaleWidth      =   8790
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   6420
      TabIndex        =   4
      Top             =   7080
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   7620
      TabIndex        =   5
      Top             =   7080
      Width           =   1095
   End
   Begin VB.Frame fraConsultas 
      Caption         =   "Consultas"
      Height          =   5475
      Left            =   60
      TabIndex        =   8
      Top             =   1560
      Width           =   8655
      Begin VB.ListBox lstConsultas 
         Height          =   5010
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   3
         Top             =   300
         Width           =   8415
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Descripción del lote"
      Height          =   1455
      Left            =   60
      TabIndex        =   6
      Top             =   60
      Width           =   8655
      Begin VB.CommandButton cmdHlpSolicitante 
         Caption         =   "..."
         Height          =   315
         Index           =   0
         Left            =   3000
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   960
         Width           =   255
      End
      Begin VB.CheckBox chkIndAsignaSolicitante 
         Alignment       =   1  'Right Justify
         Caption         =   "Asignar solicitante al Lote automáticamente"
         Height          =   195
         Left            =   3780
         TabIndex        =   2
         Top             =   1020
         Width           =   3435
      End
      Begin VB.TextBox txtNomSolicitante 
         Height          =   315
         Left            =   1020
         TabIndex        =   1
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox txtNomLote 
         Height          =   315
         Left            =   1020
         TabIndex        =   0
         Top             =   600
         Width           =   7455
      End
      Begin VB.Label lblSolicitante 
         AutoSize        =   -1  'True
         Caption         =   "Solicitante :"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   1020
         Width           =   825
      End
      Begin VB.Label lblNumLote 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1020
         TabIndex        =   10
         Top             =   300
         Width           =   735
      End
      Begin VB.Label lblTitNumConsulta 
         AutoSize        =   -1  'True
         Caption         =   "Id :"
         Height          =   195
         Left            =   540
         TabIndex        =   9
         Top             =   300
         Width           =   225
      End
      Begin VB.Label lblTitNomUsuario 
         AutoSize        =   -1  'True
         Caption         =   "Nombre :"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   660
         Width           =   645
      End
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   840
      Top             =   7140
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   13
      ImageHeight     =   13
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   2
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmEditarLote.frx":038A
            Key             =   "Blanco"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmEditarLote.frx":08AC
            Key             =   "Marcado"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuOrdenar 
      Caption         =   "&Ordenar"
      Visible         =   0   'False
      Begin VB.Menu mnuOrdenarPor 
         Caption         =   "Ordenar por &Nombre"
         Index           =   0
      End
      Begin VB.Menu mnuOrdenarPor 
         Caption         =   "Ordenar por &Id"
         Index           =   1
      End
   End
End
Attribute VB_Name = "frmEditarLote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim bOk     As Boolean
Dim rsData  As ADODB.Recordset

Sub FormResize()

End Sub

Sub HelpUsuario()
    Dim sGlosaCampos    As String
    Dim sCampos         As String
    Dim sGlosaWhere     As String
    Dim sGlosaOrder     As String
    Dim sGlsAyuda       As String
    Dim nVal            As Integer
    Dim nX              As Integer

    gsQueryLookUp = ""
    sCampos = ""
    sGlosaWhere = ""
    sGlosaOrder = ""
    
    ' Carga recordset con Ayuda de Valores
    sCampos = "nom_usuario"
    sGlosaOrder = "nom_usuario"
    sGlsAyuda = "select nom_usuario from usuarios order by nom_usuario"
    
    OpenMyDataBase
    Set grsLookUp = New ADODB.Recordset
    grsLookUp.CursorLocation = adUseClient
    grsLookUp.Open sGlsAyuda, Cnn_Satelite, adOpenForwardOnly, adLockReadOnly
    

    ' Posiciona el formulario de ayuda
    frmListaOpciones.Top = Me.Top + Me.txtNomSolicitante.Top + 330
    frmListaOpciones.Left = Me.Left + Me.txtNomSolicitante.Left + Me.txtNomSolicitante.Width + 150
    frmListaOpciones.Caption = frmListaOpciones.Caption & Me.lblSolicitante
    
    gsCampoLookUp = "nom_usuario"
    frmListaOpciones.Show vbModal

    grsLookUp.Close
    Set grsLookUp = Nothing
    CloseMyDataBase

    If Not gbCancelar Then
        Me.txtNomSolicitante = gsResultLookUp
        Me.txtNomSolicitante.SetFocus
    End If
    gbCancelar = False
End Sub

Sub OrdenaConsultas()
    Dim nItem           As Long
    Dim sConsultas      As String
    
    Screen.MousePointer = vbHourglass
    
    ' Guarda las consultas que ya estaba seleccionadas
    sConsultas = ""
    For nItem = 0 To Me.lstConsultas.ListCount - 1
        If Me.lstConsultas.Selected(nItem) Then
            sConsultas = sConsultas & "<" & Trim(Me.lstConsultas.ItemData(nItem)) & ">"
        End If
    Next nItem
    
    ' Carga las consultas nuevamente en el orden solicitado
    lstConsultas.Clear
    If Me.mnuOrdenarPor(0).Checked Then
        rsData.Sort = "nom_consulta"
    Else
        rsData.Sort = "num_consulta"
    End If
    
    While Not rsData.EOF
        lstConsultas.AddItem "" & rsData!nom_consulta & " (Id. " & rsData!num_consulta & ")"
        lstConsultas.ItemData(lstConsultas.ListCount - 1) = rsData!num_consulta
        
        If InStr(sConsultas, "<" & Trim(rsData!num_consulta) & ">") > 0 Then
            lstConsultas.Selected(lstConsultas.ListCount - 1) = True
        End If
        
        rsData.MoveNext
    Wend

    Screen.MousePointer = vbNormal
End Sub

Private Sub cmdAceptar_Click()
    If GrabarLote() Then
        Unload Me
    End If
End Sub
Function GrabarLote() As Boolean
    Dim sNumLote        As String
    Dim sNomLote        As String
    Dim sNomSolicitante As String
    Dim sIndAsignarLote As String
    Dim sGlsConsultas   As String
    Dim nItem           As Long
    Dim nCta            As Long
    
    On Error GoTo ErrGrabarLote
        
    sNumLote = IIf(gsNumLote = "", "0", gsNumLote)
    sNomSolicitante = Me.txtNomSolicitante
    sIndAsignarLote = IIf(Me.chkIndAsignaSolicitante.Value = 0, "N", "S")
    
    ' Valida consistencia de informacion
    sNomLote = Trim(Me.txtNomLote)
    If sNomLote = "" Then
        MsgBox "No ha ingresado el nombre del lote", vbCritical, App.Title
        GrabarLote = False
        Exit Function
    End If
    
    ' Valida que existe al menos una consulta asignada al lote
    nCta = 0
    For nItem = 0 To Me.lstConsultas.ListCount - 1
        If Me.lstConsultas.Selected(nItem) Then
            nCta = nCta + 1
            Exit For
        End If
    Next nItem
    
    If nCta = 0 Then
        MsgBox "No ha seleccionado ninguna consulta al lote", vbCritical, App.Title
        GrabarLote = False
        Exit Function
    End If
    
    Screen.MousePointer = vbHourglass
    
    ' Genera XML con las consultas seleccionadas
    sGlsConsultas = "<ROOT>"
    For nItem = 0 To Me.lstConsultas.ListCount - 1
        If Me.lstConsultas.Selected(nItem) Then
            sGlsConsultas = sGlsConsultas & "<ConsLote "
            sGlsConsultas = sGlsConsultas & " num_consulta=""" & Trim(Me.lstConsultas.ItemData(nItem)) & """"
            sGlsConsultas = sGlsConsultas & "/>"
        End If
    Next nItem
    sGlsConsultas = sGlsConsultas & "</ROOT>"
    
    ' Graba informacion
    bOk = db_GrabaLote(sNumLote, sNomLote, sNomSolicitante, sIndAsignarLote, sGlsConsultas)
    
    Screen.MousePointer = vbNormal
    
    If Not bOk Then
        GrabarLote = False
    Else
        If sNumLote <> gsNumLote Then
            MsgBox "Lote fue creado con el número " & sNumLote
            gsNumLote = sNumLote
        End If
        
        cmdAceptar.Enabled = False
        GrabarLote = True
        gbCancelar = False
    End If
    
    Exit Function
    
ErrGrabarLote:
    Screen.MousePointer = vbNormal
    MsgBox Err.Description, vbCritical, App.Title
    GrabarLote = False
    Exit Function
    Resume
End Function

Private Sub cmdHlpSolicitante_Click(Index As Integer)
    HelpUsuario
End Sub

Private Sub Command1_Click()
    gbCancelar = True
    cmdAceptar.Enabled = False
    Unload Me
End Sub

Private Sub Form_Load()
    IniciaForm
    CargaFormulario
End Sub

Sub VerificaAceptar()
    cmdAceptar.Enabled = (Trim(Me.txtNomLote.Text) <> "") And (lstConsultas.ListCount > 0)
End Sub

Sub IniciaForm()
    Me.Left = (Screen.Width - Me.Width) \ 2
    Me.Top = (Screen.Height - Me.Height) \ 2
    
    Me.mnuOrdenarPor(0).Checked = True
    Me.mnuOrdenarPor(1).Checked = False
End Sub
Sub CargaFormulario()
    Dim sIcono  As String
    
    On Error GoTo ErrCargaFormulario
    
    Screen.MousePointer = vbHourglass
    
    lblNumLote = gsNumLote
    txtNomLote = gsNomLote
    txtNomSolicitante = gsNomSolicitante
    Me.chkIndAsignaSolicitante.Enabled = (gsNumLote = "")
    
    ' Carga todas las consultas y un indicador de las consultas que el lote ya tiene para marcarlas al momento de mostrarla
    If db_LeeTodasConsultasPorLote(gsNumLote, rsData) Then
        If Me.mnuOrdenarPor(0).Checked Then
            rsData.Sort = "nom_consulta"
        Else
            rsData.Sort = "num_consulta"
        End If
        
        While Not rsData.EOF
            lstConsultas.AddItem "" & rsData!nom_consulta & " (Id. " & rsData!num_consulta & ")"
            lstConsultas.ItemData(lstConsultas.ListCount - 1) = rsData!num_consulta
            
            If "" & rsData!ind_lote = "S" Then
                lstConsultas.Selected(lstConsultas.ListCount - 1) = True
            End If
            
            rsData.MoveNext
        Wend
    End If

    Screen.MousePointer = vbNormal

    Me.txtNomLote = gsNomLote
    Me.lblNumLote = gsNumLote
    Me.cmdAceptar.Enabled = False

    If lstConsultas.ListCount <= 0 Then
        MsgBox "No hay consultas para asignar a este lote.", vbCritical, App.Title
    End If

    Exit Sub

ErrCargaFormulario:
    Screen.MousePointer = vbNormal
    MsgBox Error, vbCritical, App.Title
    Exit Sub
End Sub

Private Sub Form_Resize()
    FormResize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim nResp   As Integer
    
    If cmdAceptar.Enabled Then
        nResp = MsgBox("Desea grabar los cambios realizados sobre esta consulta", vbYesNoCancel + vbDefaultButton1, App.Title)
        If nResp = vbYes Then
            If GrabarLote() Then
                Set rsData = Nothing
            Else
                Cancel = True
            End If
        ElseIf nResp = vbCancel Then
            Cancel = True
        Else
            gbCancelar = True
            Set rsData = Nothing
        End If
    End If
End Sub


Private Sub lstConsultas_Click()
    VerificaAceptar
End Sub

Private Sub lstConsultas_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button And vbRightButton Then
        PopupMenu Me.mnuOrdenar
    End If
End Sub


Private Sub mnuOrdenarPor_Click(Index As Integer)
    Select Case Index
    Case 0
        If Not Me.mnuOrdenarPor(Index).Checked Then
            Me.mnuOrdenarPor(0).Checked = True
            Me.mnuOrdenarPor(1).Checked = False
            Call OrdenaConsultas
        End If
    Case 1
        If Not Me.mnuOrdenarPor(Index).Checked Then
            Me.mnuOrdenarPor(1).Checked = True
            Me.mnuOrdenarPor(0).Checked = False
            Call OrdenaConsultas
        End If
    End Select
End Sub

Private Sub txtNomLote_Change()
    VerificaAceptar
End Sub


