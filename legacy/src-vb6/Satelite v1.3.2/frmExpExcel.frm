VERSION 5.00
Begin VB.Form frmExpExcel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exportar a excel"
   ClientHeight    =   2805
   ClientLeft      =   6195
   ClientTop       =   4890
   ClientWidth     =   7875
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   7875
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   5520
      TabIndex        =   0
      Top             =   2340
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   6720
      TabIndex        =   1
      Top             =   2340
      Width           =   1095
   End
   Begin VB.Frame fraDestino 
      Caption         =   "Destino"
      Height          =   1035
      Left            =   60
      TabIndex        =   11
      Top             =   0
      Width           =   7755
      Begin VB.OptionButton optTipo 
         Caption         =   "Planilla &e&xistente"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   3
         Top             =   660
         Width           =   1755
      End
      Begin VB.OptionButton optTipo 
         Caption         =   "&Nueva planilla"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   1755
      End
   End
   Begin VB.Frame fraArchivo 
      Caption         =   "Exportar a"
      Height          =   1215
      Left            =   60
      TabIndex        =   8
      Top             =   1080
      Width           =   7755
      Begin VB.OptionButton optHoja 
         Caption         =   "Existente"
         Height          =   195
         Index           =   1
         Left            =   2100
         TabIndex        =   6
         Top             =   780
         Width           =   975
      End
      Begin VB.OptionButton optHoja 
         Caption         =   "Nueva"
         Height          =   195
         Index           =   0
         Left            =   900
         TabIndex        =   5
         Top             =   780
         Width           =   1035
      End
      Begin VB.ComboBox cboHoja 
         Height          =   315
         Left            =   3120
         Style           =   2  'Dropdown List
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   720
         Width           =   4515
      End
      Begin VB.TextBox txtNomArchivo 
         Height          =   315
         Left            =   900
         TabIndex        =   4
         Top             =   300
         Width           =   6375
      End
      Begin VB.CommandButton cmdHelpCarpPer 
         Caption         =   "..."
         Height          =   315
         Left            =   7320
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   300
         Width           =   315
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Hoja : "
         Height          =   195
         Left            =   180
         TabIndex        =   12
         Top             =   780
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Planilla : "
         Height          =   195
         Left            =   180
         TabIndex        =   10
         Top             =   360
         Width           =   630
      End
   End
   Begin VB.Label lblMensaje 
      Caption         =   "Mensaje"
      Height          =   195
      Left            =   60
      TabIndex        =   13
      Top             =   2340
      Width           =   5055
   End
End
Attribute VB_Name = "frmExpExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public msNomConsulta    As String

Dim msHojaOld           As String
Sub ActivaControlesHoja(Index As Integer)
    Select Case Index
    Case 0
        cboHoja.Enabled = False
    Case 1
        cboHoja.Enabled = (cboHoja.ListCount > 0)
    End Select
End Sub

Sub CargaHojas()
    Dim Excel       As Object
    Dim sFile       As String
    Dim i           As Integer
    Dim nItemOld    As String
    
    On Error GoTo ErrCargaHojas
    
    Me.cboHoja.Clear
    nItemOld = -1
    
    Screen.MousePointer = vbHourglass
    lblMensaje = "Cargando hojas ..."

    sFile = Trim(Me.txtNomArchivo.Text)
    If Exist(sFile) Then
        Set Excel = CreateObject("Excel.Application")
        Excel.Workbooks.Open FileName:=sFile
        
        For i = 1 To Excel.Sheets.Count
            cboHoja.AddItem Excel.Sheets(i).Name
            If msHojaOld = Excel.Sheets(i).Name Then
                nItemOld = i - 1
            End If
        Next
        
        Excel.Workbooks.Close
        Set Excel = Nothing
    End If

    If cboHoja.ListCount > 0 Then
        If nItemOld >= 0 Then
            cboHoja.ListIndex = nItemOld
        Else
            cboHoja.ListIndex = 0
        End If
        cboHoja.Enabled = True
    End If
    
    lblMensaje = ""
    Screen.MousePointer = vbNormal
    Exit Sub
    
ErrCargaHojas:
    On Error Resume Next
    Excel.Workbooks.Close
    Set Excel = Nothing
    cboHoja.Enabled = False
    lblMensaje = ""
    
    Screen.MousePointer = vbNormal
    MsgBox "No fue posible cargar las hojas de esta planilla", vbCritical, App.Title
    Exit Sub
End Sub

Sub ActivaControles(Index As Integer)
    Select Case Index
    Case 0
        txtNomArchivo.Enabled = False
        cmdHelpCarpPer.Enabled = False
        cboHoja.Enabled = False
        
        fraArchivo.Visible = False
        cmdAceptar.Top = fraDestino.Top + fraDestino.Height + 30
        cmdCancelar.Top = cmdAceptar.Top
    Case 1
        txtNomArchivo.Enabled = True
        cmdHelpCarpPer.Enabled = True
        cboHoja.Enabled = (Exist(txtNomArchivo.Text))
    
        fraArchivo.Visible = True
        cmdAceptar.Top = fraArchivo.Top + fraArchivo.Height + 30
        cmdCancelar.Top = cmdAceptar.Top
    End Select

    Me.Height = cmdAceptar.Top + cmdAceptar.Height + 410
End Sub
Sub CargaInfoAnterior()
    Dim sFileOld    As String
    
    sFileOld = GetSetting(App.Title, "ExcelFileRecent", msNomConsulta, "")
    msHojaOld = GetSetting(App.Title, "ExcelSheetRecent", msNomConsulta, "")
    
    If sFileOld <> "" And Exist(sFileOld) Then
        txtNomArchivo = sFileOld
        If msHojaOld <> "" Then
            Call CargaHojas
        End If
        
        Me.optTipo(1).Value = True
        If Me.cboHoja.ListIndex >= 0 Then
            Me.optHoja(1).Value = True
        End If
    End If

End Sub

Sub VerAceptar()
    If Not txtNomArchivo.Enabled Then
        cmdAceptar.Enabled = True
    Else
        If optHoja(0).Value Then
            cmdAceptar.Enabled = (Trim(txtNomArchivo.Text) <> "" And Exist(txtNomArchivo.Text))
        Else
            cmdAceptar.Enabled = (Trim(txtNomArchivo.Text) <> "" And Exist(txtNomArchivo.Text) And cboHoja.ListCount > 0)
        End If
    End If
    
End Sub

Private Sub cmdAceptar_Click()
    If ExportarArchivo Then
        gbCancelar = False
        Unload Me
    End If
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub


Private Sub cmdHelpCarpPer_Click()
    HelpNombreArchivo
End Sub


Sub HelpNombreArchivo()
    On Error GoTo ErrHelpNombreArchivo
    
    frmMdiPadre.CommonDialog1.CancelError = True
    frmMdiPadre.CommonDialog1.FileName = "*.xls"
    frmMdiPadre.CommonDialog1.Filter = "Planillas Excel (*.xls)|*.*"
    'frmMdiPadre.CommonDialog1.DefaultExt = "*.xls"
    frmMdiPadre.CommonDialog1.DialogTitle = "Exportar consulta a planilla Excel"
    frmMdiPadre.CommonDialog1.Action = 2
    
    If frmMdiPadre.CommonDialog1.FileName <> "" Then
        Me.txtNomArchivo = frmMdiPadre.CommonDialog1.FileName
        If cmdAceptar.Enabled Then
            cmdAceptar.SetFocus
        End If
    End If
    
    Exit Sub
    
ErrHelpNombreArchivo:
    Exit Sub
End Sub


Private Sub Form_Load()
    IniciaForm
End Sub
Function ExportarArchivo() As Boolean
    Dim sFile   As String
    Dim nFila   As Long
    Dim sCons   As String
    Dim nPos    As Integer
    Dim nCheck  As Integer
        
    On Error GoTo ErrExportarArchivo
        
    sFile = Trim(Me.txtNomArchivo)
    
    If Me.optTipo(0).Value Then
        gsNomArchivoExportar = ""
        gsNomHojaExportar = ""
    
        SaveSetting App.Title, "ExcelFileRecent", msNomConsulta, ""
        SaveSetting App.Title, "ExcelSheetRecent", msNomConsulta, ""
    Else
        If sFile = "" Then
            MsgBox "Debe especificar un nombre de archivo excel", vbCritical, App.Title
            ExportarArchivo = False
            Exit Function
        End If
        
        If Not Exist(sFile) Then
            MsgBox "Archivo " & sFile & " no existe", vbCritical, App.Title
            ExportarArchivo = False
            Exit Function
        End If
        
        gsNomArchivoExportar = sFile
        If optHoja(0).Value = True Then
            gsNomHojaExportar = ""
        Else
            gsNomHojaExportar = cboHoja.Text
        End If
    
        SaveSetting App.Title, "ExcelFileRecent", msNomConsulta, gsNomArchivoExportar
        SaveSetting App.Title, "ExcelSheetRecent", msNomConsulta, gsNomHojaExportar
    End If
    
    ExportarArchivo = True
    Exit Function
    
ErrExportarArchivo:
    Screen.MousePointer = vbNormal
    MsgBox Error, vbCritical, App.Title
    
    ExportarArchivo = False
End Function


Sub IniciaForm()
    Me.HelpContextID = 2
    Me.Left = (Screen.Width - Me.Width) \ 2
    Me.Top = (Screen.Height - Me.Height) \ 2
    Me.cmdAceptar.Enabled = False
    lblMensaje = ""
    msHojaOld = ""
    
    optTipo(0).Value = True
    optHoja(0).Value = True
    gbCancelar = True
    
    CargaInfoAnterior
    '<V1.3.0>
    ' Si viene por ejecucion de lote, solo activa la opcion de Archivo existente
    optTipo(0).Enabled = True
    optTipo(1).Enabled = True
    
    If gbEjecutandoLote Then
        Me.Caption = gsNomConsulta
        optTipo(1).Value = True
        optTipo(0).Enabled = False
    End If
    '</V1.3.0>
End Sub

Private Sub optHoja_Click(Index As Integer)
    If Index = 1 Then
        Call CargaHojas
    End If
        
    Call ActivaControlesHoja(Index)
    VerAceptar
End Sub

Private Sub optTipo_Click(Index As Integer)
    If Me.optTipo(Index).Value = True Then
        Call ActivaControles(Index)
    End If
    
    VerAceptar
End Sub

Private Sub txtNomArchivo_Change()
    If Trim(txtNomArchivo.Text) <> "" And Exist(txtNomArchivo.Text) Then
        If optHoja(1).Value = True Then
            Call CargaHojas
        End If
    End If
    
    VerAceptar
End Sub


