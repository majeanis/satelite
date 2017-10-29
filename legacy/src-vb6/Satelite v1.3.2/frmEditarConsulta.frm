VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmEditarConsulta 
   Caption         =   "Edición de Consulta"
   ClientHeight    =   8640
   ClientLeft      =   4305
   ClientTop       =   1875
   ClientWidth     =   9165
   HelpContextID   =   30
   Icon            =   "frmEditarConsulta.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8640
   ScaleWidth      =   9165
   Begin VB.Frame Frame1 
      Caption         =   "Descripción de la consulta"
      Height          =   1515
      Left            =   60
      TabIndex        =   30
      Top             =   0
      Width           =   9075
      Begin VB.ComboBox cboCodNegocio 
         Height          =   315
         Left            =   5940
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1020
         Width           =   3015
      End
      Begin VB.ComboBox cboCodArea 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1020
         Width           =   3015
      End
      Begin VB.ComboBox cboBaseDatos 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   660
         Width           =   4935
      End
      Begin VB.TextBox txtNomConsulta 
         Height          =   315
         Left            =   1320
         TabIndex        =   0
         Top             =   300
         Width           =   7635
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Negocio : "
         Height          =   195
         Left            =   5040
         TabIndex        =   41
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Área :"
         Height          =   195
         Left            =   120
         TabIndex        =   40
         Top             =   1080
         Width           =   420
      End
      Begin VB.Label lblBaseDatos 
         AutoSize        =   -1  'True
         Caption         =   "Base de Datos :"
         Height          =   195
         Left            =   120
         TabIndex        =   32
         Top             =   720
         Width           =   1140
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nombre :"
         Height          =   195
         Left            =   120
         TabIndex        =   31
         Top             =   360
         Width           =   645
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6255
      Left            =   60
      TabIndex        =   27
      Top             =   1560
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   11033
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Query"
      TabPicture(0)   =   "frmEditarConsulta.frx":038A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "txtQuery"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdProbar"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdFormato"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Parametros"
      TabPicture(1)   =   "frmEditarConsulta.frx":03A6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdProbar2"
      Tab(1).Control(1)=   "cmdEditarParam"
      Tab(1).Control(2)=   "lstParametros"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Resultado y formatos"
      TabPicture(2)   =   "frmEditarConsulta.frx":03C2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "grdResultado"
      Tab(2).Control(1)=   "cmdFormato2"
      Tab(2).Control(2)=   "cmdProbar3"
      Tab(2).Control(3)=   "txtResultado"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "Horario"
      TabPicture(3)   =   "frmEditarConsulta.frx":03DE
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "cboHorFin(1)"
      Tab(3).Control(1)=   "cboHorIni(1)"
      Tab(3).Control(2)=   "cboHorFin(0)"
      Tab(3).Control(3)=   "cboHorIni(0)"
      Tab(3).Control(4)=   "optOpcionHorario(1)"
      Tab(3).Control(5)=   "optOpcionHorario(0)"
      Tab(3).Control(6)=   "Label3(1)"
      Tab(3).Control(7)=   "Label2(1)"
      Tab(3).Control(8)=   "Label3(0)"
      Tab(3).Control(9)=   "Label2(0)"
      Tab(3).ControlCount=   10
      TabCaption(4)   =   "Archivo Excel"
      TabPicture(4)   =   "frmEditarConsulta.frx":03FA
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "cboHoja"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "txtNomHoja"
      Tab(4).Control(2)=   "cmdHelpCarpPer"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "txtNomArchivo"
      Tab(4).Control(4)=   "optHoja(0)"
      Tab(4).Control(5)=   "optHoja(1)"
      Tab(4).Control(6)=   "Label4"
      Tab(4).Control(7)=   "Label2(2)"
      Tab(4).ControlCount=   8
      Begin FPSpreadADO.fpSpread grdResultado 
         Height          =   2055
         Left            =   -74880
         TabIndex        =   10
         Top             =   420
         Visible         =   0   'False
         Width           =   5595
         _Version        =   524288
         _ExtentX        =   2646
         _ExtentY        =   1323
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
         SpreadDesigner  =   "frmEditarConsulta.frx":0416
         UnitType        =   0
      End
      Begin VB.ComboBox cboHoja 
         Height          =   315
         Left            =   -71520
         Style           =   2  'Dropdown List
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   1320
         Visible         =   0   'False
         Width           =   4575
      End
      Begin VB.TextBox txtNomHoja 
         Height          =   315
         Left            =   -71520
         TabIndex        =   23
         Top             =   960
         Width           =   4215
      End
      Begin VB.CommandButton cmdHelpCarpPer 
         Caption         =   "..."
         Height          =   315
         Left            =   -67260
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   540
         Width           =   315
      End
      Begin VB.TextBox txtNomArchivo 
         Height          =   315
         Left            =   -74100
         TabIndex        =   20
         Top             =   540
         Width           =   6795
      End
      Begin VB.OptionButton optHoja 
         Caption         =   "Nueva"
         Height          =   195
         Index           =   0
         Left            =   -74100
         TabIndex        =   21
         Top             =   1020
         Width           =   1035
      End
      Begin VB.OptionButton optHoja 
         Caption         =   "Existente"
         Height          =   195
         Index           =   1
         Left            =   -72540
         TabIndex        =   22
         Top             =   1020
         Width           =   975
      End
      Begin VB.ComboBox cboHorFin 
         Height          =   315
         Index           =   1
         Left            =   -71460
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   1800
         Width           =   1215
      End
      Begin VB.ComboBox cboHorIni 
         Height          =   315
         Index           =   1
         Left            =   -73020
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   1800
         Width           =   1215
      End
      Begin VB.ComboBox cboHorFin 
         Height          =   315
         Index           =   0
         Left            =   -71460
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   1380
         Width           =   1215
      End
      Begin VB.ComboBox cboHorIni 
         Height          =   315
         Index           =   0
         Left            =   -73020
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   1380
         Width           =   1215
      End
      Begin VB.OptionButton optOpcionHorario 
         Caption         =   "En los siguientes horarios :"
         Height          =   195
         Index           =   1
         Left            =   -74580
         TabIndex        =   15
         Top             =   960
         Width           =   2835
      End
      Begin VB.OptionButton optOpcionHorario 
         Caption         =   "En todo horario"
         Height          =   195
         Index           =   0
         Left            =   -74580
         TabIndex        =   14
         Top             =   660
         Width           =   1875
      End
      Begin VB.CommandButton cmdFormato 
         Caption         =   "&Formato"
         Height          =   315
         Left            =   7860
         TabIndex        =   6
         Top             =   5820
         Width           =   1095
      End
      Begin VB.CommandButton cmdFormato2 
         Caption         =   "&Formato"
         Height          =   315
         Left            =   -67200
         TabIndex        =   13
         Top             =   5820
         Width           =   1095
      End
      Begin VB.CommandButton cmdProbar3 
         Caption         =   "&Probar"
         Height          =   315
         Left            =   -68340
         TabIndex        =   12
         Top             =   5820
         Width           =   1095
      End
      Begin VB.TextBox txtResultado 
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   -74880
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   2580
         Width           =   5595
      End
      Begin VB.CommandButton cmdProbar2 
         Caption         =   "&Probar"
         Height          =   315
         Left            =   -67200
         TabIndex        =   9
         Top             =   5460
         Width           =   1095
      End
      Begin VB.CommandButton cmdEditarParam 
         Caption         =   "&Editar"
         Height          =   315
         Left            =   -68340
         TabIndex        =   8
         Top             =   5460
         Width           =   1095
      End
      Begin VB.CommandButton cmdProbar 
         Caption         =   "&Probar"
         Height          =   315
         Left            =   6720
         TabIndex        =   5
         Top             =   5820
         Width           =   1095
      End
      Begin VB.TextBox txtQuery 
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3435
         Left            =   60
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   4
         Top             =   420
         Width           =   6135
      End
      Begin ComctlLib.ListView lstParametros 
         Height          =   3975
         Left            =   -74940
         TabIndex        =   7
         Top             =   420
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   7011
         View            =   3
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         _Version        =   327682
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Planilla : "
         Height          =   195
         Left            =   -74820
         TabIndex        =   39
         Top             =   600
         Width           =   630
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Hoja : "
         Height          =   195
         Index           =   2
         Left            =   -74820
         TabIndex        =   38
         Top             =   1020
         Width           =   465
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "o entre"
         Height          =   195
         Index           =   1
         Left            =   -73620
         TabIndex        =   36
         Top             =   1860
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "y"
         Height          =   195
         Index           =   1
         Left            =   -71640
         TabIndex        =   35
         Top             =   1860
         Width           =   75
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "entre"
         Height          =   195
         Index           =   0
         Left            =   -73500
         TabIndex        =   34
         Top             =   1440
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "y"
         Height          =   195
         Index           =   0
         Left            =   -71640
         TabIndex        =   33
         Top             =   1440
         Width           =   75
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   7980
      TabIndex        =   26
      Top             =   7860
      Width           =   1095
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   6780
      TabIndex        =   25
      Top             =   7860
      Width           =   1095
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   720
      TabIndex        =   28
      Top             =   7980
      Visible         =   0   'False
      Width           =   5475
      _ExtentX        =   9657
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   29
      Top             =   8325
      Width           =   9165
      _ExtentX        =   16166
      _ExtentY        =   556
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   5
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   5371
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "02:08 p.m."
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "23/04/2015"
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmEditarConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim msNumConsulta           As String
Dim msGlsQuery              As String
Dim mnNumBaseDatos          As Integer
Dim maRegParametros()       As rRegParametros
Dim maRegParametrosTest()   As rRegParametros
Dim msNomUsuarioLocal       As String

Dim mItem                   As ListItem
Dim bFormLoad               As Boolean

Public mrsData              As ADODB.Recordset
Public mrsFormatos          As ADODB.Recordset

Sub BuscaTabValor(sCodTabla As String, nNumValor As Long)
    '<V1.3.1>
    Dim i   As Long
    
    Select Case sCodTabla
    Case "AREA"
        For i = 1 To cboCodArea.ListCount
            If cboCodArea.ItemData(i - 1) = nNumValor Then
                cboCodArea.ListIndex = i - 1
            End If
        Next i
        
    Case "NEGOCIO"
        For i = 1 To cboCodNegocio.ListCount
            If cboCodNegocio.ItemData(i - 1) = nNumValor Then
                cboCodNegocio.ListIndex = i - 1
            End If
        Next i
    End Select
    '</V1.3.1>
End Sub

Sub CargaCodigos()
    '<V1.3.1>
    Dim rsData  As ADODB.Recordset
    
    ' Carga todas las tablas de valores
    If db_LeeTabValores("", rsData) Then
        While Not rsData.EOF
            Select Case "" & rsData!cod_tabla
            Case "AREA"
                cboCodArea.AddItem "" & rsData!gls_valor
                cboCodArea.ItemData(cboCodArea.ListCount - 1) = rsData!num_registro

            Case "NEGOCIO"
                cboCodNegocio.AddItem "" & rsData!gls_valor
                cboCodNegocio.ItemData(cboCodNegocio.ListCount - 1) = rsData!num_registro
            
            End Select
            rsData.MoveNext
        Wend
    End If
    '</V1.3.1>
End Sub

Sub CargaHorarios(sGlsHorario As String)
    Dim sGlsHoraIni1    As String
    Dim sGlsHoraFin1    As String
    Dim sGlsHoraIni2    As String
    Dim sGlsHoraFin2    As String
    
    Dim nHora           As Integer
    Dim nMinuto         As Integer
    
    On Error GoTo Err_CargaHorarios
    
    sGlsHoraIni1 = Trim(Mid(sGlsHorario, 1, 10))
    sGlsHoraFin1 = Trim(Mid(sGlsHorario, 11, 10))
    sGlsHoraIni2 = Trim(Mid(sGlsHorario, 21, 10))
    sGlsHoraFin2 = Trim(Mid(sGlsHorario, 31, 10))
    
    If sGlsHoraIni1 <> "" And sGlsHoraFin1 <> "" Then
        nHora = Val(Left(sGlsHoraIni1, 2))
        nMinuto = Val(Mid(sGlsHoraIni1, 4, 2))
        cboHorIni(0).ListIndex = (nHora * 4 + 1) + (nMinuto / 15)
    
        nHora = Val(Left(sGlsHoraFin1, 2))
        nMinuto = Val(Mid(sGlsHoraFin1, 4, 2))
        cboHorFin(0).ListIndex = (nHora * 4 + 1) + (nMinuto / 15)
    End If
    
    If sGlsHoraIni2 <> "" And sGlsHoraFin2 <> "" Then
        nHora = Val(Left(sGlsHoraIni2, 2))
        nMinuto = Val(Mid(sGlsHoraIni2, 4, 2))
        cboHorIni(1).ListIndex = (nHora * 4 + 1) + (nMinuto / 15)
    
        nHora = Val(Left(sGlsHoraFin2, 2))
        nMinuto = Val(Mid(sGlsHoraFin2, 4, 2))
        cboHorFin(1).ListIndex = (nHora * 4 + 1) + (nMinuto / 15)
    End If
    
    If cboHorIni(0).ListIndex >= 1 Or cboHorIni(1).ListIndex >= 1 Then
        Me.optOpcionHorario(1).Value = True
    Else
        Me.optOpcionHorario(0).Value = True
    End If
    
    cboHorIni(0).Tag = ""
    cboHorFin(0).Tag = ""
    cboHorIni(1).Tag = ""
    cboHorFin(1).Tag = ""
    
    Exit Sub

Err_CargaHorarios:
    Me.optOpcionHorario(0).Value = True
    Screen.MousePointer = vbNormal
    MsgBox Error, vbInformation, App.Title
    Exit Sub
End Sub

Sub CargaInformacionArchivoExcel(ByVal sGlsArchivoSalida As String, ByVal sNomHojaSalida As String)
    '<V1.3.0>
    txtNomArchivo = sGlsArchivoSalida
    If sNomHojaSalida = "<nueva>" Or sNomHojaSalida = "" Then
        Me.optHoja(0).Value = True
    Else
        Me.optHoja(1).Value = True
        Me.txtNomHoja.Text = sNomHojaSalida
    End If
    
    If Not Exist(txtNomArchivo.Text) Then
        txtNomHoja.Visible = True
        cboHoja.Visible = False
        cboHoja.Clear
    Else
        If Me.optHoja(0).Value Then
            Me.txtNomHoja.Visible = False
            Me.cboHoja.Visible = True
        Else
            If CargaHojas(sNomHojaSalida) Then
                Me.txtNomHoja.Visible = False
                Me.cboHoja.Visible = True
            Else
                txtNomHoja.Visible = True
                cboHoja.Visible = False
                cboHoja.Clear
            End If
        End If
    End If
    '/<V1.3.0>
End Sub

Sub CargaParametrosLocal(ByVal sGlsQuery As String)
    Dim nPosIni         As Integer
    Dim nPosFin         As Integer
    Dim sParametro      As String
    Dim nIndice         As Integer
    Dim bEncontrado     As Boolean
    Dim nX              As Integer
    Dim sTipoDato       As String
    Dim sTipoAyuda      As String
    Dim sGlsAyuda       As String
    Dim sIndOpcional    As String
    Dim sGlsParametro   As String
    Dim nTotParametros  As Integer
    
    ReDim maRegParametrosTest(0) As rRegParametros
    nTotParametros = 0
    
    ' Busca todos los parámetros que estén actualmente en el query (@parametro@)
    nPosIni = InStr(1, sGlsQuery, "@")
    nIndice = nPosIni
    Do While nPosIni > 0
        nPosFin = InStr(nPosIni + 1, sGlsQuery, "@")
        If nPosIni < nPosFin Then
            sParametro = LCase(Mid(sGlsQuery, nPosIni + 1, nPosFin - nPosIni - 1))
            bEncontrado = False
            If nTotParametros > 0 Then
                For nX = 1 To nTotParametros
                    If maRegParametrosTest(nX).Nombre = sParametro Then
                        bEncontrado = True
                    End If
                Next
            End If
            
            If Not bEncontrado Then
                ' Busca parametro en Panel
                sGlsParametro = FormatoTitulo(sParametro)
                sTipoDato = "Texto"
                sGlsAyuda = ""
                For nX = 1 To lstParametros.ListItems.Count
                    If LCase(lstParametros.ListItems(nX)) = sParametro Then
                        sGlsParametro = lstParametros.ListItems(nX).SubItems(1)
                        sTipoDato = lstParametros.ListItems(nX).SubItems(2)
                        sTipoAyuda = lstParametros.ListItems(nX).SubItems(3)
                        sIndOpcional = lstParametros.ListItems(nX).SubItems(5)
                        sGlsAyuda = lstParametros.ListItems(nX).SubItems(6)
                    End If
                Next nX

                nTotParametros = nTotParametros + 1
                ReDim Preserve maRegParametrosTest(nTotParametros) As rRegParametros
                maRegParametrosTest(nTotParametros).Nombre = sParametro
                maRegParametrosTest(nTotParametros).Descripcion = sGlsParametro
                maRegParametrosTest(nTotParametros).Opcional = IIf(sIndOpcional = "S", True, False)
                maRegParametrosTest(nTotParametros).Tipo = sTipoDato
                maRegParametrosTest(nTotParametros).TipoAyuda = sTipoAyuda
                maRegParametrosTest(nTotParametros).Ayuda = sGlsAyuda
            End If
        End If
        
        nIndice = nPosFin + 1
        nPosIni = InStr(nIndice, sGlsQuery, "@")
    Loop
End Sub

Sub FormatoColumna()
    Dim nCol        As Integer
    
    ' Primero verifica que la consulta se haya ejecutado
    If Not grdResultado.Visible Then
        ProbarConsulta
    End If
    
    ' Si no hay registros, no se puede dar formato
    If Not grdResultado.Visible Then
        MsgBox "Consulta debe tener registros para poder dar formato a las columnas", vbInformation, App.Title
        Exit Sub
    End If

    ' Valida que todos los campos tengan un alias
    For nCol = 1 To mrsData.Fields.Count
        If mrsData(nCol - 1).Name = "" Then
            MsgBox "Todas las columnas de su consulta deben tener un nombre o alias", vbInformation, App.Title
            Exit Sub
        End If
    Next nCol
    
    ' Llama al formulario de formatos
    nCol = Me.grdResultado.Col - 1
    If nCol >= 0 Then
        frmDefFormato.msNumConsulta = ""
        Set frmDefFormato.mrsFormatos = mrsFormatos
        frmDefFormato.Show vbModal
        
        If Not gbCancelar Then
            Call CargarResultadoEnGrilla(mrsData, mrsFormatos, grdResultado, txtResultado, ProgressBar1)
            
            VerificaAceptar
        End If
    End If
End Sub

Function ExisteParametro(sNombre As String) As Boolean
    Dim bExiste As Boolean
    Dim nItem   As Integer
    
    bExiste = False
    For nItem = 1 To Me.lstParametros.ListItems.Count
        Set mItem = lstParametros.ListItems.Item(nItem)
        If LCase(mItem) = LCase(sNombre) Then
            bExiste = True
            Exit For
        End If
    Next nItem

    ExisteParametro = bExiste
End Function

Sub IniciaCombos()
    cboHorIni(0).AddItem "(ninguno) "
    cboHorIni(0).AddItem "00:00 a.m."
    cboHorIni(0).AddItem "00:15 a.m."
    cboHorIni(0).AddItem "00:30 a.m."
    cboHorIni(0).AddItem "00:45 a.m."
    cboHorIni(0).AddItem "01:00 a.m."
    cboHorIni(0).AddItem "01:15 a.m."
    cboHorIni(0).AddItem "01:30 a.m."
    cboHorIni(0).AddItem "01:45 a.m."
    cboHorIni(0).AddItem "02:00 a.m."
    cboHorIni(0).AddItem "02:15 a.m."
    cboHorIni(0).AddItem "02:30 a.m."
    cboHorIni(0).AddItem "02:45 a.m."
    cboHorIni(0).AddItem "03:00 a.m."
    cboHorIni(0).AddItem "03:15 a.m."
    cboHorIni(0).AddItem "03:30 a.m."
    cboHorIni(0).AddItem "03:45 a.m."
    cboHorIni(0).AddItem "04:00 a.m."
    cboHorIni(0).AddItem "04:15 a.m."
    cboHorIni(0).AddItem "04:30 a.m."
    cboHorIni(0).AddItem "04:45 a.m."
    cboHorIni(0).AddItem "05:00 a.m."
    cboHorIni(0).AddItem "05:15 a.m."
    cboHorIni(0).AddItem "05:30 a.m."
    cboHorIni(0).AddItem "05:45 a.m."
    cboHorIni(0).AddItem "06:00 a.m."
    cboHorIni(0).AddItem "06:15 a.m."
    cboHorIni(0).AddItem "06:30 a.m."
    cboHorIni(0).AddItem "06:45 a.m."
    cboHorIni(0).AddItem "07:00 a.m."
    cboHorIni(0).AddItem "07:15 a.m."
    cboHorIni(0).AddItem "07:30 a.m."
    cboHorIni(0).AddItem "07:45 a.m."
    cboHorIni(0).AddItem "08:00 a.m."
    cboHorIni(0).AddItem "08:15 a.m."
    cboHorIni(0).AddItem "08:30 a.m."
    cboHorIni(0).AddItem "08:45 a.m."
    cboHorIni(0).AddItem "09:00 a.m."
    cboHorIni(0).AddItem "09:15 a.m."
    cboHorIni(0).AddItem "09:30 a.m."
    cboHorIni(0).AddItem "09:45 a.m."
    cboHorIni(0).AddItem "10:00 a.m."
    cboHorIni(0).AddItem "10:15 a.m."
    cboHorIni(0).AddItem "10:30 a.m."
    cboHorIni(0).AddItem "10:45 a.m."
    cboHorIni(0).AddItem "11:00 a.m."
    cboHorIni(0).AddItem "11:15 a.m."
    cboHorIni(0).AddItem "11:30 a.m."
    cboHorIni(0).AddItem "11:45 a.m."
    cboHorIni(0).AddItem "12:00 p.m."
    cboHorIni(0).AddItem "12:15 p.m."
    cboHorIni(0).AddItem "12:30 p.m."
    cboHorIni(0).AddItem "12:45 p.m."
    cboHorIni(0).AddItem "13:00 p.m."
    cboHorIni(0).AddItem "13:15 p.m."
    cboHorIni(0).AddItem "13:30 p.m."
    cboHorIni(0).AddItem "13:45 p.m."
    cboHorIni(0).AddItem "14:00 p.m."
    cboHorIni(0).AddItem "14:15 p.m."
    cboHorIni(0).AddItem "14:30 p.m."
    cboHorIni(0).AddItem "14:45 p.m."
    cboHorIni(0).AddItem "15:00 p.m."
    cboHorIni(0).AddItem "15:15 p.m."
    cboHorIni(0).AddItem "15:30 p.m."
    cboHorIni(0).AddItem "15:45 p.m."
    cboHorIni(0).AddItem "16:00 p.m."
    cboHorIni(0).AddItem "16:15 p.m."
    cboHorIni(0).AddItem "16:30 p.m."
    cboHorIni(0).AddItem "16:45 p.m."
    cboHorIni(0).AddItem "17:00 p.m."
    cboHorIni(0).AddItem "17:15 p.m."
    cboHorIni(0).AddItem "17:30 p.m."
    cboHorIni(0).AddItem "17:45 p.m."
    cboHorIni(0).AddItem "18:00 p.m."
    cboHorIni(0).AddItem "18:15 p.m."
    cboHorIni(0).AddItem "18:30 p.m."
    cboHorIni(0).AddItem "18:45 p.m."
    cboHorIni(0).AddItem "19:00 p.m."
    cboHorIni(0).AddItem "19:15 p.m."
    cboHorIni(0).AddItem "19:30 p.m."
    cboHorIni(0).AddItem "19:45 p.m."
    cboHorIni(0).AddItem "20:00 p.m."
    cboHorIni(0).AddItem "20:15 p.m."
    cboHorIni(0).AddItem "20:30 p.m."
    cboHorIni(0).AddItem "20:45 p.m."
    cboHorIni(0).AddItem "21:00 p.m."
    cboHorIni(0).AddItem "21:15 p.m."
    cboHorIni(0).AddItem "21:30 p.m."
    cboHorIni(0).AddItem "21:45 p.m."
    cboHorIni(0).AddItem "22:00 p.m."
    cboHorIni(0).AddItem "22:15 p.m."
    cboHorIni(0).AddItem "22:30 p.m."
    cboHorIni(0).AddItem "22:45 p.m."
    cboHorIni(0).AddItem "23:00 p.m."
    cboHorIni(0).AddItem "23:15 p.m."
    cboHorIni(0).AddItem "23:30 p.m."
    cboHorIni(0).AddItem "23:45 p.m."
    cboHorIni(0).AddItem "24:00 p.m."

    cboHorIni(1).AddItem "(ninguno) "
    cboHorIni(1).AddItem "00:00 a.m."
    cboHorIni(1).AddItem "00:15 a.m."
    cboHorIni(1).AddItem "00:30 a.m."
    cboHorIni(1).AddItem "00:45 a.m."
    cboHorIni(1).AddItem "01:00 a.m."
    cboHorIni(1).AddItem "01:15 a.m."
    cboHorIni(1).AddItem "01:30 a.m."
    cboHorIni(1).AddItem "01:45 a.m."
    cboHorIni(1).AddItem "02:00 a.m."
    cboHorIni(1).AddItem "02:15 a.m."
    cboHorIni(1).AddItem "02:30 a.m."
    cboHorIni(1).AddItem "02:45 a.m."
    cboHorIni(1).AddItem "03:00 a.m."
    cboHorIni(1).AddItem "03:15 a.m."
    cboHorIni(1).AddItem "03:30 a.m."
    cboHorIni(1).AddItem "03:45 a.m."
    cboHorIni(1).AddItem "04:00 a.m."
    cboHorIni(1).AddItem "04:15 a.m."
    cboHorIni(1).AddItem "04:30 a.m."
    cboHorIni(1).AddItem "04:45 a.m."
    cboHorIni(1).AddItem "05:00 a.m."
    cboHorIni(1).AddItem "05:15 a.m."
    cboHorIni(1).AddItem "05:30 a.m."
    cboHorIni(1).AddItem "05:45 a.m."
    cboHorIni(1).AddItem "06:00 a.m."
    cboHorIni(1).AddItem "06:15 a.m."
    cboHorIni(1).AddItem "06:30 a.m."
    cboHorIni(1).AddItem "06:45 a.m."
    cboHorIni(1).AddItem "07:00 a.m."
    cboHorIni(1).AddItem "07:15 a.m."
    cboHorIni(1).AddItem "07:30 a.m."
    cboHorIni(1).AddItem "07:45 a.m."
    cboHorIni(1).AddItem "08:00 a.m."
    cboHorIni(1).AddItem "08:15 a.m."
    cboHorIni(1).AddItem "08:30 a.m."
    cboHorIni(1).AddItem "08:45 a.m."
    cboHorIni(1).AddItem "09:00 a.m."
    cboHorIni(1).AddItem "09:15 a.m."
    cboHorIni(1).AddItem "09:30 a.m."
    cboHorIni(1).AddItem "09:45 a.m."
    cboHorIni(1).AddItem "10:00 a.m."
    cboHorIni(1).AddItem "10:15 a.m."
    cboHorIni(1).AddItem "10:30 a.m."
    cboHorIni(1).AddItem "10:45 a.m."
    cboHorIni(1).AddItem "11:00 a.m."
    cboHorIni(1).AddItem "11:15 a.m."
    cboHorIni(1).AddItem "11:30 a.m."
    cboHorIni(1).AddItem "11:45 a.m."
    cboHorIni(1).AddItem "12:00 p.m."
    cboHorIni(1).AddItem "12:15 p.m."
    cboHorIni(1).AddItem "12:30 p.m."
    cboHorIni(1).AddItem "12:45 p.m."
    cboHorIni(1).AddItem "13:00 p.m."
    cboHorIni(1).AddItem "13:15 p.m."
    cboHorIni(1).AddItem "13:30 p.m."
    cboHorIni(1).AddItem "13:45 p.m."
    cboHorIni(1).AddItem "14:00 p.m."
    cboHorIni(1).AddItem "14:15 p.m."
    cboHorIni(1).AddItem "14:30 p.m."
    cboHorIni(1).AddItem "14:45 p.m."
    cboHorIni(1).AddItem "15:00 p.m."
    cboHorIni(1).AddItem "15:15 p.m."
    cboHorIni(1).AddItem "15:30 p.m."
    cboHorIni(1).AddItem "15:45 p.m."
    cboHorIni(1).AddItem "16:00 p.m."
    cboHorIni(1).AddItem "16:15 p.m."
    cboHorIni(1).AddItem "16:30 p.m."
    cboHorIni(1).AddItem "16:45 p.m."
    cboHorIni(1).AddItem "17:00 p.m."
    cboHorIni(1).AddItem "17:15 p.m."
    cboHorIni(1).AddItem "17:30 p.m."
    cboHorIni(1).AddItem "17:45 p.m."
    cboHorIni(1).AddItem "18:00 p.m."
    cboHorIni(1).AddItem "18:15 p.m."
    cboHorIni(1).AddItem "18:30 p.m."
    cboHorIni(1).AddItem "18:45 p.m."
    cboHorIni(1).AddItem "19:00 p.m."
    cboHorIni(1).AddItem "19:15 p.m."
    cboHorIni(1).AddItem "19:30 p.m."
    cboHorIni(1).AddItem "19:45 p.m."
    cboHorIni(1).AddItem "20:00 p.m."
    cboHorIni(1).AddItem "20:15 p.m."
    cboHorIni(1).AddItem "20:30 p.m."
    cboHorIni(1).AddItem "20:45 p.m."
    cboHorIni(1).AddItem "21:00 p.m."
    cboHorIni(1).AddItem "21:15 p.m."
    cboHorIni(1).AddItem "21:30 p.m."
    cboHorIni(1).AddItem "21:45 p.m."
    cboHorIni(1).AddItem "22:00 p.m."
    cboHorIni(1).AddItem "22:15 p.m."
    cboHorIni(1).AddItem "22:30 p.m."
    cboHorIni(1).AddItem "22:45 p.m."
    cboHorIni(1).AddItem "23:00 p.m."
    cboHorIni(1).AddItem "23:15 p.m."
    cboHorIni(1).AddItem "23:30 p.m."
    cboHorIni(1).AddItem "23:45 p.m."
    cboHorIni(1).AddItem "24:00 p.m."

    cboHorFin(0).AddItem "(ninguno) "
    cboHorFin(0).AddItem "00:00 a.m."
    cboHorFin(0).AddItem "00:15 a.m."
    cboHorFin(0).AddItem "00:30 a.m."
    cboHorFin(0).AddItem "00:45 a.m."
    cboHorFin(0).AddItem "01:00 a.m."
    cboHorFin(0).AddItem "01:15 a.m."
    cboHorFin(0).AddItem "01:30 a.m."
    cboHorFin(0).AddItem "01:45 a.m."
    cboHorFin(0).AddItem "02:00 a.m."
    cboHorFin(0).AddItem "02:15 a.m."
    cboHorFin(0).AddItem "02:30 a.m."
    cboHorFin(0).AddItem "02:45 a.m."
    cboHorFin(0).AddItem "03:00 a.m."
    cboHorFin(0).AddItem "03:15 a.m."
    cboHorFin(0).AddItem "03:30 a.m."
    cboHorFin(0).AddItem "03:45 a.m."
    cboHorFin(0).AddItem "04:00 a.m."
    cboHorFin(0).AddItem "04:15 a.m."
    cboHorFin(0).AddItem "04:30 a.m."
    cboHorFin(0).AddItem "04:45 a.m."
    cboHorFin(0).AddItem "05:00 a.m."
    cboHorFin(0).AddItem "05:15 a.m."
    cboHorFin(0).AddItem "05:30 a.m."
    cboHorFin(0).AddItem "05:45 a.m."
    cboHorFin(0).AddItem "06:00 a.m."
    cboHorFin(0).AddItem "06:15 a.m."
    cboHorFin(0).AddItem "06:30 a.m."
    cboHorFin(0).AddItem "06:45 a.m."
    cboHorFin(0).AddItem "07:00 a.m."
    cboHorFin(0).AddItem "07:15 a.m."
    cboHorFin(0).AddItem "07:30 a.m."
    cboHorFin(0).AddItem "07:45 a.m."
    cboHorFin(0).AddItem "08:00 a.m."
    cboHorFin(0).AddItem "08:15 a.m."
    cboHorFin(0).AddItem "08:30 a.m."
    cboHorFin(0).AddItem "08:45 a.m."
    cboHorFin(0).AddItem "09:00 a.m."
    cboHorFin(0).AddItem "09:15 a.m."
    cboHorFin(0).AddItem "09:30 a.m."
    cboHorFin(0).AddItem "09:45 a.m."
    cboHorFin(0).AddItem "10:00 a.m."
    cboHorFin(0).AddItem "10:15 a.m."
    cboHorFin(0).AddItem "10:30 a.m."
    cboHorFin(0).AddItem "10:45 a.m."
    cboHorFin(0).AddItem "11:00 a.m."
    cboHorFin(0).AddItem "11:15 a.m."
    cboHorFin(0).AddItem "11:30 a.m."
    cboHorFin(0).AddItem "11:45 a.m."
    cboHorFin(0).AddItem "12:00 p.m."
    cboHorFin(0).AddItem "12:15 p.m."
    cboHorFin(0).AddItem "12:30 p.m."
    cboHorFin(0).AddItem "12:45 p.m."
    cboHorFin(0).AddItem "13:00 p.m."
    cboHorFin(0).AddItem "13:15 p.m."
    cboHorFin(0).AddItem "13:30 p.m."
    cboHorFin(0).AddItem "13:45 p.m."
    cboHorFin(0).AddItem "14:00 p.m."
    cboHorFin(0).AddItem "14:15 p.m."
    cboHorFin(0).AddItem "14:30 p.m."
    cboHorFin(0).AddItem "14:45 p.m."
    cboHorFin(0).AddItem "15:00 p.m."
    cboHorFin(0).AddItem "15:15 p.m."
    cboHorFin(0).AddItem "15:30 p.m."
    cboHorFin(0).AddItem "15:45 p.m."
    cboHorFin(0).AddItem "16:00 p.m."
    cboHorFin(0).AddItem "16:15 p.m."
    cboHorFin(0).AddItem "16:30 p.m."
    cboHorFin(0).AddItem "16:45 p.m."
    cboHorFin(0).AddItem "17:00 p.m."
    cboHorFin(0).AddItem "17:15 p.m."
    cboHorFin(0).AddItem "17:30 p.m."
    cboHorFin(0).AddItem "17:45 p.m."
    cboHorFin(0).AddItem "18:00 p.m."
    cboHorFin(0).AddItem "18:15 p.m."
    cboHorFin(0).AddItem "18:30 p.m."
    cboHorFin(0).AddItem "18:45 p.m."
    cboHorFin(0).AddItem "19:00 p.m."
    cboHorFin(0).AddItem "19:15 p.m."
    cboHorFin(0).AddItem "19:30 p.m."
    cboHorFin(0).AddItem "19:45 p.m."
    cboHorFin(0).AddItem "20:00 p.m."
    cboHorFin(0).AddItem "20:15 p.m."
    cboHorFin(0).AddItem "20:30 p.m."
    cboHorFin(0).AddItem "20:45 p.m."
    cboHorFin(0).AddItem "21:00 p.m."
    cboHorFin(0).AddItem "21:15 p.m."
    cboHorFin(0).AddItem "21:30 p.m."
    cboHorFin(0).AddItem "21:45 p.m."
    cboHorFin(0).AddItem "22:00 p.m."
    cboHorFin(0).AddItem "22:15 p.m."
    cboHorFin(0).AddItem "22:30 p.m."
    cboHorFin(0).AddItem "22:45 p.m."
    cboHorFin(0).AddItem "23:00 p.m."
    cboHorFin(0).AddItem "23:15 p.m."
    cboHorFin(0).AddItem "23:30 p.m."
    cboHorFin(0).AddItem "23:45 p.m."
    cboHorFin(0).AddItem "24:00 p.m."

    cboHorFin(1).AddItem "(ninguno) "
    cboHorFin(1).AddItem "00:00 a.m."
    cboHorFin(1).AddItem "00:15 a.m."
    cboHorFin(1).AddItem "00:30 a.m."
    cboHorFin(1).AddItem "00:45 a.m."
    cboHorFin(1).AddItem "01:00 a.m."
    cboHorFin(1).AddItem "01:15 a.m."
    cboHorFin(1).AddItem "01:30 a.m."
    cboHorFin(1).AddItem "01:45 a.m."
    cboHorFin(1).AddItem "02:00 a.m."
    cboHorFin(1).AddItem "02:15 a.m."
    cboHorFin(1).AddItem "02:30 a.m."
    cboHorFin(1).AddItem "02:45 a.m."
    cboHorFin(1).AddItem "03:00 a.m."
    cboHorFin(1).AddItem "03:15 a.m."
    cboHorFin(1).AddItem "03:30 a.m."
    cboHorFin(1).AddItem "03:45 a.m."
    cboHorFin(1).AddItem "04:00 a.m."
    cboHorFin(1).AddItem "04:15 a.m."
    cboHorFin(1).AddItem "04:30 a.m."
    cboHorFin(1).AddItem "04:45 a.m."
    cboHorFin(1).AddItem "05:00 a.m."
    cboHorFin(1).AddItem "05:15 a.m."
    cboHorFin(1).AddItem "05:30 a.m."
    cboHorFin(1).AddItem "05:45 a.m."
    cboHorFin(1).AddItem "06:00 a.m."
    cboHorFin(1).AddItem "06:15 a.m."
    cboHorFin(1).AddItem "06:30 a.m."
    cboHorFin(1).AddItem "06:45 a.m."
    cboHorFin(1).AddItem "07:00 a.m."
    cboHorFin(1).AddItem "07:15 a.m."
    cboHorFin(1).AddItem "07:30 a.m."
    cboHorFin(1).AddItem "07:45 a.m."
    cboHorFin(1).AddItem "08:00 a.m."
    cboHorFin(1).AddItem "08:15 a.m."
    cboHorFin(1).AddItem "08:30 a.m."
    cboHorFin(1).AddItem "08:45 a.m."
    cboHorFin(1).AddItem "09:00 a.m."
    cboHorFin(1).AddItem "09:15 a.m."
    cboHorFin(1).AddItem "09:30 a.m."
    cboHorFin(1).AddItem "09:45 a.m."
    cboHorFin(1).AddItem "10:00 a.m."
    cboHorFin(1).AddItem "10:15 a.m."
    cboHorFin(1).AddItem "10:30 a.m."
    cboHorFin(1).AddItem "10:45 a.m."
    cboHorFin(1).AddItem "11:00 a.m."
    cboHorFin(1).AddItem "11:15 a.m."
    cboHorFin(1).AddItem "11:30 a.m."
    cboHorFin(1).AddItem "11:45 a.m."
    cboHorFin(1).AddItem "12:00 p.m."
    cboHorFin(1).AddItem "12:15 p.m."
    cboHorFin(1).AddItem "12:30 p.m."
    cboHorFin(1).AddItem "12:45 p.m."
    cboHorFin(1).AddItem "13:00 p.m."
    cboHorFin(1).AddItem "13:15 p.m."
    cboHorFin(1).AddItem "13:30 p.m."
    cboHorFin(1).AddItem "13:45 p.m."
    cboHorFin(1).AddItem "14:00 p.m."
    cboHorFin(1).AddItem "14:15 p.m."
    cboHorFin(1).AddItem "14:30 p.m."
    cboHorFin(1).AddItem "14:45 p.m."
    cboHorFin(1).AddItem "15:00 p.m."
    cboHorFin(1).AddItem "15:15 p.m."
    cboHorFin(1).AddItem "15:30 p.m."
    cboHorFin(1).AddItem "15:45 p.m."
    cboHorFin(1).AddItem "16:00 p.m."
    cboHorFin(1).AddItem "16:15 p.m."
    cboHorFin(1).AddItem "16:30 p.m."
    cboHorFin(1).AddItem "16:45 p.m."
    cboHorFin(1).AddItem "17:00 p.m."
    cboHorFin(1).AddItem "17:15 p.m."
    cboHorFin(1).AddItem "17:30 p.m."
    cboHorFin(1).AddItem "17:45 p.m."
    cboHorFin(1).AddItem "18:00 p.m."
    cboHorFin(1).AddItem "18:15 p.m."
    cboHorFin(1).AddItem "18:30 p.m."
    cboHorFin(1).AddItem "18:45 p.m."
    cboHorFin(1).AddItem "19:00 p.m."
    cboHorFin(1).AddItem "19:15 p.m."
    cboHorFin(1).AddItem "19:30 p.m."
    cboHorFin(1).AddItem "19:45 p.m."
    cboHorFin(1).AddItem "20:00 p.m."
    cboHorFin(1).AddItem "20:15 p.m."
    cboHorFin(1).AddItem "20:30 p.m."
    cboHorFin(1).AddItem "20:45 p.m."
    cboHorFin(1).AddItem "21:00 p.m."
    cboHorFin(1).AddItem "21:15 p.m."
    cboHorFin(1).AddItem "21:30 p.m."
    cboHorFin(1).AddItem "21:45 p.m."
    cboHorFin(1).AddItem "22:00 p.m."
    cboHorFin(1).AddItem "22:15 p.m."
    cboHorFin(1).AddItem "22:30 p.m."
    cboHorFin(1).AddItem "22:45 p.m."
    cboHorFin(1).AddItem "23:00 p.m."
    cboHorFin(1).AddItem "23:15 p.m."
    cboHorFin(1).AddItem "23:30 p.m."
    cboHorFin(1).AddItem "23:45 p.m."
    cboHorFin(1).AddItem "24:00 p.m."
End Sub

Sub IniciaFormatos()
    Screen.MousePointer = vbHourglass
    
    ' Abre base datos
    OpenMyDataBase
    
    Call db_LeeFormatos("-1", mrsFormatos)
        
    ' Cierra base datos
    CloseMyDataBase

    Screen.MousePointer = vbNormal
End Sub

Sub ParametroEditar()
    Dim sGlsAyuda   As String
    
    sGlsAyuda = mItem.SubItems(6)
    sGlsAyuda = Replace(sGlsAyuda, Chr(13), "")
    sGlsAyuda = Replace(sGlsAyuda, Chr(10), Chr(13) & Chr(10))
    
    frmDefParametros.msNomParametro = mItem.Text
    frmDefParametros.msGlsParametro = mItem.SubItems(1)
    frmDefParametros.msTipoDato = mItem.SubItems(2)
    frmDefParametros.msTipoAyuda = mItem.SubItems(3)
    frmDefParametros.msIndOpcional = mItem.SubItems(5)
    frmDefParametros.msGlsAyuda = sGlsAyuda
    frmDefParametros.Show vbModal
    
    If gsTipoDato <> "" Then
        mItem.SubItems(1) = gsGlsParametro
        mItem.SubItems(2) = gsTipoDato
        mItem.SubItems(3) = gsTipoAyuda
        mItem.SubItems(4) = fsConvierteTextoToLinea(gsGlsAyuda)
        mItem.SubItems(5) = gsIndOpcional
        mItem.SubItems(6) = gsGlsAyuda
        
        VerificaAceptar
    End If
End Sub

Sub FormResize()
    On Error Resume Next
    
    SSTab1.Width = Me.Width - 240
    SSTab1.Height = Me.Height - SSTab1.Top - 840 - Me.cmdAceptar.Height
    
    cmdCancelar.Top = SSTab1.Top + SSTab1.Height + 60
    cmdCancelar.Left = SSTab1.Left + SSTab1.Width - cmdCancelar.Width
    
    cmdAceptar.Top = SSTab1.Top + SSTab1.Height + 60
    cmdAceptar.Left = cmdCancelar.Left - cmdAceptar.Width - 30
    
    txtQuery.Width = SSTab1.Width - 120
    txtQuery.Height = SSTab1.Height - txtQuery.Top - 180 - cmdProbar.Height
    
    cmdFormato.Top = txtQuery.Top + txtQuery.Height + 60
    cmdFormato.Left = txtQuery.Left + txtQuery.Width - cmdFormato.Width

    cmdProbar.Top = cmdFormato.Top
    cmdProbar.Left = cmdFormato.Left - cmdFormato.Width - 30

    lstParametros.Width = SSTab1.Width - 120
    lstParametros.Height = SSTab1.Height - txtQuery.Top - 180 - cmdEditarParam.Height
    
    cmdProbar2.Top = lstParametros.Top + lstParametros.Height + 60
    cmdProbar2.Left = lstParametros.Left + lstParametros.Width - cmdEditarParam.Width
    
    cmdEditarParam.Top = cmdProbar2.Top
    cmdEditarParam.Left = cmdProbar2.Left - cmdProbar2.Width - 30
    
    'cmdNuevoParam.Top = cmdEditarParam.Top
    'cmdNuevoParam.Left = cmdEditarParam.Left - cmdEditarParam.Width - 30
    
    'lstColumnas.Width = SSTab1.Width - 120
    'lstColumnas.Height = SSTab1.Height - txtQuery.Top - 180 - cmdNuevoFormato.Height
    
    'cmdProbar3.Top = lstColumnas.Top + lstColumnas.Height + 60
    'cmdProbar3.Left = lstColumnas.Left + lstColumnas.Width - cmdEditarFormato.Width
    
    'cmdEditarFormato.Top = cmdProbar3.Top
    'cmdEditarFormato.Left = cmdProbar3.Left - cmdProbar3.Width - 30
    
    'cmdNuevoFormato.Top = cmdEditarFormato.Top
    'cmdNuevoFormato.Left = cmdEditarFormato.Left - cmdEditarFormato.Width - 30
    
    txtResultado.Width = SSTab1.Width - 120
    txtResultado.Height = SSTab1.Height - txtResultado.Top - 180 - cmdFormato2.Height
    
    grdResultado.Width = txtResultado.Width
    grdResultado.Height = txtResultado.Height
    grdResultado.Top = txtResultado.Top
    grdResultado.Left = txtResultado.Left

    cmdFormato2.Top = grdResultado.Top + grdResultado.Height + 60
    cmdFormato2.Left = grdResultado.Left + grdResultado.Width - cmdFormato2.Width
    
    cmdProbar3.Top = cmdFormato2.Top
    cmdProbar3.Left = cmdFormato2.Left - cmdFormato2.Width - 30
End Sub

Function GrabarConsulta() As Boolean
    Dim sNumConsulta        As String
    Dim sNomConsulta        As String
    Dim nNumBaseDatos       As Integer
    Dim sGlsQuery           As String
    Dim sGlsParametros      As String
    Dim sGlsFormatos        As String
    Dim sFormatoIn          As String
    Dim sFormatoOut         As String
    Dim sGlsHorarios        As String
    Dim nItem               As Integer
    Dim bOk                 As Boolean
    '<V1.3.0>
    Dim sGlsArchivoSalida   As String
    Dim sNomHojaSalida      As String
    '</V1.3.0>
    '<V1.3.1>
    Dim nNumArea            As String
    Dim nNumNegocio         As String
    '</V1.3.1>
    
    On Error GoTo ErrGrabarConsulta
        
    sNumConsulta = IIf(msNumConsulta = "", "0", msNumConsulta)
    
    ' Valida consistencia de informacion
    If cboBaseDatos.ListIndex < 0 Then
        MsgBox "No ha seleccionado la Base de Datos de esta consulta", vbCritical, App.Title
        GrabarConsulta = False
        Exit Function
    End If
    nNumBaseDatos = cboBaseDatos.ItemData(cboBaseDatos.ListIndex)
    
    sGlsQuery = Trim(txtQuery)
    If sGlsQuery = "" Then
        MsgBox "No ha detallado el Query a ejecutar", vbCritical, App.Title
        GrabarConsulta = False
        Exit Function
    End If
    
    sNomConsulta = Trim(Me.txtNomConsulta)
    If sNomConsulta = "" Then
        MsgBox "No ha ingresado el nombre de la consulta", vbCritical, App.Title
        GrabarConsulta = False
        Exit Function
    End If
    
    If Me.optOpcionHorario(1).Value = True Then
        If (cboHorIni(0).ListIndex > 0 And cboHorFin(0).ListIndex <= 0) Or _
           (cboHorIni(0).ListIndex <= 0 And cboHorFin(0).ListIndex > 0) Then
            MsgBox "Horario entre " & cboHorIni(0) & " y " & cboHorFin(0) & " no es correcto. Debe especificar un horario correcto", vbCritical, App.Title
            GrabarConsulta = False
            Exit Function
        End If
        
        If (cboHorIni(1).ListIndex > 0 And cboHorFin(1).ListIndex <= 0) Or _
           (cboHorIni(1).ListIndex <= 0 And cboHorFin(1).ListIndex > 0) Then
            MsgBox "Horario entre " & cboHorIni(1) & " y " & cboHorFin(1) & " no es correcto. Debe especificar un horario correcto", vbCritical, App.Title
            GrabarConsulta = False
            Exit Function
        End If
        
        If (cboHorIni(0).ListIndex <= 0 Or cboHorFin(0).ListIndex <= 0) And _
           (cboHorIni(1).ListIndex <= 0 Or cboHorFin(1).ListIndex <= 0) Then
            MsgBox "Debe indicar al menos un horario de ejecución correcto, o bien seleccionar la opción ""En todo horario""", vbCritical, App.Title
            GrabarConsulta = False
            Exit Function
        End If
    End If
    
    '<V1.3.0>
    If Trim(Me.txtNomHoja) <> "" And Trim(Me.txtNomArchivo) = "" Then
        MsgBox "Debe indicar el nombre del archivo excel al cual va a exportar la información, o limpiar el campo de la hoja", vbCritical, App.Title
        GrabarConsulta = False
        Exit Function
    End If
    
    If Me.optHoja(1).Value And Trim(Me.txtNomHoja) = "" Then
        MsgBox "Debe indicar el nombre de la hoja excel", vbCritical, App.Title
        GrabarConsulta = False
        Exit Function
    End If
    '</V1.3.0>
    
    Screen.MousePointer = vbHourglass
    
    ' Genera XML con los parámetros de la consulta
    Call CargaParametrosLocal(sGlsQuery)
    
    sGlsParametros = "<ROOT>"
    For nItem = 1 To UBound(maRegParametrosTest)
        sGlsParametros = sGlsParametros & "<Parametros "
        sGlsParametros = sGlsParametros & " num_consulta=""" & sNumConsulta & """"
        sGlsParametros = sGlsParametros & " nom_parametro=""" & maRegParametrosTest(nItem).Nombre & """"
        sGlsParametros = sGlsParametros & " gls_parametro=""" & Replace(maRegParametrosTest(nItem).Descripcion, "<", gsSignoMenor) & """"
        sGlsParametros = sGlsParametros & " cod_tipo_dato=""" & maRegParametrosTest(nItem).Tipo & """"
        sGlsParametros = sGlsParametros & " cod_tipo_ayuda=""" & maRegParametrosTest(nItem).TipoAyuda & """"
        sGlsParametros = sGlsParametros & " gls_ayuda_valores=""" & Replace(maRegParametrosTest(nItem).Ayuda, "<", gsSignoMenor) & """"
        sGlsParametros = sGlsParametros & " ind_opcional=""" & IIf(maRegParametrosTest(nItem).Opcional = True, "S", "N") & """"
        sGlsParametros = sGlsParametros & "/>"
    Next nItem
    sGlsParametros = sGlsParametros & "</ROOT>"
    
    ' XML para formatos de columnas
    sGlsFormatos = "<ROOT>"
    mrsFormatos.Filter = ""
    If Not mrsFormatos.EOF Then
        mrsFormatos.MoveFirst
        While Not mrsFormatos.EOF
            sFormatoIn = Replace("" & mrsFormatos!gls_formato_entrada, "<", gsSignoMenor)
            sFormatoOut = Replace("" & mrsFormatos!gls_formato_salida, "<", gsSignoMenor)
            sFormatoIn = Replace(sFormatoIn, """", gsSignoComillas)
            sFormatoOut = Replace(sFormatoOut, """", gsSignoComillas)
            
            sGlsFormatos = sGlsFormatos & "<Formatos"
            sGlsFormatos = sGlsFormatos & " nom_columna=""" & "" & mrsFormatos!nom_columna & """"
            sGlsFormatos = sGlsFormatos & " cod_tipo_dato_salida=""" & "" & mrsFormatos!cod_tipo_dato_salida & """"
            sGlsFormatos = sGlsFormatos & " ind_separador_miles=""" & "" & mrsFormatos!ind_separador_miles & """"
            sGlsFormatos = sGlsFormatos & " num_decimales=""" & "" & mrsFormatos!num_decimales & """"
            sGlsFormatos = sGlsFormatos & " gls_formato_entrada=""" & sFormatoIn & """"
            sGlsFormatos = sGlsFormatos & " gls_formato_salida=""" & sFormatoOut & """"
            sGlsFormatos = sGlsFormatos & "/>"
            
            mrsFormatos.MoveNext
        Wend
    End If
    sGlsFormatos = sGlsFormatos & "</ROOT>"
    
    ' Horarios
    If Me.optOpcionHorario(0).Value = True Then
        sGlsHorarios = ""
    Else
        sGlsHorarios = Left(cboHorIni(0).Text & Space(10), 10) & Left(cboHorFin(0).Text & Space(10), 10) & Left(cboHorIni(1).Text & Space(10), 10) & Left(cboHorFin(1).Text & Space(10), 10)
        sGlsHorarios = Replace(sGlsHorarios, "(ninguno)", "         ")
    End If
    
    '<V1.3.0>
    ' Archivo Excel
    sGlsArchivoSalida = Trim(Me.txtNomArchivo.Text)
    sNomHojaSalida = IIf(optHoja(0).Value, "<nueva>", Trim(Me.txtNomHoja.Text))
    '</V1.3.0>

    '<V1.3.1>
    nNumArea = -1
    nNumNegocio = -1
    If cboCodArea.ListIndex >= 0 Then
        nNumArea = cboCodArea.ItemData(cboCodArea.ListIndex)
    End If
    If cboCodNegocio.ListIndex >= 0 Then
        nNumNegocio = cboCodNegocio.ItemData(cboCodNegocio.ListIndex)
    End If
    '</V1.3.1>

    ' Graba informacion
    bOk = db_GrabaConsulta(sNumConsulta, sNomConsulta, nNumBaseDatos, sGlsQuery, sGlsParametros, sGlsFormatos, sGlsHorarios, _
                           sGlsArchivoSalida, sNomHojaSalida, nNumArea, nNumNegocio, msNomUsuarioLocal)
    
    Screen.MousePointer = vbNormal
    
    If Not bOk Then
        GrabarConsulta = False
    Else
        If sNumConsulta <> msNumConsulta Then
            MsgBox "Consulta fue creada con el número " & sNumConsulta
            gsNumConsulta = sNumConsulta
        End If
        
        cmdAceptar.Enabled = False
        GrabarConsulta = True
        gbCancelar = False
    End If
    
    Exit Function
    
ErrGrabarConsulta:
    Screen.MousePointer = vbNormal
    MsgBox Err.Description, vbCritical, App.Title
    GrabarConsulta = False
    Exit Function
    Resume
End Function

Sub IniciaLista()
    Me.lstParametros.ColumnHeaders.Add , , "Nombre", 2000
    Me.lstParametros.ColumnHeaders.Add , , "Título", 2000
    Me.lstParametros.ColumnHeaders.Add , , "Tipo", 1000
    Me.lstParametros.ColumnHeaders.Add , , "Ayuda", 1000
    Me.lstParametros.ColumnHeaders.Add , , "Detalle Ayuda", 5000
    Me.lstParametros.ColumnHeaders.Add , , "Opcional", 1200
    Me.lstParametros.ColumnHeaders.Add , , "", 0
End Sub

Sub CargaConsulta()
    Dim rsData              As ADODB.Recordset
    Dim i                   As Integer
    Dim sGlsHorario         As String
    '<V1.3.0>
    Dim sGlsArchivoSalida   As String
    Dim sNomHojaSalida      As String
    '</V1.3.0>
    '<V1.3.1>
    Dim nNumArea            As Long
    Dim nNumNegocio         As Long
    '</V1.3.1>
        
    On Error GoTo ErrCargaConsulta
            
    Screen.MousePointer = vbHourglass
    
    ' Abre base datos
    OpenMyDataBase
    
    ' Lee consulta
    If db_LeeConsulta(msNumConsulta, rsData) Then
        If Not rsData.EOF Then
            mnNumBaseDatos = Val(rsData!num_basedatos)
            msGlsQuery = "" & rsData!gls_query
            sGlsHorario = "" & rsData!gls_horario_ejecucion
            '<V1.3.0>
            sGlsArchivoSalida = "" & rsData!gls_archivo_salida
            sNomHojaSalida = "" & rsData!nom_hoja_salida
            '</V1.3.0>
            '<V1.3.1>
            nNumArea = IIf(IsNull(rsData!num_area), -1, rsData!num_area)
            Call BuscaTabValor("AREA", nNumArea)
            
            nNumNegocio = IIf(IsNull(rsData!num_negocio), -1, rsData!num_negocio)
            Call BuscaTabValor("NEGOCIO", nNumNegocio)
            '</V1.3.1>
        
            Call CargaParametros(msNumConsulta, msGlsQuery, maRegParametros())
            Call db_LeeFormatos(msNumConsulta, mrsFormatos)
        End If
    End If
        
    ' Cierra base datos
    CloseMyDataBase
    
    ' Carga informacion en la pantalla
    Me.txtQuery.Text = msGlsQuery
    
    For i = 1 To UBound(maRegParametros)
        Set mItem = Me.lstParametros.ListItems.Add(, , LCase(maRegParametros(i).Nombre))
        mItem.SubItems(1) = maRegParametros(i).Descripcion
        mItem.SubItems(2) = maRegParametros(i).Tipo
        mItem.SubItems(3) = maRegParametros(i).TipoAyuda
        mItem.SubItems(4) = fsConvierteTextoToLinea(maRegParametros(i).Ayuda)
        mItem.SubItems(5) = IIf(maRegParametros(i).Opcional, "S", "N")
        mItem.SubItems(6) = maRegParametros(i).Ayuda
    Next i
    
    If lstParametros.ListItems.Count > 0 Then
        Set mItem = Me.lstParametros.ListItems.Item(1)
        mItem.Selected = True
        mItem.EnsureVisible
        Me.cmdEditarParam.Enabled = True
    Else
        Me.cmdEditarParam.Enabled = False
    End If

    For i = 0 To Me.cboBaseDatos.ListCount - 1
        If cboBaseDatos.ItemData(i) = mnNumBaseDatos Then
            cboBaseDatos.ListIndex = i
            Exit For
        End If
    Next i
    
    Call CargaHorarios(sGlsHorario)
    
    '<V1.3.0>
    Call CargaInformacionArchivoExcel(sGlsArchivoSalida, sNomHojaSalida)
    '</V1.3.0>
    
    Screen.MousePointer = vbNormal
    Exit Sub
    
ErrCargaConsulta:
    Screen.MousePointer = vbNormal
    MsgBox Error, vbInformation, App.Title
    Exit Sub
End Sub


Sub ParametroNuevo()
    Dim sGlsAyuda   As String
    
    frmDefParametros.msNomParametro = ""
    frmDefParametros.msTipoDato = ""
    frmDefParametros.msGlsAyuda = ""
    frmDefParametros.Show vbModal
    
    If gsTipoDato <> "" Then
        Set mItem = Me.lstParametros.ListItems.Add(, , LCase(gsNomParametro))
        mItem.SubItems(1) = gsGlsParametro
        mItem.SubItems(2) = gsTipoDato
        mItem.SubItems(3) = gsTipoAyuda
        mItem.SubItems(4) = fsConvierteTextoToLinea(gsGlsAyuda)
        mItem.SubItems(5) = gsIndOpcional
        mItem.SubItems(6) = gsGlsAyuda
    
        mItem.Selected = True
        mItem.EnsureVisible
        Me.cmdEditarParam.Enabled = True
        
        VerificaAceptar
    End If
End Sub
Sub RefrescaParametros()
    Dim sFile               As String
    Dim sGlsQuery           As String
    Dim sGlsSetup           As String
    Dim sGlsAyuda           As String
    Dim aRegParametros()    As rRegParametros
    Dim nItem               As Integer
    
    ' Abre base de datos
    OpenMyDataBase
    
    ' Carga parametros
    Call CargaParametros(msNumConsulta, Me.txtQuery, aRegParametros())
    
    ' Cierra base de datos
    CloseMyDataBase
    
    For nItem = 1 To UBound(aRegParametros)
        If Not ExisteParametro(aRegParametros(nItem).Nombre) Then
            Set mItem = Me.lstParametros.ListItems.Add(, , LCase(aRegParametros(nItem).Nombre))
            mItem.SubItems(1) = aRegParametros(nItem).Descripcion
            mItem.SubItems(2) = aRegParametros(nItem).Tipo
            mItem.SubItems(3) = aRegParametros(nItem).TipoAyuda
            mItem.SubItems(4) = fsConvierteTextoToLinea(aRegParametros(nItem).Ayuda)
            mItem.SubItems(5) = aRegParametros(nItem).Opcional
            mItem.SubItems(6) = aRegParametros(nItem).Ayuda
        End If
    Next nItem
End Sub

Sub VerificaAceptar()
    cmdAceptar.Enabled = (Trim(txtNomConsulta.Text) <> "" And Trim(txtQuery.Text) <> "" And cboBaseDatos.ListIndex >= 0)
End Sub

Private Sub cboCodArea_Click()
    VerificaAceptar
End Sub


Private Sub cboCodNegocio_Click()
    VerificaAceptar
End Sub


Private Sub cboHoja_Click()
    '<V1.3.0>
    If cboHoja.ListIndex >= 0 Then
        txtNomHoja = cboHoja.Text
        
        VerificaAceptar
    End If
    '</V1.3.0>
End Sub


Private Sub cboHorFin_Click(Index As Integer)
    If cboHorIni(Index).ListIndex < 0 Or cboHorIni(Index).Tag = "*" Then
        cboHorIni(Index).Tag = "*"
        
        If cboHorFin(Index).ListIndex = 0 Then
            cboHorIni(Index).ListIndex = 0
        ElseIf cboHorFin(Index).ListIndex = 1 Then
            cboHorIni(Index).ListIndex = cboHorFin(Index).ListCount - 2
        Else
            cboHorIni(Index).ListIndex = cboHorFin(Index).ListIndex - 1
        End If
    End If
    VerificaAceptar
End Sub


Private Sub cboHorFin_LostFocus(Index As Integer)
    cboHorIni(Index).Tag = ""
End Sub


Private Sub cboHorIni_Click(Index As Integer)
    If cboHorFin(Index).ListIndex < 0 Or cboHorFin(Index).Tag = "*" Then
        cboHorFin(Index).Tag = "*"
        
        If cboHorIni(Index).ListIndex = 0 Then
            cboHorFin(Index).ListIndex = 0
        ElseIf cboHorIni(Index).ListIndex = cboHorIni(Index).ListCount - 1 Then
            cboHorFin(Index).ListIndex = 2
        Else
            cboHorFin(Index).ListIndex = cboHorIni(Index).ListIndex + 1
        End If
    End If
    VerificaAceptar
End Sub


Private Sub cboHorIni_LostFocus(Index As Integer)
    cboHorFin(Index).Tag = ""
End Sub


Private Sub cmdAceptar_Click()
    If GrabarConsulta() Then
        Unload Me
    End If
End Sub


Private Sub cmdCancelar_Click()
    gbCancelar = True
    cmdAceptar.Enabled = False
    Unload Me
End Sub


Private Sub cmdEditarParam_Click()
    ParametroEditar
End Sub

Private Sub cmdFormato_Click()
    FormatoColumna
End Sub

Private Sub cmdFormato2_Click()
    FormatoColumna
End Sub


Private Sub cmdHelpCarpPer_Click()
    '<V1.3.0>
    HelpNombreArchivo
    '</V1.3.0>
End Sub

Private Sub cmdHelpHoja_Click()
    '<V1.3.0>
    If Trim(txtNomArchivo.Text) <> "" And Exist(txtNomArchivo.Text) Then
        If optHoja(0).Value = True Then
            Me.txtNomHoja.Visible = False
            Me.cboHoja.Visible = True
        Else
            If CargaHojas("") Then
                Me.txtNomHoja.Visible = False
                Me.cboHoja.Visible = True
            Else
                txtNomHoja.Visible = True
                cboHoja.Visible = False
                cboHoja.Clear
            End If
        End If
    End If
    '</V1.3.0>
End Sub

Private Sub cmdProbar_Click()
    ProbarConsulta
End Sub

Private Sub cmdProbar2_Click()
    ProbarConsulta
End Sub

Private Sub cmdProbar3_Click()
    ProbarConsulta
End Sub


Private Sub Form_Activate()
    If bFormLoad = False Then
        If msNumConsulta <> "" Then
            txtQuery.SetFocus
        End If
        cmdAceptar.Enabled = False
        bFormLoad = True
    End If
End Sub

Private Sub Form_Load()
    IniciaForm
    IniciaLista
    IniciaCombos
    '<V1.3.1>
    CargaCodigos
    '</V1.3.1>
    
    If msNumConsulta <> "" Then
        CargaConsulta
    Else
        IniciaFormatos
    End If
End Sub

Sub IniciaForm()
    Dim nX  As Integer
    
    Me.Left = (Screen.Width - Me.Width) \ 2
    Me.Top = (Screen.Height - Me.Height) \ 3
    
    grsBaseDatos.Filter = ""
    If Not grsBaseDatos.EOF Then
        grsBaseDatos.MoveFirst
        While Not grsBaseDatos.EOF
            cboBaseDatos.AddItem "" & grsBaseDatos!nom_basedatos
            cboBaseDatos.ItemData(cboBaseDatos.ListCount - 1) = grsBaseDatos!num_basedatos
            grsBaseDatos.MoveNext
        Wend
    End If
    
    msNumConsulta = gsNumConsulta
    txtNomConsulta = gsNomConsulta
    msNomUsuarioLocal = gsNomUsuarioLocal
    StatusBar1.Panels(1).Text = UCase(msNomUsuarioLocal)
    
    If msNumConsulta = "" Then
        txtNomConsulta.Enabled = True
    Else
        txtNomConsulta.Enabled = False
    End If
    Me.grdResultado.Visible = False
    Me.txtResultado.Visible = False
    
    Me.txtQuery.Left = 60
    Me.txtQuery.Top = 420
    Me.lstParametros.Left = 60
    Me.lstParametros.Top = 420
    Me.grdResultado.Top = 420
    Me.grdResultado.Left = 60
    Me.txtResultado.Left = 60
    Me.txtResultado.Top = 420
    
    optOpcionHorario(0).Value = True
    
    '<V1.3.0>
    cboHoja.Top = Me.txtNomHoja.Top
    cboHoja.Left = Me.txtNomHoja.Left
    '</V1.3.0>
    
    gbCancelar = True
    bFormLoad = False
    Me.SSTab1.Tab = 0
End Sub

Private Sub cboBaseDatos_Click()
    VerificaAceptar
End Sub

Private Sub Form_Resize()
    FormResize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim nResp   As Integer
    
    If cmdAceptar.Enabled Then
        nResp = MsgBox("Desea grabar los cambios realizados sobre esta consulta", vbYesNoCancel + vbDefaultButton1, App.Title)
        If nResp = vbYes Then
            If Not GrabarConsulta() Then
                Cancel = True
            End If
        ElseIf nResp = vbCancel Then
            Cancel = True
        Else
            gbCancelar = True
        End If
    End If
End Sub

Private Sub grdResultado_DblClick(ByVal Col As Long, ByVal Row As Long)
    FormatoColumna
End Sub


Private Sub lstParametros_DblClick()
    If lstParametros.ListItems.Count > 0 Then
        ParametroEditar
    End If
End Sub

Private Sub lstParametros_ItemClick(ByVal Item As ComctlLib.ListItem)
    Set mItem = Item
End Sub


Private Sub optHoja_Click(Index As Integer)
    '<V1.3.0>
    Me.txtNomHoja.Enabled = (Index = 1)
    Me.cboHoja.Enabled = (Index = 1)
    If Index = 1 Then
        If cboHoja.Visible And cboHoja.ListCount = 0 Then
            If CargaHojas("") Then
                Me.txtNomHoja.Visible = False
                Me.cboHoja.Visible = True
            Else
                txtNomHoja.Visible = True
                cboHoja.Visible = False
                cboHoja.Clear
            End If
        End If
    End If
    
    VerificaAceptar
    '</V1.3.0>
End Sub

Private Sub optOpcionHorario_Click(Index As Integer)
    If optOpcionHorario(0).Value Then
        cboHorIni(0).Enabled = False
        cboHorFin(0).Enabled = False
        cboHorIni(1).Enabled = False
        cboHorFin(1).Enabled = False
    Else
        cboHorIni(0).Enabled = True
        cboHorFin(0).Enabled = True
        cboHorIni(1).Enabled = True
        cboHorFin(1).Enabled = True
    End If
    VerificaAceptar
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    If SSTab1.Tab <> PreviousTab Then
        txtQuery.Visible = False
        cmdFormato.Visible = False
        cmdProbar.Visible = False
        lstParametros.Visible = False
        'cmdNuevoParam.Visible = False
        cmdEditarParam.Visible = False
        cmdProbar2.Visible = False
        cmdFormato2.Visible = False
        cmdProbar3.Visible = False

        Select Case SSTab1.Tab
        Case 0
            grdResultado.Visible = False
            txtQuery.Visible = True
            cmdFormato.Visible = True
            cmdProbar.Visible = True
            Me.HelpContextID = 19
        Case 1
            grdResultado.Visible = False
            lstParametros.Visible = True
            'cmdNuevoParam.Visible = True
            cmdEditarParam.Visible = True
            cmdProbar2.Visible = True
            Me.HelpContextID = 20
            
            RefrescaParametros
        Case 2
            grdResultado.Visible = (Val(grdResultado.Tag) > 0)
            cmdFormato2.Visible = True
            cmdProbar3.Visible = True
            cmdFormato2.Enabled = True
            Me.HelpContextID = 21
        End Select
    End If
End Sub

Private Sub txtNomArchivo_Change()
    '<V1.3.0>
    If Not Exist(txtNomArchivo.Text) Then
        txtNomHoja.Visible = True
        cboHoja.Visible = False
        cboHoja.Clear
    Else
        If Me.optHoja(0).Value Then
            Me.txtNomHoja.Visible = False
            Me.cboHoja.Visible = True
        ElseIf Me.optHoja(1).Value Then
            If CargaHojas("") Then
                Me.cboHoja.Visible = True
                Me.txtNomHoja.Visible = False
            Else
                txtNomHoja.Visible = True
                cboHoja.Visible = False
                cboHoja.Clear
            End If
        End If
    End If
    
    VerificaAceptar
    '</V1.3.0>
End Sub

Private Sub txtNomConsulta_Change()
    VerificaAceptar
End Sub

Private Sub txtQuery_Change()
    VerificaAceptar
    grdResultado.Visible = False
    grdResultado.Tag = ""
End Sub

Sub ProbarConsulta()
    Dim nTotRegistros       As Long
    Dim sGlsQuery           As String
    Dim aRegParametros()    As rRegParametros
    Dim nNumBaseDatos       As Integer
    
    If cboBaseDatos.ListIndex < 0 Then
        Screen.MousePointer = vbNormal
        MsgBox "No ha seleccionado la Base de Datos de esta consulta", vbCritical, App.Title
        Exit Sub
    End If
    nNumBaseDatos = cboBaseDatos.ItemData(cboBaseDatos.ListIndex)
    
    sGlsQuery = Trim(txtQuery)
    If sGlsQuery = "" Then
        Screen.MousePointer = vbNormal
        MsgBox "No ha detallado el Query a ejecutar", vbCritical, App.Title
        Exit Sub
    End If
        
    Call CargaParametrosLocal(sGlsQuery)
    
    If EjecutaConsulta(msNumConsulta, nNumBaseDatos, sGlsQuery, maRegParametrosTest, mrsData, mrsFormatos, nTotRegistros, True, grdResultado, Me.txtResultado, StatusBar1, ProgressBar1) Then
        Me.grdResultado.Tag = nTotRegistros
        If nTotRegistros = 0 Then
            Me.grdResultado.Visible = False
        Else
            Me.SSTab1.Tab = 2
        End If
        
        cmdFormato.Enabled = (Me.grdResultado.Visible)
        cmdFormato2.Enabled = (Me.grdResultado.Visible)
    End If
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

Function CargaHojas(sNomHoja As String) As Boolean
    '<V1.3.0>
    Dim Excel       As Object
    Dim sFile       As String
    Dim i           As Integer
    Dim nItemOld    As String
    
    On Error GoTo ErrCargaHojas
    
    Me.cboHoja.Clear
    nItemOld = -1
    
    Screen.MousePointer = vbHourglass

    sFile = Trim(Me.txtNomArchivo.Text)
    If Exist(sFile) Then
        Set Excel = CreateObject("Excel.Application")
        Excel.Workbooks.Open FileName:=sFile
        
        For i = 1 To Excel.Sheets.Count
            cboHoja.AddItem Excel.Sheets(i).Name
            If sNomHoja = Excel.Sheets(i).Name Then
                nItemOld = i
            End If
        Next
        
        Excel.Workbooks.Close
        Set Excel = Nothing
    End If

    If cboHoja.ListCount > 0 Then
        If nItemOld > 0 Then
            cboHoja.ListIndex = nItemOld - 1
        Else
            cboHoja.ListIndex = 0
        End If
        cboHoja.Enabled = True
    End If
    
    Screen.MousePointer = vbNormal
    CargaHojas = True
    
    Exit Function
    
ErrCargaHojas:
    On Error Resume Next
    Excel.Workbooks.Close
    Set Excel = Nothing
    
    Screen.MousePointer = vbNormal
    CargaHojas = False
    
    Exit Function
    '</V1.3.0>
End Function


