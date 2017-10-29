VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmExcel 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exportar a Excel"
   ClientHeight    =   825
   ClientLeft      =   5790
   ClientTop       =   5445
   ClientWidth     =   3735
   Icon            =   "frmExcel.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   825
   ScaleWidth      =   3735
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   0
      Top             =   660
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Generando planilla excel ..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3270
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   720
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   4
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmExcel.frx":058A
            Key             =   "K1"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmExcel.frx":0764
            Key             =   "K2"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmExcel.frx":093E
            Key             =   "K3"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmExcel.frx":0B18
            Key             =   "K4"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim nImagen     As Integer
Dim gSysTray    As clsSysTray
Private Sub Form_Load()
    Dim sFile   As String
    
    On Error GoTo ErrCarga
    
    sFile = Command()
    If sFile = "" Then
        End
    End If
    
    Set gSysTray = New clsSysTray
    Set gSysTray.SourceWindow = Me
    
    nImagen = 1
    gSysTray.Icon = ImageList1.ListImages("K1").Picture
    gSysTray.ChangeToolTip "Generando planilla"
    
    Me.WindowState = vbMinimized
    gSysTray.MinToSysTray

    DoEvents
    Call ExportToExcelFromFile(sFile)

    gSysTray.RemoveFromSysTray
    Set gSysTray = Nothing
    End
    
ErrCarga:
End Sub

Private Sub Timer1_Timer()
    Dim sIcono  As String
    
    nImagen = nImagen + 1
    If nImagen = 5 Then nImagen = 1
    sIcono = "K" & Format(nImagen, "&")
    gSysTray.Icon = ImageList1.ListImages(sIcono).Picture
End Sub

