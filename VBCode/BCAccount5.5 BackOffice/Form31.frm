VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "CRYSTL32.OCX"
Begin VB.Form Form31 
   Caption         =   "พิมพ์เอกสารขาย"
   ClientHeight    =   9000
   ClientLeft      =   2025
   ClientTop       =   1920
   ClientWidth     =   11970
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form31.frx":0000
   ScaleHeight     =   9000
   ScaleWidth      =   11970
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CrystalReport312 
      Left            =   2610
      Top             =   6525
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowNavigationCtls=   -1  'True
      WindowShowCancelBtn=   -1  'True
      WindowShowPrintBtn=   -1  'True
      WindowShowExportBtn=   -1  'True
      WindowShowZoomCtl=   -1  'True
      WindowShowProgressCtls=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.CommandButton CMD312 
      Caption         =   "พิมพ์บิลเงินเชื่อ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   4050
      TabIndex        =   7
      Top             =   5175
      Width           =   1770
   End
   Begin VB.CheckBox CKBill 
      BackColor       =   &H00FFFFFF&
      Caption         =   "บิลเงินเชื่อ"
      Height          =   330
      Left            =   3240
      TabIndex        =   6
      Top             =   3375
      Width           =   1410
   End
   Begin Crystal.CrystalReport CrystalReport311 
      Left            =   1215
      Top             =   5895
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.OptionButton Opt312 
      Caption         =   "ฟอร์ม ชื่อย่อ"
      ForeColor       =   &H8000000D&
      Height          =   300
      Left            =   3225
      Picture         =   "Form31.frx":72FB
      TabIndex        =   5
      Top             =   2925
      Width           =   1440
   End
   Begin VB.OptionButton Opt311 
      Caption         =   "ฟอร์ม ชื่อเต็ม"
      ForeColor       =   &H8000000D&
      Height          =   300
      Left            =   3225
      Picture         =   "Form31.frx":E5F6
      TabIndex        =   4
      Top             =   2475
      Value           =   -1  'True
      Width           =   1440
   End
   Begin VB.CommandButton CMD311 
      Caption         =   "พิมพ์เงินสด"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4050
      TabIndex        =   1
      Top             =   4455
      Width           =   1785
   End
   Begin VB.TextBox TXT311 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3225
      TabIndex        =   0
      Top             =   1800
      Width           =   2565
   End
   Begin VB.Label LBL313 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "เลือกฟอร์มที่พิมพ์ :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   390
      Left            =   1305
      TabIndex        =   3
      Top             =   2430
      Width           =   1830
   End
   Begin VB.Label LBL312 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "เลขที่เอกสาร :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   390
      Left            =   1620
      TabIndex        =   2
      Top             =   1845
      Width           =   1545
   End
End
Attribute VB_Name = "Form31"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMD311_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocNo As String
Dim vRepID As Integer
Dim vRepType As String
Dim vShow As String, vReportName As String

On Error GoTo Errdescription

vDocNo = Trim(TXT311.Text)
vRepID = 2
vRepType = "INV"

If Opt311.Value = True Then
    vShow = 1
ElseIf Opt312.Value = True Then
    vShow = 0
End If

vQuery = "select reportname from bcvat.dbo.bcreportname where repid = " & vRepID & " and reptype = '" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vReportName = Trim(vRecordset.Fields("reportname").Value)
    
    With CrystalReport311
            .ReportFileName = vReportName & ".rpt"
            .ParameterFields(0) = "@DocNo;" & vDocNo & ";true"
            .Destination = crptToWindow
            .WindowState = crptMaximized
            .Action = 1
    End With
End If
vRecordset.Close

Errdescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If

End Sub

Private Sub CMD312_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocNo As String
Dim vRepID As Integer
Dim vRepType As String
Dim vShow As String, vReportName As String

On Error GoTo Errdescription

vDocNo = Trim(TXT311.Text)
vRepID = 31
vRepType = "INV"


vQuery = "select reportname from bcvat.dbo.bcreportname where repid = " & vRepID & " and reptype = '" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vReportName = Trim(vRecordset.Fields("reportname").Value)
    
    With CrystalReport311
            .ReportFileName = vReportName & ".rpt"
            .ParameterFields(0) = "@DocNo;" & vDocNo & ";true"
            .Destination = crptToWindow
            .WindowState = crptMaximized
            .Action = 1
    End With
End If
vRecordset.Close

Errdescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub
