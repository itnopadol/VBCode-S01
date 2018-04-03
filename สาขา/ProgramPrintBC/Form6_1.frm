VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "CRYSTL32.OCX"
Begin VB.Form Form6_1 
   Caption         =   "พิมพ์ตั๋วแลกเงิน"
   ClientHeight    =   8355
   ClientLeft      =   2460
   ClientTop       =   825
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form6_1.frx":0000
   ScaleHeight     =   8355
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CrystalReport6_1 
      Left            =   1920
      Top             =   6120
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
      WindowShowCloseBtn=   -1  'True
      WindowShowProgressCtls=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   315
      Left            =   2850
      TabIndex        =   11
      Top             =   2475
      Width           =   2190
      _ExtentX        =   3863
      _ExtentY        =   556
      _Version        =   393216
      Format          =   69599233
      CurrentDate     =   38491
   End
   Begin VB.CheckBox Check101 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "พิมพ์คำร้องตั๋วแลกเงิน"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   2850
      TabIndex        =   9
      Top             =   1350
      Width           =   2190
   End
   Begin VB.CommandButton CMD101 
      Caption         =   "พิมพ์ใบคำร้องตั๋ว"
      Enabled         =   0   'False
      Height          =   390
      Left            =   3525
      TabIndex        =   8
      Top             =   3975
      Width           =   1515
   End
   Begin VB.ComboBox CMBBank 
      Height          =   315
      Left            =   6525
      TabIndex        =   6
      Top             =   3075
      Width           =   2640
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   2850
      TabIndex        =   5
      Top             =   3075
      Width           =   2190
      _ExtentX        =   3863
      _ExtentY        =   556
      _Version        =   393216
      Format          =   69599233
      CurrentDate     =   38143
   End
   Begin VB.CommandButton CMD6_11 
      Caption         =   "พิมพ์ตั๋วแลกเงิน"
      Height          =   390
      Left            =   7650
      TabIndex        =   2
      Top             =   3975
      Width           =   1515
   End
   Begin VB.TextBox TXT6_11 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2850
      TabIndex        =   0
      Top             =   1875
      Width           =   2190
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "วันที่ตั๋วแลกเงิน"
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   1275
      TabIndex        =   10
      Top             =   3075
      Width           =   1515
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "เลือกธนาคาร"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   5550
      TabIndex        =   7
      Top             =   3075
      Width           =   915
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "วันที่พิมพ์ใบขอรับรองตั๋ว"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   300
      TabIndex        =   4
      Top             =   2475
      Width           =   2490
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "พิมพ์ตั๋วแลกเงิน"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   14.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   465
      Left            =   2475
      TabIndex        =   3
      Top             =   300
      Width           =   7515
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "เลขที่ตั๋วแลกเงิน"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   465
      Left            =   1725
      TabIndex        =   1
      Top             =   1875
      Width           =   1065
   End
End
Attribute VB_Name = "Form6_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check101_Click()
If Check101.Value = 1 Then
    CMD101.Enabled = True
    CMBBank.Enabled = False
    CMD6_11.Enabled = False
    TXT6_11.SetFocus
End If
If Check101.Value = 0 Then
    CMD101.Enabled = False
    CMBBank.Enabled = True
    CMD6_11.Enabled = True
End If
End Sub

Private Sub CMD101_Click()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vDocNo As String, vReportName As String
Dim vRepType As String
Dim vRepID As Integer
Dim vDocDate1, vDocDate2 As Date


On Error GoTo ErrDescription

vDocNo = Trim(TXT6_11.Text)

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
'vQuery = "select reportname from dbo.bcreportname where reptype = 'chq' and repid = 227 "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vReportName = Trim(vRecordset.Fields("reportname").Value)
End If
vRecordset.Close

vDocDate2 = DTPicker2.Day & "/" & DTPicker2.Month & "/" & DTPicker2.Year
vDocDate1 = DTPicker1.Day & "/" & DTPicker1.Month & "/" & DTPicker1.Year
With CrystalReport6_1
.ReportFileName = vReportName & ".rpt"
.ParameterFields(0) = "ChqNumber;" & vDocNo & ";true"
.ParameterFields(1) = "docdate;" & vDocDate1 & ";true"
.Destination = crptToWindow
.WindowState = crptMaximized
.Action = 1
End With


ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description, vbCritical, "ข้อความเตือน"
Exit Sub
End If
End Sub

Private Sub CMD6_11_Click()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vDocNo As String, vReportName As String
Dim vRepType As String
Dim vRepID As Integer
Dim Date1 As Date

On Error GoTo ErrDescription
vDocNo = Trim(TXT6_11.Text)

If CMBBank.Text = "" Then
MsgBox "กรุณาเลือกธนาคารที่จะพิมพ์ด้วยนะครับ"
Exit Sub
ElseIf CMBBank.Text = Trim("ธนาคารเอเชีย") Then
vRepType = "BD"
vRepID = 124
ElseIf CMBBank.Text = Trim("ธนาคารกรุงเทพฯ") Then
vRepType = "CHQ"
vRepID = 210
End If
'Date1 = Trim(TXT6_12.Text)

Date1 = DTPicker1.Day & "/" & DTPicker1.Month & "/" & DTPicker1.Year + 543
vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
'vQuery = "select reportname from dbo.bcreportname where reptype = '" & vRepType & "'and repid = " & vRepID & " "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vReportName = Trim(vRecordset.Fields("reportname").Value)
End If
vRecordset.Close

With CrystalReport6_1
.ReportFileName = vReportName & ".rpt"
.ParameterFields(0) = "ExchangeDate;" & Date1 & ";true"
.ParameterFields(1) = "@ChqNumber;" & vDocNo & ";true"
.Destination = crptToWindow
.WindowState = crptMaximized
.Action = 1
End With

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description, vbCritical, "ข้อความเตือน"
Exit Sub
End If
End Sub

Private Sub Form_Load()

CMBBank.AddItem Trim("ธนาคารกรุงเทพฯ")
CMBBank.AddItem Trim("ธนาคารเอเชีย")

End Sub
