VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form Form113 
   Caption         =   "พิมพ์รายงาน ติดตามการแก้ไขสินค้าติดลบประจำวัน"
   ClientHeight    =   7605
   ClientLeft      =   3075
   ClientTop       =   1050
   ClientWidth     =   11835
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form113.frx":0000
   ScaleHeight     =   7605
   ScaleWidth      =   11835
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport Crystal101 
      Left            =   1530
      Top             =   5130
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
   Begin VB.CommandButton CMDPrintDailyReport 
      Caption         =   "พิมพ์รายงานประจำวัน"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   6885
      TabIndex        =   5
      Top             =   2655
      Width           =   2175
   End
   Begin VB.CheckBox CHK101 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "เลือก Section Manager :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   2340
      TabIndex        =   3
      Top             =   1485
      Width           =   2805
   End
   Begin VB.CommandButton CMDPrint 
      Caption         =   "พิมพ์รายงาน"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   5220
      TabIndex        =   2
      Top             =   2655
      Width           =   1500
   End
   Begin MSComCtl2.DTPicker DTPDocDate 
      Height          =   330
      Left            =   5220
      TabIndex        =   1
      Top             =   2070
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   61472769
      CurrentDate     =   39592
   End
   Begin VB.ComboBox CMBSection 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5220
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1485
      Width           =   3840
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "เลือกวันที่ดูรายงาน สินค้าติดลบ :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2475
      TabIndex        =   4
      Top             =   2070
      Width           =   2670
   End
End
Attribute VB_Name = "Form113"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CHK101_Click()
If Me.CHK101.Value = 1 Then
   Me.CMBSection.Enabled = True
Else
   Me.CMBSection.Enabled = False
End If
End Sub

Private Sub CMDPrint_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vRepID As Integer
Dim vRepType As String
Dim vReportName As String
Dim vSectionID As String
Dim vDocDate As String

Dim vLen As Integer
Dim vPosition As Integer

On Error GoTo ErrDescription


vRepID = 407
vRepType = "IC"

If Me.CHK101.Value = 1 Then
   vLen = Len(Me.CMBSection.Text)
   vPosition = InStr(1, Me.CMBSection.Text, "/")
   vSectionID = Right(Me.CMBSection.Text, vLen - vPosition)
Else
   vSectionID = ""
End If
vDocDate = Day(Me.DTPDocDate.Value) & "/" & Month(Me.DTPDocDate.Value) & "/" & Year(Me.DTPDocDate.Value)

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "'"
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
   vReportName = vRecordset.Fields("reportname").Value
End If
vRecordset.Close

With Me.Crystal101
.ReportFileName = vReportName & ".rpt"
.ParameterFields(0) = "@vSectionID;" & vSectionID & ";true "
.ParameterFields(1) = "@vDocDate;" & vDocDate & ";true "
.Destination = crptToWindow
.WindowState = crptMaximized
.Action = 1
End With


ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If

End Sub


Private Sub CMDPrintDailyReport_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vRepID As Integer
Dim vRepType As String
Dim vReportName As String
Dim vSectionID As String
Dim vDocDate As String


On Error GoTo ErrDescription


vRepID = 422
vRepType = "IC"

vSectionID = ""
vDocDate = Day(Me.DTPDocDate.Value) & "/" & Month(Me.DTPDocDate.Value) & "/" & Year(Me.DTPDocDate.Value)

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "'"
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
   vReportName = vRecordset.Fields("reportname").Value
End If
vRecordset.Close

With Me.Crystal101
.ReportFileName = vReportName & ".rpt"
.ParameterFields(0) = "@vSectionID;" & vSectionID & ";true "
.ParameterFields(1) = "@vDocDate;" & vDocDate & ";true "
.Destination = crptToWindow
.WindowState = crptMaximized
.Action = 1
End With


ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Form_Load()
Me.DTPDocDate.Value = Now

Call vGetSection
End Sub

Private Sub vGetSection()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vSecManList As ListItem

On Error GoTo ErrDescription

Me.CMBSection.Clear
vQuery = "exec USP_PM_FindSecMan"
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    While Not vRecordset.EOF
    Me.CMBSection.AddItem (Trim(vRecordset.Fields("secmanname").Value) & "/" & Trim(vRecordset.Fields("salecode").Value))
    vRecordset.MoveNext
    Wend
End If
vRecordset.Close

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If

End Sub

