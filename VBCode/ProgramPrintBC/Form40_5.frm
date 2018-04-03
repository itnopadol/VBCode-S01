VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "CRYSTL32.OCX"
Begin VB.Form Form40_5 
   Caption         =   "รายงาน เคลื่อนไหวเจ้าหนี้ ตามช่วงวันที่"
   ClientHeight    =   8130
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form40_5.frx":0000
   ScaleHeight     =   8130
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport Crystal101 
      Left            =   2040
      Top             =   5640
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
   Begin VB.CommandButton CMD101 
      Caption         =   "พิมพ์รายงาน"
      Height          =   465
      Left            =   4500
      TabIndex        =   3
      Top             =   3450
      Width           =   1290
   End
   Begin MSComCtl2.DTPicker DTPicker102 
      Height          =   315
      Left            =   2700
      TabIndex        =   2
      Top             =   2850
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      _Version        =   393216
      Format          =   66584577
      CurrentDate     =   38463
   End
   Begin MSComCtl2.DTPicker DTPicker101 
      Height          =   315
      Left            =   2700
      TabIndex        =   1
      Top             =   2325
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      _Version        =   393216
      Format          =   66584577
      CurrentDate     =   38463
   End
   Begin VB.ComboBox CMB101 
      Height          =   315
      Left            =   2700
      TabIndex        =   0
      Top             =   1650
      Width           =   4440
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "ถึงวันที่"
      ForeColor       =   &H8000000E&
      Height          =   240
      Left            =   2100
      TabIndex        =   6
      Top             =   2850
      Width           =   540
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "จากวันที่"
      ForeColor       =   &H8000000E&
      Height          =   240
      Left            =   2025
      TabIndex        =   5
      Top             =   2325
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "รหัสเจ้าหนี้"
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   1875
      TabIndex        =   4
      Top             =   1650
      Width           =   840
   End
End
Attribute VB_Name = "Form40_5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMD101_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vReportName As String
Dim vApCode As String
Dim vDate1, vDate2 As Date
Dim StrCount As Integer
Dim vRepID As Integer
Dim vRepType As String

On Error GoTo ErrDescription

StrCount = InStr(Trim(CMB101.Text), "/")
vApCode = Trim(Left(CMB101.Text, StrCount - 1))
If vApCode <> "" Then
vDate1 = DTPicker101.Day & "/" & DTPicker101.Month & "/" & DTPicker101.Year
vDate2 = DTPicker102.Day & "/" & DTPicker102.Month & "/" & DTPicker102.Year

vRepID = 214
vRepType = "AP"
vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
'vQuery = "select reportname from bcnp.dbo.bcreportname where repid = 214  and reptype = 'AP' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vReportName = Trim(vRecordset.Fields("reportname").Value)
End If
vRecordset.Close

With Crystal101
.ReportFileName = vReportName & ".rpt"
.ParameterFields(0) = "@Apcode;" & vApCode & ";true"
.ParameterFields(1) = "@Start;" & vDate1 & ";true"
.ParameterFields(2) = "@End;" & vDate2 & ";true"
.Destination = crptToWindow
.WindowState = crptMaximized
.Action = 1
End With
Else
MsgBox "กรุณาเลือก รหัสเจ้าหนี้ที่ต้องการดูรายงานด้วยครับ"
Exit Sub
End If


ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Form_Load()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String

On Error Resume Next

DTPicker101 = Now
DTPicker102 = Now

vQuery = "select code+'/'+name1 as apname from bcnp.dbo.bcap order by code"
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    While Not vRecordset.EOF
        CMB101.AddItem Trim(vRecordset.Fields("apname").Value)
    vRecordset.MoveNext
    Wend
End If
vRecordset.Close
End Sub
