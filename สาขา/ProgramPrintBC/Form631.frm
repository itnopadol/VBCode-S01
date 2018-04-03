VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "CRYSTL32.OCX"
Begin VB.Form Form631 
   Caption         =   "รายงาน ประวัติเช็คคืน"
   ClientHeight    =   8385
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form631.frx":0000
   ScaleHeight     =   8385
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CrystalReport631 
      Left            =   1440
      Top             =   6240
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
   Begin VB.ComboBox CMBChqStatus 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2625
      TabIndex        =   10
      Top             =   2400
      Visible         =   0   'False
      Width           =   4365
   End
   Begin VB.ComboBox CMBReportCondition 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2625
      TabIndex        =   8
      Top             =   1650
      Width           =   4365
   End
   Begin VB.CommandButton CMD6311 
      Caption         =   "พิมพ์รายงาน"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5025
      TabIndex        =   7
      Top             =   4125
      Width           =   1965
   End
   Begin MSComCtl2.DTPicker DTP6312 
      Height          =   465
      Left            =   2625
      TabIndex        =   4
      Top             =   3150
      Visible         =   0   'False
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   820
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   69599233
      CurrentDate     =   38142
   End
   Begin MSComCtl2.DTPicker DTP6311 
      Height          =   465
      Left            =   2625
      TabIndex        =   3
      Top             =   2400
      Visible         =   0   'False
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   820
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   69599233
      CurrentDate     =   38142
   End
   Begin VB.ComboBox CMBArcode 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2625
      TabIndex        =   1
      Top             =   2400
      Visible         =   0   'False
      Width           =   4365
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "สถานะเช็ค"
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
      Left            =   1875
      TabIndex        =   11
      Top             =   2400
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "รูปแบบรายงาน"
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
      Left            =   1500
      TabIndex        =   9
      Top             =   1725
      Width           =   1065
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "ถึงวันที่"
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
      Height          =   390
      Left            =   1875
      TabIndex        =   6
      Top             =   3225
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "จากวันที่"
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
      Height          =   390
      Left            =   1875
      TabIndex        =   5
      Top             =   2400
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "รหัสลูกค้า"
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
      Height          =   390
      Left            =   1875
      TabIndex        =   2
      Top             =   2400
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "รายงาน ประวัติเช็คคืน"
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
      Height          =   540
      Left            =   2625
      TabIndex        =   0
      Top             =   225
      Width           =   7365
   End
End
Attribute VB_Name = "Form631"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMBReportCondition_Click()

If Trim(CMBReportCondition.Text) = "ดูรายงานทั้งหมด" Then
    CMBArcode.Visible = False
    Label2.Visible = False
    Label3.Visible = False
    Label4.Visible = False
    Label6.Visible = False
    CMBChqStatus.Visible = False
    DTP6311.Visible = False
    DTP6312.Visible = False
ElseIf Trim(CMBReportCondition.Text) = "ดูรายงานตาม รหัสลูกค้า" Then
    CMBArcode.Visible = True
    Label2.Visible = True
    Label3.Visible = False
    Label4.Visible = False
    DTP6311.Visible = False
    DTP6312.Visible = False
    Label6.Visible = False
    CMBChqStatus.Visible = False
ElseIf Trim(CMBReportCondition.Text) = "ดูรายงานตาม วันที่รับคืนเช็ค" Then
    CMBArcode.Visible = False
    Label2.Visible = False
    Label3.Visible = True
    Label4.Visible = True
    DTP6311.Visible = True
    DTP6312.Visible = True
    Label6.Visible = False
    CMBChqStatus.Visible = False
ElseIf Trim(CMBReportCondition.Text) = "ดูรายงานตาม วันที่ครบกำหนดเช็ค" Then
    CMBArcode.Visible = False
    Label2.Visible = False
    Label3.Visible = True
    Label4.Visible = True
    DTP6311.Visible = True
    DTP6312.Visible = True
    Label6.Visible = False
    CMBChqStatus.Visible = False
ElseIf Trim(CMBReportCondition.Text) = "ดูรายงานตาม สถานะเช็ค" Then
    CMBArcode.Visible = False
    Label2.Visible = False
    Label3.Visible = False
    Label4.Visible = False
    DTP6311.Visible = False
    DTP6312.Visible = False
    Label6.Visible = True
    CMBChqStatus.Visible = True
End If
End Sub

Private Sub CMD6311_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String, vRepType As String
Dim Date1 As Date, Date2 As Date
Dim vReportName As String, vARCode As String
Dim vRepID As Integer, StrCount As Integer

On Error GoTo ErrDescription

If Trim(CMBReportCondition.Text) = "ดูรายงานทั้งหมด" Then
Call Print_ChqReturn_All
ElseIf Trim(CMBReportCondition.Text) = "ดูรายงานตาม รหัสลูกค้า" Then
Call Print_ChqReturn_ArCode
ElseIf Trim(CMBReportCondition.Text) = "ดูรายงานตาม วันที่รับคืนเช็ค" Then
Call Print_ChqReturn_RetDate
ElseIf Trim(CMBReportCondition.Text) = "ดูรายงานตาม วันที่ครบกำหนดเช็ค" Then
Call Print_ChqReturn_ChqDueDate
ElseIf Trim(CMBReportCondition.Text) = "ดูรายงานตาม สถานะเช็ค" Then
Call Print_ChqReturn_ChqStatus
Else
MsgBox "กรุณาเลือกการดูรายงานด้วยนะครับ", vbCritical, "ข้อความเตือน"
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description, vbCritical, "ข้อความเตือน"
Exit Sub
End If
End Sub

Private Sub Form_Load()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String

On Error Resume Next

vQuery = "select distinct arcode+' = '+name1 as arname from TB_AR_ChqinRetHist a left outer join bcar b on a.arcode = b.code "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    While Not vRecordset.EOF
    CMBArcode.AddItem Trim(vRecordset.Fields("arname").Value)
    vRecordset.MoveNext
    Wend
End If
vRecordset.Close

CMBReportCondition.AddItem "ดูรายงานทั้งหมด"
CMBReportCondition.AddItem "ดูรายงานตาม รหัสลูกค้า"
CMBReportCondition.AddItem "ดูรายงานตาม วันที่รับคืนเช็ค"
CMBReportCondition.AddItem "ดูรายงานตาม วันที่ครบกำหนดเช็ค"
CMBReportCondition.AddItem "ดูรายงานตาม สถานะเช็ค"

CMBChqStatus.AddItem Trim("0 = เช็คในมือ")
CMBChqStatus.AddItem Trim("1 = เช็คฝาก")
CMBChqStatus.AddItem Trim("2 = เช็คผ่าน")
CMBChqStatus.AddItem Trim("3 = เช็คคืน")
CMBChqStatus.AddItem Trim("4 = เช็คยกเลิก")
CMBChqStatus.AddItem Trim("5 = ขายลดเช็ค")
CMBChqStatus.AddItem Trim("6 = เช็คคืนฝากธนาคาร")
End Sub

Private Sub Option1_Click()
DTP6311.Visible = False
DTP6312.Visible = False
Label3.Visible = False
Label4.Visible = False
End Sub

Private Sub Option2_Click()
DTP6311.Visible = True
DTP6312.Visible = True
Label3.Visible = True
Label4.Visible = True
End Sub

Public Function Print_ChqReturn_All()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String, vRepType As String
Dim vReportName As String
Dim vRepID As Integer

vRepType = "CHQ"
vRepID = 137
vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
'vQuery = "select reportname from bcreportname where reptype = '" & vRepType & "' and repid = " & vRepID & " "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
With CrystalReport631
.ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
.Destination = crptToWindow
.WindowState = crptMaximized
.Action = 1
End With
End If
vRecordset.Close

End Function

Public Function Print_ChqReturn_ArCode()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String, vRepType As String
Dim Date1 As Date, Date2 As Date
Dim vReportName As String, vARCode As String
Dim vRepID As Integer, StrCount As Integer

vRepType = "CHQ"
vRepID = 138
StrCount = InStr(Trim(CMBArcode.Text), "=")
vARCode = Trim(Left(CMBArcode.Text, StrCount - 1))

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
'vQuery = "select reportname from bcreportname where reptype = '" & vRepType & "' and repid = " & vRepID & " "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
With CrystalReport631
.ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
.ParameterFields(0) = "ArCode;" & vARCode & ";true"
.Destination = crptToWindow
.WindowState = crptMaximized
.Action = 1
End With
End If
vRecordset.Close
End Function

Public Function Print_ChqReturn_RetDate()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String, vRepType As String
Dim Date1 As Date, Date2 As Date
Dim vReportName As String
Dim vRepID As Integer

Date1 = DTP6311.Day & "/" & DTP6311.Month & "/" & DTP6311.Year
Date2 = DTP6312.Day & "/" & DTP6312.Month & "/" & DTP6312.Year
vRepType = "CHQ"
vRepID = 139

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
'vQuery = "select reportname from bcreportname where reptype = '" & vRepType & "' and repid = " & vRepID & " "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
With CrystalReport631
.ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
.ParameterFields(0) = "@Date1;" & Date1 & ";true"
.ParameterFields(1) = "@Date2;" & Date2 & ";true"
.Destination = crptToWindow
.WindowState = crptMaximized
.Action = 1
End With
End If
vRecordset.Close
End Function

Public Function Print_ChqReturn_ChqDueDate()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String, vRepType As String
Dim Date1 As Date, Date2 As Date
Dim vReportName As String
Dim vRepID As Integer

Date1 = DTP6311.Day & "/" & DTP6311.Month & "/" & DTP6311.Year
Date2 = DTP6312.Day & "/" & DTP6312.Month & "/" & DTP6312.Year
vRepType = "CHQ"
vRepID = 140

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
'vQuery = "select reportname from bcreportname where reptype = '" & vRepType & "' and repid = " & vRepID & " "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
With CrystalReport631
.ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
.ParameterFields(0) = "@Date1;" & Date1 & ";true"
.ParameterFields(1) = "@Date2;" & Date2 & ";true"
.Destination = crptToWindow
.WindowState = crptMaximized
.Action = 1
End With
End If
vRecordset.Close
End Function

Public Function Print_ChqReturn_ChqStatus()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String, vRepType As String, vChqStatus As String
Dim vReportName As String
Dim vRepID As Integer, StrCount As Integer

vRepType = "CHQ"
vRepID = 141
StrCount = InStr(Trim(CMBChqStatus.Text), "=")
vChqStatus = Trim(Left(CMBChqStatus.Text, StrCount - 1))

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
'vQuery = "select reportname from bcreportname where reptype = '" & vRepType & "' and repid = " & vRepID & " "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
With CrystalReport631
.ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
.ParameterFields(0) = "@ChqStatus;" & vChqStatus & ";true"
.Destination = crptToWindow
.WindowState = crptMaximized
.Action = 1
End With
End If
vRecordset.Close
End Function
