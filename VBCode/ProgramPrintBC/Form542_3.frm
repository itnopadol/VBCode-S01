VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form Form542_3 
   Caption         =   "หน้าพิมพ์รายงานยอดลูกหนี้ประจำเดือนตามรหัสประเภทลูกค้า"
   ClientHeight    =   8355
   ClientLeft      =   2190
   ClientTop       =   690
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form542_3.frx":0000
   ScaleHeight     =   8355
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CMD101 
      Caption         =   "พิมพ์"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   3600
      TabIndex        =   8
      Top             =   4230
      Width           =   1995
   End
   Begin Crystal.CrystalReport Crystal101 
      Left            =   675
      Top             =   7155
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin MSComCtl2.DTPicker DTPicker101 
      Height          =   420
      Left            =   3600
      TabIndex        =   3
      Top             =   3510
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   741
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   51642369
      CurrentDate     =   38888
   End
   Begin VB.ComboBox CMBDebt102 
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
      Left            =   3600
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   2790
      Width           =   6990
   End
   Begin VB.ComboBox CMBDebt101 
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
      Left            =   3600
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   2115
      Width           =   6990
   End
   Begin VB.ComboBox CMBReportType 
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
      Left            =   3600
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1440
      Width           =   6990
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ณ วันที่ :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   420
      Left            =   1080
      TabIndex        =   7
      Top             =   3555
      Width           =   2445
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ถึงรหัสลูกหนี้ :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   420
      Left            =   585
      TabIndex        =   6
      Top             =   2835
      Width           =   2940
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "จากรหัสลูกหนี้ :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   420
      Left            =   720
      TabIndex        =   5
      Top             =   2160
      Width           =   2805
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ประเภทรายงาน :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   855
      TabIndex        =   4
      Top             =   1440
      Width           =   2670
   End
End
Attribute VB_Name = "Form542_3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMD101_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vReportType As Integer
Dim vBegDebt As String
Dim vEndDebt As String
Dim vReportName As String
Dim vRepType As String
Dim vRepID As Integer
Dim vAtDate As String

On Error GoTo ErrDescription

If CMBReportType.Text <> "" And CMBDebt101.Text <> "" And CMBDebt102.Text <> "" Then
  Select Case CMBReportType.ListIndex
  Case 0:
      vReportType = 0
  Case 1:
      vReportType = 1
  End Select
  
  vBegDebt = Left(Trim(Trim(CMBDebt101.Text)), InStr(1, Trim(CMBDebt101.Text), "//") - 1)
  vEndDebt = Left(Trim(Trim(CMBDebt102.Text)), InStr(1, Trim(CMBDebt102.Text), "//") - 1)
  vAtDate = CDate(DTPicker101.Day & "/" & DTPicker101.Month & "/" & DTPicker101.Year)
  
  vRepType = "AR"
  vRepID = 317
  
  vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
  'vQuery = "select reportname from dbo.bcreportname where reptype = '" & vRepType & "' and repid = " & vRepID & " "
  If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vReportName = Trim(vRecordset.Fields("reportname").Value)
  End If
  vRecordset.Close
  
  With Crystal101
  .ReportFileName = vReportName & ".rpt"
  .ParameterFields(0) = "@IsPresent;" & vReportType & ";true"
  .ParameterFields(1) = "@TypeCode1;" & vBegDebt & ";true"
  .ParameterFields(2) = "@TypeCode2;" & vEndDebt & ";true"
  .ParameterFields(3) = "@AtDate;" & vAtDate & ";true"
  .Destination = crptToWindow
  .WindowState = crptMaximized
  .Action = 1
  End With

Else
  MsgBox "กรุณาใส่เลือกข้อมูลดูรายงานให้ครบด้วย", vbCritical, "ข้อความเตือน"
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
Dim vGroupARItems As ListItem

On Error Resume Next

DTPicker101 = Now
vQuery = "select (code+'//'+name) as debtname from bcardebtgroup order by code"
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    While Not vRecordset.EOF
        CMBDebt101.AddItem Trim(vRecordset.Fields("debtname").Value)
        CMBDebt102.AddItem Trim(vRecordset.Fields("debtname").Value)
    vRecordset.MoveNext
    Wend
End If
vRecordset.Close


'vQuery = "select  distinct  isnull(pressmencode,'0000') as pressmencode ,isnull(name,'ไม่มีพนักงานเร่งรัด') as pressmenname from  dbo.bcar a left join dbo.bcsale b on a.pressmencode = b.code order by pressmenname"
'If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
 '   vRecordset.MoveFirst
  '  While Not vRecordset.EOF
   '     CMBPressMen.AddItem Trim(vRecordset.Fields("pressmenname").Value & "/" & vRecordset.Fields("pressmencode").Value)
    'vRecordset.MoveNext
    'Wend
'End If
'vRecordset.Close



CMBReportType.AddItem Trim("รายงานยอดลูกหนี้ที่รวมปีที่ไม่ยกยอด")
CMBReportType.AddItem Trim("รายงานยอดลูกหนี้รวมปีปัจจุบัน")

End Sub
