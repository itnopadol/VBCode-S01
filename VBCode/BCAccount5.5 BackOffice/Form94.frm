VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "CRYSTL32.OCX"
Begin VB.Form Form94 
   Caption         =   "รายงาน สรุปค่าใช้จ่าย"
   ClientHeight    =   8955
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11985
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form94.frx":0000
   ScaleHeight     =   8955
   ScaleWidth      =   11985
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport Crystal2 
      Left            =   4680
      Top             =   6120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin Crystal.CrystalReport Crystal1 
      Left            =   1935
      Top             =   5850
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.CommandButton CMD941 
      Caption         =   "พิมพ์รายงาน"
      Height          =   615
      Left            =   5550
      TabIndex        =   5
      Top             =   3525
      Width           =   1665
   End
   Begin VB.ComboBox CMB943 
      Enabled         =   0   'False
      Height          =   315
      Left            =   4275
      TabIndex        =   4
      Top             =   2925
      Width           =   2940
   End
   Begin VB.ComboBox CMBMonth2 
      Height          =   315
      Left            =   4275
      TabIndex        =   3
      Top             =   2475
      Width           =   2940
   End
   Begin VB.ComboBox CMBMonth1 
      Height          =   315
      Left            =   4275
      TabIndex        =   2
      Top             =   2025
      Width           =   2940
   End
   Begin VB.ComboBox CMB942 
      Height          =   315
      Left            =   4275
      TabIndex        =   1
      Top             =   1200
      Width           =   2940
   End
   Begin VB.ComboBox CMB941 
      Height          =   315
      Left            =   1125
      TabIndex        =   0
      Top             =   1200
      Width           =   1740
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ถึงเดือน"
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   3000
      TabIndex        =   9
      Top             =   2475
      Width           =   1065
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "จากเดือน"
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   3000
      TabIndex        =   8
      Top             =   2025
      Width           =   1065
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "ประเภทรายงาน"
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   3000
      TabIndex        =   7
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ปี"
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   150
      TabIndex        =   6
      Top             =   1200
      Width           =   690
   End
End
Attribute VB_Name = "Form94"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMD941_Click()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vYear As Integer, vMonth1 As Integer, vMonth2 As Integer
Dim vRepType As String
Dim vRepID As String, vReportName As String
Dim vDetails As String

On Error GoTo Errdescription

If CMB942.Text = "สรุปค่าใช้จ่ายประจำปีแบบรวม" Then
    Call PrintPayMent_Total
ElseIf CMB942.Text = "สรุปค่าใช้จ่ายประจำปีแบบแยก" Then
    Call PrintPayMent_Depart
Else
    Call PrintPayMent_Depart2
End If

Errdescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Form_Load()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset

CMB942.AddItem "สรุปค่าใช้จ่ายประจำปีแบบรวม"
'CMB942.AddItem "สรุปค่าใช้จ่ายประจำปีแบบแยก"
'CMB942.AddItem "สรุปค่าใช้จ่ายประจำปีแบบแยกแผนก"

vQuery = "select distinct year(docdate) as PaymentYear from dbo.BCPAYMENT order by paymentyear desc "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    While Not vRecordset.EOF
        CMB941.AddItem Trim(vRecordset.Fields("paymentyear").Value)
        vRecordset.MoveNext
    Wend
End If
vRecordset.Close

vQuery = "select code+'    '+name as name from dbo.BCdepartment order by code  "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    While Not vRecordset.EOF
        CMB943.AddItem Trim(vRecordset.Fields("name").Value)
        vRecordset.MoveNext
    Wend
End If
vRecordset.Close

vQuery = "select distinct month(docdate) as PaymentMonth from dbo.BCPAYMENT order by paymentMonth  "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    While Not vRecordset.EOF
        CMBMonth1.AddItem Trim(vRecordset.Fields("PaymentMonth").Value)
        CMBMonth2.AddItem Trim(vRecordset.Fields("PaymentMonth").Value)
        vRecordset.MoveNext
    Wend
End If
vRecordset.Close
End Sub

Public Sub PrintPayMent_Total()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vYear As Integer, vMonth1 As Integer, vMonth2 As Integer
Dim vRepType As String
Dim vRepID As String, vReportName As String
Dim vDetails As String

On Error GoTo Errdescription

    vRepType = "GL"
    vRepID = 20
    vMonth1 = Trim(CMBMonth1.Text)
    vMonth2 = Trim(CMBMonth2.Text)
    vYear = CMB941.Text

vQuery = "select reportname from bcvat.dbo.bcreportname where repid = '" & vRepID & "' and reptype = '" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vReportName = Trim(vRecordset.Fields("reportname").Value)
End If
vRecordset.Close

With Crystal2
    .ReportFileName = Trim(vReportName) & ".rpt"
    .ParameterFields(0) = "@year;" & vYear & ";true"
    .ParameterFields(1) = "@Mb;" & vMonth1 & ";true"
    .ParameterFields(2) = "@Me;" & vMonth2 & ";true"
    Call ReportSetLocation(Crystal2)
    .WindowState = crptMaximized
    .Destination = crptToWindow
    .Action = 1
End With

Errdescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Public Sub PrintPayMent_Depart()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vYear As Integer, vMonth1 As Integer, vMonth2 As Integer
Dim vRepType As String
Dim vRepID As String, vReportName As String
Dim vDetails As String

On Error GoTo Errdescription

    vRepType = "GL"
    vRepID = 21
    vDetails = Trim(Left(CMB943.Text, 2))
    vYear = CMB941.Text

vQuery = "select reportname from bcvat.dbo.bcreportname where repid = '" & vRepID & "' and reptype = '" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vReportName = Trim(vRecordset.Fields("reportname").Value)
End If
vRecordset.Close

With Crystal1
    .ReportFileName = Trim(vReportName) & ".rpt"
    .ParameterFields(0) = "@year;" & vYear & ";true"
    .ParameterFields(1) = "@DepartCode;" & vDetails & ";true"
    .WindowState = crptMaximized
    .Destination = crptToWindow
    .Action = 1
End With

Errdescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Public Sub PrintPayMent_Depart2()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vFDate As Integer, vLDate As Integer, vFYear As Integer
Dim vRepType As String
Dim vRepID As String, vReportName As String
Dim vDetails As String

On Error GoTo Errdescription

    vRepType = "GL"
    vRepID = 22
    vFDate = CMBMonth1.Text
    vLDate = CMBMonth2.Text
    vFYear = CMB941.Text

vQuery = "select reportname from bcvat.dbo.bcreportname where repid = '" & vRepID & "' and reptype = '" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vReportName = Trim(vRecordset.Fields("reportname").Value)
End If
vRecordset.Close

With Crystal1
    .ReportFileName = Trim(vReportName) & ".rpt"
    .ParameterFields(0) = "@FDate;" & vFDate & ";true"
    .ParameterFields(1) = "@LDate;" & vLDate & ";true"
    .ParameterFields(2) = "@FYear;" & vFYear & ";true"
    .WindowState = crptMaximized
    .Destination = crptToWindow
    .Action = 1
End With

Errdescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub



