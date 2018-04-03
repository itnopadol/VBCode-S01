VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "CRYSTL32.OCX"
Begin VB.Form Form99 
   Caption         =   "รายงาน การจ่ายเงินประจำวัน"
   ClientHeight    =   7980
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form99.frx":0000
   ScaleHeight     =   7980
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CMD101 
      Caption         =   "พิมพ์"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7155
      TabIndex        =   3
      Top             =   2835
      Width           =   960
   End
   Begin Crystal.CrystalReport Crystal101 
      Left            =   945
      Top             =   6030
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
   Begin MSComCtl2.DTPicker DTPicker101 
      Height          =   285
      Left            =   4365
      TabIndex        =   2
      Top             =   1935
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   503
      _Version        =   393216
      Format          =   69795841
      CurrentDate     =   38826
   End
   Begin VB.ComboBox CMB101 
      Height          =   315
      Left            =   4365
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1350
      Width           =   3795
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ณ วันที่ :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   330
      Left            =   2970
      TabIndex        =   4
      Top             =   1935
      Width           =   1275
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ประเภทรายงาน :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   285
      Left            =   2250
      TabIndex        =   0
      Top             =   1350
      Width           =   1995
   End
End
Attribute VB_Name = "Form99"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMD101_Click()

If CMB101.ListIndex = 0 Then
    Call PrintPrivate
ElseIf CMB101.ListIndex = 1 Then
    Call PrintCompany
End If

End Sub

Private Sub Form_Load()
DTPicker101 = Now
CMB101.AddItem Trim("รายงาน การจ่ายเงินประจำวัน-ส่วนตัว")
CMB101.AddItem Trim("รายงาน การจ่ายเงินประจำวัน-บริษัท")
CMB101.Text = CMB101.List(0)
End Sub

Private Sub PrintCompany()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vRepType As String
Dim vReportName As String
Dim vRepID As Integer
Dim Date1 As Date

On Error GoTo ErrDescription

Date1 = DTPicker101.Day & "/" & DTPicker101.Month & "/" & DTPicker101.Year
vRepID = 311
vRepType = "PM"

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
'vQuery = "select reportname from bcreportname where repid = " & vRepID & " and reptype = '" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vReportName = Trim(vRecordset.Fields("reportname").Value)
    With Crystal101
        .ReportFileName = vReportName & ".rpt"
        .ParameterFields(0) = "@vCreateDate1;" & Date1 & ";true"
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .Action = 1
    End With
End If
vRecordset.Close

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub PrintPrivate()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vRepType As String
Dim vReportName As String
Dim vRepID As Integer
Dim Date1 As Date

On Error GoTo ErrDescription

Date1 = DTPicker101.Day & "/" & DTPicker101.Month & "/" & DTPicker101.Year
vRepID = 312
vRepType = "PM"

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
'vQuery = "select reportname from bcreportname where repid = " & vRepID & " and reptype = '" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vReportName = Trim(vRecordset.Fields("reportname").Value)
    With Crystal101
        .ReportFileName = vReportName & ".rpt"
        .ParameterFields(0) = "@vCreateDate1;" & Date1 & ";true"
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .Action = 1
    End With
End If
vRecordset.Close

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

