VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "CRYSTL32.OCX"
Begin VB.Form Form393 
   Caption         =   "รายงาน สรุปบิลขายลดราคาสินค้า"
   ClientHeight    =   8430
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form393.frx":0000
   ScaleHeight     =   8430
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport Crystal102 
      Left            =   3150
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
   Begin Crystal.CrystalReport Crystal101 
      Left            =   1440
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
   Begin VB.ComboBox CMB102 
      Height          =   315
      Left            =   3000
      TabIndex        =   6
      Text            =   "ขายส่ง"
      Top             =   2175
      Width           =   1740
   End
   Begin VB.ComboBox CMB101 
      Height          =   315
      Left            =   3000
      TabIndex        =   4
      Text            =   "รายงานสรุปบิลขายลดราคาสินค้ารายวันแบบแสดงรายละเอียด"
      Top             =   1425
      Width           =   4215
   End
   Begin VB.CommandButton CMD101 
      Caption         =   "พิมพ์รายงาน"
      Height          =   390
      Left            =   3600
      TabIndex        =   2
      Top             =   3330
      Width           =   1140
   End
   Begin MSComCtl2.DTPicker DTPicker101 
      Height          =   315
      Left            =   3000
      TabIndex        =   1
      Top             =   2700
      Width           =   1740
      _ExtentX        =   3069
      _ExtentY        =   556
      _Version        =   393216
      Format          =   69533697
      CurrentDate     =   38527
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ประเภทการขาย"
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
      Height          =   240
      Left            =   1575
      TabIndex        =   5
      Top             =   2250
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ประเภทรายงาน"
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
      Height          =   315
      Left            =   1500
      TabIndex        =   3
      Top             =   1425
      Width           =   1290
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ณ วันที่"
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
      Height          =   240
      Left            =   1800
      TabIndex        =   0
      Top             =   2700
      Width           =   990
   End
End
Attribute VB_Name = "Form393"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMD101_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vReportName As String
Dim vDate1 As Date
Dim vSaleType As Integer
Dim vRepID As Integer
Dim vRepType As String

On Error GoTo ErrDescription

If CMB101.Text <> "" And CMB102.Text <> "" Then
If DTPicker101.Year = Year(Now) + 543 Then
vDate1 = DTPicker101.Day & "/" & DTPicker101.Month & "/" & DTPicker101.Year - 543
Else
vDate1 = DTPicker101.Day & "/" & DTPicker101.Month & "/" & DTPicker101.Year
End If

If CMB102.Text = Trim("ขายส่ง") Then
    vSaleType = 0
ElseIf CMB102.Text = Trim("ขายปลีก") Then
    vSaleType = 1
End If

If CMB101.Text = Trim("รายงานสรุปบิลขายลดราคาสินค้ารายวันแบบแสดงรายละเอียด") Then
vRepID = 235
vRepType = "AR"
vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
    'vQuery = "select reportname from bcreportname where repid = 235 and reptype = 'AR' "
ElseIf CMB101.Text = Trim("รายงานสรุปบิลขายลดราคาสินค้ารายวันแบบแสดงรายสรุป") Then
vRepID = 236
vRepType = "AR"
vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
    'vQuery = "select reportname from bcreportname where repid = 236 and reptype = 'AR' "
ElseIf CMB101.Text = Trim("รายงานสรุปบิลขายลดราคาสินค้ารายเดือน") Then
    Call PrintDiscountItemMonth
    Exit Sub
ElseIf CMB101.Text = Trim("รายงานลดราคาใบสั่งขายสั่งจองประจำวัน") Then
vRepID = 303
vRepType = "AR"
vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
    'vQuery = "select reportname from bcreportname where repid = 303 and reptype = 'AR' "
ElseIf CMB101.Text = Trim("รายงานสรุปลดราคาใบสั่งขายสั่งจองประจำวัน") Then
vRepID = 304
vRepType = "AR"
vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
    'vQuery = "select reportname from bcreportname where repid = 304 and reptype = 'AR' "
End If
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vReportName = Trim(vRecordset.Fields("reportname").Value)
End If
vRecordset.Close

With Crystal101
.ReportFileName = Trim(vReportName & ".rpt")
.ParameterFields(0) = "@Date;" & vDate1 & ";true"
.ParameterFields(1) = "@IsWholeSale;" & vSaleType & ";true"
.Destination = crptToWindow
.WindowState = crptMaximized
.Action = 1
End With
Else
    MsgBox "กรุณาเลือก ประเภทของรายงาน และ ประเภทการขาย ด้วยนะครับ"
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

CMB101.AddItem Trim("รายงานลดราคาใบสั่งขายสั่งจองประจำวัน")
CMB101.AddItem Trim("รายงานสรุปลดราคาใบสั่งขายสั่งจองประจำวัน")
CMB101.AddItem Trim("รายงานสรุปบิลขายลดราคาสินค้ารายวันแบบแสดงรายละเอียด")
CMB101.AddItem Trim("รายงานสรุปบิลขายลดราคาสินค้ารายวันแบบแสดงรายสรุป")
CMB101.AddItem Trim("รายงานสรุปบิลขายลดราคาสินค้ารายเดือน")
DTPicker101.Value = Now

CMB102.AddItem Trim("ขายส่ง")
CMB102.AddItem Trim("ขายปลีก")
End Sub

Public Sub PrintDiscountItemMonth()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vReportName As String
Dim vYear  As String
Dim vMonth As String
Dim vSaleType As Integer
Dim vRepID As Integer
Dim vRepType As String

vRepID = 295
vRepType = "INV"
vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
'vQuery = "select reportname from bcnp.dbo.bcreportname where repid = 295 and reptype = 'INV' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vReportName = Trim(vRecordset.Fields("reportname").Value)
End If
vRecordset.Close

vYear = DTPicker101.Year
vMonth = DTPicker101.Month
If CMB102.Text = Trim("ขายส่ง") Then
    vSaleType = 0
ElseIf CMB102.Text = Trim("ขายปลีก") Then
    vSaleType = 1
End If

With Crystal102
.ReportFileName = Trim(vReportName) & ".rpt"
.ParameterFields(0) = "@Year;" & vYear & ";true"
.ParameterFields(1) = "@Month;" & vMonth & ";true"
.ParameterFields(2) = "@IsWholeSale;" & vSaleType & ";true"
.Destination = crptToWindow
.WindowState = crptMaximized
.Action = 1
End With
End Sub
