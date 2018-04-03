VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form Form96 
   Caption         =   "รายงานจัดลำดับ"
   ClientHeight    =   8415
   ClientLeft      =   5415
   ClientTop       =   450
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form96.frx":0000
   ScaleHeight     =   8415
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport Crystal101 
      Left            =   1920
      Top             =   5880
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
   Begin VB.CommandButton Cmd101 
      Caption         =   "พิมพ์รายงาน"
      Height          =   465
      Left            =   3300
      TabIndex        =   3
      Top             =   3000
      Width           =   1665
   End
   Begin MSComCtl2.DTPicker DTPicker102 
      Height          =   315
      Left            =   2625
      TabIndex        =   2
      Top             =   2475
      Width           =   2340
      _ExtentX        =   4128
      _ExtentY        =   556
      _Version        =   393216
      Format          =   61800449
      CurrentDate     =   38327
   End
   Begin MSComCtl2.DTPicker DTPicker101 
      Height          =   315
      Left            =   2625
      TabIndex        =   1
      Top             =   2025
      Width           =   2340
      _ExtentX        =   4128
      _ExtentY        =   556
      _Version        =   393216
      Format          =   61800449
      CurrentDate     =   38327
   End
   Begin VB.ComboBox Cmb101 
      Height          =   315
      Left            =   2625
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1350
      Width           =   2340
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "ถึงวันที่"
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   1800
      TabIndex        =   6
      Top             =   2475
      Width           =   690
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "จากวันที่"
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   1800
      TabIndex        =   5
      Top             =   2025
      Width           =   765
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ประเภทของรายงาน"
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   1125
      TabIndex        =   4
      Top             =   1350
      Width           =   1440
   End
End
Attribute VB_Name = "Form96"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub CMD101_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vRepType As String
Dim vRepID As Integer
Dim vReportName As String
Dim vDate1 As Date, vDate2 As Date


If CMB101.Text <> "" Then
vDate1 = DTPicker101.Day & "/" & DTPicker101.Month & "/" & DTPicker101.Year
vDate2 = DTPicker102.Day & "/" & DTPicker102.Month & "/" & DTPicker102.Year
Select Case CMB101.ListIndex
Case 0
    vRepType = "AP"
    vRepID = 189
Case 1
    vRepType = "AR"
    vRepID = 190
Case 2
    vRepType = "AP"
    vRepID = 191
Case 3
    vRepType = "AR"
        vRepID = 192
End Select

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
'vQuery = "select reportname from bcnp.dbo.bcreportname where reptype = '" & vRepType & "' and repid = " & vRepID & " "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vReportName = Trim(vRecordset.Fields("reportname").Value)
End If
vRecordset.Close

Else
MsgBox "กรุณาเลือกประเภท รายงาน"
Exit Sub
End If

With Crystal101
.ReportFileName = vReportName & ".rpt"
.ParameterFields(0) = "@StartDate;" & vDate1 & ";true"
.ParameterFields(1) = "@EndDate;" & vDate2 & ";true"
.Destination = crptToWindow
.WindowState = crptMaximized
.Action = 1
End With

End Sub

Private Sub Form_Load()
CMB101.AddItem Trim("รายงานจัดลำดับยอดซื้อ")
CMB101.AddItem Trim("รายงานจัดลำดับยอดขาย")
CMB101.AddItem Trim("รายงานจัดลำดับยอดจ่าย")
CMB101.AddItem Trim("รายงานจัดลำดับยอดรับ")
End Sub





