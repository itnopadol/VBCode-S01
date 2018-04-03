VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "CRYSTL32.OCX"
Begin VB.Form Form811 
   Caption         =   "พิมพ์รายงาน การเปลี่ยนรหัสสินค้าและราคาสินค้า"
   ClientHeight    =   7980
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form811.frx":0000
   ScaleHeight     =   7980
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport Crystal101 
      Left            =   2280
      Top             =   5400
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
   Begin VB.ComboBox CMB101 
      Height          =   315
      Left            =   2400
      TabIndex        =   2
      Top             =   1575
      Width           =   2565
   End
   Begin VB.CommandButton CMD101 
      Caption         =   "พิมพ์รายงาน"
      Height          =   540
      Left            =   3375
      TabIndex        =   1
      Top             =   2925
      Width           =   1590
   End
   Begin MSComCtl2.DTPicker DTPicker101 
      Height          =   315
      Left            =   2400
      TabIndex        =   0
      Top             =   2175
      Width           =   2565
      _ExtentX        =   4524
      _ExtentY        =   556
      _Version        =   393216
      Format          =   69795841
      CurrentDate     =   38425
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ณ วันที่"
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   1200
      TabIndex        =   4
      Top             =   2175
      Width           =   1140
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ประเภทรายงาน"
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   1125
      TabIndex        =   3
      Top             =   1575
      Width           =   1215
   End
End
Attribute VB_Name = "Form811"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMD101_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vRepID As Integer
Dim vRepType As String, vReportName As String
Dim vDate1  As Date

On Error GoTo ErrDescription

If Trim(CMB101.Text) <> "" Then
    If CMB101.Text = Trim("รายงาน การเปลี่ยนรหัสสินค้า") Then
        vRepID = 208
        vRepType = "IV"
    ElseIf CMB101.Text = Trim("รายงาน การเปลี่ยนราคาสินค้า") Then
        vRepID = 209
        vRepType = "IV"
    End If
    vDate1 = DTPicker101.Day & "/" & DTPicker101.Month & "/" & DTPicker101.Year
    
    vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
    'vQuery = "select reportname from bcnp.dbo.bcreportname where repid = " & vRepID & " and reptype = '" & vRepType & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vReportName = Trim(vRecordset.Fields("reportname").Value)
    End If
    vRecordset.Close
    
    With Crystal101
    .ReportFileName = vReportName & ".rpt"
    .ParameterFields(0) = "@Date1;" & vDate1 & ";true"
    .Destination = crptToWindow
    .WindowState = crptMaximized
    .Action = 1
    End With
Else
    MsgBox "ยังไม่ได้เลือกประเภทของรายงาน กรุณาเลือกประเภทของรายงานด้วยครับ"
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub

Private Sub Form_Load()
CMB101.AddItem Trim("รายงาน การเปลี่ยนรหัสสินค้า")
CMB101.AddItem Trim("รายงาน การเปลี่ยนราคาสินค้า")
DTPicker101.Value = Now
End Sub
