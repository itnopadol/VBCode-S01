VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form Form812 
   Caption         =   "รายงาน การพิมพ์ใบจ่ายและใบหยิบสินค้า"
   ClientHeight    =   8190
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form812.frx":0000
   ScaleHeight     =   8190
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport Crystal101 
      Left            =   2280
      Top             =   5880
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
      Left            =   3075
      TabIndex        =   6
      Top             =   1200
      Width           =   5040
   End
   Begin MSComCtl2.DTPicker DTPicker102 
      Height          =   315
      Left            =   3075
      TabIndex        =   4
      Top             =   2550
      Width           =   1740
      _ExtentX        =   3069
      _ExtentY        =   556
      _Version        =   393216
      Format          =   64552961
      CurrentDate     =   38548
   End
   Begin MSComCtl2.DTPicker DTPicker101 
      Height          =   315
      Left            =   3075
      TabIndex        =   3
      Top             =   1875
      Width           =   1740
      _ExtentX        =   3069
      _ExtentY        =   556
      _Version        =   393216
      Format          =   64552961
      CurrentDate     =   38548
   End
   Begin VB.CommandButton CMD101 
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
      Left            =   3600
      TabIndex        =   2
      Top             =   3150
      Width           =   1215
   End
   Begin VB.Label Label3 
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
      Left            =   1800
      TabIndex        =   5
      Top             =   1200
      Width           =   1290
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "ถึงวันที่"
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
      Left            =   2400
      TabIndex        =   1
      Top             =   2550
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "จากวันที่"
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
      Left            =   2325
      TabIndex        =   0
      Top             =   1875
      Width           =   690
   End
End
Attribute VB_Name = "Form812"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMD101_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vRepID As Integer
Dim vRepType As String
Dim vReportName As String
Dim vDate1 As Date, vDate2 As Date

On Error GoTo ErrDescription

If CMB101.Text <> "" Then
    If CMB101.ListIndex = 0 Then
        vRepID = 240
    ElseIf CMB101.ListIndex = 1 Then
        vRepID = 239
    ElseIf CMB101.ListIndex = 2 Then
        vRepID = 241
    ElseIf CMB101.ListIndex = 3 Then
        vRepID = 242
    ElseIf CMB101.ListIndex = 4 Then
        vRepID = 243
    ElseIf CMB101.ListIndex = 5 Then
        vRepID = 253
    End If
    vRepType = "SO"
    
    vDate1 = Trim(DTPicker101.Day & "/" & DTPicker101.Month & "/" & DTPicker101.Year)
    vDate2 = Trim(DTPicker102.Day & "/" & DTPicker102.Month & "/" & DTPicker102.Year)
    
    vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
    'vQuery = "select reportname from bcnp.dbo.bcreportname where repid = " & vRepID & " and reptype = '" & vRepType & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vReportName = Trim(vRecordset.Fields("reportname").Value)
    End If
    vRecordset.Close
     
     With Crystal101
     .ReportFileName = vReportName & ".rpt"
     .ParameterFields(0) = "@vStartDate;" & vDate1 & ";true"
     .ParameterFields(1) = "@vEndDate;" & vDate2 & ";true"
     .Destination = crptToWindow
     .WindowState = crptMaximized
     .Action = 1
     End With
Else
    MsgBox "กรุณา เลือกประเภทรายงานสินค้าด้วยครับ"
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub

Private Sub Form_Load()
On Error GoTo ErrDescription

DTPicker101.Value = Now
DTPicker102.Value = Now
CMB101.Clear
CMB101.AddItem Trim("รายงานการพิมพ์ใบหยิบสินค้า การขนส่งรับเอง")
CMB101.AddItem Trim("รายงาน การพิมพ์ใบหยิบสินค้า การขนส่งส่งให้")
CMB101.AddItem Trim("รายงาน เวลาการจัดสินค้าจากใบหยิบ ตามพนักงาน")
CMB101.AddItem Trim("รายงาน เวลาการจัดสินค้าจากใบหยิบ ตามโซนสินค้า")
CMB101.AddItem Trim("รายงาน การพิมพ์ใบจ่ายสินค้า ตามชื่อผู้พิมพ์")
CMB101.AddItem Trim("รายงาน การพิมพ์ใบหยิบสินค้าแบบรายละเอียด ตามพนักงานขาย")

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub
