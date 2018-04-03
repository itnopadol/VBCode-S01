VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form Form412 
   Caption         =   "รายงานเคลื่อนไหวเจ้าหนี้"
   ClientHeight    =   8985
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   Picture         =   "Form412.frx":0000
   ScaleHeight     =   8985
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CrystalReport412 
      Left            =   3735
      Top             =   5895
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
   Begin Crystal.CrystalReport CrystalReport411 
      Left            =   1800
      Top             =   5895
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
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   315
      Left            =   2550
      TabIndex        =   5
      Top             =   2925
      Width           =   2640
      _ExtentX        =   4657
      _ExtentY        =   556
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
      Format          =   16646145
      CurrentDate     =   38139
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   2550
      TabIndex        =   4
      Top             =   2400
      Width           =   2640
      _ExtentX        =   4657
      _ExtentY        =   556
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
      Format          =   16646145
      CurrentDate     =   38139
   End
   Begin VB.CommandButton CMD4121 
      Caption         =   "ดูรายงาน"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   3750
      TabIndex        =   1
      Top             =   3600
      Width           =   1440
   End
   Begin VB.TextBox TXT4121 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   2550
      TabIndex        =   0
      Top             =   1500
      Width           =   2640
   End
   Begin VB.Label Label3 
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
      Height          =   315
      Left            =   1575
      TabIndex        =   7
      Top             =   2925
      Width           =   915
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "ตั้งแต่วันที่"
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
      Left            =   1575
      TabIndex        =   6
      Top             =   2400
      Width           =   915
   End
   Begin VB.Label LBL4121 
      BackStyle       =   0  'Transparent
      Caption         =   "รหัสเจ้าหนี้"
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
      Left            =   1575
      TabIndex        =   3
      Top             =   1500
      Width           =   1065
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "รายงานเคลื่อนไหวเจ้าหนี้"
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
      Left            =   2550
      TabIndex        =   2
      Top             =   225
      Width           =   7515
   End
End
Attribute VB_Name = "Form412"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMD4121_Click()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vApCode As String, vRepType As String, vReportName As String
Dim vRepID As Integer
Dim Date1 As Date, Date2 As Date

vApCode = Trim(TXT4121.Text)
If vApCode <> "" Then
    vRepType = "AP"
    vRepID = 4
    
    Date1 = DTPicker1.Day & "/" & DTPicker1.Month & "/" & DTPicker1.Year
    Date2 = DTPicker2.Day & "/" & DTPicker2.Month & "/" & DTPicker2.Year
    vQuery = "select reportname from bcreportname where repid = " & vRepID & "  and reptype = '" & vRepType & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vReportName = Trim(vRecordset.Fields("reportname").Value)
        With CrystalReport412
        .ReportFileName = vReportName & ".rpt"
        .ParameterFields(0) = "@ApCode;" & vApCode & ";true"
        .ParameterFields(1) = "@Date1;" & Date1 & ";true"
        .ParameterFields(2) = "@Date2;" & Date2 & ";true"
        .WindowState = crptMaximized
        .Destination = crptToWindow
        .Action = 1
        End With
    
    End If
    vRecordset.Close
Else
    MsgBox "กรุณาใส่ข้อมูลดูรายงานให้ครบด้วยครับ", vbInformation + vbCritical, "ข้อความเตือน"
End If
End Sub


Private Sub Form_Load()
Me.DTPicker1.Value = Now
Me.DTPicker2.Value = Now
End Sub

Private Sub TXT4121_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    CMD4121_Click
End If
End Sub
