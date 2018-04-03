VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "CRYSTL32.OCX"
Begin VB.Form Form40_3 
   Caption         =   "หน้าเคลื่อนไหวเจ้าหนี้ - ทั้งหมด"
   ClientHeight    =   8340
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form40_3.frx":0000
   ScaleHeight     =   8340
   ScaleWidth      =   12000
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CrystalReport40_31 
      Left            =   1440
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
   Begin VB.TextBox TXT40_31 
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
      Left            =   2850
      TabIndex        =   1
      Top             =   1875
      Width           =   2565
   End
   Begin VB.CommandButton CMD40_31 
      Caption         =   "ดูรายงาน"
      Height          =   465
      Left            =   3840
      TabIndex        =   0
      Top             =   3075
      Width           =   1590
   End
   Begin VB.Label LBL40_32 
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
      Left            =   1875
      TabIndex        =   3
      Top             =   1875
      Width           =   915
   End
   Begin VB.Label LBL40_31 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "เคลื่อนไหวเจ้าหนี้"
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
      Height          =   615
      Left            =   2550
      TabIndex        =   2
      Top             =   225
      Width           =   7440
   End
End
Attribute VB_Name = "Form40_3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMD40_31_Click()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vApCode As String, vRepType As String, vReportName As String
Dim vRepID As Integer

vApCode = Trim(TXT40_31.Text)
If vApCode <> "" Then
    vRepType = "AP"
    vRepID = 85
    
    vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
    'vQuery = "select reportname from bcreportname where repid = " & vRepID & "  and reptype = '" & vRepType & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vReportName = Trim(vRecordset.Fields("reportname").Value)
        With CrystalReport40_31
        .ReportFileName = vReportName & ".rpt"
        .ParameterFields(0) = "@apcode;" & vApCode & ";true"
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

Private Sub TXT40_31_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    CMD40_31_Click
End If
End Sub
