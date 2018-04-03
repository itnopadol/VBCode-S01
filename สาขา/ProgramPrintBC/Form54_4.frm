VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "CRYSTL32.OCX"
Begin VB.Form Form54_4 
   Caption         =   "หน้าเคลื่อนไหวลูกหนี้"
   ClientHeight    =   8370
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form54_4.frx":0000
   ScaleHeight     =   8370
   ScaleWidth      =   12000
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CrystalReport54_41 
      Left            =   2760
      Top             =   5040
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
   Begin VB.CommandButton CMD54_41 
      Caption         =   "ดูรายงาน"
      Height          =   465
      Left            =   3825
      TabIndex        =   1
      Top             =   2850
      Width           =   1440
   End
   Begin VB.TextBox TXT54_41 
      Height          =   465
      Left            =   2775
      TabIndex        =   0
      Top             =   1800
      Width           =   2490
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "รายงานเคลื่อนไหวลูกหนี้"
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
      TabIndex        =   3
      Top             =   225
      Width           =   7365
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "รหัสลูกหนี้"
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
      Left            =   1800
      TabIndex        =   2
      Top             =   1800
      Width           =   915
   End
End
Attribute VB_Name = "Form54_4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMD54_41_Click()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vARCode As String, vRepType As String, vReportName As String
Dim vRepID As Integer

vARCode = Trim(TXT54_41.Text)
If vARCode <> "" Then
    vRepType = "AR"
    vRepID = 86
    
    vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
    'vQuery = "select reportname from bcreportname where repid = " & vRepID & "  and reptype = '" & vRepType & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vReportName = Trim(vRecordset.Fields("reportname").Value)
        With CrystalReport54_41
        .ReportFileName = vReportName & ".rpt"
        .ParameterFields(0) = "@arcode;" & vARCode & ";true"
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

Private Sub TXT54_41_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    CMD54_41_Click
End If
End Sub
