VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form Form511 
   Caption         =   "รายงานเคลื่อนไหวลูกหนี้"
   ClientHeight    =   9000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11985
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form511.frx":0000
   ScaleHeight     =   9000
   ScaleWidth      =   11985
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CrystalReport511 
      Left            =   1080
      Top             =   5850
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
   Begin VB.CommandButton CMD511 
      Caption         =   "ดูรายงาน"
      Height          =   540
      Left            =   3825
      TabIndex        =   1
      Top             =   2700
      Width           =   1440
   End
   Begin VB.TextBox TXT511 
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
      Top             =   1725
      Width           =   2715
   End
   Begin VB.Label LBL511 
      BackStyle       =   0  'Transparent
      Caption         =   "รหัสลูกค้า"
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
      Left            =   1650
      TabIndex        =   3
      Top             =   1725
      Width           =   690
   End
   Begin VB.Label Label1 
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
      Left            =   2550
      TabIndex        =   2
      Top             =   225
      Width           =   7515
   End
End
Attribute VB_Name = "Form511"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMD511_Click()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vARCode As String, vRepType As String, vReportName As String
Dim vRepID As Integer

vARCode = Trim(TXT511.Text)
If vARCode <> "" Then
    vRepType = "AR"
    vRepID = 3
    
    vQuery = "select reportname from bcreportname where repid = " & vRepID & "  and reptype = '" & vRepType & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vReportName = Trim(vRecordset.Fields("reportname").Value)
        With CrystalReport511
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

Private Sub TXT511_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    CMD511_Click
End If
End Sub
