VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form Form6_3 
   Caption         =   "พิมพ์ใบแลกเปลี่ยนเช็ค"
   ClientHeight    =   9000
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form6_3.frx":0000
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport Crystal101 
      Left            =   450
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
   Begin VB.TextBox TextDocno 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   4455
      TabIndex        =   0
      Top             =   1620
      Width           =   2085
   End
   Begin VB.CommandButton CMDPrint 
      Caption         =   "พิมพ์"
      Height          =   420
      Left            =   5445
      TabIndex        =   1
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "เลขที่เอกสาร :"
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
      Left            =   3330
      TabIndex        =   2
      Top             =   1620
      Width           =   1095
   End
End
Attribute VB_Name = "Form6_3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMDPrint_Click()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vDocNo As String
Dim vReportName As String
Dim vRepID As Integer
Dim vRepType As String

On Error GoTo ErrDescription

If Me.TextDocno.Text <> "" Then
  vDocNo = Me.TextDocno.Text
  vRepID = 362
  vRepType = "CHQ"
  vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
  If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
   vReportName = Trim(vRecordset.Fields("reportname").Value)
  End If
  vRecordset.Close
  
  
  With Me.Crystal101
  .ReportFileName = vReportName & ".rpt"
  .ParameterFields(0) = "@vDocno;" & vDocNo & ";true "
  .Destination = crptToWindow
  .WindowState = crptMaximized
  .Action = 1
  End With
  Me.TextDocno.Text = ""
  Me.TextDocno.SetFocus
Else
 MsgBox "กรุณากรอกเลขที่เอกสารยกเลิกเช็ค", vbCritical, "Send Error Message"
 Me.TextDocno.SetFocus
End If


ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description, vbCritical, "ข้อความเตือน"
Exit Sub
End If
End Sub

Private Sub TextDocno_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Call CMDPrint_Click
End If
End Sub
