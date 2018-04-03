VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "CRYSTL32.OCX"
Begin VB.Form Form6_2 
   Caption         =   "พิมพ์ใบยกเลิกเช็ครับ"
   ClientHeight    =   9000
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form6_2.frx":0000
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin VB.TextBox TextDocNo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4095
      TabIndex        =   0
      Top             =   1575
      Width           =   2670
   End
   Begin Crystal.CrystalReport Crystal101 
      Left            =   540
      Top             =   5805
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.CommandButton CMDPrint 
      Caption         =   "พิมพ์เอกสาร"
      Height          =   465
      Left            =   5580
      TabIndex        =   1
      Top             =   2385
      Width           =   1185
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "เลขที่เอกสาร :"
      ForeColor       =   &H8000000E&
      Height          =   285
      Left            =   2970
      TabIndex        =   2
      Top             =   1575
      Width           =   1140
   End
End
Attribute VB_Name = "Form6_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMDPrint_Click()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vDocno As String
Dim vReportName As String
Dim vRepID As Integer
Dim vRepType As String

On Error GoTo ErrDescription

If Me.TextDocNo.Text <> "" Then
  vDocno = Me.TextDocNo.Text
  vRepID = 361
  vRepType = "CHQ"
  vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
  If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
   vReportName = Trim(vRecordset.Fields("reportname").Value)
  End If
  vRecordset.Close
  
  
  With Me.Crystal101
  .ReportFileName = vReportName & ".rpt"
  .ParameterFields(0) = "@vDocno;" & vDocno & ";true "
  .Destination = crptToWindow
  .WindowState = crptMaximized
  .Action = 1
  End With
  Me.TextDocNo.Text = ""
  Me.TextDocNo.SetFocus
Else
 MsgBox "กรุณากรอกเลขที่เอกสารยกเลิกเช็ค", vbCritical, "Send Error Message"
 Me.TextDocNo.SetFocus
End If


ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description, vbCritical, "ข้อความเตือน"
Exit Sub
End If
End Sub

Private Sub TextDocNo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Call CMDPrint_Click
End If
End Sub
