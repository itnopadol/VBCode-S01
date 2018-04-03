VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form Form52 
   Caption         =   "พิมพ์ใบตั้งลูกหนี้อื่น ๆ"
   ClientHeight    =   9000
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form52.frx":0000
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport Crystal101 
      Left            =   2115
      Top             =   5445
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
   Begin VB.CommandButton CMD101 
      Caption         =   "พิมพ์"
      Height          =   375
      Left            =   4635
      TabIndex        =   1
      Top             =   2385
      Width           =   1095
   End
   Begin VB.TextBox Text101 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   3960
      TabIndex        =   0
      Top             =   1755
      Width           =   1770
   End
   Begin VB.Label Label1 
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
      Height          =   330
      Left            =   2835
      TabIndex        =   2
      Top             =   1800
      Width           =   1590
   End
End
Attribute VB_Name = "Form52"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Cmd101_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocNo As String
Dim vReportName As String

On Error GoTo Errdescription

If Text101.Text <> "" Then
        vDocNo = Trim(Text101.Text)
        vQuery = "select  docno from dbo.bcarotherdebt  where docno = '" & vDocNo & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vDocNo = Trim(vRecordset.Fields("docno").Value)
        Else
        MsgBox "ไม่มีเลขที่เอกสาร " & vDocNo & " นี้ในระบบ กรุณาตรวจสอบด้วยนะครับ"
        Text101.Text = ""
        Text101.SetFocus
        Exit Sub
        End If
        vRecordset.Close
        
        vQuery = "select reportname from bcvat.dbo.bcreportname where repid = 27  and reptype = 'PM' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vReportName = Trim(vRecordset.Fields("reportname").Value)
        End If
        vRecordset.Close
        
        With Crystal101
            .ReportFileName = vReportName & ".rpt"
            .Destination = crptToWindow
            .WindowState = crptMaximized
            .ParameterFields(0) = "@vdocno;" & vDocNo & ";true"
            .Action = 1
        End With
Else
        MsgBox "กรุณาใส่เลขที่เอกสารที่จะพิมพ์ด้วยนะครับ"
        Text101.SetFocus
End If

Errdescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

