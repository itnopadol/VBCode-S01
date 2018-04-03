VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "CRYSTL32.OCX"
Begin VB.Form Form41 
   Caption         =   "พิมพ์ใบสำคัญตั้งเจ้าหนี้อื่น ๆ"
   ClientHeight    =   7980
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form41.frx":0000
   ScaleHeight     =   7980
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   2160
      Top             =   6120
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
   Begin VB.CommandButton CMD101 
      Caption         =   "พิมพ์เอกสาร"
      Height          =   465
      Left            =   3375
      TabIndex        =   1
      Top             =   2175
      Width           =   1590
   End
   Begin VB.TextBox TXT101 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   2400
      TabIndex        =   0
      Top             =   1350
      Width           =   2565
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "เลขที่เอกสารตั้งเจ้าหนี้"
      ForeColor       =   &H8000000E&
      Height          =   390
      Left            =   750
      TabIndex        =   2
      Top             =   1350
      Width           =   1590
   End
End
Attribute VB_Name = "Form41"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMD101_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocNo As String
Dim vReportName As String
Dim vRepID As Integer
Dim vRepType As String

If TXT101.Text <> "" Then
        vDocNo = Trim(TXT101.Text)
        vQuery = "select  docno from bcnp.dbo.bcapotherdebt  where docno = '" & vDocNo & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vDocNo = Trim(vRecordset.Fields("docno").Value)
        Else
        MsgBox "ไม่มีเลขที่เอกสาร " & vDocNo & " นี้ในระบบ กรุณาตรวจสอบด้วยนะครับ"
        TXT101.Text = ""
        TXT101.SetFocus
        Exit Sub
        End If
        vRecordset.Close
        
        vRepID = 204
        vRepType = "PM"
        vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
        'vQuery = "select reportname from bcnp.dbo.bcreportname where repid = 204 and reptype = 'PM' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vReportName = Trim(vRecordset.Fields("reportname").Value)
        End If
        vRecordset.Close
        
        With CrystalReport1
            .ReportFileName = vReportName & ".rpt"
            .Destination = crptToWindow
            .WindowState = crptMaximized
            .ParameterFields(0) = "@vdocno;" & vDocNo & ";true"
            .Action = 1
        End With
Else
        MsgBox "กรุณาใส่เลขที่เอกสารที่จะพิมพ์ด้วยนะครับ"
        TXT101.SetFocus
End If
End Sub

