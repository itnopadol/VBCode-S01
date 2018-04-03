VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "CRYSTL32.OCX"
Begin VB.Form Form54_6 
   Caption         =   "พิมพ์จดหมายทวงหนี้"
   ClientHeight    =   8055
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form54_6.frx":0000
   ScaleHeight     =   8055
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport Crystal103 
      Left            =   2250
      Top             =   6165
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
   Begin VB.CommandButton CMD104 
      Caption         =   "พิมพ์ใบแจ้งชำระเงิน ธนาคารกรุงเทพ"
      Height          =   465
      Left            =   9270
      TabIndex        =   9
      Top             =   3420
      Width           =   1590
   End
   Begin Crystal.CrystalReport Crystal102 
      Left            =   1710
      Top             =   6165
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
   Begin VB.CommandButton CMD103 
      Caption         =   "พิมพ์ใบขออนุมัติ"
      Height          =   465
      Left            =   4005
      TabIndex        =   8
      Top             =   3420
      Visible         =   0   'False
      Width           =   1590
   End
   Begin Crystal.CrystalReport Crystal101 
      Left            =   1125
      Top             =   6165
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.CommandButton CMD102 
      Caption         =   "พิมพ์ใบนำฝากเงินไทยพาณิชย์"
      Height          =   465
      Left            =   7515
      TabIndex        =   7
      Top             =   3420
      Width           =   1605
   End
   Begin Crystal.CrystalReport CrystalReport54_6 
      Left            =   540
      Top             =   6165
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.CommandButton Cmd101 
      Caption         =   "กำหนดเลขที่จดหมายใหม่"
      Height          =   405
      Left            =   2250
      TabIndex        =   6
      Top             =   1080
      Width           =   2010
   End
   Begin VB.TextBox TXT54_6 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2250
      TabIndex        =   3
      Top             =   1620
      Width           =   2025
   End
   Begin VB.CommandButton Cmd54_6 
      Caption         =   "พิมพ์จดหมาย"
      Height          =   465
      Left            =   5760
      TabIndex        =   2
      Top             =   3420
      Width           =   1605
   End
   Begin VB.ComboBox Cmb2 
      Height          =   315
      Left            =   2250
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   2550
      Width           =   3345
   End
   Begin VB.ComboBox Cmb1 
      Height          =   315
      Left            =   2250
      TabIndex        =   0
      Top             =   2100
      Width           =   7470
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "ประเภทจดหมาย"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   600
      TabIndex        =   5
      Top             =   2550
      Width           =   1590
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "รหัสลูกค้า"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   390
      Left            =   900
      TabIndex        =   4
      Top             =   1620
      Width           =   1215
   End
End
Attribute VB_Name = "Form54_6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Cmb1_Click()
TXT54_6.Text = Left(Cmb1.Text, InStr(Cmb1.Text, "/") - 1)
End Sub

Private Sub CMD101_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vAutoNumber As String

On Error GoTo ErrDescription

vQuery = "Update npmaster.dbo.np_Generate_Docno set autonumber = autonumber+1 where headertype= 7 "
gConnection.Execute vQuery

vQuery = "select autonumber from npmaster.dbo.np_Generate_docno where headertype = 7 "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vAutoNumber = Trim(vRecordset.Fields("autonumber").Value)
End If
vRecordset.Close

MsgBox "ได้กำหนดเลขที่จดหมายใหม่เป็น เลขที่ " & vAutoNumber & ""

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub

Private Sub CMD102_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vARCode As String
Dim vReportName As String
Dim vCountStr As Integer
Dim vRepID As Integer
Dim vRepType As String

On Error GoTo ErrDescription

If TXT54_6.Text <> "" Then
  vARCode = Trim(TXT54_6.Text)
  
  vRepID = 314
  vRepType = "AR"
  vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
  'vQuery = "select reportname from bcnp.dbo.bcreportname where repid = 314 and reptype = 'AR' "
  If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
      vReportName = Trim(vRecordset.Fields("reportname").Value)
  End If
  vRecordset.Close
  
  With Crystal101
  .ReportFileName = vReportName & ".rpt"
  .ParameterFields(0) = "@ARCode;" & vARCode & ";true"
  .Destination = crptToWindow
  .WindowState = crptMaximized
  .Action = 1
  End With
  
  vRepID = 315
  vRepType = "AR"
  vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
  'vQuery = "select reportname from bcnp.dbo.bcreportname where repid = 315 and reptype = 'AR' "
  If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
      vReportName = Trim(vRecordset.Fields("reportname").Value)
  End If
  vRecordset.Close
  
  With Crystal101
  .ReportFileName = vReportName & ".rpt"
  .ParameterFields(0) = "@ARCode;" & vARCode & ";true"
  .Destination = crptToWindow
  .WindowState = crptMaximized
  .Action = 1
  End With
Else
  MsgBox "กรุณา กรอกรหัสสินค้าที่ต้องการจะพิมพ์เอกสารด้วย", vbCritical, "Send Error"
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub

Private Sub CMD103_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vAutoNumber As Integer
Dim vGenNumber As String
Dim vYear As String
Dim vMaxNumber As Integer
Dim vDocdate As Date
Dim vARCode As String
Dim vRepID As Integer
Dim vRepType As String


If TXT54_6.Text <> "" Then
vQuery = "exec dbo.USP_CD_GenerateConfirmNumber "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
  vYear = Trim(vRecordset.Fields("year1").Value)
  vMaxNumber = Trim(vRecordset.Fields("maxnumber").Value)
End If
vRecordset.Close
vGenNumber = Format(vMaxNumber, "0000") & "/" & vYear

vDocdate = CDate(Day(Now) & "/" & Month(Now) & "/" & Year(Now))
vARCode = Trim(TXT54_6.Text)
vQuery = "exec dbo.USP_CD_InsertConfirmSaleOrderRequestLogs '" & vGenNumber & "','" & vARCode & "','" & vDocdate & "','" & vUserID & "' "
gConnection.Execute vQuery

vRepID = 328
vRepType = "CD"
vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
 'vQuery = "select reportname from dbo.bcreportname where repid = 328 and reptype = 'CD' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    With Crystal102
        .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
        .ParameterFields(0) = "@vArCode;" & vARCode & ";true"
        .Formulas(0) = "vDocNo ='" & vGenNumber & "' "
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .Action = 1
    End With
End If
vRecordset.Close

TXT54_6.Text = ""
Cmb1.Text = ""

Else
  MsgBox "กรุณา กรอกรหัสสินค้าที่ต้องการจะพิมพ์เอกสารด้วย", vbCritical, "Send Error"
End If
End Sub

Private Sub CMD104_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vARCode As String
Dim vReportName As String
Dim vCountStr As Integer
Dim vRepID As Integer
Dim vRepType As String

On Error GoTo ErrDescription

If TXT54_6.Text <> "" Then
   vARCode = Trim(TXT54_6.Text)
  vRepID = 370
  vRepType = "AR"
  vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
  'vQuery = "select reportname from bcnp.dbo.bcreportname where repid = 370 and reptype = 'AR' "
  If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
      vReportName = Trim(vRecordset.Fields("reportname").Value)
  End If
  vRecordset.Close
  
  With Crystal103
  .ReportFileName = vReportName & ".rpt"
  .ParameterFields(0) = "@vArCode;" & vARCode & ";true"
  .Destination = crptToWindow
  .WindowState = crptMaximized
  .Action = 1
  End With
  
  End If
  
ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub

Private Sub Cmd54_6_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vAutoName As String
Dim vAutoNumber As Integer
Dim vGenNo As String, vGenNo1 As String
Dim vYear As String
Dim vLetterNo As String, vReportName As String
Dim vARCode As String
Dim vCountStr As Integer
Dim vRepID As Integer
Dim vRepType As String

On Error GoTo ErrDescription

If TXT54_6.Text <> "" Then

  vQuery = "select   header,autonumber ,right(year(getdate())+543,2) as year1 from npmaster.dbo.NP_Generate_DocNo where headertype = 7"
  If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
      vAutoNumber = Trim(vRecordset.Fields("autonumber").Value)
      vAutoName = Trim(vRecordset.Fields("header").Value)
      vYear = Trim(vRecordset.Fields("year1").Value)
  End If
  vRecordset.Close
  
  vGenNo = Format(vAutoNumber, "000")
  vGenNo1 = vAutoName & " " & vGenNo & "/" & vYear
  vARCode = Trim(TXT54_6.Text)
  
  vRepType = "AR"
  If Cmb2.Text = Trim("จดหมายดึงข้อมูลจากใบแจ้งหนี้") Then
  vRepID = 203
  vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
  'vQuery = "select reportname from bcnp.dbo.bcreportname where repid = '203' and reptype = 'AR' "
  ElseIf Cmb2.Text = Trim("จดหมายดึงข้อมูลจากบิลขาย") Then
  vRepID = 202
  vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
  'vQuery = "select reportname from bcnp.dbo.bcreportname where repid = '202' and reptype = 'AR' "
  Else
  MsgBox "กรุณาเลือก ประเภทการพิมพ์จดหมายแจ้งหนี้ด้วยนะครับ"
  Exit Sub
  End If
  If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
      vReportName = Trim(vRecordset.Fields("reportname").Value)
  End If
  vRecordset.Close
  
  With CrystalReport54_6
  .ReportFileName = vReportName & ".rpt"
  .Destination = crptToWindow
  .WindowState = crptMaximized
  .ParameterFields(0) = "@ARCode;" & vARCode & ";true"
  .Formulas(0) = "AutoNumber='" & vGenNo1 & "' "
  .Action = 1
  End With

Else
  MsgBox "กรุณา กรอกรหัสสินค้าที่ต้องการจะพิมพ์เอกสารด้วย", vbCritical, "Send Error"
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub


Private Sub Form_Load()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String

Cmb2.AddItem "จดหมายดึงข้อมูลจากใบแจ้งหนี้"
Cmb2.AddItem "จดหมายดึงข้อมูลจากบิลขาย"

vQuery = "select code+'/'+name1 as arname  from bcnp.dbo.bcar order by code "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    While Not vRecordset.EOF
    Cmb1.AddItem Trim(vRecordset.Fields("arname").Value)
    vRecordset.MoveNext
    Wend
End If
vRecordset.Close

'If vDepartment <> "CD" And vDepartment <> "AC" Then
 '   CMD101.Visible = False
  '  Cmb2.Visible = False
   ' Label2.Visible = False
    'Cmd54_6.Visible = False
    'CMD102.Visible = False
'End If

End Sub
