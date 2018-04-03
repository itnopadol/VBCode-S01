VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "CRYSTL32.OCX"
Begin VB.Form Form72 
   Caption         =   "พิมพ์เช็คและใบนำฝากจากหน้าโอนเงินระหว่างธนาคาร"
   ClientHeight    =   8370
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form72.frx":0000
   ScaleHeight     =   8370
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CrystalReport72 
      Left            =   2760
      Top             =   5760
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
   Begin VB.CommandButton CMD721 
      Caption         =   "พิมพ์เอกสาร"
      Height          =   690
      Left            =   3900
      TabIndex        =   4
      Top             =   3525
      Width           =   1815
   End
   Begin VB.ComboBox CMB721 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3300
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   2400
      Width           =   4290
   End
   Begin VB.TextBox TXT721 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   3300
      TabIndex        =   0
      Top             =   1650
      Width           =   2415
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "พิมพ์ใบนำฝากและพิมพ์เช็ค จากการโอนเงินระหว่างธนาคาร"
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
      Height          =   465
      Left            =   2550
      TabIndex        =   5
      Top             =   300
      Width           =   7440
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "ฟอร์มที่พิมพ์"
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
      Left            =   2175
      TabIndex        =   3
      Top             =   2400
      Width           =   990
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "เลขที่เอกสาร"
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
      Left            =   2175
      TabIndex        =   2
      Top             =   1650
      Width           =   915
   End
End
Attribute VB_Name = "Form72"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMD721_Click()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vRepType As String, vReportName As String
Dim vRepID As Integer
Dim vDocNo As String, vCheck As String, vBook As String, vChqNumber As String

vDocNo = Trim(TXT721.Text)
vCheck = Trim(CMB721.Text)

vQuery = "select tobook ,substring(mydescription,1,7) as ChqNumber  from BCBANKTRANSFER where docno = '" & vDocNo & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vBook = Trim(vRecordset.Fields("tobook").Value)
    If IsNull(vRecordset.Fields("chqnumber").Value) Then
    vChqNumber = "ไม่มีเช็ค"
    Else
     vChqNumber = Trim(vRecordset.Fields("chqnumber").Value)
    End If
End If
vRecordset.Close
Call CheckNumeric(vChqNumber)
vRepType = "TRF"
If vCheck = "พิมพ์ใบนำฝากธนาคารกรุงเทพ" And (vBook = "253-3-09888-1" Or vBook = "253-0-81240-9" Or vBook = "253-3-04147-7" Or vBook = "253-4-18282-5") Then
    If vChqNumber <> "" And vCheckValue = True Then
    vRepID = 156
    Else
    vRepID = 160
    End If
ElseIf vCheck = "พิมพ์ใบนำฝากธนาคารเอเชีย" And vBook = "109-1-00190-7" Then
    If vChqNumber <> "" And vCheckValue = True Then
    vRepID = 157
    Else
    vRepID = 161
    End If
ElseIf vCheck = "พิมพ์เช็คธนาคารกรุงเทพ" Then
    vRepID = 158
ElseIf vCheck = "พิมพ์เช็คธนาคารเอเชีย" Then
    vRepID = 159
ElseIf vCheck = "พิมพ์เช็คธนาคารกสิกรไทย" Then
    vRepID = 408
ElseIf vCheck = "พิมพ์เช็คธนาคารกรุงไทย" Then
    vRepID = 409
ElseIf vCheck = "พิมพ์เช็คธนาคารทหารไทย" Then
    vRepID = 410
ElseIf vCheck = "พิมพ์เช็คธนาคารไทยพาณิชย์" Then
    vRepID = 411
ElseIf vCheck = "พิมพ์เช็คธนาคารกรุงศรีอยุธยา" Then
    vRepID = 412
Else
    MsgBox "คุณเลือกธนาคารไม่ถูกต้อง กรุณาเลือกใหม่อีกครั้ง", vbInformation, "ข้อความแจ้งเตือน"
    Exit Sub
End If

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
'vQuery = "select reportname  from bcnp.dbo.bcreportname where reptype = '" & vRepType & "' and repid = " & vRepID & " "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vReportName = Trim(vRecordset.Fields("reportname").Value)
End If
vRecordset.Close

With CrystalReport72
.ReportFileName = vReportName & ".rpt"
.ParameterFields(0) = "@DocNo;" & vDocNo & ";true"
.Destination = crptToWindow
.WindowState = crptMaximized
.Action = 1
End With
TXT721.Text = ""
End Sub

Private Sub Form_Load()
CMB721.AddItem Trim("พิมพ์ใบนำฝากธนาคารกรุงเทพ")
CMB721.AddItem Trim("พิมพ์ใบนำฝากธนาคารเอเชีย")
CMB721.AddItem Trim("พิมพ์เช็คธนาคารกสิกรไทย")
CMB721.AddItem Trim("พิมพ์เช็คธนาคารกรุงเทพ")
CMB721.AddItem Trim("พิมพ์เช็คธนาคารกรุงไทย")
CMB721.AddItem Trim("พิมพ์เช็คธนาคารทหารไทย")
CMB721.AddItem Trim("พิมพ์เช็คธนาคารไทยพาณิชย์")
CMB721.AddItem Trim("พิมพ์เช็คธนาคารกรุงศรีอยุธยา")
CMB721.AddItem Trim("พิมพ์เช็คธนาคารเอเชีย")
End Sub
