VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "CRYSTL32.OCX"
Begin VB.Form Form22 
   Caption         =   "รายงาน ตรวจสอบเอกสารใบสั่งซื้อสินค้า"
   ClientHeight    =   8010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form22.frx":0000
   ScaleHeight     =   8010
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport Crystal101 
      Left            =   1080
      Top             =   6600
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
      Caption         =   "พิมพ์รายงาน"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   5850
      TabIndex        =   4
      Top             =   3675
      Width           =   1590
   End
   Begin VB.TextBox Text101 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   315
      Left            =   2475
      TabIndex        =   3
      Top             =   2625
      Width           =   2490
   End
   Begin VB.ComboBox CMB101 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   315
      Left            =   2475
      TabIndex        =   2
      Top             =   2100
      Width           =   4965
   End
   Begin VB.OptionButton Opt102 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "ดูตามรหัสเจ้าหนี้"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   2475
      TabIndex        =   1
      Top             =   1575
      Width           =   1590
   End
   Begin VB.OptionButton Opt101 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "ดูตามเลขที่เอกสาร"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   2475
      TabIndex        =   0
      Top             =   1200
      Width           =   1590
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "เลขที่ใบสั่งซื้อ"
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
      Height          =   315
      Left            =   1425
      TabIndex        =   6
      Top             =   2625
      Width           =   1065
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "รหัสเจ้าหนี้"
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
      Height          =   240
      Left            =   1425
      TabIndex        =   5
      Top             =   2100
      Width           =   1065
   End
End
Attribute VB_Name = "Form22"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMD101_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vReportName As String
Dim vCountStr As Integer
Dim vApCode As String
Dim vDocNo  As String
Dim vRepID As Integer
Dim vRepType As String

On Error GoTo ErrDescription

If Opt101.Value = True Then
    If Text101.Text <> "" Then
        vDocNo = Trim(Text101.Text)
        vApCode = ""
    Else
        MsgBox "กรุณาใส่เลขที่ใบสั่งซื้อด้วย"
        Exit Sub
    End If
ElseIf Opt102.Value = True Then
    If CMB101.Text <> "" Then
        vCountStr = InStr(Trim(CMB101.Text), "/")
        vApCode = Trim(Left(CMB101.Text, vCountStr - 1))
        vDocNo = ""
    Else
        MsgBox "กรุณาเลือกรหัสเจ้าหนี้ด้วยครับ"
        Exit Sub
    End If
End If

vRepID = 215
vRepType = "AP"
vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
'vQuery = "select reportname from bcnp.dbo.bcreportname where repid = 215 and reptype = 'AP' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vReportName = Trim(vRecordset.Fields("reportname").Value)
End If
vRecordset.Close

With Crystal101
.ReportFileName = vReportName & ".rpt"
.ParameterFields(0) = "@vDocno;" & vDocNo & ";true"
.ParameterFields(1) = "@vapcode;" & vApCode & ";true"
.Destination = crptToWindow
.WindowState = crptMaximized
.Action = 1
End With

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub

Private Sub Form_Load()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String

vQuery = "select code+'/'+name1 as apname from bcnp.dbo.bcap order by code"
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    While Not vRecordset.EOF
        CMB101.AddItem Trim(vRecordset.Fields("apname").Value)
    vRecordset.MoveNext
    Wend
End If
vRecordset.Close

End Sub

Private Sub Opt101_Click()
If Opt101.Value = True Then
    CMB101.Enabled = False
    Text101.Enabled = True
    Text101.SetFocus
End If
End Sub

Private Sub Opt102_Click()
If Opt102.Value = True Then
    CMB101.Enabled = True
    Text101.Enabled = False
    CMB101.SetFocus
End If
End Sub
