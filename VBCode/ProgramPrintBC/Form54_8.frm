VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form Form54_8 
   Caption         =   "รายงาน เคลื่อนไหวลูกหนี้ ตามช่วงเวลา"
   ClientHeight    =   7995
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form54_8.frx":0000
   ScaleHeight     =   7995
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport Crystal101 
      Left            =   3120
      Top             =   6075
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
   Begin VB.ComboBox CMB102 
      Height          =   315
      Left            =   2475
      TabIndex        =   0
      Top             =   1290
      Width           =   3915
   End
   Begin VB.CommandButton CMD101 
      Caption         =   "พิมพ์รายงาน"
      Height          =   465
      Left            =   3150
      TabIndex        =   4
      Top             =   4215
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker DTP102 
      Height          =   315
      Left            =   2475
      TabIndex        =   3
      Top             =   3240
      Width           =   1890
      _ExtentX        =   3334
      _ExtentY        =   556
      _Version        =   393216
      Format          =   57737217
      CurrentDate     =   38435
   End
   Begin MSComCtl2.DTPicker DTP101 
      Height          =   315
      Left            =   2475
      TabIndex        =   2
      Top             =   2565
      Width           =   1890
      _ExtentX        =   3334
      _ExtentY        =   556
      _Version        =   393216
      Format          =   57737217
      CurrentDate     =   38435
   End
   Begin VB.ComboBox CMB101 
      Height          =   315
      Left            =   2475
      TabIndex        =   1
      Top             =   1890
      Width           =   3915
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ประเภทรายงาน"
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   375
      TabIndex        =   8
      Top             =   1290
      Width           =   1965
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "ถึงวันที่"
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
      Left            =   1800
      TabIndex        =   7
      Top             =   3240
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "จากวันที่"
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
      Left            =   1725
      TabIndex        =   6
      Top             =   2565
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ลูกค้า"
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
      Left            =   1950
      TabIndex        =   5
      Top             =   1890
      Width           =   615
   End
End
Attribute VB_Name = "Form54_8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMD101_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vARCode1 As String
Dim vDate1 As Date, vDate2 As Date
Dim vReportName As String
Dim StrCount As Integer
Dim vRepID As Integer
Dim vRepType As String

On Error GoTo ErrDescription

If CMB102.Text <> "" Then
        StrCount = InStr(Trim(CMB101.Text), "/")
        vARCode1 = Trim(Left(CMB101.Text, StrCount - 1))
        vDate1 = DTP101.Day & "/" & DTP101.Month & "/" & DTP101.Year
        vDate2 = DTP102.Day & "/" & DTP102.Month & "/" & DTP102.Year
    
    If CMB102.Text = Trim("ไม่รวมมัดจำ") Then
       vRepType = "AR"
       vRepID = 212
       vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
        'vQuery = "select reportname from bcnp.dbo.bcreportname where repid = 212 and reptype = 'AR' "
    Else
        vRepID = 226
        vRepType = "AR"
        vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
        'vQuery = "select reportname from bcnp.dbo.bcreportname where repid = 226 and reptype = 'AR' "
    End If
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vReportName = Trim(vRecordset.Fields("reportname").Value)
        End If
        vRecordset.Close
    
    With Crystal101
    .ReportFileName = vReportName & ".rpt"
    .ParameterFields(0) = "@ARCode;" & vARCode1 & ";true"
    .ParameterFields(2) = "@Start;" & vDate1 & ";true"
    .ParameterFields(3) = "@End;" & vDate2 & ";true"
    .Destination = crptToWindow
    .WindowState = crptMaximized
    .Action = 1
    End With
Else
    MsgBox "กรุณาเลือก ประเภทรายงานด้วยนะครับ"
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

On Error Resume Next

DTP101 = Now
DTP102 = Now

CMB102.AddItem Trim("รวมมัดจำ")
CMB102.AddItem Trim("ไม่รวมมัดจำ")

vQuery = "select code+'/'+name1 as arname from bcnp.dbo.bcar order by code"
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    While Not vRecordset.EOF
        CMB101.AddItem Trim(vRecordset.Fields("arname").Value)
    vRecordset.MoveNext
    Wend
End If
vRecordset.Close
End Sub
