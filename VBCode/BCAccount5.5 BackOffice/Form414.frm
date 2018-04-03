VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form Form414 
   Caption         =   "รายงาน เคลื่อนไหวเจ้าหนี้ ตามช่วงเวลา"
   ClientHeight    =   8970
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11955
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form414.frx":0000
   ScaleHeight     =   8970
   ScaleWidth      =   11955
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport Crystal101 
      Left            =   1260
      Top             =   6390
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
      Caption         =   "พิมพ์รายงาน"
      Height          =   465
      Left            =   4995
      TabIndex        =   4
      Top             =   3525
      Width           =   1365
   End
   Begin MSComCtl2.DTPicker DTPicker102 
      Height          =   315
      Left            =   3150
      TabIndex        =   3
      Top             =   3000
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      _Version        =   393216
      Format          =   16580609
      CurrentDate     =   38463
   End
   Begin MSComCtl2.DTPicker DTPicker101 
      Height          =   315
      Left            =   3150
      TabIndex        =   2
      Top             =   2400
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      _Version        =   393216
      Format          =   16580609
      CurrentDate     =   38463
   End
   Begin VB.ComboBox CMB102 
      Height          =   315
      Left            =   3150
      TabIndex        =   1
      Top             =   1875
      Width           =   5415
   End
   Begin VB.ComboBox CMB101 
      Height          =   315
      Left            =   3150
      TabIndex        =   0
      Top             =   1350
      Width           =   1815
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ถึงวันที่"
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   1575
      TabIndex        =   8
      Top             =   3000
      Width           =   1440
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "จากวันที่"
      ForeColor       =   &H8000000E&
      Height          =   240
      Left            =   2025
      TabIndex        =   7
      Top             =   2400
      Width           =   990
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "รหัสเจ้าหนี้"
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   2175
      TabIndex        =   6
      Top             =   1875
      Width           =   840
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "เลือกบริษัท"
      ForeColor       =   &H8000000E&
      Height          =   240
      Left            =   2100
      TabIndex        =   5
      Top             =   1350
      Width           =   915
   End
End
Attribute VB_Name = "Form414"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Cmd101_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vReportName As String
Dim vApCode As String
Dim vDate1, vDate2 As Date
Dim StrCount As Integer
Dim vRepID As Integer
Dim vRepType As String

StrCount = InStr(Trim(CMB102.Text), "/")
vApCode = Trim(Left(CMB102.Text, StrCount - 1))
If vApCode <> "" Then
vDate1 = DTPicker101.Day & "/" & DTPicker101.Month & "/" & DTPicker101.Year
vDate2 = DTPicker102.Day & "/" & DTPicker102.Month & "/" & DTPicker102.Year
If CMB101.Text = "Vat47" Then
    vRepID = 17
ElseIf CMB101.Text = "Vat48" Then
    vRepID = 18
End If
        vQuery = "select reportname from bcvat.dbo.bcreportname where repid = " & vRepID & "  and reptype = 'AP' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vReportName = Trim(vRecordset.Fields("reportname").Value)
        End If
        vRecordset.Close

With Crystal101
.ReportFileName = vReportName & ".rpt"
.ParameterFields(0) = "@Apcode;" & vApCode & ";true"
.ParameterFields(1) = "@Start;" & vDate1 & ";true"
.ParameterFields(2) = "@End;" & vDate2 & ";true"
.Destination = crptToWindow
.WindowState = crptMaximized
.Action = 1
End With
Else
MsgBox "กรุณาเลือก รหัสเจ้าหนี้ที่ต้องการดูรายงานด้วยครับ"
Exit Sub
End If
End Sub

Private Sub Form_Load()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String

DTPicker101 = Now
DTPicker102 = Now

vQuery = "select code+'/'+name1 as apname from bcvat.dbo.bcap order by code"
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    While Not vRecordset.EOF
        CMB102.AddItem Trim(vRecordset.Fields("apname").Value)
    vRecordset.MoveNext
    Wend
End If
vRecordset.Close

CMB101.AddItem Trim("Vat47")
CMB101.AddItem Trim("Vat48")

End Sub
