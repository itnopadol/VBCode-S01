VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "CRYSTL32.OCX"
Begin VB.Form Form542_1 
   Caption         =   "หน้าพิมพ์รายงานยอดลูกหนี้ประจำเดือนตามรหัสลูกหนี้"
   ClientHeight    =   8235
   ClientLeft      =   2205
   ClientTop       =   1110
   ClientWidth     =   12000
   Icon            =   "Form541_2.frx":0000
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   Picture         =   "Form541_2.frx":08CA
   ScaleHeight     =   8235
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CrystalReport541_21 
      Left            =   2160
      Top             =   6360
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
   Begin VB.CommandButton CMD541_21 
      Caption         =   "พิมพ์"
      Height          =   540
      Left            =   2925
      TabIndex        =   3
      Top             =   3900
      Width           =   1665
   End
   Begin MSComCtl2.DTPicker DTP541_21 
      Height          =   315
      Left            =   2025
      TabIndex        =   2
      Top             =   3150
      Width           =   2565
      _ExtentX        =   4524
      _ExtentY        =   556
      _Version        =   393216
      Format          =   66584577
      CurrentDate     =   38027
   End
   Begin VB.ComboBox CMB541_22 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2025
      TabIndex        =   1
      Top             =   2400
      Width           =   2565
   End
   Begin VB.ComboBox CMB541_21 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2025
      TabIndex        =   0
      Top             =   1725
      Width           =   2565
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "พิมพ์รายงานยอดลูกหนี้ประจำเดือนตามรหัสลูกหนี้"
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
      Height          =   615
      Left            =   2550
      TabIndex        =   4
      Top             =   300
      Width           =   7440
   End
End
Attribute VB_Name = "Form542_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMD541_21_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vARCode1 As String, vARCode2 As String
Dim vDate1 As Date
Dim vRepID As Integer
Dim vRepType As String

On Error GoTo ErrDescription

vARCode1 = Trim(CMB541_21.Text)
vARCode2 = Trim(CMB541_22.Text)
vDate1 = DTP541_21.Day & "/" & DTP541_21.Month & "/" & DTP541_21.Year
If CMB541_21.Text <> "" And CMB541_22.Text <> "" Then

vRepID = 30
vRepType = "AR"
vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
        'vQuery = "select reportname from bcreportname where repid = 30 and reptype = 'AR' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            With CrystalReport541_21
                .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
                .ParameterFields(0) = "@ARCODE1;" & vARCode1 & " ;true"
                .ParameterFields(1) = "@ARCODE2;" & vARCode2 & " ;true"
                .ParameterFields(2) = "@AtDate;" & vDate1 & ";true"
                .Destination = crptToWindow
                .WindowState = crptMaximized
                .Action = 1
            End With
        End If
        vRecordset.Close
Else
        MsgBox "คุณใส่เงื่อนไขไม่ครบ", vbInformation + vbCritical, "ข้อความเตือน"
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



On Error GoTo ErrDescription

DTP541_21 = Now

vQuery = "select distinct   code from bcar where activestatus = 1  order by code"
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    While Not vRecordset.EOF
        CMB541_21.AddItem Trim(vRecordset.Fields("code").Value)
        CMB541_22.AddItem Trim(vRecordset.Fields("code").Value)
    vRecordset.MoveNext
    Wend
End If
vRecordset.Close


ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
End If

End Sub
