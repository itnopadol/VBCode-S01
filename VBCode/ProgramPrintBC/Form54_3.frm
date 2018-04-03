VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "CRYSTL32.OCX"
Begin VB.Form Form54_3 
   Caption         =   "หน้ายอดเคลื่อนไหวลูกหนี้"
   ClientHeight    =   8340
   ClientLeft      =   4875
   ClientTop       =   3630
   ClientWidth     =   12000
   Icon            =   "Form54_3.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form54_3.frx":08CA
   ScaleHeight     =   8340
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CrystalReport54_31 
      Left            =   2040
      Top             =   5640
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
   Begin VB.OptionButton Opt54_33 
      BackColor       =   &H80000001&
      Caption         =   "Option3"
      Height          =   195
      Left            =   6675
      TabIndex        =   9
      Top             =   2400
      Width           =   240
   End
   Begin VB.OptionButton Opt54_32 
      BackColor       =   &H80000001&
      Caption         =   "Option2"
      Height          =   195
      Left            =   6675
      TabIndex        =   8
      Top             =   1950
      Width           =   240
   End
   Begin VB.OptionButton Opt54_31 
      BackColor       =   &H80000001&
      Caption         =   "Option1"
      Height          =   225
      Left            =   6675
      TabIndex        =   7
      Top             =   1500
      Value           =   -1  'True
      Width           =   240
   End
   Begin MSComCtl2.DTPicker DTP54_32 
      Height          =   390
      Left            =   3525
      TabIndex        =   5
      Top             =   2775
      Width           =   2340
      _ExtentX        =   4128
      _ExtentY        =   688
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   66584577
      CurrentDate     =   38044
   End
   Begin MSComCtl2.DTPicker DTP54_31 
      Height          =   390
      Left            =   3525
      TabIndex        =   4
      Top             =   2250
      Width           =   2340
      _ExtentX        =   4128
      _ExtentY        =   688
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   66584577
      CurrentDate     =   38044
   End
   Begin VB.CommandButton CMD54_31 
      Caption         =   "พิมพ์รายงาน"
      Height          =   540
      Left            =   4275
      TabIndex        =   3
      Top             =   3900
      Width           =   1515
   End
   Begin VB.ComboBox CMB54_31 
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
      Left            =   3525
      TabIndex        =   2
      Top             =   1500
      Width           =   2340
   End
   Begin VB.Label LBL54_37 
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
      Left            =   2700
      TabIndex        =   13
      Top             =   2775
      Width           =   690
   End
   Begin VB.Label LBL54_36 
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
      Left            =   2700
      TabIndex        =   12
      Top             =   2250
      Width           =   765
   End
   Begin VB.Label LBL54_35 
      BackStyle       =   0  'Transparent
      Caption         =   "แสดงเฉพาะไม่มียอดคงค้าง"
      ForeColor       =   &H8000000E&
      Height          =   240
      Left            =   6975
      TabIndex        =   11
      Top             =   2400
      Width           =   1890
   End
   Begin VB.Label LBL54_34 
      BackStyle       =   0  'Transparent
      Caption         =   "แสดงเฉพาะมียอดคงค้าง"
      ForeColor       =   &H8000000E&
      Height          =   240
      Left            =   6975
      TabIndex        =   10
      Top             =   1950
      Width           =   1815
   End
   Begin VB.Label LBL54_33 
      BackStyle       =   0  'Transparent
      Caption         =   "แสดงทั้งหมด"
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
      Left            =   6975
      TabIndex        =   6
      Top             =   1500
      Width           =   1440
   End
   Begin VB.Label LBL54_32 
      BackStyle       =   0  'Transparent
      Caption         =   "รหัสลูกค้า"
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
      Left            =   2700
      TabIndex        =   1
      Top             =   1500
      Width           =   840
   End
   Begin VB.Label LBL54_31 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "พิมพ์รายงานยอดเคลื่อนไหวลูกหนี้"
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
      Height          =   540
      Left            =   2700
      TabIndex        =   0
      Top             =   225
      Width           =   7290
   End
End
Attribute VB_Name = "Form54_3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMD54_31_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vARCode1 As String, vARCode2 As String
Dim vDate1 As Date
Dim vCheck  As Integer
Dim Date1 As Date, Date2 As Date
Dim vRepID As Integer
Dim vRepType As String

On Error GoTo ErrDescription

vARCode1 = Trim(CMB54_31.Text)
If CMB54_31.Text <> "" Then
Date1 = DTP54_31.Day & "/" & DTP54_31.Month & "/" & DTP54_31.Year
Date2 = DTP54_32.Day & "/" & DTP54_32.Month & "/" & DTP54_32.Year

If Opt54_31.Value = True Then
    vCheck = 0
ElseIf Opt54_32.Value = True Then
    vCheck = 1
ElseIf Opt54_33.Value = True Then
    vCheck = 2
End If

        vRepID = 48
        vRepType = "AR"
        vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
        'vQuery = "select reportname from bcreportname where repid = 48 and reptype = 'AR' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            With CrystalReport54_31
                .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
                .ParameterFields(0) = "@ARCODE;" & vARCode1 & " ;true"
                .ParameterFields(1) = "@Show;" & vCheck & " ;true"
                .ParameterFields(2) = "@FDate;" & Date1 & " ;true"
                .ParameterFields(3) = "@LDate;" & Date2 & " ;true"
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
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset

DTP54_31 = Now
DTP54_32 = Now

vQuery = "select distinct code from bcar order by code"
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    While Not vRecordset.EOF
        CMB54_31.AddItem Trim(vRecordset.Fields("code").Value)
        vRecordset.MoveNext
    Wend
End If

vRecordset.Close
End Sub

