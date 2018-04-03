VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form FrmOrder403 
   Caption         =   "Form403 รายงานคิดค่าขนส่งพนักงานจัดส่ง"
   ClientHeight    =   9000
   ClientLeft      =   4035
   ClientTop       =   450
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form403.frx":0000
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport Crystal101 
      Left            =   1620
      Top             =   6435
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
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   5715
      TabIndex        =   6
      Top             =   3825
      Width           =   1275
   End
   Begin VB.ComboBox CMB101 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4275
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1890
      Width           =   2715
   End
   Begin MSComCtl2.DTPicker DTPicker102 
      Height          =   330
      Left            =   4275
      TabIndex        =   4
      Top             =   3060
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   582
      _Version        =   393216
      Format          =   16121857
      CurrentDate     =   38730
   End
   Begin MSComCtl2.DTPicker DTPicker101 
      Height          =   330
      Left            =   4275
      TabIndex        =   3
      Top             =   2565
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   582
      _Version        =   393216
      Format          =   16121857
      CurrentDate     =   38730
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ถึงวันที่"
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
      Height          =   285
      Left            =   2520
      TabIndex        =   2
      Top             =   3105
      Width           =   1590
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "จากวันที่"
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
      Left            =   2520
      TabIndex        =   1
      Top             =   2565
      Width           =   1590
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "เลือกประเภท รายงาน"
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
      Height          =   285
      Left            =   1800
      TabIndex        =   0
      Top             =   1890
      Width           =   2310
   End
End
Attribute VB_Name = "FrmOrder403"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMD101_Click()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vDate1 As Date
Dim vDate2 As Date
Dim vRepID As Integer
Dim vRepType As String
Dim vReportName As String

On Error GoTo ErrDescription

If CMB101.Text <> "" Then
    If CMB101.Text = Trim("แบบ แสดงรายละเอียด") Then
        vRepID = 301
    ElseIf CMB101.Text = Trim("แบบ แสดงสรุป") Then
        vRepID = 302
    End If
    vRepType = "DO"
    
    vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
    'vQuery = "select reportname from bcnp.dbo.bcreportname where repid = " & vRepID & " and reptype = '" & vRepType & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vReportName = Trim(vRecordset.Fields("reportname").Value)
    End If
    vRecordset.Close
    
        vDate1 = DTPicker101.Day & "/" & DTPicker101.Month & "/" & DTPicker101.Year
        vDate2 = DTPicker102.Day & "/" & DTPicker102.Month & "/" & DTPicker102.Year
    
        With Crystal101
        .ReportFileName = vReportName & ".rpt"
        .ParameterFields(0) = "@BeginDate;" & vDate1 & ";true"
        .ParameterFields(1) = "@EndDate;" & vDate2 & ";true"
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .Action = 1
        End With
Else
    MsgBox "กรุณาเลือกประเภทรายงาน", vbCritical, "Send Error"
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Private Sub Form_Load()
CMB101.AddItem Trim("แบบ แสดงรายละเอียด")
CMB101.AddItem Trim("แบบ แสดงสรุป")
DTPicker101 = Now
DTPicker102 = Now
End Sub
