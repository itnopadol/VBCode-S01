VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form Form411 
   Caption         =   "หน้ารายงาน ค่าใช้จ่าย แยกตามเจ้าหนี้"
   ClientHeight    =   9000
   ClientLeft      =   1860
   ClientTop       =   2010
   ClientWidth     =   11970
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form411.frx":0000
   ScaleHeight     =   9000
   ScaleWidth      =   11970
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CrystalReport411 
      Left            =   1485
      Top             =   6480
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
   Begin VB.CommandButton CMD411 
      Caption         =   "พิมพ์รายงาน"
      Height          =   540
      Left            =   4050
      TabIndex        =   7
      Top             =   4425
      Width           =   1365
   End
   Begin MSComCtl2.DTPicker DTP412 
      Height          =   465
      Left            =   2625
      TabIndex        =   3
      Top             =   3225
      Width           =   2790
      _ExtentX        =   4921
      _ExtentY        =   820
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
      Format          =   16580609
      CurrentDate     =   38057
   End
   Begin MSComCtl2.DTPicker DTP411 
      Height          =   465
      Left            =   2625
      TabIndex        =   2
      Top             =   2625
      Width           =   2790
      _ExtentX        =   4921
      _ExtentY        =   820
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
      Format          =   16580609
      CurrentDate     =   38057
   End
   Begin VB.TextBox TXT411 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   2625
      TabIndex        =   1
      Top             =   1725
      Width           =   2790
   End
   Begin VB.Label LBL414 
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
      Height          =   390
      Left            =   1875
      TabIndex        =   6
      Top             =   3225
      Width           =   765
   End
   Begin VB.Label LBL413 
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
      Height          =   390
      Left            =   1875
      TabIndex        =   5
      Top             =   2625
      Width           =   765
   End
   Begin VB.Label LBL412 
      BackStyle       =   0  'Transparent
      Caption         =   "รหัสเจ้าหนี้"
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
      Height          =   390
      Left            =   1725
      TabIndex        =   4
      Top             =   1725
      Width           =   840
   End
   Begin VB.Label LBL411 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "รายงาน ค่าใช้จ่าย แยกตามเจ้าหนี้"
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
      Left            =   2625
      TabIndex        =   0
      Top             =   300
      Width           =   7365
   End
End
Attribute VB_Name = "Form411"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMD411_Click()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vApCode As String
Dim Date1 As Date, Date2 As Date
Dim vRepID As Integer
Dim vRepType As String
Dim vReportName As String

On Error GoTo Errdescription
Date1 = DTP411.Day & "/" & DTP411.Month & "/" & DTP411.Year
Date2 = DTP412.Day & "/" & DTP412.Month & "/" & DTP412.Year
 vApCode = Trim(TXT411.Text)
vRepID = 1
vRepType = "PM"

vQuery = "select reportname from bcvat.dbo.bcreportname where repid = " & vRepID & " and reptype = '" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vReportName = Trim(vRecordset.Fields("reportname").Value)
    
    With CrystalReport411
            .ReportFileName = vReportName & ".rpt"
            .ParameterFields(0) = "@StartDate;" & Date1 & ";true"
            .ParameterFields(1) = "@EndDate;" & Date2 & ";true"
            .ParameterFields(2) = "@APCode;" & vApCode & ";true"
            .Destination = crptToWindow
            .WindowState = crptMaximized
            .Action = 1
    End With
End If
vRecordset.Close


Errdescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If

End Sub

Private Sub Form_Load()
DTP411 = Now
DTP412 = Now
End Sub
