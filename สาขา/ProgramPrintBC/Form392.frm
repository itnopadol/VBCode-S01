VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "CRYSTL32.OCX"
Begin VB.Form Form392 
   Caption         =   "รายงานใบมัดจำ แยกตามลูกค้า"
   ClientHeight    =   8355
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form392.frx":0000
   ScaleHeight     =   8355
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CrystalReport392_1 
      Left            =   1560
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
      WindowShowCloseBtn=   -1  'True
      WindowShowProgressCtls=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.CommandButton CMD392_2 
      Caption         =   "ตามพนักงาน"
      Height          =   465
      Left            =   4125
      TabIndex        =   8
      Top             =   4800
      Width           =   1290
   End
   Begin VB.CommandButton CMD392_1 
      Caption         =   "ตามลูกค้า"
      Height          =   465
      Left            =   4125
      TabIndex        =   6
      Top             =   4125
      Width           =   1290
   End
   Begin VB.TextBox TXT392_1 
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
      Height          =   390
      Left            =   3150
      TabIndex        =   2
      Top             =   1650
      Width           =   2265
   End
   Begin MSComCtl2.DTPicker DTP392_2 
      Height          =   390
      Left            =   3150
      TabIndex        =   1
      Top             =   3225
      Width           =   2265
      _ExtentX        =   3995
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
      Format          =   69795841
      CurrentDate     =   38056
   End
   Begin MSComCtl2.DTPicker DTP392_1 
      Height          =   390
      Left            =   3150
      TabIndex        =   0
      Top             =   2475
      Width           =   2265
      _ExtentX        =   3995
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
      Format          =   69795841
      CurrentDate     =   38056
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "รายงานใบมัดจำ แยกตามลูกค้า"
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
      Left            =   2775
      TabIndex        =   7
      Top             =   300
      Width           =   7290
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
      Height          =   390
      Left            =   2400
      TabIndex        =   5
      Top             =   3225
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
      Height          =   390
      Left            =   2400
      TabIndex        =   4
      Top             =   2475
      Width           =   690
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "รหัสลูกค้า/พนักงานขาย"
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
      Left            =   1350
      TabIndex        =   3
      Top             =   1650
      Width           =   1740
   End
End
Attribute VB_Name = "Form392"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMD392_1_Click()
Dim vQuery As String, vARCode As String
Dim vRecordset As New ADODB.Recordset
Dim Date1 As Date, Date2 As Date
Dim vRepType As String, vReportName As String
Dim vRepID As Integer

On Error GoTo ErrDescription

vARCode = Trim(TXT392_1.Text)
Date1 = DTP392_1.Day & "/" & DTP392_1.Month & "/" & DTP392_1.Year
Date2 = DTP392_2.Day & "/" & DTP392_2.Month & "/" & DTP392_2.Year
vRepID = 78
vRepType = "DP"

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
'vQuery = "select reportname from bcreportname where repid = " & vRepID & " and reptype = '" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vReportName = Trim(vRecordset.Fields("reportname").Value)

    With CrystalReport392_1
    .ReportFileName = vReportName & ".rpt"
    .ParameterFields(0) = "@ARCode;" & vARCode & ";true"
    .ParameterFields(1) = "@StartDate;" & Date1 & ";true"
    .ParameterFields(2) = "@EndDate;" & Date2 & ";true"
    .Destination = crptToWindow
    .WindowState = crptMaximized
    .Action = 1
    
    End With
End If

vRecordset.Close

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If

End Sub

Private Sub CMD392_2_Click()
Dim vQuery As String, vSaleCode As String
Dim vRecordset As New ADODB.Recordset
Dim Date1 As Date, Date2 As Date
Dim vRepType As String, vReportName As String
Dim vRepID As Integer

On Error GoTo ErrDescription

vSaleCode = Trim(TXT392_1.Text)
Date1 = DTP392_1.Day & "/" & DTP392_1.Month & "/" & DTP392_1.Year
Date2 = DTP392_2.Day & "/" & DTP392_2.Month & "/" & DTP392_2.Year
vRepID = 261
vRepType = "DP"

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
'vQuery = "select reportname from bcreportname where repid = " & vRepID & " and reptype = '" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vReportName = Trim(vRecordset.Fields("reportname").Value)

    With CrystalReport392_1
    .ReportFileName = vReportName & ".rpt"
    .ParameterFields(0) = "@vSaleCode;" & vSaleCode & ";true"
    .ParameterFields(1) = "@StartDate;" & Date1 & ";true"
    .ParameterFields(2) = "@EndDate;" & Date2 & ";true"
    .Destination = crptToWindow
    .WindowState = crptMaximized
    .Action = 1
    
    End With
End If

vRecordset.Close

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Form_Load()
DTP392_1 = Now
DTP392_2 = Now
End Sub
