VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "CRYSTL32.OCX"
Begin VB.Form Form88 
   Caption         =   "Form1"
   ClientHeight    =   8370
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form88.frx":0000
   ScaleHeight     =   8370
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CrystalReport88 
      Left            =   960
      Top             =   6600
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
   Begin VB.TextBox Text2 
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
      Height          =   360
      Left            =   3450
      TabIndex        =   3
      Top             =   2115
      Width           =   2265
   End
   Begin VB.TextBox Text1 
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
      Height          =   375
      Left            =   3465
      TabIndex        =   2
      Top             =   1395
      Width           =   2265
   End
   Begin MSComCtl2.DTPicker DTPDate2 
      Height          =   375
      Left            =   3450
      TabIndex        =   1
      Top             =   2850
      Width           =   2265
      _ExtentX        =   3995
      _ExtentY        =   661
      _Version        =   393216
      Format          =   67108865
      CurrentDate     =   38206
   End
   Begin VB.CommandButton Command1 
      Caption         =   "พิมพ์รายงาน"
      Height          =   525
      Left            =   4530
      TabIndex        =   0
      Top             =   3675
      Width           =   1185
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "ณ วันที่"
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
      Height          =   240
      Left            =   2325
      TabIndex        =   7
      Top             =   2850
      Width           =   690
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "ถึงรหัสสินค้า"
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
      Left            =   2325
      TabIndex        =   6
      Top             =   2175
      Width           =   990
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "จากรหัสสินค้า"
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
      Left            =   2325
      TabIndex        =   5
      Top             =   1425
      Width           =   1065
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "รายงาน สต็อกการ์ด GP"
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
      Left            =   2475
      TabIndex        =   4
      Top             =   300
      Width           =   7515
   End
End
Attribute VB_Name = "Form88"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim vQuery As String
Dim vRecordset1 As New ADODB.Recordset
Dim vNowDate As Date
Dim vReportName As String
Dim Date2 As Date
Dim vItem1, vItem2 As String
Dim conn As New ADODB.Connection
Dim vRepID As Integer
Dim vRepType As String

conn.Open "Provider=SQLOLEDB.1;Data Source=nebula;Initial Catalog=BCNP;User ID=vbuser;Password=132"
Date2 = DTPDate2.Day & "/" & DTPDate2.Month & "/" & DTPDate2.Year + 543
vItem1 = Trim(Text1.Text)
vItem2 = Trim(Text2.Text)

vRepID = 165
vRepType = "StockGP"
vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
'vQuery = "select reportname  from bcnp.dbo.bcreportname where reptype = 'StockGP' and repid = 165 "
vRecordset1.Open vQuery, conn, adOpenDynamic, adLockOptimistic
 If Not vRecordset1.EOF Then
    vReportName = Trim(vRecordset1.Fields("reportname").Value)
End If
vRecordset1.Close

With CrystalReport88
.ReportFileName = vReportName & ".rpt"
.ParameterFields(0) = "Itemnumber1;" & vItem1 & ";true"
.ParameterFields(1) = "Itemnumber2;" & vItem2 & ";true"
.ParameterFields(2) = "Date2;" & Date2 & ";true"
.Destination = crptToWindow
.WindowState = crptMaximized
.Action = 1
End With

End Sub


Private Sub Form_Load()
Me.DTPDate2.Value = "31/12/2003"
End Sub
