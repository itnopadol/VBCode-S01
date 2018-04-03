VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "CRYSTL32.OCX"
Begin VB.Form Form3_82 
   Caption         =   "Form1"
   ClientHeight    =   8355
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form3_82.frx":0000
   ScaleHeight     =   8355
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CrystalReport3_82 
      Left            =   1440
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
   Begin VB.CommandButton Command1 
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
      Height          =   465
      Left            =   4425
      TabIndex        =   9
      Top             =   4125
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker DTPDate2 
      Height          =   390
      Left            =   2925
      TabIndex        =   3
      Top             =   3375
      Width           =   2715
      _ExtentX        =   4789
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
      CurrentDate     =   38285
   End
   Begin MSComCtl2.DTPicker DTPDate1 
      Height          =   390
      Left            =   2925
      TabIndex        =   2
      Top             =   2700
      Width           =   2715
      _ExtentX        =   4789
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
      CurrentDate     =   38285
   End
   Begin VB.ComboBox Combo2 
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
      Left            =   2925
      TabIndex        =   1
      Top             =   2025
      Width           =   2715
   End
   Begin VB.ComboBox Combo1 
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
      Left            =   2925
      TabIndex        =   0
      Top             =   1500
      Width           =   2715
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "ถึง วันที่"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   2100
      TabIndex        =   8
      Top             =   3375
      Width           =   915
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "ณ วันที่"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   2100
      TabIndex        =   7
      Top             =   2700
      Width           =   765
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "ถึง จุดขายที่ "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   1800
      TabIndex        =   6
      Top             =   2025
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "จาก จุดขายที่"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   1800
      TabIndex        =   5
      Top             =   1500
      Width           =   1290
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "รายงาน การขายสินค้า ณ จุดขายต่าง ๆ"
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
      Left            =   2400
      TabIndex        =   4
      Top             =   300
      Width           =   7665
   End
End
Attribute VB_Name = "Form3_82"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDate1 As Date, vDate2 As Date
Dim vPositionSale1 As String, vpositionsale2 As String
Dim vReportName As String
Dim vRepID As Integer
Dim vRepType As String

vDate1 = DTPDate1.Day & "/" & DTPDate1.Month & "/" & DTPDate1.Year
vDate2 = DTPDate2.Day & "/" & DTPDate2.Month & "/" & DTPDate2.Year
vPositionSale1 = Trim(Combo1.Text)
vpositionsale2 = Trim(Combo2.Text)

vRepID = 177
vRepType = "SO"
vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
'vQuery = "select reportname from bcnp.dbo.bcreportname where repid = '177' and reptype = 'SO' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vReportName = Trim(vRecordset.Fields("reportname").Value)
End If
vRecordset.Close

    With CrystalReport3_82
    .ReportFileName = vReportName & ".rpt"
    .ParameterFields(0) = "@MachineB;" & vPositionSale1 & ";true"
    .ParameterFields(1) = "@MachineE;" & vpositionsale2 & ";true"
    .ParameterFields(2) = "@BDocdate;" & vDate1 & ";true"
    .ParameterFields(3) = "@EDocdate;" & vDate2 & ";true"
    .Destination = crptToWindow
    .WindowState = crptMaximized
    .Action = 1
    End With

End Sub

Private Sub Form_Load()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vMachine As String

DTPDate1 = Now
DTPDate2 = Now

vQuery = "select machineno from bcnp.dbo.bpsmachine order by machineno"
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    While Not vRecordset.EOF
    Combo1.AddItem Trim(vRecordset.Fields("machineno").Value)
    Combo2.AddItem Trim(vRecordset.Fields("machineno").Value)
    vRecordset.MoveNext
    Wend
End If
vRecordset.Close
End Sub
