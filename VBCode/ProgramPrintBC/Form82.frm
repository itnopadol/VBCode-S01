VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "CRYSTL32.OCX"
Begin VB.Form Form82 
   Caption         =   "รายงาน สินค้าขายดี ตามคลัง"
   ClientHeight    =   8355
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   Picture         =   "Form82.frx":0000
   ScaleHeight     =   8355
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CrystalPrint 
      Left            =   720
      Top             =   7080
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
   Begin VB.CommandButton CMDPrint 
      Caption         =   "พิมพ์รายงาน"
      Height          =   615
      Left            =   3360
      TabIndex        =   6
      Top             =   5280
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.ComboBox CMBItemType 
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
      Left            =   2775
      TabIndex        =   5
      Text            =   "ชนิดสินค้า"
      Top             =   4650
      Visible         =   0   'False
      Width           =   2190
   End
   Begin VB.ComboBox CMBWHCode 
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
      Left            =   2775
      TabIndex        =   4
      Text            =   "คลัง"
      Top             =   4050
      Visible         =   0   'False
      Width           =   2190
   End
   Begin VB.TextBox TXTLevelSale 
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
      Left            =   2775
      TabIndex        =   3
      Text            =   "อันดับขายดี"
      Top             =   3450
      Visible         =   0   'False
      Width           =   2190
   End
   Begin MSComCtl2.DTPicker DTPDate2 
      Height          =   390
      Left            =   2775
      TabIndex        =   2
      Top             =   2850
      Visible         =   0   'False
      Width           =   2190
      _ExtentX        =   3863
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
      Format          =   67108865
      CurrentDate     =   38129
   End
   Begin MSComCtl2.DTPicker DTPDate1 
      Height          =   390
      Left            =   2775
      TabIndex        =   1
      Top             =   2250
      Visible         =   0   'False
      Width           =   2190
      _ExtentX        =   3863
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
      Format          =   67108865
      CurrentDate     =   38129
   End
   Begin VB.ComboBox CMBSelect 
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
      Left            =   2775
      TabIndex        =   0
      Top             =   1500
      Width           =   5040
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ประเภทสินค้า"
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
      Left            =   1125
      TabIndex        =   12
      Top             =   4650
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "คลัง"
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
      TabIndex        =   11
      Top             =   4050
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "จำนวนอันดับสินค้า"
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
      Left            =   1425
      TabIndex        =   10
      Top             =   3450
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
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
      Left            =   1725
      TabIndex        =   9
      Top             =   2850
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
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
      Left            =   1725
      TabIndex        =   8
      Top             =   2250
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "เลือกรายงาน"
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
      TabIndex        =   7
      Top             =   1500
      Width           =   990
   End
End
Attribute VB_Name = "Form82"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vCmbSelect As String
Private Sub CMBSelect_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String

vCmbSelect = Trim(CMBSelect.Text)
CMDPrint.Visible = True
If vCmbSelect = "รายงาน สินค้าขายดีรวมคลัง" Then
    DTPDate1.Visible = True
    DTPDate2.Visible = True
    TXTLevelSale.Visible = True
    CMBWHCode.Visible = False
    CMBItemType.Visible = False
    Label2.Visible = True
    Label3.Visible = True
    Label4.Visible = True
    Label5.Visible = False
    Label6.Visible = False
ElseIf vCmbSelect = "รายงาน สินค้าขายดีแยกคลัง" Then
    DTPDate1.Visible = True
    DTPDate2.Visible = True
    TXTLevelSale.Visible = True
    CMBWHCode.Visible = True
    CMBItemType.Visible = False
    Label2.Visible = True
    Label3.Visible = True
    Label4.Visible = True
    Label5.Visible = True
    Label6.Visible = False
            vQuery = "SELECT    code From dbo.BCWarehouse "
                If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                    vRecordset.MoveFirst
                    While Not vRecordset.EOF
                    CMBWHCode.AddItem Trim(vRecordset.Fields("code").Value)
                    vRecordset.MoveNext
                    Wend
                End If
            vRecordset.Close
Else
    DTPDate1.Visible = True
    DTPDate2.Visible = True
    TXTLevelSale.Visible = True
    CMBWHCode.Visible = True
    CMBItemType.Visible = True
    Label2.Visible = True
    Label3.Visible = True
    Label4.Visible = True
    Label5.Visible = True
    Label6.Visible = True
                vQuery = "SELECT    code From dbo.bcitemtype "
                If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                    vRecordset.MoveFirst
                    While Not vRecordset.EOF
                    CMBItemType.AddItem Trim(vRecordset.Fields("code").Value)
                    vRecordset.MoveNext
                    Wend
                End If
            vRecordset.Close
End If

End Sub

Private Sub CMDPrint_Click()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vRepType As String, vReportName As String
Dim vRepID As Integer
Dim Date1 As Date, Date2 As Date
Dim vLevel As Integer
Dim vWHCode As String
Dim vItemType As String

vRepType = "item"

If vCmbSelect = "รายงาน สินค้าขายดีรวมคลัง" Then
    vRepID = 119
ElseIf vCmbSelect = "รายงาน สินค้าขายดีแยกคลัง" Then
    vRepID = 120
Else
    vRepID = 121
End If

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
'vQuery = "select reportname from bcnp.dbo.bcreportname where reptype = '" & vRepType & "' and repid ='" & vRepID & "'  "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vReportName = Trim(vRecordset.Fields("reportname").Value)
End If
vRecordset.Close

Date1 = DTPDate1.Day & "/" & DTPDate1.Month & "/" & DTPDate1.Year
Date2 = DTPDate2.Day & "/" & DTPDate2.Month & "/" & DTPDate2.Year
vLevel = Trim(TXTLevelSale.Text)
If CMBWHCode.Visible = True Then
    vWHCode = Trim(CMBWHCode.Text)
End If
If CMBItemType.Visible = True Then
    vItemType = Trim(CMBItemType.Text)
End If

With CrystalPrint
.ReportFileName = Trim(vReportName) & ".rpt"
.Destination = crptToWindow
.WindowState = crptMaximized
.ParameterFields(0) = "@SDate;" & Date1 & ";true"
.ParameterFields(1) = "@EDate;" & Date2 & ";true"
.ParameterFields(2) = "@TOP;" & vLevel & ";true"
If CMBWHCode.Visible = True Then
.ParameterFields(3) = "@WHCode;" & vWHCode & ";true"
End If
If CMBItemType.Visible = True Then
.ParameterFields(4) = "@ItemType;" & vItemType & ";true"
End If
.Action = 1
End With

End Sub

Private Sub Form_Load()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset

On Error Resume Next

CMBSelect.AddItem Trim("รายงาน สินค้าขายดีรวมคลัง")
CMBSelect.AddItem Trim("รายงาน สินค้าขายดีแยกคลัง")
CMBSelect.AddItem Trim("รายงาน สินค้าขายดีแยกคลัง_ชนิดสินค้า")

End Sub

