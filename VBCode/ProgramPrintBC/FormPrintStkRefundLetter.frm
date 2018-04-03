VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form FormPrintStkRefundLetter 
   Caption         =   "พิมพ์จดหมายส่งคืนสินค้าให้กับเจ้าหนี้"
   ClientHeight    =   11490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15375
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "FormPrintStkRefundLetter.frx":0000
   ScaleHeight     =   11490
   ScaleWidth      =   15375
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport Crystal101 
      Left            =   810
      Top             =   6435
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.CommandButton CMDPrint 
      Caption         =   "พิมพ์จดหมาย"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4725
      TabIndex        =   2
      Top             =   2520
      Width           =   1995
   End
   Begin VB.TextBox TXTDocNo 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   14.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   3420
      TabIndex        =   0
      Top             =   1665
      Width           =   3300
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "เลขที่ใบส่งคืนสินค้า :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   14.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   465
      Left            =   1080
      TabIndex        =   1
      Top             =   1710
      Width           =   2400
   End
End
Attribute VB_Name = "FormPrintStkRefundLetter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMDPrint_Click()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vDocNo As String
Dim vRepID  As Integer
Dim vRepType  As String


On Error Resume Next

vDocNo = Me.TXTDocNo.Text

vRepID = 506
vRepType = "AP"

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    With Crystal101
        .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
        .ParameterFields(0) = "@vDocNo;" & vDocNo & ";true"
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .Action = 1
    End With
End If
vRecordset.Close
    

End Sub
