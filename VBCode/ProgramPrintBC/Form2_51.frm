VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form Form2_51 
   Caption         =   "รายงาน แสดงเลขที่ใบเสนอซื้อสินค้าที่ได้จากใบ Back Order"
   ClientHeight    =   7980
   ClientLeft      =   4155
   ClientTop       =   1800
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form2_51.frx":0000
   ScaleHeight     =   7980
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport Crystal101 
      Left            =   4950
      Top             =   6660
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
   Begin VB.OptionButton Option103 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "รหัสลูกค้า"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   1800
      TabIndex        =   6
      Top             =   1950
      Width           =   1665
   End
   Begin VB.OptionButton Option102 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "รหัสพนักงาน"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   1800
      TabIndex        =   5
      Top             =   1500
      Width           =   1665
   End
   Begin VB.OptionButton Option101 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "เลขที่ Back Order"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   1800
      TabIndex        =   4
      Top             =   1050
      Value           =   -1  'True
      Width           =   1665
   End
   Begin VB.CommandButton CMD101 
      Caption         =   "ดูรายงาน"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1800
      TabIndex        =   3
      Top             =   6300
      Width           =   1665
   End
   Begin VB.TextBox Text101 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1800
      TabIndex        =   1
      Top             =   2400
      Width           =   1665
   End
   Begin MSComctlLib.ListView ListView101 
      Height          =   3240
      Left            =   1800
      TabIndex        =   0
      Top             =   2850
      Width           =   8340
      _ExtentX        =   14711
      _ExtentY        =   5715
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "เลขที่ Back Order"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "เลขที่ใบเสนอซื้อ"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "วันที่ทำรายการ"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "ชื่อลูกค้า"
         Object.Width           =   6174
      EndProperty
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ข้อมูล"
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
      Height          =   240
      Left            =   150
      TabIndex        =   2
      Top             =   2400
      Width           =   1590
   End
End
Attribute VB_Name = "Form2_51"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMD101_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vReportName As String
Dim vStatus1, vStatus2 As String
Dim vRepID As Integer
Dim vRepType As String
Dim vPrNO As String
Dim vBackOrder As String
Dim vSaleCode As String
Dim vARCode As String

On Error GoTo ErrDescription

If Text101.Text <> "" Then
vRepType = "PR"
vRepID = 233

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
'vQuery = "select reportname from bcreportname  where reptype = '" & vRepType & "'  and repid = " & vRepID & " "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vReportName = Trim(vRecordset.Fields("reportname").Value)
End If
vRecordset.Close

If Option101.Value = True Then
    vBackOrder = Trim(Text101.Text)
    vSaleCode = ""
    vARCode = ""
ElseIf Option102.Value = True Then
    vBackOrder = ""
    vSaleCode = Trim(Text101.Text)
    vARCode = ""
ElseIf Option103.Value = True Then
    vBackOrder = ""
    vSaleCode = ""
    vARCode = Trim(Text101.Text)
End If

 With Crystal101
 .ReportFileName = vReportName & ".rpt"
 .ParameterFields(1) = "@vSaleCode;" & vSaleCode & ";true"
 .ParameterFields(2) = "@vBackOrderNo;" & vBackOrder & ";true"
 .ParameterFields(3) = "@vARCode;" & vARCode & ";true"
 .Destination = crptToWindow
 .WindowState = crptMaximized
 .Action = 1
 End With
Else
    MsgBox "ยังไม่ได้กรอกข้อมูล "
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub

Private Sub Option101_Click()
Text101.SetFocus
Text101.Text = ""
ListView101.ListItems.Clear
End Sub

Private Sub Option102_Click()
Text101.SetFocus
Text101.Text = ""
ListView101.ListItems.Clear
End Sub

Private Sub Option103_Click()
Text101.SetFocus
Text101.Text = ""
ListView101.ListItems.Clear
End Sub

Private Sub Text101_KeyPress(KeyAscii As Integer)
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim PRListItem As ListItem
Dim vData As String

If KeyAscii = 13 Then
    If Text101.Text <> "" Then
    ListView101.ListItems.Clear
        vData = Trim(Text101.Text)
    If Option101.Value = True Then
        vQuery = "execute USP_AP_GeneratePR '','" & vData & "' ,'' "
    ElseIf Option102.Value = True Then
        vQuery = "execute USP_AP_GeneratePR  '" & vData & "','' ,'' "
    ElseIf Option103.Value = True Then
    vQuery = "execute USP_AP_GeneratePR  '','','" & vData & "'  "
    End If
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vRecordset.MoveFirst
        While Not vRecordset.EOF
            Set PRListItem = ListView101.ListItems.Add(, , Trim(vRecordset.Fields("backorderno").Value))
            PRListItem.SubItems(1) = Trim(vRecordset.Fields("PRNO").Value)
            PRListItem.SubItems(2) = Trim(vRecordset.Fields("createdatetime").Value)
            PRListItem.SubItems(3) = Trim(vRecordset.Fields("arname").Value)
        vRecordset.MoveNext
        Wend
    Else
        MsgBox "ไม่มีข้อมูลที่ต้องการดู"
    End If
    vRecordset.Close
Else
MsgBox "กรุณาใส่ข้อมูลในการดูข้อมูลด้วยครับ"
End If
End If

End Sub
