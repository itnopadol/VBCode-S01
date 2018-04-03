VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "CRYSTL32.OCX"
Begin VB.Form Form2_52 
   Caption         =   "รายงาน แสดงสถานะของใบเสนอซื้อสินค้า"
   ClientHeight    =   7995
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   Picture         =   "Form2_52.frx":0000
   ScaleHeight     =   7995
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport Crystal101 
      Left            =   7800
      Top             =   1440
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
   Begin MSComCtl2.DTPicker DTPicker102 
      Height          =   315
      Left            =   4875
      TabIndex        =   11
      Top             =   2100
      Width           =   1740
      _ExtentX        =   3069
      _ExtentY        =   556
      _Version        =   393216
      Format          =   69795841
      CurrentDate     =   38577
   End
   Begin MSComCtl2.DTPicker DTPicker101 
      Height          =   315
      Left            =   4875
      TabIndex        =   10
      Top             =   1650
      Width           =   1740
      _ExtentX        =   3069
      _ExtentY        =   556
      _Version        =   393216
      Format          =   69795841
      CurrentDate     =   38577
   End
   Begin VB.ComboBox CMB101 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4800
      Style           =   1  'Simple Combo
      TabIndex        =   8
      Text            =   "รายงาน ตรวจสอบสถานะคงค้างใบเสนอซื้อสินค้า (PR) สำหรับบริหารสินค้า"
      Top             =   2625
      Width           =   5640
   End
   Begin VB.OptionButton Option104 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "ดูตาม รหัสลูกค้า"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   1875
      TabIndex        =   7
      Top             =   2175
      Width           =   2040
   End
   Begin VB.OptionButton Option103 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "ดูตาม รหัสพนักงาน"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   1875
      TabIndex        =   6
      Top             =   1800
      Width           =   2040
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3840
      Left            =   300
      TabIndex        =   5
      Top             =   3600
      Width           =   10140
      _ExtentX        =   17886
      _ExtentY        =   6773
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Back Order"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "เลขที่ PR"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "สถานะของเอกสาร"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "เลขที่อนุมัติ"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "เลขที่ใบสั่งซื้อ"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "เลขที่รับเข้า"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "พนักงาน"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "ชื่อลูกค้า"
         Object.Width           =   2540
      EndProperty
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
      Left            =   9225
      TabIndex        =   3
      Top             =   3075
      Width           =   1215
   End
   Begin VB.TextBox Text101 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1875
      TabIndex        =   0
      Top             =   2625
      Width           =   2040
   End
   Begin VB.OptionButton Option102 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "ดูตาม เลขที่ใบเสนอซื้อ"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   1875
      TabIndex        =   2
      Top             =   1425
      Width           =   2040
   End
   Begin VB.OptionButton Option101 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "ดูตาม เลขที่ Back Order"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   1875
      TabIndex        =   1
      Top             =   1050
      Value           =   -1  'True
      Width           =   2040
   End
   Begin VB.Label Label4 
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
      ForeColor       =   &H80000014&
      Height          =   240
      Left            =   4125
      TabIndex        =   13
      Top             =   2100
      Width           =   690
   End
   Begin VB.Label Label3 
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
      ForeColor       =   &H80000014&
      Height          =   240
      Left            =   4050
      TabIndex        =   12
      Top             =   1650
      Width           =   765
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "รายงาน"
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
      Height          =   315
      Left            =   3975
      TabIndex        =   9
      Top             =   2625
      Width           =   765
   End
   Begin VB.Label Label1 
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
      Left            =   1200
      TabIndex        =   4
      Top             =   2625
      Width           =   615
   End
End
Attribute VB_Name = "Form2_52"
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
Dim vStartDate As Date
Dim vEndDate As Date

On Error GoTo ErrDescription

If Text101.Text <> "" Then
vRepType = "PR"
vRepID = 232
vStartDate = DTPicker101.Day & "/" & DTPicker101.Month & "/" & DTPicker101.Year
vEndDate = DTPicker102.Day & "/" & DTPicker102.Month & "/" & DTPicker102.Year

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
'vQuery = "select reportname from bcreportname  where reptype = '" & vRepType & "'  and repid = " & vRepID & " "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vReportName = Trim(vRecordset.Fields("reportname").Value)
End If
vRecordset.Close

 With Crystal101
 .ReportFileName = vReportName & ".rpt"
 .ParameterFields(0) = "@vDocDate1;" & vStartDate & ";true "
 .ParameterFields(1) = "@vDocDate2;" & vEndDate & ";true "
 .Destination = crptToWindow
 .WindowState = crptMaximized
 .Action = 1
 End With
Else
    MsgBox "ยังไม่ได้เลือก ประเภทของรายงาน"
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub

Private Sub Form_Load()
CMB101.AddItem Trim("รายงาน ตรวจสอบสถานะคงค้างใบเสนอซื้อสินค้า (PR) สำหรับบริหารสินค้า")
DTPicker101.Value = Now
DTPicker102.Value = Now
End Sub

Private Sub Option101_Click()
Text101.SetFocus
ListView1.ListItems.Clear
End Sub

Private Sub Option102_Click()
Text101.SetFocus
Text101.Text = ""
ListView1.ListItems.Clear
End Sub

Private Sub Option103_Click()
Text101.SetFocus
Text101.Text = ""
ListView1.ListItems.Clear
End Sub

Private Sub Option104_Click()
Text101.SetFocus
Text101.Text = ""
ListView1.ListItems.Clear
End Sub

Private Sub Text101_KeyPress(KeyAscii As Integer)
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim PRListItem As ListItem
Dim vData As String
Dim vPRStatus As Integer
Dim vPOStatus As Integer
Dim vRVStatus As Integer
Dim i As Integer

On Error GoTo ErrDescription

If KeyAscii = 13 Then
    If Text101.Text <> "" Then
    ListView1.ListItems.Clear
    i = 0
    vData = Trim(Text101.Text)
    If Option101.Value = True Then
        vQuery = "execute USP_PRG_PRApprove '" & vData & "','' ,'','' "
    ElseIf Option102.Value = True Then
        vQuery = "execute USP_PRG_PRApprove  '','" & vData & "','' ,'' "
    ElseIf Option103.Value = True Then
    vQuery = "execute USP_PRG_PRApprove  '','','" & vData & "',''  "
    ElseIf Option104.Value = True Then
    vQuery = "execute USP_PRG_PRApprove  '','','','" & vData & "' "
    End If
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vRecordset.MoveFirst
        While Not vRecordset.EOF
            i = i + 1
            Set PRListItem = ListView1.ListItems.Add(, , Trim(vRecordset.Fields("backorderno").Value))
            If vRecordset.Fields("APNO").Value = Trim("NotApprove") Then
                vPRStatus = 0
            Else
                vPRStatus = 1
            End If
            If vRecordset.Fields("PONO").Value = Trim("NotBuy") Then
                vPOStatus = 0
            Else
                vPOStatus = 1
            End If
            If vRecordset.Fields("RVNO").Value = Trim("NotReceive") Then
                vRVStatus = 0
            Else
                vRVStatus = 1
            End If
            PRListItem.SubItems(1) = Trim(vRecordset.Fields("PRNO").Value)
            If vPRStatus = 0 Then
                PRListItem.SubItems(2) = Trim("ค้างอนุมัติ")
            ElseIf vPOStatus = 0 Then
                PRListItem.SubItems(2) = Trim("ค้างทำ PO")
            ElseIf vRVStatus = 0 Then
                PRListItem.SubItems(2) = Trim("ยังไม่ได้รับเข้า")
            ElseIf vRVStatus = 1 Then
                PRListItem.SubItems(2) = Trim("รับเข้าแล้ว")
            End If
            ListView1.ListItems(i).ListSubItems(2).ForeColor = "&H000000FF"
            If vRecordset.Fields("APNO").Value = Trim("NotApprove") Then
                PRListItem.SubItems(3) = " - "
            Else
                PRListItem.SubItems(3) = Trim(vRecordset.Fields("apno").Value)
            End If
            If vRecordset.Fields("PONO").Value = Trim("NotBuy") Then
                PRListItem.SubItems(4) = " - "
            Else
                PRListItem.SubItems(4) = Trim(vRecordset.Fields("pono").Value)
            End If
            If vRecordset.Fields("rvno").Value = Trim("NotReceive") Then
                PRListItem.SubItems(5) = " - "
            Else
                PRListItem.SubItems(5) = Trim(vRecordset.Fields("rvno").Value)
            End If
            PRListItem.SubItems(6) = Trim(vRecordset.Fields("salename").Value)
            PRListItem.SubItems(7) = Trim(vRecordset.Fields("arname").Value)
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

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub
