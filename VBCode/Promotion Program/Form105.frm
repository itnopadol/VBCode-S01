VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form Form105 
   Caption         =   "สร้างใบขอยกเลิกสินค้าโปรโมชั่น"
   ClientHeight    =   8985
   ClientLeft      =   5715
   ClientTop       =   1020
   ClientWidth     =   14385
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8985
   ScaleWidth      =   14385
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Pic101 
      Height          =   7980
      Left            =   225
      ScaleHeight     =   7920
      ScaleWidth      =   11475
      TabIndex        =   19
      Top             =   180
      Visible         =   0   'False
      Width           =   11535
      Begin VB.CommandButton CMD108 
         Caption         =   "ยกเลิก"
         Height          =   420
         Left            =   9720
         TabIndex        =   24
         Top             =   5805
         Width           =   1095
      End
      Begin VB.CommandButton CMD107 
         Caption         =   "ตกลง"
         Height          =   420
         Left            =   8325
         TabIndex        =   23
         Top             =   5805
         Width           =   1095
      End
      Begin MSComctlLib.ListView ListView102 
         Height          =   3975
         Left            =   765
         TabIndex        =   22
         Top             =   1530
         Width           =   10050
         _ExtentX        =   17727
         _ExtentY        =   7011
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ลำดับ"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "เลขที่เอกสาร"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "วันที่เอกสาร"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "เหตุผล"
            Object.Width           =   9172
         EndProperty
      End
      Begin VB.TextBox TextSearch101 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1485
         TabIndex        =   21
         Top             =   675
         Width           =   2580
      End
      Begin VB.Label Label5 
         Caption         =   "ค้นหา :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   810
         TabIndex        =   20
         Top             =   720
         Width           =   825
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2520
      Top             =   7695
   End
   Begin Crystal.CrystalReport Crystal101 
      Left            =   675
      Top             =   7650
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
   Begin VB.CommandButton CMD106 
      Caption         =   "พิมพ์"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   10395
      TabIndex        =   10
      Top             =   7515
      Width           =   1275
   End
   Begin VB.CommandButton CMD105 
      Caption         =   "บันทึก"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   8910
      TabIndex        =   9
      Top             =   7515
      Width           =   1275
   End
   Begin VB.Frame Frame3 
      Caption         =   "ส่วนของเหตุผลการขอยกเลิก"
      Height          =   1590
      Left            =   270
      TabIndex        =   13
      Top             =   5715
      Width           =   11400
      Begin VB.TextBox Text103 
         Appearance      =   0  'Flat
         Height          =   825
         Left            =   1755
         MultiLine       =   -1  'True
         TabIndex        =   8
         Top             =   450
         Width           =   9105
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "เหตุผล :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   585
         TabIndex        =   16
         Top             =   450
         Width           =   960
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "ส่วนของรายการสินค้าโปรโมชั่น"
      Height          =   3795
      Left            =   270
      TabIndex        =   12
      Top             =   1800
      Width           =   11400
      Begin VB.CheckBox Check101 
         Caption         =   "เลือกทั้งหมด"
         Height          =   330
         Left            =   540
         TabIndex        =   18
         Top             =   3375
         Width           =   1500
      End
      Begin MSComctlLib.ListView ListView101 
         Height          =   2850
         Left            =   540
         TabIndex        =   7
         Top             =   450
         Width           =   10365
         _ExtentX        =   18283
         _ExtentY        =   5027
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ลำดับ"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "รหัสสินค้า"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "ชื่อสินค้า"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "ราคาปกติ"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "ราคาโปรโมชั่น"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "หน่วยนับ"
            Object.Width           =   1764
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "เอกสารขอยกเลิกโปรโมชั่น"
      Height          =   1455
      Left            =   270
      TabIndex        =   11
      Top             =   225
      Width           =   11400
      Begin VB.CommandButton CMD104 
         Height          =   330
         Left            =   9045
         Picture         =   "Form105.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   360
         Width           =   330
      End
      Begin VB.TextBox Text102 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   7470
         TabIndex        =   5
         Top             =   360
         Width           =   1545
      End
      Begin VB.CommandButton CMD103 
         Height          =   330
         Left            =   3645
         Picture         =   "Form105.frx":031A
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   360
         Width           =   330
      End
      Begin VB.CommandButton CMD102 
         Height          =   330
         Left            =   4365
         Picture         =   "Form105.frx":06FF
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   360
         Width           =   330
      End
      Begin VB.CommandButton CMD101 
         Height          =   330
         Left            =   4005
         Picture         =   "Form105.frx":0ACC
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   360
         Width           =   330
      End
      Begin MSComCtl2.DTPicker DTPicker101 
         Height          =   375
         Left            =   2070
         TabIndex        =   4
         Top             =   900
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   60555265
         CurrentDate     =   38925
      End
      Begin VB.TextBox Text101 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2070
         TabIndex        =   0
         Top             =   360
         Width           =   1545
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "เลขที่โปรโมชั่น :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5760
         TabIndex        =   17
         Top             =   405
         Width           =   1635
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "วันที่เอกสาร :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   585
         TabIndex        =   15
         Top             =   945
         Width           =   1410
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "เลขที่เอกสาร :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   540
         TabIndex        =   14
         Top             =   405
         Width           =   1455
      End
   End
End
Attribute VB_Name = "Form105"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vNEwDocno As String
Dim vQuery As String
Dim vIsOpen As Integer

Private Sub Check101_Click()
Dim vCount As Integer
Dim i As Integer

If ListView101.ListItems.Count > 0 Then
  vCount = ListView101.ListItems.Count
  
  If Check101.Value = 1 Then
    For i = 1 To vCount
      ListView101.ListItems(i).Checked = True
      ListView101.ListItems(i).ForeColor = "&H000000FF" 'Red
      ListView101.ListItems.Item(i).ListSubItems(1).ForeColor = "&H000000FF"
      ListView101.ListItems.Item(i).ListSubItems(2).ForeColor = "&H000000FF"
      ListView101.ListItems.Item(i).ListSubItems(3).ForeColor = "&H000000FF"
      ListView101.ListItems.Item(i).ListSubItems(4).ForeColor = "&H000000FF"
    Next i
  Else
      For i = 1 To vCount
      ListView101.ListItems(i).Checked = False
      ListView101.ListItems(i).ForeColor = "&H80000008" 'Black
      ListView101.ListItems.Item(i).ListSubItems(1).ForeColor = "&H80000008"
      ListView101.ListItems.Item(i).ListSubItems(2).ForeColor = "&H80000008"
      ListView101.ListItems.Item(i).ListSubItems(3).ForeColor = "&H80000008"
      ListView101.ListItems.Item(i).ListSubItems(4).ForeColor = "&H80000008"
    Next i
  End If
End If

End Sub

Private Sub CMD101_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String

On Error GoTo ErrDescription

vCheckJob1 = 1
vQuery = "execute USP_PM_RequestCancelNewDocNo"
If OpenDatabase(gConnection, vRecordset, vQuery) <> 0 Then
    vNEwDocno = Trim(vRecordset.Fields("newdocno").Value)
End If
vRecordset.Close
Text101.Text = UCase(vNEwDocno)
vIsOpen = 0

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub CMD102_Click()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vListItem As ListItem
Dim i As Integer

ListView102.ListItems.Clear
vQuery = "exec dbo.USP_PM_SearchRequestCancel 0,'' "
If OpenDatabase(gConnection, vRecordset, vQuery) <> 0 Then
vRecordset.MoveFirst
i = 1
While Not vRecordset.EOF
Set vListItem = ListView102.ListItems.Add(, , i)
vListItem.SubItems(1) = Trim(vRecordset.Fields("docno").Value)
vListItem.SubItems(2) = Trim(vRecordset.Fields("docdate").Value)
vListItem.SubItems(3) = Trim(vRecordset.Fields("CauseDescription").Value)
vRecordset.MoveNext
i = i + 1
Wend
End If
vRecordset.Close
Pic101.Visible = True
TextSearch101.SetFocus
End Sub

Private Sub CMD103_Click()
  Text102.Enabled = True
  CMD101.Enabled = True
  CMD104.Enabled = True
  DTPicker101.Enabled = True
  ListView101.Enabled = True
  Check101.Enabled = True
  CMD105.Enabled = True
  CMD106.Enabled = True
  Text103.Enabled = True
      
  Text101.Text = ""
  Text102.Text = ""
  Text103.Text = ""
  DTPicker101 = Now
  ListView101.ListItems.Clear
  vIsOpen = 0
End Sub

Private Sub CMD104_Click()
Dim vRecordset As New ADODB.Recordset
Dim vReqNo As String
Dim vItemCode As String
Dim i  As Integer
Dim vListItem As ListItem

On Error GoTo ErrDescription

vReqNo = Trim(Text102.Text)
If vReqNo <> "" Then
  vIsOpen = 0
  vQuery = "exec dbo.USP_PM_SearchItemPromotion '" & vReqNo & "' "
  If OpenDatabase(gConnection, vRecordset, vQuery) <> 0 Then
  ListView101.ListItems.Clear
  vRecordset.MoveFirst
  i = 1
  While Not vRecordset.EOF
  Set vListItem = ListView101.ListItems.Add(, , i)
  vListItem.SubItems(1) = Trim(vRecordset.Fields("itemcode").Value)
  vListItem.SubItems(2) = Trim(vRecordset.Fields("itemname").Value)
  vListItem.SubItems(3) = Format(Trim(vRecordset.Fields("price").Value), "##,##0.00")
  vListItem.SubItems(4) = Format(Trim(vRecordset.Fields("promoprice").Value), "##,##0.00")
  vListItem.SubItems(5) = Trim(vRecordset.Fields("unitcode").Value)
  i = i + 1
  vRecordset.MoveNext
  Wend
  ListView101.SetFocus
  Else
    MsgBox "ไม่สามารถค้นหารายการสินค้าของรายการโปรโมชั่นที่กรอกเข้ามาได้ สินค้าที่ต้องการยกเลิกต้องอนุมัติเรียบร้อยแล้ว กรุณาตรวจอีกครั้ง", vbCritical, "Send Message"
    ListView101.ListItems.Clear
    Text102.Text = ""
  End If
  vRecordset.Close

Else
  MsgBox "กรุณากรอกข้อมูลเลขที่ใบเสนอสินค้าโปรโมชั่นด้วย", vbCritical, "Send Message"
End If


ErrDescription:
If Err.Description <> "" Then
  MsgBox Err.Description
  Exit Sub
End If
End Sub

Private Sub CMD105_Click()
Dim vRecordset As New ADODB.Recordset
Dim vDocno As String
Dim vDocDate As String
Dim vPromoCode As String
Dim vCancelCode As String
Dim vCauseDescription As String
Dim i As Integer
Dim vCount As Integer
Dim vRequestNo As String
Dim vLine As Integer
Dim vItemCode As String
Dim vItemName As String
Dim vPrice As Currency
Dim vUnitCode As String
Dim vLineNumber As Integer


vCount = 0
vRequestNo = Trim(Text102.Text)

If Text101.Text = "" Then
  MsgBox "กรุณากดรันเลขที่เอกสารขอยกเลิกโปรโมชั่นด้วย", vbCritical, "Send Error"
  Exit Sub
Else
  vDocno = Trim(Text101.Text)
End If

For i = 1 To ListView101.ListItems.Count
  If ListView101.ListItems.Item(i).Checked = True Then
    vCount = vCount + 1
  End If
Next i

If Text103.Text = "" Then
  MsgBox "กรุณากรอกเหตุผลของการขอยกเลิกโปรโมชั่นด้วย", vbCritical, "Send Error"
  Exit Sub
Else
  vCauseDescription = Trim(Text103.Text)
End If

If vCount > 0 Then
  
  vDocDate = DTPicker101
  
  vQuery = "begin tran"
  gConnection.Execute vQuery
  
  vQuery = "exec dbo.USP_PM_InsertRequestCancel '" & vDocno & "','" & vDocDate & "','" & vRequestNo & "','" & vUserID & "','" & vCauseDescription & "' "
  gConnection.Execute vQuery

  vLine = -1
  For i = 1 To ListView101.ListItems.Count
    If ListView101.ListItems.Item(i).Checked = True Then
      vItemCode = Trim(ListView101.ListItems.Item(i).SubItems(1))
      vItemName = Trim(ListView101.ListItems.Item(i).SubItems(2))
      vPrice = Trim(ListView101.ListItems.Item(i).SubItems(3))
      vUnitCode = Trim(ListView101.ListItems.Item(i).SubItems(5))
      vLine = vLine + 1
      vQuery = "dbo.USP_PM_InsertRequestSubCancel '" & vDocno & "','" & vDocDate & "','" & vItemCode & "','" & vItemName & "'," & vPrice & ",'" & vUnitCode & "'," & vLine & " "
      gConnection.Execute vQuery
    End If
  Next i
  
  vQuery = "commit tran"
  gConnection.Execute vQuery
  
  MsgBox "", vbInformation, "Send Information"
  
ErrDescription:
  If Err.Description <> "" Then
    vQuery = "rollback tran"
    gConnection.Execute vQuery
    MsgBox Err.Description
  End If
  
  Call ClearScreen
  
Else
  MsgBox "ยังไม่ได้เลือกสินค้าที่จะยกเลิก"
End If

End Sub

Private Sub CMD106_Click()
Dim vRecordset As New ADODB.Recordset
Dim vRepType As String
Dim vRepID As Integer
Dim vDocno As String
Dim vReportName As String

If Text101.Text <> "" And ListView101.ListItems.Count > 0 And vIsOpen = 1 Then
  vDocno = Trim(Text101.Text)
  vRepType = "PM"
  vRepID = 326
    vQuery = "select reportname from bcnp.dbo.bcreportname where repid = " & vRepID & " and reptype = '" & vRepType & "' "
    If OpenDatabase(gConnection, vRecordset, vQuery) <> 0 Then
        vReportName = Trim(vRecordset.Fields("reportname").Value)
    End If
    vRecordset.Close
    
    With Crystal101
    .ReportFileName = vReportName & ".rpt"
    .ParameterFields(0) = "@vDocno;" & vDocno & ";true"
    .Destination = crptToWindow
    .WindowState = crptMaximized
    .Action = 1
  End With
  End If
End Sub

Private Sub CMD107_Click()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vListItem As ListItem
Dim i As Integer
Dim vDocno As String
Dim vIndex As Integer

If ListView102.ListItems.Count > 0 Then
  vDocno = Trim(ListView102.SelectedItem.SubItems(1))
  vQuery = "exec dbo.USP_PM_SearchRequestDetails '" & vDocno & "' "
  If OpenDatabase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    vIsOpen = 1
    Text101.Text = Trim(vRecordset.Fields("docno").Value)
    Text102.Text = Trim(vRecordset.Fields("promocode").Value)
    Text103.Text = Trim(vRecordset.Fields("CauseDescription").Value)
    DTPicker101.Value = Trim(vRecordset.Fields("docdate").Value)
    i = 1
    ListView101.ListItems.Clear
    While Not vRecordset.EOF
      Set vListItem = ListView101.ListItems.Add(, , i)
      vListItem.SubItems(1) = Trim(vRecordset.Fields("itemcode").Value)
      vListItem.SubItems(2) = Trim(vRecordset.Fields("itemname").Value)
      vListItem.SubItems(3) = Format(Trim(vRecordset.Fields("price").Value), "##,##0.00")
      vListItem.SubItems(4) = Trim(vRecordset.Fields("promoprice").Value)
      vListItem.SubItems(5) = Trim(vRecordset.Fields("unitcode").Value)
      ListView101.ListItems.Item(i).Checked = True
    vRecordset.MoveNext
    i = i + 1
    Wend
    
  End If
  vRecordset.Close
  Text102.Enabled = False
  CMD101.Enabled = False
  CMD104.Enabled = False
  Pic101.Visible = False
End If
End Sub

Private Sub CMD108_Click()
Pic101.Visible = False
End Sub

Private Sub Form_Load()
DTPicker101 = Now
End Sub

Private Sub ListView101_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim vIndex As Integer

vIndex = Item.Index

If ListView101.ListItems.Item(vIndex).Checked = True Then
ListView101.ListItems(vIndex).ForeColor = "&H000000FF" 'Red
ListView101.ListItems.Item(vIndex).ListSubItems(1).ForeColor = "&H000000FF"
ListView101.ListItems.Item(vIndex).ListSubItems(2).ForeColor = "&H000000FF"
ListView101.ListItems.Item(vIndex).ListSubItems(3).ForeColor = "&H000000FF"
ListView101.ListItems.Item(vIndex).ListSubItems(4).ForeColor = "&H000000FF"
Else
  ListView101.ListItems(vIndex).ForeColor = "&H80000008"  'Black
  ListView101.ListItems.Item(vIndex).ListSubItems(1).ForeColor = "&H80000008"
  ListView101.ListItems.Item(vIndex).ListSubItems(2).ForeColor = "&H80000008"
  ListView101.ListItems.Item(vIndex).ListSubItems(3).ForeColor = "&H80000008"
  ListView101.ListItems.Item(vIndex).ListSubItems(4).ForeColor = "&H80000008"
End If
End Sub


Private Sub ListView102_DblClick()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vListItem As ListItem
Dim i As Integer
Dim vDocno As String
Dim vIndex As Integer
Dim vIsApprove As Integer

If ListView102.ListItems.Count > 0 Then
  vDocno = Trim(ListView102.SelectedItem.SubItems(1))
  vQuery = "exec dbo.USP_PM_SearchRequestDetails '" & vDocno & "' "
  If OpenDatabase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    vIsOpen = 1
    vIsApprove = Trim(vRecordset.Fields("iscancelapprove").Value)
    Text101.Text = Trim(vRecordset.Fields("docno").Value)
    Text102.Text = Trim(vRecordset.Fields("promocode").Value)
    Text103.Text = Trim(vRecordset.Fields("CauseDescription").Value)
    DTPicker101.Value = Trim(vRecordset.Fields("docdate").Value)
    i = 1
    ListView101.ListItems.Clear
    While Not vRecordset.EOF
      Set vListItem = ListView101.ListItems.Add(, , i)
      vListItem.SubItems(1) = Trim(vRecordset.Fields("itemcode").Value)
      vListItem.SubItems(2) = Trim(vRecordset.Fields("itemname").Value)
      vListItem.SubItems(3) = Format(Trim(vRecordset.Fields("price").Value), "##,##0.00")
      vListItem.SubItems(4) = Trim(vRecordset.Fields("promoprice").Value)
      vListItem.SubItems(5) = Trim(vRecordset.Fields("unitcode").Value)
      ListView101.ListItems.Item(i).Checked = True
    vRecordset.MoveNext
    i = i + 1
    Wend
    
  End If
  vRecordset.Close
  
  If vIsApprove = 1 Then
      DTPicker101.Enabled = False
      ListView101.Enabled = False
      Check101.Enabled = False
      CMD105.Enabled = False
      CMD106.Enabled = False
      Text103.Enabled = False
  Else
      DTPicker101.Enabled = True
      ListView101.Enabled = True
      Check101.Enabled = True
      CMD105.Enabled = True
      CMD106.Enabled = True
      Text103.Enabled = True
  End If
  
  Text102.Enabled = False
  CMD101.Enabled = False
  CMD104.Enabled = False
  Pic101.Visible = False
End If


End Sub

Private Sub ListView102_KeyPress(KeyAscii As Integer)
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vListItem As ListItem
Dim i As Integer
Dim vDocno As String
Dim vIndex As Integer

If KeyAscii = 13 Then
    If ListView102.ListItems.Count > 0 Then
      vDocno = Trim(ListView102.SelectedItem.SubItems(1))
      vQuery = "exec dbo.USP_PM_SearchRequestDetails '" & vDocno & "' "
      If OpenDatabase(gConnection, vRecordset, vQuery) <> 0 Then
        vRecordset.MoveFirst
        vIsOpen = 1
        Text101.Text = Trim(vRecordset.Fields("docno").Value)
        Text102.Text = Trim(vRecordset.Fields("promocode").Value)
        Text103.Text = Trim(vRecordset.Fields("CauseDescription").Value)
        DTPicker101.Value = Trim(vRecordset.Fields("docdate").Value)
        i = 1
        ListView101.ListItems.Clear
        While Not vRecordset.EOF
          Set vListItem = ListView101.ListItems.Add(, , i)
          vListItem.SubItems(1) = Trim(vRecordset.Fields("itemcode").Value)
          vListItem.SubItems(2) = Trim(vRecordset.Fields("itemname").Value)
          vListItem.SubItems(3) = Format(Trim(vRecordset.Fields("price").Value), "##,##0.00")
          vListItem.SubItems(4) = Trim(vRecordset.Fields("promoprice").Value)
          vListItem.SubItems(5) = Trim(vRecordset.Fields("unitcode").Value)
          ListView101.ListItems.Item(i).Checked = True
        vRecordset.MoveNext
        i = i + 1
        Wend
        
      End If
      vRecordset.Close
      Text102.Enabled = False
      CMD101.Enabled = False
      CMD104.Enabled = False
      Pic101.Visible = False
    End If
End If
End Sub

Private Sub Text102_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If Text102.Text <> "" Then
    Call CMD104_Click
  End If
End If
End Sub

Private Sub TextSearch101_Change()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vListItem As ListItem
Dim i As Integer
Dim vDocno As String

ListView102.ListItems.Clear
If TextSearch101.Text <> "" Then
  vDocno = Trim(TextSearch101.Text)
  vQuery = "exec dbo.USP_PM_SearchRequestCancel 1,'" & vDocno & "' "
Else
  vQuery = "exec dbo.USP_PM_SearchRequestCancel 0,'' "
End If
If OpenDatabase(gConnection, vRecordset, vQuery) <> 0 Then
vRecordset.MoveFirst
i = 1
While Not vRecordset.EOF
Set vListItem = ListView102.ListItems.Add(, , i)
vListItem.SubItems(1) = Trim(vRecordset.Fields("docno").Value)
vListItem.SubItems(2) = Trim(vRecordset.Fields("docdate").Value)
vListItem.SubItems(3) = Trim(vRecordset.Fields("CauseDescription").Value)
vRecordset.MoveNext
i = i + 1
Wend
End If
vRecordset.Close

End Sub

Private Sub Timer1_Timer()
If Check101.Value = 1 Then
  Check101.Value = 0
Else
  Check101.Value = 1
End If
End Sub

Public Sub ClearScreen()
Text101.Text = ""
Text102.Text = ""
Text103.Text = ""
ListView101.ListItems.Clear
End Sub
