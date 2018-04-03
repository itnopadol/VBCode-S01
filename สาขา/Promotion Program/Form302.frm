VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form302 
   Caption         =   "ยกเลิกโปรโมชั่น"
   ClientHeight    =   9000
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Pic101 
      Height          =   6945
      Left            =   900
      ScaleHeight     =   6885
      ScaleWidth      =   10080
      TabIndex        =   17
      Top             =   675
      Visible         =   0   'False
      Width           =   10140
      Begin VB.CommandButton CMD105 
         Caption         =   "ออก"
         Height          =   420
         Left            =   8865
         TabIndex        =   22
         Top             =   4635
         Width           =   825
      End
      Begin VB.CommandButton CMD104 
         Caption         =   "ตกลง"
         Height          =   420
         Left            =   7875
         TabIndex        =   21
         Top             =   4635
         Width           =   825
      End
      Begin MSComctlLib.ListView ListView102 
         Height          =   3345
         Left            =   360
         TabIndex        =   20
         Top             =   1080
         Width           =   9330
         _ExtentX        =   16457
         _ExtentY        =   5900
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
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ลำดับ"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "เลขที่ขอยกเลิก"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "วันที่"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "ผู้ขอยกเลิก"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "เหตุผลการขอยกเลิก"
            Object.Width           =   8819
         EndProperty
      End
      Begin VB.TextBox TextSearch101 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1080
         TabIndex        =   19
         Top             =   405
         Width           =   2310
      End
      Begin VB.Label Label8 
         Caption         =   "ค้นหา :"
         Height          =   330
         Left            =   405
         TabIndex        =   18
         Top             =   450
         Width           =   600
      End
   End
   Begin VB.CommandButton CMD103 
      Caption         =   "ยกเลิก"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10170
      TabIndex        =   5
      Top             =   6930
      Width           =   825
   End
   Begin VB.CommandButton CMD102 
      Caption         =   "ตกลง"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9000
      TabIndex        =   4
      Top             =   6930
      Width           =   825
   End
   Begin VB.CommandButton CMD101 
      Height          =   285
      Left            =   4860
      Picture         =   "Form302.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1215
      Width           =   375
   End
   Begin MSComctlLib.ListView ListView101 
      Height          =   3255
      Left            =   990
      TabIndex        =   3
      Top             =   3330
      Width           =   10005
      _ExtentX        =   17648
      _ExtentY        =   5741
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
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ลำดับ"
         Object.Width           =   1147
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "รหัสสินค้า"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "ชื่อสินค้า"
         Object.Width           =   8819
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "ราคาโปรโมชั่น"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "หน่วย"
         Object.Width           =   2205
      EndProperty
   End
   Begin VB.TextBox Text101 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3060
      TabIndex        =   1
      Top             =   1215
      Width           =   1770
   End
   Begin VB.ComboBox CMB101 
      Height          =   315
      Left            =   3060
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   720
      Width           =   2985
   End
   Begin VB.Label LBLRequest 
      Height          =   285
      Left            =   3060
      TabIndex        =   16
      Top             =   1755
      Width           =   1725
   End
   Begin VB.Label Label6 
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
      Height          =   240
      Left            =   990
      TabIndex        =   15
      Top             =   1755
      Width           =   1995
   End
   Begin VB.Label LBLStop 
      Height          =   240
      Left            =   5940
      TabIndex        =   14
      Top             =   2565
      Width           =   1725
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "วันที่สิ้นสุด :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4770
      TabIndex        =   13
      Top             =   2565
      Width           =   1050
   End
   Begin VB.Label LBLStart 
      Height          =   240
      Left            =   3060
      TabIndex        =   12
      Top             =   2565
      Width           =   1590
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "วันที่เริ่มโปรโมชั่น :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   945
      TabIndex        =   11
      Top             =   2565
      Width           =   2040
   End
   Begin VB.Label LBLPMName 
      Height          =   240
      Left            =   3060
      TabIndex        =   10
      Top             =   2160
      Width           =   7980
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "ชื่อโปรโมชั่น :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   990
      TabIndex        =   9
      Top             =   2160
      Width           =   1995
   End
   Begin VB.Label Label3 
      Caption         =   "รายการสินค้า"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   990
      TabIndex        =   8
      Top             =   3060
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "เลขที่ขอยกเลิกโปรโมชั่น :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   990
      TabIndex        =   7
      Top             =   1260
      Width           =   1995
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "เลือกประเภทการยกเลิก :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   990
      TabIndex        =   6
      Top             =   765
      Width           =   1995
   End
End
Attribute VB_Name = "Form302"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vQuery As String


Private Sub CMD101_Click()
Dim vRecordset As New ADODB.Recordset
Dim i  As Integer
Dim vLIstItem As ListItem

vQuery = "exec dbo.USP_PM_SearchRequestCancelNoApprove 0,'' "
  If OpenDatabase(gConnection, vRecordset, vQuery) <> 0 Then
  ListView102.ListItems.Clear
  vRecordset.MoveFirst
  i = 1
  While Not vRecordset.EOF
  Set vLIstItem = ListView102.ListItems.Add(, , i)
  vLIstItem.SubItems(1) = Trim(vRecordset.Fields("docno").Value)
  vLIstItem.SubItems(2) = Trim(vRecordset.Fields("docdate").Value)
  vLIstItem.SubItems(3) = Trim(vRecordset.Fields("cancelcode").Value)
  vLIstItem.SubItems(4) = Trim(vRecordset.Fields("CauseDescription").Value)
  i = i + 1
  vRecordset.MoveNext
  Wend
  End If
  vRecordset.Close

Pic101.Visible = True
TextSearch101.SetFocus
End Sub

Private Sub CMD102_Click()
Dim vRecordset As New ADODB.Recordset
Dim vReqNo As String
Dim vItemCode As String
Dim i  As Integer
Dim vLIstItem As ListItem
Dim vAnswer As Integer
Dim vCancelNo As String

On Error GoTo ErrDescription

vAnswer = MsgBox("คุณต้องการยกเลิกรหัสสินค้าที่เข้าร่วมโปรโมชั่นใช่หรือไม่?", vbYesNo, "Send Massege")
If vAnswer = 6 Then
If ListView101.ListItems.Count <> 0 Then
    vCancelNo = UCase(Trim(Text101.Text))
    For i = 1 To ListView101.ListItems.Count
        vReqNo = Trim(LBLRequest.Caption)
        vItemCode = Trim(ListView101.ListItems.Item(i).SubItems(1))
        vQuery = "exec dbo.USP_PM_PromotionCancel '" & vReqNo & "' ,'" & vItemCode & "','" & vUserID & "' "
        gConnection.Execute vQuery
    Next i
    vQuery = "exec dbo.USP_PM_UpdateCancelPromotion '" & vCancelNo & "' "
    gConnection.Execute vQuery
End If
  LBLRequest.Caption = ""
  LBLPMName.Caption = ""
  LBLStart.Caption = ""
  LBLStop.Caption = ""
  ListView101.ListItems.Clear
  Text101.Text = ""
  MsgBox "ยกเลิกสินค้าโปรโมชั่นตามที่ได้เลือกเรียบร้อยแล้ว", vbInformation, "Send Message"
Else
  Exit Sub
End If


ErrDescription:
If Err.Description <> "" Then
  MsgBox Err.Description
  Exit Sub
End If

End Sub

Private Sub CMD103_Click()
Unload Form302
End Sub

Private Sub Command2_Click()

End Sub

Private Sub CMD104_Click()
Dim vRecordset As New ADODB.Recordset
Dim vReqNo As String
Dim vItemCode As String
Dim i  As Integer
Dim vLIstItem As ListItem

On Error GoTo ErrDescription


vReqNo = Trim(ListView102.SelectedItem.SubItems(1))
If vReqNo <> "" Then
  vQuery = "exec dbo.USP_PM_SearchItemCancelPromotion '" & vReqNo & "' "
  If OpenDatabase(gConnection, vRecordset, vQuery) <> 0 Then
  ListView101.ListItems.Clear
  vRecordset.MoveFirst
  i = 1
  LBLRequest.Caption = Trim(vRecordset.Fields("requestno").Value)
  LBLPMName.Caption = Trim(vRecordset.Fields("pmname").Value)
  LBLStart.Caption = Trim(vRecordset.Fields("datestart").Value)
  LBLStop.Caption = Trim(vRecordset.Fields("dateend").Value)
  While Not vRecordset.EOF
  Set vLIstItem = ListView101.ListItems.Add(, , i)
  vLIstItem.SubItems(1) = Trim(vRecordset.Fields("itemcode").Value)
  vLIstItem.SubItems(2) = Trim(vRecordset.Fields("itemname").Value)
  vLIstItem.SubItems(3) = Format(Trim(vRecordset.Fields("price").Value), "##,##0.00")
  vLIstItem.SubItems(4) = Trim(vRecordset.Fields("unitcode").Value)
  i = i + 1
  vRecordset.MoveNext
  Wend
  Pic101.Visible = False
  Else
    MsgBox "ไม่สามารถค้นหารายการสินค้าของใบขอยกเลิกโปรโมชั่นที่กรอกเข้ามาได้ อาจเกิดจากพิมพ์เลขที่ไม่ถูกต้อง หรือได้ทำการยกเลิกเอกสารดังกล่าวไปแล้ว กรุณาตรวจอีกครั้ง", vbCritical, "Send Message"
    LBLRequest.Caption = ""
    LBLPMName.Caption = ""
    LBLStart.Caption = ""
    LBLStop.Caption = ""
    ListView101.ListItems.Clear
    Text101.Text = ""
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
Pic101.Visible = False
End Sub

Private Sub Form_Load()
CMB101.AddItem Trim("ลบสินค้ารายตัว")
CMB101.Text = Trim("ลบสินค้ารายตัว")
End Sub

Private Sub ListView102_DblClick()
Dim vRecordset As New ADODB.Recordset
Dim vReqNo As String
Dim vItemCode As String
Dim i  As Integer
Dim vLIstItem As ListItem

On Error GoTo ErrDescription


vReqNo = Trim(ListView102.SelectedItem.SubItems(1))
If vReqNo <> "" Then
  vQuery = "exec dbo.USP_PM_SearchItemCancelPromotion '" & vReqNo & "' "
  If OpenDatabase(gConnection, vRecordset, vQuery) <> 0 Then
  ListView101.ListItems.Clear
  vRecordset.MoveFirst
  i = 1
  LBLRequest.Caption = Trim(vRecordset.Fields("requestno").Value)
  LBLPMName.Caption = Trim(vRecordset.Fields("pmname").Value)
  LBLStart.Caption = Trim(vRecordset.Fields("datestart").Value)
  LBLStop.Caption = Trim(vRecordset.Fields("dateend").Value)
  While Not vRecordset.EOF
  Set vLIstItem = ListView101.ListItems.Add(, , i)
  vLIstItem.SubItems(1) = Trim(vRecordset.Fields("itemcode").Value)
  vLIstItem.SubItems(2) = Trim(vRecordset.Fields("itemname").Value)
  vLIstItem.SubItems(3) = Format(Trim(vRecordset.Fields("price").Value), "##,##0.00")
  vLIstItem.SubItems(4) = Trim(vRecordset.Fields("unitcode").Value)
  i = i + 1
  vRecordset.MoveNext
  Wend
  Pic101.Visible = False
  Else
    MsgBox "ไม่สามารถค้นหารายการสินค้าของใบขอยกเลิกโปรโมชั่นที่กรอกเข้ามาได้ อาจเกิดจากพิมพ์เลขที่ไม่ถูกต้อง หรือได้ทำการยกเลิกเอกสารดังกล่าวไปแล้ว กรุณาตรวจอีกครั้ง", vbCritical, "Send Message"
    LBLRequest.Caption = ""
    LBLPMName.Caption = ""
    LBLStart.Caption = ""
    LBLStop.Caption = ""
    ListView101.ListItems.Clear
    Text101.Text = ""
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

Private Sub Text101_KeyPress(KeyAscii As Integer)
Dim vRecordset As New ADODB.Recordset
Dim vReqNo As String
Dim vItemCode As String
Dim i  As Integer
Dim vLIstItem As ListItem

On Error GoTo ErrDescription

If KeyAscii = 13 Then
  vReqNo = Trim(Text101.Text)
If vReqNo <> "" Then
  vQuery = "exec dbo.USP_PM_SearchItemCancelPromotion '" & vReqNo & "' "
  If OpenDatabase(gConnection, vRecordset, vQuery) <> 0 Then
  ListView101.ListItems.Clear
  vRecordset.MoveFirst
  i = 1
  LBLRequest.Caption = Trim(vRecordset.Fields("requestno").Value)
  LBLPMName.Caption = Trim(vRecordset.Fields("pmname").Value)
  LBLStart.Caption = Trim(vRecordset.Fields("datestart").Value)
  LBLStop.Caption = Trim(vRecordset.Fields("dateend").Value)
  While Not vRecordset.EOF
  Set vLIstItem = ListView101.ListItems.Add(, , i)
  vLIstItem.SubItems(1) = Trim(vRecordset.Fields("itemcode").Value)
  vLIstItem.SubItems(2) = Trim(vRecordset.Fields("itemname").Value)
  vLIstItem.SubItems(3) = Format(Trim(vRecordset.Fields("price").Value), "##,##0.00")
  vLIstItem.SubItems(4) = Trim(vRecordset.Fields("unitcode").Value)
  i = i + 1
  vRecordset.MoveNext
  Wend
  ListView101.SetFocus
  Else
    MsgBox "ไม่สามารถค้นหารายการสินค้าของใบขอยกเลิกโปรโมชั่นที่กรอกเข้ามาได้ อาจเกิดจากพิมพ์เลขที่ไม่ถูกต้อง หรือได้ทำการยกเลิกเอกสารดังกล่าวไปแล้ว กรุณาตรวจอีกครั้ง", vbCritical, "Send Message"
    LBLRequest.Caption = ""
    LBLPMName.Caption = ""
    LBLStart.Caption = ""
    LBLStop.Caption = ""
    ListView101.ListItems.Clear
    Text101.Text = ""
  End If
  vRecordset.Close

Else
  MsgBox "กรุณากรอกข้อมูลเลขที่ใบเสนอสินค้าโปรโมชั่นด้วย", vbCritical, "Send Message"
End If
End If


ErrDescription:
If Err.Description <> "" Then
  MsgBox Err.Description
  Exit Sub
End If
End Sub

Private Sub TextSearch101_Change()
Dim vRecordset As New ADODB.Recordset
Dim i  As Integer
Dim vLIstItem As ListItem
Dim vType As Integer
Dim vSearch As String

If TextSearch101.Text = "" Then
  vType = 0
  vSearch = ""
Else
  vType = 1
  vSearch = Trim(TextSearch101.Text)
End If
  
ListView102.ListItems.Clear

vQuery = "exec dbo.USP_PM_SearchRequestCancelNoApprove " & vType & ",'" & vSearch & "' "
  If OpenDatabase(gConnection, vRecordset, vQuery) <> 0 Then
  vRecordset.MoveFirst
  i = 1
  While Not vRecordset.EOF
  Set vLIstItem = ListView102.ListItems.Add(, , i)
  vLIstItem.SubItems(1) = Trim(vRecordset.Fields("docno").Value)
  vLIstItem.SubItems(2) = Trim(vRecordset.Fields("docdate").Value)
  vLIstItem.SubItems(3) = Trim(vRecordset.Fields("cancelcode").Value)
  vLIstItem.SubItems(4) = Trim(vRecordset.Fields("CauseDescription").Value)
  i = i + 1
  vRecordset.MoveNext
  Wend
  End If
  vRecordset.Close
End Sub
