VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form2_5 
   Caption         =   "ดึงทะเบียนสินค้าที่ยกเลิกขายนำกลับมาใช้ใหม่"
   ClientHeight    =   8115
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form2_5.frx":0000
   ScaleHeight     =   8115
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin VB.PictureBox PICSearchItem 
      BackColor       =   &H00FFFFFF&
      Height          =   11625
      Left            =   -90
      ScaleHeight     =   11565
      ScaleWidth      =   12060
      TabIndex        =   11
      Top             =   1215
      Visible         =   0   'False
      Width           =   12120
      Begin VB.CommandButton CMDClose 
         Caption         =   "ปิด"
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
         Left            =   10080
         TabIndex        =   17
         Top             =   5760
         Width           =   1320
      End
      Begin VB.CommandButton CMDSelect 
         Caption         =   "เลือก"
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
         Left            =   8550
         TabIndex        =   16
         Top             =   5760
         Width           =   1320
      End
      Begin VB.CommandButton BTNClickSearchItem 
         Height          =   285
         Left            =   6345
         Picture         =   "Form2_5.frx":72FB
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   405
         Width           =   285
      End
      Begin MSComctlLib.ListView ListViewItemCode 
         Height          =   4380
         Left            =   720
         TabIndex        =   14
         Top             =   1080
         Width           =   10680
         _ExtentX        =   18838
         _ExtentY        =   7726
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ลำดับ"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "รหัสสินค้า"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "ชื่อสินค้า"
            Object.Width           =   7585
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "หน่วย"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "สถานะสินค้า"
            Object.Width           =   3528
         EndProperty
      End
      Begin VB.TextBox TXTSearchItem 
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
         Height          =   300
         Left            =   1665
         TabIndex        =   13
         Top             =   405
         Width           =   4605
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "คำที่ค้นหา :"
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
         Left            =   720
         TabIndex        =   12
         Top             =   450
         Width           =   960
      End
   End
   Begin MSComctlLib.ListView ListViewStatus 
      Height          =   1275
      Left            =   2475
      TabIndex        =   19
      Top             =   2700
      Visible         =   0   'False
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   2249
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "สถานะ"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.CommandButton CMDStatus 
      Height          =   285
      Left            =   4275
      Picture         =   "Form2_5.frx":76C8
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   2430
      Width           =   285
   End
   Begin VB.CommandButton CMDExit 
      Caption         =   "เคลียร์หน้าจอ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   4410
      TabIndex        =   10
      Top             =   4500
      Width           =   1770
   End
   Begin VB.CommandButton CMDUpdate 
      Caption         =   "ปรับข้อมูล"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   2520
      TabIndex        =   9
      Top             =   4500
      Width           =   1770
   End
   Begin VB.CommandButton CMDSearchItem 
      Height          =   285
      Left            =   4275
      Picture         =   "Form2_5.frx":7A39
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1575
      Width           =   285
   End
   Begin VB.TextBox TXTItem 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2475
      TabIndex        =   7
      Top             =   1575
      Width           =   1770
   End
   Begin VB.Line Line1 
      BorderWidth     =   5
      Index           =   1
      X1              =   0
      X2              =   12015
      Y1              =   3645
      Y2              =   3645
   End
   Begin VB.Line Line1 
      BorderWidth     =   5
      Index           =   0
      X1              =   0
      X2              =   12015
      Y1              =   1170
      Y2              =   1170
   End
   Begin VB.Label LBLActiveStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   2475
      TabIndex        =   6
      Top             =   2880
      Width           =   1770
   End
   Begin VB.Label LBLItemStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   2475
      TabIndex        =   5
      Top             =   2430
      Width           =   1770
   End
   Begin VB.Label LBLItemName 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   2475
      TabIndex        =   4
      Top             =   1980
      Width           =   8340
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "สถานะสินค้า :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   990
      TabIndex        =   3
      Top             =   2880
      Width           =   1320
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "สถานะการขาย :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   450
      TabIndex        =   2
      Top             =   2430
      Width           =   1860
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ชื่อสินค้า :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   1125
      TabIndex        =   1
      Top             =   1980
      Width           =   1185
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "รหัสสินค้า :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1170
      TabIndex        =   0
      Top             =   1575
      Width           =   1140
   End
End
Attribute VB_Name = "Form2_5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vQuery As String
Dim vMemSelectIndex As Integer

Private Sub BTNClickSearchItem_Click()
Dim vRecordset As New ADODB.Recordset
Dim vSearch As String
Dim vListItems As ListItem
Dim i As Integer

On Error Resume Next

If Me.TXTSearchItem.Text <> "" Then
   vSearch = Me.TXTSearchItem.Text
Else
   Exit Sub
End If

Me.ListViewItemCode.ListItems.Clear
vQuery = "exec dbo.USP_MB_SearchItemUpdateActive 0 ,'" & vSearch & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
   vRecordset.MoveFirst
   While Not vRecordset.EOF
       i = i + 1
       Set vListItems = Me.ListViewItemCode.ListItems.Add(, , i)
               vListItems.SubItems(1) = Trim(vRecordset.Fields("itemcode").Value)
               vListItems.SubItems(2) = Trim(vRecordset.Fields("itemname").Value)
               vListItems.SubItems(3) = Trim(vRecordset.Fields("defstkunitcode").Value)
               vListItems.SubItems(4) = Trim(vRecordset.Fields("statusname").Value)
               vRecordset.MoveNext
   Wend
End If
vRecordset.Close

End Sub

Private Sub CMDClose_Click()
Me.PICSearchItem.Visible = False
End Sub

Private Sub CMDExit_Click()
      Me.TXTItem.Text = ""
      Me.LBLItemName.Caption = ""
      Me.LBLItemStatus.Caption = ""
      Me.LBLActiveStatus.Caption = ""
End Sub

Private Sub CMDSearchItem_Click()
Me.PICSearchItem.Visible = True
Me.TXTSearchItem.SetFocus
End Sub

Private Sub CMDSelect_Click()
On Error Resume Next

If Me.ListViewItemCode.ListItems.Count > 0 Then
   Me.TXTItem.Text = Me.ListViewItemCode.ListItems(vMemSelectIndex).SubItems(1)
   Me.LBLItemName.Caption = Me.ListViewItemCode.ListItems(vMemSelectIndex).SubItems(2)
   Me.LBLItemStatus.Caption = Me.ListViewItemCode.ListItems(vMemSelectIndex).SubItems(4)
   Me.LBLActiveStatus.Caption = "สินค้ายกเลิก"
   Me.PICSearchItem.Visible = False
End If
End Sub

Private Sub CMDStatus_Click()
ListViewStatus.Visible = True
End Sub

Private Sub CMDUpdate_Click()
Dim vAnswer As Integer
Dim vItemCode As String
Dim vItemStatus As Integer

On Error GoTo ErrDescription


If Me.TXTItem.Text <> "" And Me.LBLItemName.Caption <> "" Then
   vAnswer = MsgBox("คุณต้องการดึงข้อมูลสินค้ากลับมาใช้ใช่หรือไม่", vbYesNo, "Send Question Message")
   If vAnswer = 6 Then
      vItemCode = Me.TXTItem.Text
     If Me.LBLItemStatus.Caption = "หยุดขาย" Then
      vItemStatus = 0
    ElseIf Me.LBLItemStatus.Caption = "สต็อกขาย" Then
      vItemStatus = 1
    ElseIf Me.LBLItemStatus.Caption = "หยุดซื้อ/เลิกผลิต" Then
      vItemStatus = 2
    ElseIf Me.LBLItemStatus.Caption = "สั่งพิเศษ" Then
      vItemStatus = 3
    ElseIf Me.LBLItemStatus.Caption = "ของแถม" Then
      vItemStatus = 4
    End If

      vQuery = "exec dbo.USP_IV_UpdateItemActiveStatus '" & vItemCode & "'," & vItemStatus & ",'" & vUserID & "' "
      gConnection.Execute (vQuery)
      
      MsgBox "ได้ทำการดึงข้อมูลสินค้ากลับมาให้ใช้ใหม่เรียบร้อยแล้ว", vbInformation, "Send Information Message"
      Me.TXTItem.Text = ""
      Me.LBLItemName.Caption = ""
      Me.LBLItemStatus.Caption = ""
      Me.LBLActiveStatus.Caption = ""
   End If
End If

ErrDescription:
If Err.Description <> "" Then
   MsgBox Err.Description
   Exit Sub
End If
End Sub

Private Sub Form_Load()
Call vGetStatus
End Sub

Private Sub vGetStatus()
Dim vListStatus As ListItem

On Error Resume Next

Set vListStatus = Me.ListViewStatus.ListItems.Add(, , "หยุดขาย")
Set vListStatus = Me.ListViewStatus.ListItems.Add(, , "สต็อกขาย")
Set vListStatus = Me.ListViewStatus.ListItems.Add(, , "หยุดซื้อ/เลิกผลิต")
Set vListStatus = Me.ListViewStatus.ListItems.Add(, , "สั่งพิเศษ")
Set vListStatus = Me.ListViewStatus.ListItems.Add(, , "ของแถม")

End Sub
Private Sub ListViewItemCode_DblClick()
On Error Resume Next

If Me.ListViewItemCode.ListItems.Count > 0 Then
   vMemSelectIndex = Me.ListViewItemCode.SelectedItem.Index
   Me.TXTItem.Text = Me.ListViewItemCode.ListItems(vMemSelectIndex).SubItems(1)
   Me.LBLItemName.Caption = Me.ListViewItemCode.ListItems(vMemSelectIndex).SubItems(2)
   Me.LBLItemStatus.Caption = Me.ListViewItemCode.ListItems(vMemSelectIndex).SubItems(4)
   Me.LBLActiveStatus.Caption = "สินค้ายกเลิก"
   Me.PICSearchItem.Visible = False
End If
End Sub

Private Sub ListViewItemCode_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error Resume Next

If Me.ListViewItemCode.ListItems.Count > 0 Then
   vMemSelectIndex = Me.ListViewItemCode.SelectedItem.Index
End If
End Sub

Private Sub ListViewStatus_DblClick()
Dim vIndex As Integer

On Error Resume Next

If Me.ListViewStatus.ListItems.Count > 0 Then
   vIndex = Me.ListViewStatus.SelectedItem.Index
   Me.LBLItemStatus.Caption = Me.ListViewStatus.ListItems(vIndex).Text
   Me.ListViewStatus.Visible = False
End If
End Sub

Private Sub TXTItem_Change()
Dim vRecordset As New ADODB.Recordset
Dim vSearch As String

On Error Resume Next

vSearch = Me.TXTItem.Text
vQuery = "exec dbo.USP_MB_SearchItemUpdateActive 1 ,'" & vSearch & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
   Me.LBLItemName.Caption = Trim(vRecordset.Fields("itemname").Value)
   Me.LBLItemStatus.Caption = Trim(vRecordset.Fields("statusname").Value)
   Me.LBLActiveStatus.Caption = "สินค้ายกเลิก"
Else
   Me.LBLItemName.Caption = ""
   Me.LBLItemStatus.Caption = ""
   Me.LBLActiveStatus.Caption = ""
End If
vRecordset.Close
End Sub

Private Sub TXTSearchItem_KeyPress(KeyAscii As Integer)
Dim vRecordset As New ADODB.Recordset
Dim vSearch As String
Dim vListItems As ListItem
Dim i As Integer

On Error Resume Next

If KeyAscii = 13 Then
   If Me.TXTSearchItem.Text <> "" Then
      vSearch = Me.TXTSearchItem.Text
   Else
      Exit Sub
   End If
   
   Me.ListViewItemCode.ListItems.Clear
   vQuery = "exec dbo.USP_MB_SearchItemUpdateActive 0 ,'" & vSearch & "' "
   If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
      vRecordset.MoveFirst
      While Not vRecordset.EOF
          i = i + 1
          Set vListItems = Me.ListViewItemCode.ListItems.Add(, , i)
                  vListItems.SubItems(1) = Trim(vRecordset.Fields("itemcode").Value)
                  vListItems.SubItems(2) = Trim(vRecordset.Fields("itemname").Value)
                  vListItems.SubItems(3) = Trim(vRecordset.Fields("defstkunitcode").Value)
                  vListItems.SubItems(4) = Trim(vRecordset.Fields("statusname").Value)
                  vRecordset.MoveNext
      Wend
   End If
   vRecordset.Close
End If
End Sub
