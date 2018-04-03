VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FormCancelAdjustStock 
   Caption         =   "ยกเลิกใบปรับปรุงหลังการตรวจนับ"
   ClientHeight    =   7605
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11835
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "FormCancelAdjustStock.frx":0000
   ScaleHeight     =   7605
   ScaleWidth      =   11835
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CMDSearchAdjust 
      Height          =   330
      Left            =   3645
      Picture         =   "FormCancelAdjustStock.frx":72FB
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1080
      Width           =   330
   End
   Begin VB.CheckBox CHKSelectAll 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "เลือกทั้งเอกสาร"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   495
      TabIndex        =   8
      Top             =   6165
      Width           =   1365
   End
   Begin VB.CommandButton CMDExit 
      Caption         =   "ออก"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   9855
      TabIndex        =   7
      Top             =   6255
      Width           =   1410
   End
   Begin VB.CommandButton CMDSave 
      Caption         =   "บันทึกยกเลิก"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   8190
      TabIndex        =   6
      Top             =   6255
      Width           =   1410
   End
   Begin MSComctlLib.ListView ListViewItemCode 
      Height          =   3615
      Left            =   495
      TabIndex        =   4
      Top             =   2475
      Width           =   10770
      _ExtentX        =   18997
      _ExtentY        =   6376
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
      Appearance      =   1
      NumItems        =   7
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
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "คลัง"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "ที่เก็บ"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "หน่วยนับ"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "ยอดปรับปรุง"
         Object.Width           =   2646
      EndProperty
   End
   Begin VB.TextBox TXTDocno 
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
      Height          =   330
      Left            =   1755
      TabIndex        =   2
      Top             =   1080
      Width           =   1860
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "รายการสินค้า "
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
      Left            =   495
      TabIndex        =   5
      Top             =   2160
      Width           =   1635
   End
   Begin VB.Label LBLDocDate 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   1755
      TabIndex        =   3
      Top             =   1530
      Width           =   1860
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "วันที่ใบปรับปรุง :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   90
      TabIndex        =   1
      Top             =   1530
      Width           =   1545
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "เลขที่ใบปรับปรุง :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   135
      TabIndex        =   0
      Top             =   1080
      Width           =   1545
   End
End
Attribute VB_Name = "FormCancelAdjustStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CHKSelectAll_Click()
Dim i As Integer

On Error Resume Next
   
If Me.ListViewItemCode.ListItems.Count > 0 Then
   If Me.CHKSelectAll.Value = 1 Then
      For i = 1 To Me.ListViewItemCode.ListItems.Count
              Me.ListViewItemCode.ListItems.Item(i).Checked = True
      Next i
   End If

   If Me.CHKSelectAll.Value = 0 Then
      For i = 1 To Me.ListViewItemCode.ListItems.Count
              Me.ListViewItemCode.ListItems.Item(i).Checked = False
      Next i
   End If
End If
End Sub

Private Sub CMDExit_Click()
Unload Me
End Sub

Private Sub CMDSave_Click()
Dim vQuery As String
Dim vDocNo As String
Dim vItemCode As String
Dim vTypeSelect As Integer

Dim vSelect As Integer
Dim i As Integer
Dim vCount As Integer

Dim n  As Integer

On Error GoTo ErrDescription

If Me.ListViewItemCode.ListItems.Count > 0 And Me.TXTDocno.Text <> "" Then
   vDocNo = Me.TXTDocno.Text
   vCount = Me.ListViewItemCode.ListItems.Count
   vSelect = 0
   For i = 1 To Me.ListViewItemCode.ListItems.Count
      If Me.ListViewItemCode.ListItems(i).Checked = True Then
         vSelect = vSelect + 1
      End If
   Next i
   
   If vSelect >= 1 Then
      If vSelect = vCount Then
      vTypeSelect = 1
      Else
      vTypeSelect = 0
      End If
      
      If vTypeSelect = 1 Then
         vItemCode = ""
         vQuery = "exec dbo.USP_MB_CancelStockAdjust '" & vDocNo & "','" & vItemCode & "'," & vTypeSelect & " "
         gConnection.Execute vQuery
      End If
      
      If vTypeSelect = 0 Then
         For n = 1 To Me.ListViewItemCode.ListItems.Count
         If Me.ListViewItemCode.ListItems(n).Checked = True Then
            vItemCode = Me.ListViewItemCode.ListItems(n).ListSubItems(1)
            vQuery = "exec dbo.USP_MB_CancelStockAdjust '" & vDocNo & "','" & vItemCode & "'," & vTypeSelect & " "
            gConnection.Execute vQuery
         End If
         Next n
      End If
      MsgBox "ยกเลิกรายการสินค้าเรียบร้อยแล้ว", vbInformation, "Send Information Message"
      Me.TXTDocno.Text = ""
      Me.LBLDocDate.Caption = ""
      Me.ListViewItemCode.ListItems.Clear
      Me.TXTDocno.SetFocus
   End If
End If

ErrDescription:
If Err.Description <> "" Then
   MsgBox Err.Description
   Exit Sub
End If
End Sub

Private Sub ListViewItemCode_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim i As Integer
Dim vCount As Integer
Dim vSelect As Integer

On Error Resume Next

If Me.ListViewItemCode.ListItems.Count > 0 Then
   vSelect = 0
   vCount = Me.ListViewItemCode.ListItems.Count
   For i = 1 To Me.ListViewItemCode.ListItems.Count
           If Me.ListViewItemCode.ListItems.Item(i).Checked = True Then
              vSelect = vSelect + 1
           End If
   Next i
   
   If vSelect = vCount Then
       Me.CHKSelectAll.Value = 1
   End If
End If
End Sub

Private Sub TXTDocno_KeyPress(KeyAscii As Integer)
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vDocNo As String
Dim vListDocNo As ListItem
Dim i As Integer

On Error Resume Next

If KeyAscii = 13 Then
   If Me.TXTDocno.Text <> "" Then
      vDocNo = UCase(Me.TXTDocno.Text)
      vQuery = "exec dbo.USP_MB_SearchStockAdjust '" & vDocNo & "' "
      If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
         Me.LBLDocDate.Caption = vRecordset.Fields("docdate").Value
         vRecordset.MoveFirst
         i = 1
         While Not vRecordset.EOF
         Set vListDocNo = Me.ListViewItemCode.ListItems.Add(, , i)
         vListDocNo.SubItems(1) = vRecordset.Fields("itemcode").Value
         vListDocNo.SubItems(2) = vRecordset.Fields("itemname").Value
         vListDocNo.SubItems(3) = vRecordset.Fields("whcode").Value
         vListDocNo.SubItems(4) = vRecordset.Fields("shelfcode").Value
         vListDocNo.SubItems(5) = vRecordset.Fields("unitcode").Value
         vListDocNo.SubItems(6) = vRecordset.Fields("qty").Value
         i = i + 1
         vRecordset.MoveNext
         Wend
      Else
         Me.ListViewItemCode.ListItems.Clear
         Me.LBLDocDate.Caption = ""
      End If
      vRecordset.Close
   End If
End If
End Sub
