VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form3_15 
   Caption         =   "ยกเลิกใบสั่งขาย"
   ClientHeight    =   8955
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form3_15.frx":0000
   ScaleHeight     =   8955
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Cmd101 
      Caption         =   "ยกเลิกเอกสาร"
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
      Left            =   2175
      TabIndex        =   2
      Top             =   5250
      Width           =   1365
   End
   Begin VB.TextBox Text101 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   2175
      TabIndex        =   0
      Top             =   1350
      Width           =   2490
   End
   Begin MSComctlLib.ListView ListView101 
      Height          =   3015
      Left            =   2175
      TabIndex        =   1
      Top             =   1950
      Width           =   8490
      _ExtentX        =   14975
      _ExtentY        =   5318
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "รหัสสินค้า"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ชื่อสินค้า"
         Object.Width           =   5115
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "จำนวน"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "ราคาสินค้า"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "ราคารวม"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "หน่วยนับ"
         Object.Width           =   1764
      EndProperty
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "เลขที่ใบสั่งขาย"
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
      Left            =   1050
      TabIndex        =   3
      Top             =   1350
      Width           =   1065
   End
End
Attribute VB_Name = "Form3_15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMD101_Click()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vItemCode As String
Dim i As Integer, vAmount As Integer, vCount As Integer
Dim vDocNo As String, vDiffQty As String
Dim vCheck As Boolean
Dim vListItems As ListItem

On Error GoTo ErrDescription

For i = 1 To ListView101.ListItems.Count
vCheck = ListView101.ListItems(i).Checked
If vCheck = True Then
                    vDocNo = Trim(Text101.Text)
                    vItemCode = ListView101.ListItems.Item(i).Text
                    vQuery = "Update bcnp.dbo.bcsaleordersub set remainqty = 0 ,iscancel = 1 where docno = '" & vDocNo & "' and itemcode = '" & vItemCode & "' "
                    gConnection.Execute vQuery
End If
Next i
Text101.Text = ""
ListView101.ListItems.Clear
Text101.SetFocus
MsgBox "ยกเลิกเอกสารขาย เรียบร้อยแล้วครับ"


ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If

End Sub

Private Sub Text101_KeyPress(KeyAscii As Integer)
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocNo As String
Dim vItemList As ListItem
Dim vAnswer As Integer

On Error GoTo ErrDescription
If KeyAscii = 13 Then
                vDocNo = Trim(Text101.Text)
                vQuery = "select docno,docdate,itemcode,itemname,qty,price,amount,unitcode from bcsaleordersub  where docno = '" & vDocNo & "' and remainqty <> 0 and iscancel = 0"
                If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                    vRecordset.MoveFirst
                   Do Until vRecordset.EOF
                    Set vItemList = ListView101.ListItems.Add(, , vRecordset.Fields("itemcode").Value)
                    vItemList.SubItems(1) = vRecordset.Fields("itemname").Value
                    vItemList.SubItems(2) = vRecordset.Fields("qty").Value
                    vItemList.SubItems(3) = vRecordset.Fields("price").Value
                    vItemList.SubItems(4) = vRecordset.Fields("amount").Value
                    vItemList.SubItems(5) = vRecordset.Fields("unitcode").Value
                    vRecordset.MoveNext
                    Loop
                Else
                MsgBox "ไม่มีเลขที่เอกสาร เลขที่ " & vDocNo & " ที่ต้องการทำการยกเลิก"
                End If
                vRecordset.Close
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If

End Sub
