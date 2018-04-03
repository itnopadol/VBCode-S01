VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form2_7 
   Caption         =   "ลบรายตัวสินค้า ใบเสนอซื้อที่ไม่ได้อนุมัติ"
   ClientHeight    =   8310
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form2_7.frx":0000
   ScaleHeight     =   8310
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CMD101 
      Caption         =   "ยกเลิก"
      Height          =   465
      Left            =   7725
      TabIndex        =   2
      Top             =   5775
      Width           =   1590
   End
   Begin VB.TextBox Text101 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2700
      TabIndex        =   0
      Top             =   1275
      Width           =   2415
   End
   Begin MSComctlLib.ListView ListView101 
      Height          =   3540
      Left            =   975
      TabIndex        =   1
      Top             =   2025
      Width           =   8340
      _ExtentX        =   14711
      _ExtentY        =   6244
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "เลขที่ใบเสนอซื้อ"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "วันที่เอกสาร"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "รหัสสินค้า"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "จำนวนที่ไม่ได้อนุมัติ"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "ชื่อสินค้า"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "หน่วยนับ"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "จำนวนที่เสนอซื้อ"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "จำนวนอนุมัติ"
         Object.Width           =   2646
      EndProperty
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "เลขที่ใบเสนอซื้อสินค้า"
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
      Left            =   975
      TabIndex        =   3
      Top             =   1350
      Width           =   1740
   End
End
Attribute VB_Name = "Form2_7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vRefreshNo As String

Private Sub CMD101_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vCount As Integer
Dim vCheckQty As Integer
Dim vDocno As String
Dim vItemCode As String
Dim i As Integer

On Error GoTo ErrDescription

vCount = ListView101.ListItems.Count
For i = 1 To vCount
If ListView101.ListItems(i).Checked = True Then
    vCheckQty = Trim(ListView101.ListItems(i).SubItems(7))
    vDocno = Trim(ListView101.ListItems(i).Text)
    vItemCode = Trim(ListView101.ListItems(i).SubItems(2))
    If vCheckQty = 0 Then
        vQuery = "delete bcnp.dbo.bcstkrequestsub where docno = '" & vDocno & "' and itemcode = '" & vItemCode & "' "
        gConnection.Execute vQuery
    ElseIf vCheckQty > 0 Then
        vQuery = "exec USP_PR_UpDateRemainQTY '" & vDocno & "' ,'" & vItemCode & "' "
        gConnection.Execute vQuery
    End If
End If
Next i
vRefreshNo = vDocno
MsgBox "ทำการแก้ไขข้อมูลของใบเสนอซื้อเลขที่ " & vDocno & " เรียบร้อยแล้วครับ "
Call RefreshData
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Text101_KeyPress(KeyAscii As Integer)
Dim vQuery  As String
Dim vRecordset As New ADODB.Recordset
Dim vDocno As String
Dim vItemList As ListItem

On Error GoTo ErrDescription

If KeyAscii = 13 Then
    vDocno = Text101.Text
    If vDocno <> "" Then
        ListView101.ListItems.Clear
        vQuery = "exec bcnp.dbo.USP_AP_GetItemApprovePR  '" & vDocno & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vRecordset.MoveFirst
            While Not vRecordset.EOF
                Set vItemList = ListView101.ListItems.Add(, , Trim(vRecordset.Fields("docno").Value))
                vItemList.SubItems(1) = Trim(vRecordset.Fields("docdate").Value)
                vItemList.SubItems(2) = Trim(vRecordset.Fields("itemcode").Value)
                vItemList.SubItems(3) = Trim(vRecordset.Fields("remainqty").Value)
                vItemList.SubItems(4) = Trim(vRecordset.Fields("itemname").Value)
                vItemList.SubItems(5) = Trim(vRecordset.Fields("unitcode").Value)
                vItemList.SubItems(6) = Trim(vRecordset.Fields("qty").Value)
                vItemList.SubItems(7) = Trim(vRecordset.Fields("qty").Value) - Trim(vRecordset.Fields("remainqty").Value)
            vRecordset.MoveNext
            Wend
            Text101.Text = ""
            Text101.SetFocus
        Else
            MsgBox "เอกสารใบเสนอซื้อสินค้าเลขที่ " & vDocno & " ไม่มีสินค้าค้างอนุมัติ"
        End If
        vRecordset.Close
    Else
        MsgBox "กรุณาใส่เลขที่ใบเสนอซื้อสินค้าที่ทำการอนุมัติแล้ว แต่ต้องการจะลบรายการสินค้าข้างใน"
        Text101.SetFocus
    End If
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If

End Sub

Public Sub RefreshData()
Dim vQuery  As String
Dim vRecordset As New ADODB.Recordset
Dim vItemList As ListItem

On Error GoTo ErrDescription

    If vRefreshNo <> "" Then
        ListView101.ListItems.Clear
        vQuery = "exec bcnp.dbo.USP_AP_GetItemApprovePR  '" & vRefreshNo & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vRecordset.MoveFirst
            While Not vRecordset.EOF
                Set vItemList = ListView101.ListItems.Add(, , Trim(vRecordset.Fields("docno").Value))
                vItemList.SubItems(1) = Trim(vRecordset.Fields("docdate").Value)
                vItemList.SubItems(2) = Trim(vRecordset.Fields("itemcode").Value)
                vItemList.SubItems(3) = Trim(vRecordset.Fields("remainqty").Value)
                vItemList.SubItems(4) = Trim(vRecordset.Fields("itemname").Value)
                vItemList.SubItems(5) = Trim(vRecordset.Fields("unitcode").Value)
                vItemList.SubItems(6) = Trim(vRecordset.Fields("qty").Value)
                vItemList.SubItems(7) = Trim(vRecordset.Fields("qty").Value) - Trim(vRecordset.Fields("remainqty").Value)
            vRecordset.MoveNext
            Wend
            Text101.Text = ""
            Text101.SetFocus
        End If
        vRecordset.Close
    Else
        MsgBox "กรุณาใส่เลขที่ใบเสนอซื้อสินค้าที่ทำการอนุมัติแล้ว แต่ต้องการจะลบรายการสินค้าข้างใน"
        Text101.SetFocus
    End If


ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub
