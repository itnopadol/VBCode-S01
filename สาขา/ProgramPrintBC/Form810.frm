VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form810 
   Caption         =   "ยกเลิกสินค้าในใบขอโอนสินค้า"
   ClientHeight    =   8370
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form810.frx":0000
   ScaleHeight     =   8370
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "ยกเลิกเอกสาร"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   1200
      TabIndex        =   2
      Top             =   6600
      Width           =   2040
   End
   Begin VB.TextBox Text1 
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
      Height          =   390
      Left            =   1200
      TabIndex        =   0
      Top             =   1200
      Width           =   3090
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4590
      Left            =   1200
      TabIndex        =   1
      Top             =   1875
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   8096
      View            =   3
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
      NumItems        =   12
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "รหัสสินค้า"
         Object.Width           =   2734
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ชื่อสินค้า"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "จำนวนที่ขอโอน"
         Object.Width           =   2205
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "หน่วย"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "จากคลัง"
         Object.Width           =   1323
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "จากชั้นเก็บ"
         Object.Width           =   1676
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "เข้าคลัง"
         Object.Width           =   1323
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "เข้าชั้นเก็บ"
         Object.Width           =   1676
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "ใบขอโอน"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "วันที่ขอโอน"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "ใบโอน"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "โอนแล้ว"
         Object.Width           =   1764
      EndProperty
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "เลขที่เอกสาร"
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
      Left            =   225
      TabIndex        =   3
      Top             =   1200
      Width           =   840
   End
End
Attribute VB_Name = "Form810"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocno As String, vItemCode As String
Dim i As Integer
Dim vQTY As Integer
Dim vTransferNo As String, vTransferNo1 As String

On Error GoTo ErrDescription

vDocno = Trim(Text1.Text)
vQuery = "select distinct refno as transferno  from bcnp.dbo.bcstktransfsub where refno = '" & vDocno & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            If IsNull(vRecordset.Fields("transferno").Value) Then
                vTransferNo = "DocNo"
            Else
                vTransferNo = Trim(vRecordset.Fields("transferno").Value)
            End If
End If
vRecordset.Close
If vTransferNo <> "DocNo" Then
        For i = 1 To ListView1.ListItems.Count
        vItemCode = ListView1.ListItems(i).Text
        If ListView1.ListItems(i).Checked = True Then
                vQuery = " select docno,itemcode,transferno,qtytransfer from bcnp.dbo.vw_tf_bcstktransfer2 where docno = '" & vDocno & "' and itemcode = '" & vItemCode & "' "
                If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                    'vTransferNo1 = Trim(vRecordset.Fields("transferno").Value)
                    vQTY = Trim(vRecordset.Fields("qtytransfer").Value)
                End If
                vRecordset.Close
                If vQTY <> 0 Then
                    vQuery = "Update bcnp.dbo.bcstktransfsub2 set qty = 0,mydescription  = " & vQTY & " from bcstktransfsub2 where docno = '" & vDocno & "' and itemcode = '" & vItemCode & "' "
                    gConnection.Execute vQuery
                Else
                    vQuery = "delete bcnp.dbo.bcstktransfsub2 where docno = '" & vDocno & "' and itemcode = '" & vItemCode & "' "
                    gConnection.Execute vQuery
                End If
        End If

        Next i
Else
    vQuery = "update bcnp.dbo.bcstktransfer2  set iscancel = 1 where docno = '" & vDocno & "' "
    gConnection.Execute vQuery
    'vQuery = "delete bcstktransfsub2 where docno = '" & vDocNo & "' "
    'gConnection.Execute vQuery
End If
Text1.Text = ""
ListView1.ListItems.Clear

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocno As String
Dim ItemList As ListItem
Dim vItemCode As String

On Error GoTo ErrDescription

If KeyAscii = 13 Then
    vDocno = Trim(Text1.Text)
    vQuery = "select docno from bcstktransfer2 where docno = '" & vDocno & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        MsgBox "โปรแกรมจะทำการคำนวณยอดจำนวนสินค้าที่สามารถทำใบโอนได้"
    Else
        MsgBox "ไม่มีเลขที่เอกสาร เลขที่ " & vDocno & "ในระบบ กรุณาตรวจสอบด้วยครับ"
        Exit Sub
    End If
    vRecordset.Close

vQuery = "select * from bcnp.dbo.vw_tf_bcstktransfer2 where docno = '" & vDocno & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vRecordset.MoveFirst
        Do Until vRecordset.EOF
                If IsNull(Trim(vRecordset.Fields("mydescription").Value)) Then
                        vItemCode = Trim(vRecordset.Fields("itemcode").Value)
                        vQuery = "update bcnp.dbo.vw_tf_bcstktransfer2 " _
                            & " set mydescription =  CONVERT(char(10), qty) where docno = '" & vDocno & "' and itemcode = '" & vItemCode & "' "
                        gConnection.Execute vQuery
                End If
        vRecordset.MoveNext
        Loop
End If
vRecordset.Close

vQuery = "select * from bcnp.dbo.vw_tf_bcstktransfer2 where docno = '" & vDocno & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vRecordset.MoveFirst
        Do Until vRecordset.EOF
                If Trim(vRecordset.Fields("qty").Value) <> 0 Then
                        vItemCode = Trim(vRecordset.Fields("itemcode").Value)
                        vQuery = "update bcnp.dbo.vw_tf_bcstktransfer2 " _
                                        & " set qty =  convert(numeric(10),mydescription)-qtytransfer  where docno = '" & vDocno & "' and itemcode = '" & vItemCode & "' "
                        gConnection.Execute vQuery
                End If
        vRecordset.MoveNext
        Loop
End If
MsgBox "ได้ทำการคำนวณจำนวนสินค้าที่ขอโอนเรียบร้อยแล้ว"
vRecordset.Close

vQuery = "select  a.docno,a.docdate,a.itemcode,b.name1,a.qty,a.unitcode,c.fromwh, " _
                & " c.fromshelf,c.towh,c.toshelf,transferno,qtytransfer " _
                & " from    vw_tf_bcstktransfer2 a " _
                & " left    join bcitem b on a.itemcode = b.code " _
                & " left    join bcstktransfsub2 c on a.docno = c.docno and a.itemcode = c.itemcode " _
                & " where a.docno = '" & vDocno & "' and a.qty >0 "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    While Not vRecordset.EOF
    Set ItemList = ListView1.ListItems.Add(, , Trim(vRecordset.Fields("itemcode").Value))
    ItemList.SubItems(1) = Trim(vRecordset.Fields("name1").Value)
    ItemList.SubItems(2) = Trim(vRecordset.Fields("qty").Value)
    ItemList.SubItems(3) = Trim(vRecordset.Fields("unitcode").Value)
    ItemList.SubItems(4) = Trim(vRecordset.Fields("fromwh").Value)
    ItemList.SubItems(5) = Trim(vRecordset.Fields("fromshelf").Value)
    ItemList.SubItems(6) = Trim(vRecordset.Fields("towh").Value)
    ItemList.SubItems(7) = Trim(vRecordset.Fields("toshelf").Value)
    ItemList.SubItems(8) = Trim(vRecordset.Fields("docno").Value)
    ItemList.SubItems(9) = Trim(vRecordset.Fields("docdate").Value)
    If IsNull(vRecordset.Fields("transferno").Value) Then
    ItemList.SubItems(10) = ""
    Else
    ItemList.SubItems(10) = Trim(vRecordset.Fields("transferno").Value)
    End If
    If IsNull(vRecordset.Fields("qtytransfer").Value) Then
    ItemList.SubItems(11) = ""
    Else
    ItemList.SubItems(11) = Trim(vRecordset.Fields("qtytransfer").Value)
    End If
vRecordset.MoveNext
Wend
End If
vRecordset.Close
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub
