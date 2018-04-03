VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form810 
   Caption         =   "ยกเลิกสินค้าค้างเบิก/โอน ในเอกสารขอเบิก/ขอโอน"
   ClientHeight    =   11490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15375
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form810.frx":0000
   ScaleHeight     =   11490
   ScaleWidth      =   15375
   WindowState     =   2  'Maximized
   Begin VB.CheckBox CHKAll 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "เลือกทั้งหมด"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   2340
      TabIndex        =   4
      Top             =   3015
      Width           =   1230
   End
   Begin VB.CommandButton CMDClearScreen 
      Caption         =   "ล้างหน้าจอ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   3915
      TabIndex        =   7
      Top             =   7020
      Width           =   1500
   End
   Begin VB.CommandButton CMDSearch 
      Height          =   400
      Left            =   4860
      Picture         =   "Form810.frx":9673
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1845
      Width           =   375
   End
   Begin VB.TextBox TXTDescription 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   2340
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   2340
      Width           =   11895
   End
   Begin MSComctlLib.ListView ListViewItemIssue 
      Height          =   3615
      Left            =   2340
      TabIndex        =   5
      Top             =   3375
      Visible         =   0   'False
      Width           =   11895
      _ExtentX        =   20981
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
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ลำดับ"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "รหัสสินค้า"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "ชื่อสินค้า"
         Object.Width           =   8819
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "ขอเบิกจำนวน"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "คงค้างจำนวน"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "หน่วย"
         Object.Width           =   2646
      EndProperty
   End
   Begin VB.PictureBox PicPoint 
      Height          =   240
      Left            =   -45
      ScaleHeight     =   180
      ScaleWidth      =   630
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.ComboBox CMBDocType 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2340
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1350
      Width           =   2895
   End
   Begin VB.CommandButton CMDCancel 
      Caption         =   "ยกเลิกเอกสาร"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   2340
      TabIndex        =   6
      Top             =   7020
      Width           =   1500
   End
   Begin VB.TextBox TXTDocNo 
      Appearance      =   0  'Flat
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
      Left            =   2340
      TabIndex        =   1
      Top             =   1845
      Width           =   2505
   End
   Begin MSComctlLib.ListView ListViewItemTransfer 
      Height          =   3285
      Left            =   2340
      TabIndex        =   8
      Top             =   3690
      Visible         =   0   'False
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   5794
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
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   13
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ลำดับ"
         Object.Width           =   2734
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "รหัสสินค้า"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "ชื่อสินค้า"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "จำนวนที่ขอโอน"
         Object.Width           =   2205
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "หน่วย"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "จากคลัง"
         Object.Width           =   1323
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "จากชั้นเก็บ"
         Object.Width           =   1676
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "เข้าคลัง"
         Object.Width           =   1323
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "เข้าชั้นเก็บ"
         Object.Width           =   1676
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "ใบขอโอน"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "วันที่ขอโอน"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "ใบโอน"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "โอนแล้ว"
         Object.Width           =   1764
      EndProperty
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "หมายเหตุ :"
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
      Height          =   375
      Left            =   945
      TabIndex        =   12
      Top             =   2340
      Width           =   1275
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ประเภทเอกสาร :"
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
      Height          =   375
      Left            =   540
      TabIndex        =   10
      Top             =   1350
      Width           =   1680
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "เลขที่เอกสาร :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   390
      Left            =   675
      TabIndex        =   9
      Top             =   1845
      Width           =   1515
   End
End
Attribute VB_Name = "Form810"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim vMemIsCancel As Integer
Dim vMemIsConfirm As Integer
Dim vMemBillStatus As Integer

'Private Sub Command1_Click()
'Dim vRecordset As New ADODB.Recordset
'Dim vQuery As String
'Dim vDocNo As String, vItemCode As String
'Dim i As Integer
'Dim vQTY As Integer
'Dim vTransferNo As String, vTransferNo1 As String

'On Error GoTo ErrDescription

'vDocNo = Trim(Text1.Text)
'vQuery = "select distinct refno as transferno  from bcnp.dbo.bcstktransfsub where refno = '" & vDocNo & "' "
'If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
 '           If IsNull(vRecordset.Fields("transferno").Value) Then
  '              vTransferNo = "DocNo"
   '         Else
    '            vTransferNo = Trim(vRecordset.Fields("transferno").Value)
     '       End If
'End If
'vRecordset.Close
'If vTransferNo <> "DocNo" Then
 '       For i = 1 To ListView1.ListItems.Count
  '      vItemCode = ListView1.ListItems(i).Text
   '     If ListView1.ListItems(i).Checked = True Then
    '            vQuery = " select docno,itemcode,transferno,qtytransfer from bcnp.dbo.vw_tf_bcstktransfer2 where docno = '" & vDocNo & "' and itemcode = '" & vItemCode & "' "
     '           If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
      '              'vTransferNo1 = Trim(vRecordset.Fields("transferno").Value)
       '             vQTY = Trim(vRecordset.Fields("qtytransfer").Value)
        '        End If
         '       vRecordset.Close
          '      If vQTY <> 0 Then
           '         vQuery = "Update bcnp.dbo.bcstktransfsub2 set qty = 0,mydescription  = " & vQTY & " from bcstktransfsub2 where docno = '" & vDocNo & "' and itemcode = '" & vItemCode & "' "
            '        gConnection.Execute vQuery
             '   Else
              '      vQuery = "delete bcnp.dbo.bcstktransfsub2 where docno = '" & vDocNo & "' and itemcode = '" & vItemCode & "' "
               '     gConnection.Execute vQuery
                'End If
        'End If

        'Next i
'Else
 '   vQuery = "update bcnp.dbo.bcstktransfer2  set iscancel = 1 where docno = '" & vDocNo & "' "
  '  gConnection.Execute vQuery
   ' 'vQuery = "delete bcstktransfsub2 where docno = '" & vDocNo & "' "
    ''gConnection.Execute vQuery
'End If
'Text1.Text = ""
'ListView1.ListItems.Clear

'ErrDescription:
'If Err.Description <> "" Then
 '   MsgBox Err.Description
  '  Exit Sub
'End If
'End Sub


'Private Sub Text1_KeyPress(KeyAscii As Integer)
'Dim vRecordset As New ADODB.Recordset
'Dim vQuery As String
'Dim vDocNo As String
'Dim ItemList As ListItem
'Dim vItemCode As String

'On Error GoTo ErrDescription

'If KeyAscii = 13 Then
 '   vDocNo = Trim(Text1.Text)
  '  vQuery = "select docno from bcstktransfer2 where docno = '" & vDocNo & "' "
   ' If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    '    MsgBox "โปรแกรมจะทำการคำนวณยอดจำนวนสินค้าที่สามารถทำใบโอนได้"
    'Else
     '   MsgBox "ไม่มีเลขที่เอกสาร เลขที่ " & vDocNo & "ในระบบ กรุณาตรวจสอบด้วยครับ"
      '  Exit Sub
    'End If
    'vRecordset.Close

'vQuery = "select * from bcnp.dbo.vw_tf_bcstktransfer2 where docno = '" & vDocNo & "' "
'If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
 '       vRecordset.MoveFirst
  '      Do Until vRecordset.EOF
   '             If IsNull(Trim(vRecordset.Fields("mydescription").Value)) Then
    '                    vItemCode = Trim(vRecordset.Fields("itemcode").Value)
     '                   vQuery = "update bcnp.dbo.vw_tf_bcstktransfer2 " _
      '                      & " set mydescription =  CONVERT(char(10), qty) where docno = '" & vDocNo & "' and itemcode = '" & vItemCode & "' "
       '                 gConnection.Execute vQuery
        '        End If
        'vRecordset.MoveNext
        'Loop
'End If
'vRecordset.Close

'vQuery = "select * from bcnp.dbo.vw_tf_bcstktransfer2 where docno = '" & vDocNo & "' "
'If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
 '       vRecordset.MoveFirst
  '      Do Until vRecordset.EOF
   '             If Trim(vRecordset.Fields("qty").Value) <> 0 Then
    '                    vItemCode = Trim(vRecordset.Fields("itemcode").Value)
     '                   vQuery = "update bcnp.dbo.vw_tf_bcstktransfer2 " _
      '                                  & " set qty =  convert(numeric(10),mydescription)-qtytransfer  where docno = '" & vDocNo & "' and itemcode = '" & vItemCode & "' "
       '                 gConnection.Execute vQuery
        '        End If
        'vRecordset.MoveNext
        'Loop
'End If
'MsgBox "ได้ทำการคำนวณจำนวนสินค้าที่ขอโอนเรียบร้อยแล้ว"
'vRecordset.Close

'vQuery = "select  a.docno,a.docdate,a.itemcode,b.name1,a.qty,a.unitcode,c.fromwh, " _
 '               & " c.fromshelf,c.towh,c.toshelf,transferno,qtytransfer " _
  '              & " from    vw_tf_bcstktransfer2 a " _
   '             & " left    join bcitem b on a.itemcode = b.code " _
    '            & " left    join bcstktransfsub2 c on a.docno = c.docno and a.itemcode = c.itemcode " _
     '           & " where a.docno = '" & vDocNo & "' and a.qty >0 "
'If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
 '   vRecordset.MoveFirst
  '  While Not vRecordset.EOF
   ' Set ItemList = ListView1.ListItems.Add(, , Trim(vRecordset.Fields("itemcode").Value))
    'ItemList.SubItems(1) = Trim(vRecordset.Fields("name1").Value)
    'ItemList.SubItems(2) = Trim(vRecordset.Fields("qty").Value)
    'ItemList.SubItems(3) = Trim(vRecordset.Fields("unitcode").Value)
    'ItemList.SubItems(4) = Trim(vRecordset.Fields("fromwh").Value)
    'ItemList.SubItems(5) = Trim(vRecordset.Fields("fromshelf").Value)
    'ItemList.SubItems(6) = Trim(vRecordset.Fields("towh").Value)
    'ItemList.SubItems(7) = Trim(vRecordset.Fields("toshelf").Value)
    'ItemList.SubItems(8) = Trim(vRecordset.Fields("docno").Value)
    'ItemList.SubItems(9) = Trim(vRecordset.Fields("docdate").Value)
    'If IsNull(vRecordset.Fields("transferno").Value) Then
    'ItemList.SubItems(10) = ""
    'Else
    'ItemList.SubItems(10) = Trim(vRecordset.Fields("transferno").Value)
    'End If
    'If IsNull(vRecordset.Fields("qtytransfer").Value) Then
    'ItemList.SubItems(11) = ""
    'Else
    'ItemList.SubItems(11) = Trim(vRecordset.Fields("qtytransfer").Value)
    'End If
'vRecordset.MoveNext
'Wend
'End If
'vRecordset.Close
'End If

'ErrDescription:
'If Err.Description <> "" Then
 '   MsgBox Err.Description
  '  Exit Sub
'End If
'End Sub


Private Sub CHKAll_Click()
Dim i As Integer

If Me.CHKAll.Value = 1 Then
If Me.ListViewItemIssue.ListItems.Count > 0 Then
   For i = 1 To Me.ListViewItemIssue.ListItems.Count
   Me.ListViewItemIssue.ListItems(i).Checked = True
   Next i
ElseIf Me.ListViewItemTransfer.ListItems.Count > 0 Then
   For i = 1 To Me.ListViewItemTransfer.ListItems.Count
   Me.ListViewItemTransfer.ListItems(i).Checked = True
   Next i
End If
Else
If Me.ListViewItemIssue.ListItems.Count > 0 Then
   For i = 1 To Me.ListViewItemIssue.ListItems.Count
   Me.ListViewItemIssue.ListItems(i).Checked = False
   Next i
ElseIf Me.ListViewItemTransfer.ListItems.Count > 0 Then
   For i = 1 To Me.ListViewItemTransfer.ListItems.Count
   Me.ListViewItemTransfer.ListItems(i).Checked = False
   Next i
End If
End If
End Sub

Private Sub CMBDocType_Click()
If Me.CMBDocType.ListIndex = 0 Then
   Me.ListViewItemIssue.Visible = True
   Me.ListViewItemTransfer.Visible = False
ElseIf Me.CMBDocType.ListIndex = 1 Then
   Me.ListViewItemIssue.Visible = False
   Me.ListViewItemTransfer.Visible = True
End If
End Sub

Private Sub CMBDocType_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Me.TXTDocNo.SetFocus
End If
End Sub

Private Sub CMDCancel_Click()
Dim vQuery As String

Dim vCountItem As Integer
Dim i As Integer
Dim vDocno As String
Dim vItemCode As String
Dim vDocType As String

If Me.TXTDocNo.Text <> "" Then
   
   If vMemIsCancel = 1 Then
      MsgBox "เอกสารได้ถูกยกเลิกไปแล้ว ไม่สามารถยกเลิกรายการสินค้าได้ กรุณาตรวจสอบ", vbCritical, "Send Error Message"
      Me.TXTDocNo.Enabled = True
      Call CMDClearScreen_Click
      Me.TXTDocNo.Text = ""
      Me.TXTDocNo.SetFocus
      Exit Sub
   End If
   
   If vMemIsCancel = 0 And vMemIsConfirm = 0 And vMemBillStatus = 0 Then
      MsgBox "เอกสารเป็นเอกสารใหม่ ยังไม่ได้ถูกอ้างอิง สามารถยกเลิกเอกสารได้ ณ โปรแกรม BCAccount ได้เอง กรุณาตรวจสอบ", vbCritical, "Send Error Message"
      Me.TXTDocNo.Enabled = True
      Call CMDClearScreen_Click
      Me.TXTDocNo.Text = ""
      Me.TXTDocNo.SetFocus
      Exit Sub
   End If
   
   vDocno = Me.TXTDocNo.Text
   vDocType = Me.CMBDocType.ListIndex + 1
   
   If Me.CMBDocType.ListIndex = 0 Then
      If Me.ListViewItemIssue.ListItems.Count > 0 Then
         For i = 1 To Me.ListViewItemIssue.ListItems.Count
         If Me.ListViewItemIssue.ListItems(i).Checked = True Then
         vCountItem = vCountItem + 1
         End If
         Next i
      End If
      
      If vCountItem = 0 Then
         MsgBox "ยังไม่ได้เลือกรายการสินค้าที่จะยกเลิก กรุณาตรวจสอบ", vbCritical, "Send Error Message "
         Me.ListViewItemIssue.SetFocus
         Exit Sub
      End If
      
      For i = 1 To Me.ListViewItemIssue.ListItems.Count
      If Me.ListViewItemIssue.ListItems(i).Checked = True Then
         vItemCode = Me.ListViewItemIssue.ListItems(i).SubItems(1)
         
         vQuery = "exec dbo.USP_NP_CancelRemainStockDocument '" & vDocno & "'," & vDocType & ",'" & vItemCode & "' "
         gConnection.Execute vQuery
      End If
      Next i
            
      MsgBox "ได้ทำการยกเลิกรายการสินค้าในใบขอเบิก เลขที่ " & vDocno & " เรียบร้อยแล้ว กรุณาตรวจสอบเอกสารใน BCAccount", vbExclamation, "Send Information Message"
      Call CMDClearScreen_Click
      
   End If

   If Me.CMBDocType.ListIndex = 1 Then
      If Me.ListViewItemTransfer.ListItems.Count > 0 Then
         For i = 1 To Me.ListViewItemTransfer.ListItems.Count
         If Me.ListViewItemTransfer.ListItems(i).Checked = True Then
         vCountItem = vCountItem + 1
         End If
         Next i
      End If
      
      If vCountItem = 0 Then
         MsgBox "ยังไม่ได้เลือกรายการสินค้าที่จะยกเลิก กรุณาตรวจสอบ", vbCritical, "Send Error Message "
         Me.ListViewItemTransfer.SetFocus
         Exit Sub
      End If
      
      For i = 1 To Me.ListViewItemTransfer.ListItems.Count
      If Me.ListViewItemTransfer.ListItems(i).Checked = True Then
         vItemCode = Me.ListViewItemTransfer.ListItems(i).SubItems(1)
         
         vQuery = "exec dbo.USP_NP_CancelRemainStockDocument '" & vDocno & "'," & vDocType & ",'" & vItemCode & "' "
         gConnection.Execute vQuery
      End If
      Next i
            
      MsgBox "ได้ทำการยกเลิกรายการสินค้าในใบขอโอนเลขที่ " & vDocno & " เรียบร้อยแล้ว กรุณาตรวจสอบเอกสารใน BCAccount", vbExclamation, "Send Information Message"
      Call CMDClearScreen_Click
      
   End If
   
Else
MsgBox "กรุณากรอกเลขที่เอกสาร ", vbCritical, "Send Error Message "
Me.TXTDocNo.SetFocus
End If
End Sub

Private Sub CMDClearScreen_Click()
vMemIsCancel = 0
vMemIsConfirm = 0
vMemBillStatus = 0
Me.CMBDocType.ListIndex = 0
Me.TXTDocNo.Text = ""
Me.TXTDescription.Text = ""
Me.ListViewItemIssue.ListItems.Clear
Me.ListViewItemTransfer.ListItems.Clear
Me.TXTDocNo.Enabled = True
Me.CMBDocType.Enabled = True
Me.CMDSearch.Enabled = True
Me.TXTDocNo.SetFocus
End Sub

Private Sub CMDSearch_Click()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset

Dim vDocno As String
Dim vType As Integer
Dim vListItem As ListItem
Dim i As Integer
Dim vQTY As Double
Dim vRemain As Double


If Me.TXTDocNo.Text <> "" Then
vType = Me.CMBDocType.ListIndex + 1

vDocno = Me.TXTDocNo.Text

Me.ListViewItemIssue.ListItems.Clear
Me.ListViewItemTransfer.ListItems.Clear

vQuery = "exec dbo.USP_NP_SearchStockDocument '" & vDocno & "'," & vType & " "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    Me.TXTDescription.Text = Trim(vRecordset.Fields("mydescription").Value)
    vMemIsCancel = Trim(vRecordset.Fields("iscancel").Value)
    vMemIsConfirm = Trim(vRecordset.Fields("isconfirm").Value)
    vMemBillStatus = Trim(vRecordset.Fields("billstatus").Value)
    For i = 1 To vRecordset.RecordCount
            Set vListItem = Me.ListViewItemIssue.ListItems.Add(, , i)
            vListItem.SubItems(1) = Trim(vRecordset.Fields("itemcode").Value)
            vListItem.SubItems(2) = Trim(vRecordset.Fields("itemname").Value)
            vQTY = Trim(vRecordset.Fields("qty").Value)
            vRemain = Trim(vRecordset.Fields("remainqty").Value)
            vListItem.SubItems(3) = Format(vQTY, "##,##0.00")
            vListItem.SubItems(4) = Format(vRemain, "##,##0.00")
            vListItem.SubItems(5) = Trim(vRecordset.Fields("unitcode").Value)
    vRecordset.MoveNext
Next i
Me.TXTDocNo.Enabled = False
Me.CMBDocType.Enabled = False
Me.CMDSearch.Enabled = False

If Me.CMBDocType.ListIndex = 0 Then
   Me.ListViewItemIssue.SetFocus
ElseIf Me.CMBDocType.ListIndex = 1 Then
   Me.ListViewItemTransfer.SetFocus
End If

Else
MsgBox "ไม่มีเอกสารนี้ในระบบ กรุณาตรวจสอบ", vbCritical, "Send Error Message"
Me.TXTDocNo.SetFocus
End If
vRecordset.Close
        
End If
End Sub

Private Sub Form_Load()
Call CreateDocType
Call SetListViewColor(ListViewItemIssue, PicPoint, vbWhite, vbLightGreen)
Call SetListViewColor(ListViewItemTransfer, PicPoint, vbWhite, vbLightBlue)
End Sub

Public Sub CreateDocType()

Me.CMBDocType.AddItem ("ใบขอเบิกสินค้าและวัตถุดิบ")
Me.CMBDocType.AddItem ("ใบขอโอนสินค้า")

Me.CMBDocType.ListIndex = 0
End Sub

Private Sub ListViewItemIssue_KeyDown(KeyCode As Integer, Shift As Integer)
If Me.ListViewItemIssue.ListItems.Count > 0 Then
If KeyCode = 27 Then
   Call CMDClearScreen_Click
End If
End If
End Sub

Private Sub TXTDocNo_KeyDown(KeyCode As Integer, Shift As Integer)
If Me.TXTDocNo.Text <> "" Then
   If KeyCode = 8 Then
      Call CMDClearScreen_Click
   End If
End If

If KeyCode = 27 Then
   Call CMDClearScreen_Click
End If
End Sub

Private Sub TXTDocNo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If Me.TXTDocNo.Text <> "" Then
      Call CMDSearch_Click
   End If
End If

End Sub
