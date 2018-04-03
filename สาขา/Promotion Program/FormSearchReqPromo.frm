VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FormSearchReqPromo 
   Caption         =   "ค้นหา เลขที่ใบเสนอสินค้าโปรโมชั่น"
   ClientHeight    =   5430
   ClientLeft      =   4500
   ClientTop       =   1935
   ClientWidth     =   8820
   LinkTopic       =   "Form1"
   ScaleHeight     =   5430
   ScaleWidth      =   8820
   Begin MSComctlLib.ListView ListView101 
      Height          =   4065
      Left            =   375
      TabIndex        =   2
      Top             =   825
      Width           =   8040
      _ExtentX        =   14182
      _ExtentY        =   7170
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "เลขที่เอกสาร"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "วันที่เอกสาร"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "ชื่อโปรโมชั่น"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Section Manager"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "IsConfirm"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "IsCancel"
         Object.Width           =   1764
      EndProperty
   End
   Begin VB.TextBox Text101 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2025
      TabIndex        =   1
      Top             =   300
      Width           =   3315
   End
   Begin VB.Label Label1 
      Caption         =   "ค้นหาจากข้อความ"
      Height          =   315
      Left            =   600
      TabIndex        =   0
      Top             =   375
      Width           =   1365
   End
End
Attribute VB_Name = "FormSearchReqPromo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vDocno As String

Private Sub Form_Unload(Cancel As Integer)
MDIForm1.Enabled = True
End Sub

Private Sub ListView101_DblClick()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vIndex As Integer
Dim vListItemSearch As ListItem
Dim i As Integer
Dim vItemCode As String

On Error Resume Next

    vIndex = ListView101.SelectedItem.Index
    vCheckJob1 = 0
    MDIForm1.Enabled = True
    Form201.ListView101.ListItems.Clear
    vDocno = ListView101.SelectedItem.Text
    
    If ListView101.ListItems.Item(vIndex).SubItems(4) = 0 And ListView101.ListItems.Item(vIndex).SubItems(5) = 0 Then
        vCheckStatusPrint = 0
    Else
        vCheckStatusPrint = 1
    End If
    
    If (ListView101.ListItems.Item(vIndex).SubItems(5) = 1) Then 'And ListView101.ListItems.Item(vIndex).SubItems(4) = 2) Or (ListView101.ListItems.Item(vIndex).SubItems(5) = 1 And ListView101.ListItems.Item(vIndex).SubItems(4) = 0) Then
        Form201.Image101.Visible = False
        Form201.Image102.Visible = False
        Form201.Image103.Visible = True
        Form201.CMD108.Enabled = False
    ElseIf (ListView101.ListItems.Item(vIndex).SubItems(4) = 2 Or ListView101.ListItems.Item(vIndex).SubItems(4) = 1) And ListView101.ListItems.Item(vIndex).SubItems(5) = 0 Then
        Form201.Image101.Visible = False
        Form201.Image102.Visible = True
        Form201.Image103.Visible = False
        Form201.CMD108.Enabled = False
    End If
    If ListView101.ListItems.Item(vIndex).SubItems(4) = 0 And ListView101.ListItems.Item(vIndex).SubItems(5) = 0 Then
        Form201.Image101.Visible = True
        Form201.Image102.Visible = False
        Form201.Image103.Visible = False
    End If
    i = 0
    vQuery = "execute USP_PM_RequestSubSearch '" & vDocno & "' "
    If OpenDatabase(gConnection, vRecordset, vQuery) <> 0 Then
        Form201.Text101.Text = Trim(vRecordset.Fields("docno").Value)
        Form201.DTPicker101.Value = Trim(vRecordset.Fields("docdate").Value)
        Form201.Text102.Text = Trim(vRecordset.Fields("pmname").Value)
        Form201.Text103.Text = Trim(vRecordset.Fields("secman").Value)
        vRecordset.MoveFirst
        While Not vRecordset.EOF
            i = i + 1
            Set vListItemSearch = Form201.ListView101.ListItems.Add(, , Trim(vRecordset.Fields("itemcode").Value))
            vListItemSearch.SubItems(1) = Trim(vRecordset.Fields("itemname").Value)
            vListItemSearch.SubItems(2) = Trim(vRecordset.Fields("price").Value)
            vListItemSearch.SubItems(3) = Trim(vRecordset.Fields("promoprice").Value)
            vListItemSearch.SubItems(4) = Trim(vRecordset.Fields("discount").Value) '- Trim(vRecordset.Fields("promoprice").Value) 'Trim(vRecordset.Fields("discountword").Value)
            vListItemSearch.SubItems(5) = Trim(vRecordset.Fields("discountword").Value) 'Trim(vRecordset.Fields("discounttype").Value)
            vListItemSearch.SubItems(6) = Trim(vRecordset.Fields("unitcode").Value) 'Trim(vRecordset.Fields("price").Value) - Trim(vRecordset.Fields("promoprice").Value)
            vListItemSearch.SubItems(7) = Trim(vRecordset.Fields("mydescription").Value) 'Trim(vRecordset.Fields("unitcode").Value)
            vListItemSearch.SubItems(8) = Trim(vRecordset.Fields("discounttype").Value) 'Trim(vRecordset.Fields("promomember").Value)
            vListItemSearch.SubItems(9) = Trim(vRecordset.Fields("promomember").Value) 'Trim(vRecordset.Fields("mydescription").Value)
            vListItemSearch.SubItems(10) = Trim(vRecordset.Fields("isbrochure").Value)
            vListItemSearch.SubItems(11) = Trim(vRecordset.Fields("iscancel").Value)
            vListItemSearch.SubItems(12) = Trim(vRecordset.Fields("name1").Value)
            vListItemSearch.SubItems(13) = Trim(vRecordset.Fields("promotiontype").Value)
            If Trim(vRecordset.Fields("iscancel").Value) = 1 Then
                Form201.ListView101.ListItems(i).ForeColor = "&H000000FF"
                Form201.ListView101.ListItems.Item(i).ListSubItems(1).ForeColor = "&H000000FF"
                Form201.ListView101.ListItems.Item(i).ListSubItems(2).ForeColor = "&H000000FF"
                Form201.ListView101.ListItems.Item(i).ListSubItems(3).ForeColor = "&H000000FF"
                Form201.ListView101.ListItems.Item(i).ListSubItems(4).ForeColor = "&H000000FF"
                Form201.ListView101.ListItems.Item(i).ListSubItems(5).ForeColor = "&H000000FF"
                Form201.ListView101.ListItems.Item(i).ListSubItems(6).ForeColor = "&H000000FF"
                Form201.ListView101.ListItems.Item(i).ListSubItems(7).ForeColor = "&H000000FF"
                Form201.ListView101.ListItems.Item(i).ListSubItems(8).ForeColor = "&H000000FF"
                Form201.ListView101.ListItems.Item(i).ListSubItems(9).ForeColor = "&H000000FF"
                Form201.ListView101.ListItems.Item(i).ListSubItems(10).ForeColor = "&H000000FF"
                Form201.ListView101.ListItems.Item(i).ListSubItems(11).ForeColor = "&H000000FF"
                Form201.ListView101.ListItems.Item(i).ListSubItems(12).ForeColor = "&H000000FF"
                Form201.ListView101.ListItems.Item(i).ListSubItems(13).ForeColor = "&H000000FF"
                Form201.ListView101.ListItems.Item(i).Checked = False
            Else
                Form201.ListView101.ListItems.Item(i).Checked = True
            End If
            vRecordset.MoveNext
        Wend
    End If
    vRecordset.Close
    Unload FormSearchReqPromo

End Sub

Private Sub ListView101_KeyPress(KeyAscii As Integer)
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vIndex As Integer
Dim vListItemSearch As ListItem
Dim i As Integer
Dim vItemCode As String

On Error Resume Next

If KeyAscii = 13 Then
    vIndex = ListView101.SelectedItem.Index
    vCheckJob1 = 0
    MDIForm1.Enabled = True
    Form201.ListView101.ListItems.Clear
    vDocno = ListView101.SelectedItem.Text
    If (ListView101.ListItems.Item(vIndex).SubItems(5) = 1 And ListView101.ListItems.Item(vIndex).SubItems(4) = 2) Or (ListView101.ListItems.Item(vIndex).SubItems(5) = 1 And ListView101.ListItems.Item(vIndex).SubItems(4) = 0) Then
        Form201.Image101.Visible = False
        Form201.Image102.Visible = False
        Form201.Image103.Visible = True
        Form201.CMD108.Enabled = False
    ElseIf ListView101.ListItems.Item(vIndex).SubItems(4) = 2 And ListView101.ListItems.Item(vIndex).SubItems(5) = 0 Then
        Form201.Image101.Visible = False
        Form201.Image102.Visible = True
        Form201.Image103.Visible = False
        Form201.CMD108.Enabled = False
    End If
    If ListView101.ListItems.Item(vIndex).SubItems(4) = 0 And ListView101.ListItems.Item(vIndex).SubItems(5) = 0 Then
        Form201.Image101.Visible = True
        Form201.Image102.Visible = False
        Form201.Image103.Visible = False
    End If
    i = 0
    vQuery = "execute USP_PM_RequestSubSearch '" & vDocno & "' "
    If OpenDatabase(gConnection, vRecordset, vQuery) <> 0 Then
        Form201.Text101.Text = Trim(vRecordset.Fields("docno").Value)
        Form201.DTPicker101.Value = Trim(vRecordset.Fields("docdate").Value)
        Form201.Text102.Text = Trim(vRecordset.Fields("pmname").Value)
        Form201.Text103.Text = Trim(vRecordset.Fields("secman").Value)
        vRecordset.MoveFirst
        While Not vRecordset.EOF
            i = i + 1
            Set vListItemSearch = Form201.ListView101.ListItems.Add(, , Trim(vRecordset.Fields("itemcode").Value))
            vListItemSearch.SubItems(1) = Trim(vRecordset.Fields("itemname").Value)
            vListItemSearch.SubItems(2) = Trim(vRecordset.Fields("price").Value)
            vListItemSearch.SubItems(3) = Trim(vRecordset.Fields("promoprice").Value)
            vListItemSearch.SubItems(4) = Trim(vRecordset.Fields("discount").Value) '- Trim(vRecordset.Fields("promoprice").Value) 'Trim(vRecordset.Fields("discountword").Value)
            vListItemSearch.SubItems(5) = Trim(vRecordset.Fields("discountword").Value) 'Trim(vRecordset.Fields("discounttype").Value)
            vListItemSearch.SubItems(6) = Trim(vRecordset.Fields("unitcode").Value) 'Trim(vRecordset.Fields("price").Value) - Trim(vRecordset.Fields("promoprice").Value)
            vListItemSearch.SubItems(7) = Trim(vRecordset.Fields("mydescription").Value) 'Trim(vRecordset.Fields("unitcode").Value)
            vListItemSearch.SubItems(8) = Trim(vRecordset.Fields("discounttype").Value) 'Trim(vRecordset.Fields("promomember").Value)
            vListItemSearch.SubItems(9) = Trim(vRecordset.Fields("promomember").Value) 'Trim(vRecordset.Fields("mydescription").Value)
            vListItemSearch.SubItems(10) = Trim(vRecordset.Fields("isbrochure").Value)
            vListItemSearch.SubItems(11) = Trim(vRecordset.Fields("iscancel").Value)
            vListItemSearch.SubItems(12) = Trim(vRecordset.Fields("name1").Value)
            vListItemSearch.SubItems(13) = Trim(vRecordset.Fields("promotiontype").Value)
            If Trim(vRecordset.Fields("iscancel").Value) = 1 Then
                Form201.ListView101.ListItems(i).ForeColor = "&H000000FF"
                Form201.ListView101.ListItems.Item(i).ListSubItems(1).ForeColor = "&H000000FF"
                Form201.ListView101.ListItems.Item(i).ListSubItems(2).ForeColor = "&H000000FF"
                Form201.ListView101.ListItems.Item(i).ListSubItems(3).ForeColor = "&H000000FF"
                Form201.ListView101.ListItems.Item(i).ListSubItems(4).ForeColor = "&H000000FF"
                Form201.ListView101.ListItems.Item(i).ListSubItems(5).ForeColor = "&H000000FF"
                Form201.ListView101.ListItems.Item(i).ListSubItems(6).ForeColor = "&H000000FF"
                Form201.ListView101.ListItems.Item(i).ListSubItems(7).ForeColor = "&H000000FF"
                Form201.ListView101.ListItems.Item(i).ListSubItems(8).ForeColor = "&H000000FF"
                Form201.ListView101.ListItems.Item(i).ListSubItems(9).ForeColor = "&H000000FF"
                Form201.ListView101.ListItems.Item(i).ListSubItems(10).ForeColor = "&H000000FF"
                Form201.ListView101.ListItems.Item(i).ListSubItems(11).ForeColor = "&H000000FF"
                Form201.ListView101.ListItems.Item(i).ListSubItems(12).ForeColor = "&H000000FF"
                Form201.ListView101.ListItems.Item(i).ListSubItems(13).ForeColor = "&H000000FF"
                Form201.ListView101.ListItems.Item(i).Checked = False
            Else
                Form201.ListView101.ListItems.Item(i).Checked = True
            End If
            vRecordset.MoveNext
        Wend
    End If
    vRecordset.Close
    Unload FormSearchReqPromo
End If
End Sub

Private Sub Text101_KeyPress(KeyAscii As Integer)
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vListPromotion As ListItem
Dim vSearchPromotion As String

On Error GoTo ErrDescription

If KeyAscii = 13 Then
        ListView101.ListItems.Clear
        vSearchPromotion = Trim(Text101.Text)
        vQuery = "execute USP_PM_RequestSearch  '" & vSearchPromotion & "' "
        If OpenDatabase(gConnection, vRecordset, vQuery) <> 0 Then
            vRecordset.MoveFirst
            While Not vRecordset.EOF
                Set vListPromotion = ListView101.ListItems.Add(, , Trim(vRecordset.Fields("docno").Value))
                vListPromotion.SubItems(1) = Trim(vRecordset.Fields("docdate").Value)
                vListPromotion.SubItems(2) = Trim(vRecordset.Fields("pmcode").Value)
                vListPromotion.SubItems(3) = Trim(vRecordset.Fields("secman").Value)
                vListPromotion.SubItems(4) = Trim(vRecordset.Fields("isconfirm").Value)
                vListPromotion.SubItems(5) = Trim(vRecordset.Fields("iscancel").Value)
            vRecordset.MoveNext
            Wend
        End If
        vRecordset.Close
        ListView101.SetFocus
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub
