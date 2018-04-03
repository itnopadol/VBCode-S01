VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FormSearchMainPromotion 
   Caption         =   "ค้นหาทะเบียนโปรโมชั่น"
   ClientHeight    =   3675
   ClientLeft      =   5055
   ClientTop       =   2490
   ClientWidth     =   8310
   LinkTopic       =   "Form2"
   ScaleHeight     =   3675
   ScaleWidth      =   8310
   Begin MSComctlLib.ListView ListView101 
      Height          =   2940
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   8190
      _ExtentX        =   14446
      _ExtentY        =   5186
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
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "รหัสโปรโมชั่น"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ชื่อโปรโมชั่น"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "วันที่เริ่มโปรโมชั้น"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "วันที่หมดโปรโมชั่น"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "ผู้ทำเอกสาร"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "วันที่ทำเอกสาร"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "IsCancel"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "PMName"
         Object.Width           =   3528
      EndProperty
   End
End
Attribute VB_Name = "FormSearchMainPromotion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim ProList As ListItem

On Error Resume Next

ListView101.ListItems.Clear
vQuery = "exec USP_PM_Find"
If OpenDatabase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    While Not vRecordset.EOF
    Set ProList = ListView101.ListItems.Add(, , Trim(vRecordset.Fields("pmcode").Value))
    ProList.SubItems(1) = Trim(vRecordset.Fields("pmname").Value)
    ProList.SubItems(2) = Trim(vRecordset.Fields("datestart").Value)
    ProList.SubItems(3) = Trim(vRecordset.Fields("dateend").Value)
    ProList.SubItems(4) = Trim(vRecordset.Fields("creatorcode").Value)
    ProList.SubItems(5) = Trim(vRecordset.Fields("createdate").Value)
    ProList.SubItems(6) = Trim(vRecordset.Fields("iscancel").Value)
    ProList.SubItems(7) = Trim(vRecordset.Fields("pmname1").Value)
    vRecordset.MoveNext
    Wend
End If
vRecordset.Close

End Sub

Private Sub Form_Unload(Cancel As Integer)
MDIForm1.Enabled = True
End Sub

Private Sub ListView101_DblClick()
Dim vIsCancel As Integer
Dim vIndex As Integer

    vIndex = ListView101.SelectedItem.Index
    vIsCancel = ListView101.ListItems.Item(vIndex).SubItems(6)
    If vIsCancel = 0 Then
        If vMemCommand = 1 Then
        Form201.Text102.Text = Trim(ListView101.ListItems.Item(vIndex).SubItems(7))
        Form201.LBLPromoStartDate.Caption = Trim(ListView101.ListItems.Item(vIndex).SubItems(2))
        Form201.LBLPromoStopDate.Caption = Trim(ListView101.ListItems.Item(vIndex).SubItems(3))
        ElseIf vMemCommand = 2 Then
        Form103.Text101.Text = Trim(ListView101.ListItems.Item(vIndex).SubItems(7))
        Form103.Label101.Caption = Trim(ListView101.ListItems.Item(vIndex).SubItems(2))
        Form103.Label102.Caption = Trim(ListView101.ListItems.Item(vIndex).SubItems(3))
        ElseIf vMemCommand = 3 Then
        Form401.Text102.Text = Trim(ListView101.ListItems.Item(vIndex).SubItems(7))
        End If
        MDIForm1.Enabled = True
        vMemCommand = 0
        Unload FormSearchMainPromotion
    Else
        MsgBox "ไม่สามารถเลือกโปรโมชั่นนี้ได้ เนื่องจากได้ทำการยกเลิกไปแล้ว"
    End If

End Sub

Private Sub ListView101_ItemClick(ByVal Item As MSComctlLib.ListItem)
'Dim vIsCancel As Integer

 '   vIsCancel = ListView101.ListItems.Item(Item.Index).SubItems(6)
  '  If vIsCancel = 0 Then
   '     If vMemCommand = 1 Then
    '    Form201.Text102.Text = Trim(ListView101.ListItems.Item(Item.Index).SubItems(7))
     '   ElseIf vMemCommand = 2 Then
       ' Form103.Text101.Text = Trim(ListView101.ListItems.Item(Item.Index).SubItems(7))
      '  Form103.Label101.Caption = Trim(ListView101.ListItems.Item(Item.Index).SubItems(2))
        'Form103.Label102.Caption = Trim(ListView101.ListItems.Item(Item.Index).SubItems(3))
        'ElseIf vMemCommand = 3 Then
        'Form401.Text102.Text = Trim(ListView101.ListItems.Item(Item.Index).SubItems(7))
        'End If
        'MDIForm1.Enabled = True
        'vMemCommand = 0
        'Unload FormSearchMainPromotion
    'Else
     '   MsgBox "ไม่สามารถเลือกโปรโมชั่นนี้ได้ เนื่องจากได้ทำการยกเลิกไปแล้ว"
    'End If

End Sub
