VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form2_9 
   Caption         =   "ź�͡��� ͹��ѵ���ʹͫ����Թ���"
   ClientHeight    =   9000
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form2_9.frx":0000
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CMDSearch 
      Height          =   285
      Left            =   6075
      Picture         =   "Form2_9.frx":9673
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1530
      Width           =   285
   End
   Begin VB.CommandButton CMDDelete 
      Caption         =   "ź�͡���"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1215
      TabIndex        =   3
      Top             =   5445
      Width           =   1140
   End
   Begin MSComctlLib.ListView ListViewItemList 
      Height          =   2760
      Left            =   1215
      TabIndex        =   2
      ToolTipText     =   "��¡���Թ��ҷ����͹��ѵ����"
      Top             =   2475
      Width           =   9555
      _ExtentX        =   16854
      _ExtentY        =   4868
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
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "�ӴѺ"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "�����Թ���"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "�����Թ���"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "�ӹǹ���͹��ѵ�"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "˹��¹Ѻ"
         Object.Width           =   2028
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "�Ţ����͡���"
         Object.Width           =   2646
      EndProperty
   End
   Begin VB.TextBox TextDocNo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3960
      TabIndex        =   0
      ToolTipText     =   "��͡�Ţ����͹���� ���ǡ����� Enter �����Ѻ"
      Top             =   1530
      Width           =   2085
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "��¡���Թ��� �����͹��ѵ����"
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
      Height          =   240
      Left            =   1170
      TabIndex        =   5
      Top             =   2205
      Width           =   3930
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "�Ţ����͡���͹��ѵ���ʹͫ����Թ��� :"
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
      Height          =   285
      Left            =   1170
      TabIndex        =   4
      Top             =   1530
      Width           =   2895
   End
End
Attribute VB_Name = "Form2_9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vDocNo As String
Dim vQuery As String

Private Sub CMDDelete_Click()
Dim vAnswer As Integer


If Me.TextDocno.Text <> "" And Me.ListViewItemList.ListItems.Count > 0 Then
  If UCase(Me.TextDocno.Text) = Me.ListViewItemList.ListItems(1).SubItems(5) Then
   vDocNo = Me.TextDocno.Text
   vAnswer = MsgBox("�س��ͧ���ź�͡���͹��ѵ���ʹͫ����Թ����Ţ��� " & vDocNo & "  ���������", vbYesNo, "Question Message ?")
   If vAnswer = 6 Then
     vQuery = "exec dbo.USP_NP_DeleteRequestConfirm '" & vDocNo & "' "
     gConnection.Execute vQuery
   Else
     Exit Sub
   End If
   Me.TextDocno.Text = ""
   MsgBox "ź�͡���͹��ѵ���ʹͫ����Թ����Ţ��� " & vDocNo & " ���º�������� ��سҵ�Ǩ�ͺ", vbCritical, "Send Information Message"
   Me.ListViewItemList.ListItems.Clear
  Else
    MsgBox "�Ţ����͹��ѵ��͡����ʹͫ����Թ����բ��������ç�Ѻ��¡���Թ���  ��سҵ�Ǩ�ͺ", vbCritical, "Send Error Message"
  End If
End If

End Sub

Private Sub CMDSearch_Click()
Dim vRecordset As New ADODB.Recordset
Dim i As Integer
Dim vListItem As ListItem

If Me.TextDocno.Text <> "" Then
 vDocNo = Me.TextDocno.Text
 Me.ListViewItemList.ListItems.Clear
 vQuery = "exec dbo.USP_NP_SearchRequestConfirm '" & vDocNo & "' "
   i = 1
   If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
      If vRecordset.Fields("ScriptDescription").Value = "" Then
        While Not vRecordset.EOF
         Set vListItem = Me.ListViewItemList.ListItems.Add(, , i)
         vListItem.SubItems(1) = vRecordset.Fields("itemcode").Value
         vListItem.SubItems(2) = vRecordset.Fields("itemname").Value
         vListItem.SubItems(3) = Format(vRecordset.Fields("confirmqty").Value, "##,##0.00")
         vListItem.SubItems(4) = vRecordset.Fields("unitcode").Value
         vListItem.SubItems(5) = vRecordset.Fields("docno").Value
         i = i + 1
        vRecordset.MoveNext
        Wend
      Else
        MsgBox vRecordset.Fields("ScriptDescription").Value
      End If
   End If
End If
End Sub

Private Sub TextDocno_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Call CMDSearch_Click
End If
End Sub
