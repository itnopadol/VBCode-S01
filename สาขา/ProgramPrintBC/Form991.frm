VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form991 
   Caption         =   "¡��ԡ���͹��ѵ��͡��õ�ҧ �"
   ClientHeight    =   11490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form991.frx":0000
   ScaleHeight     =   11490
   ScaleMode       =   0  'User
   ScaleWidth      =   15392.4
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CMDConfirm 
      Caption         =   "͹��ѵ�"
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
      Left            =   2070
      TabIndex        =   11
      Top             =   7920
      Width           =   1320
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   585
      Top             =   10485
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   30
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form991.frx":9673
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form991.frx":BAC5
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CheckBox CHKAll 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "���͡������"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   675
      TabIndex        =   10
      Top             =   1980
      Width           =   1500
   End
   Begin MSComctlLib.ProgressBar PGBUpdate 
      Height          =   240
      Left            =   675
      TabIndex        =   9
      Top             =   7605
      Visible         =   0   'False
      Width           =   13965
      _ExtentX        =   24633
      _ExtentY        =   423
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton CMDSave 
      Caption         =   "¡��ԡ͹��ѵ�"
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
      Left            =   675
      TabIndex        =   8
      Top             =   7920
      Width           =   1320
   End
   Begin VB.PictureBox PicPoint 
      Height          =   195
      Left            =   0
      ScaleHeight     =   135
      ScaleWidth      =   270
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   330
   End
   Begin MSComCtl2.DTPicker DTPDocDate 
      Height          =   375
      Left            =   2385
      TabIndex        =   6
      Top             =   1350
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   55967745
      CurrentDate     =   40004
   End
   Begin MSComctlLib.ListView ListViewDocNo 
      Height          =   5190
      Left            =   675
      TabIndex        =   4
      Top             =   2385
      Width           =   13965
      _ExtentX        =   24633
      _ExtentY        =   9155
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "�ӴѺ"
         Object.Width           =   2620
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "�Ţ����͡���"
         Object.Width           =   4366
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "�ѹ����͡���"
         Object.Width           =   2620
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "��ҧ�ԧ1"
         Object.Width           =   11352
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "�����˵�"
         Object.Width           =   10479
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "������ҧ�͡���"
         Object.Width           =   3493
      EndProperty
   End
   Begin VB.ComboBox CMBModule 
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
      Left            =   5850
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1350
      Width           =   2310
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
      Left            =   10035
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1350
      Width           =   4605
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "�͡��û�Ш��ѹ��� :"
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
      Left            =   675
      TabIndex        =   5
      Top             =   1350
      Width           =   1590
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "����������� :"
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
      Left            =   4185
      TabIndex        =   2
      Top             =   1350
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "�������͡��� :"
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
      Left            =   8415
      TabIndex        =   0
      Top             =   1350
      Width           =   1545
   End
End
Attribute VB_Name = "Form991"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CHKAll_Click()
Dim i As Integer

If Me.CHKAll.Value = 1 Then
For i = 1 To Me.ListViewDocNo.ListItems.Count
Me.ListViewDocNo.ListItems(i).Checked = True
Next i
End If

If Me.CHKAll.Value = 0 Then
For i = 1 To Me.ListViewDocNo.ListItems.Count
Me.ListViewDocNo.ListItems(i).Checked = False
Next i
End If
End Sub

Private Sub CMBDocType_Click()
Call CheckData
End Sub

Private Sub CMBModule_Click()
If Me.CMBModule.ListIndex = 0 Then
Call Buy
ElseIf Me.CMBModule.ListIndex = 1 Then
Call Sale
ElseIf Me.CMBModule.ListIndex = 2 Then
Call Vendor
ElseIf Me.CMBModule.ListIndex = 3 Then
Call Customer
ElseIf Me.CMBModule.ListIndex = 4 Then
Call ItemStock
End If

Call CheckData
End Sub

Private Sub CMDConfirm_Click()
Dim vAnswer As Integer
Dim vDocNo As String
Dim i As Integer
Dim vQuery As String
Dim vCountSelect As Integer
Dim vType As Integer

On Error GoTo ErrDescription


If Me.ListViewDocNo.ListItems.Count > 0 Then
   vAnswer = MsgBox("�س��ͧ��� ͹��ѵ��͡��÷�����͡������������ ?", vbYesNo, "Send Message Question ?")
   
   If vAnswer = 6 Then
   
   For i = 1 To Me.ListViewDocNo.ListItems.Count
      If Me.ListViewDocNo.ListItems(i).Checked = True Then
         vCountSelect = vCountSelect + 1
      End If
   Next i
   
   If vCountSelect = 0 Then
      MsgBox "�ѧ��������͡�͡��÷���͹��ѵ� ��سҵ�Ǩ�ͺ", vbCritical, "Send Error Message"
      Me.ListViewDocNo.SetFocus
      Exit Sub
   End If
   
   Me.PGBUpdate.Visible = True
   Me.PGBUpdate.Min = 0
   Me.PGBUpdate.Max = vCountSelect
   vType = Me.CMBDocType.ListIndex
   
   If Me.CMBModule.ListIndex = 0 Then
      For i = 1 To Me.ListViewDocNo.ListItems.Count
      If Me.ListViewDocNo.ListItems(i).Checked = True Then
      vDocNo = Me.ListViewDocNo.ListItems(i).SubItems(1)
      vQuery = "exec dbo.USP_NP_UpdateConfirmModuleBuy " & vType & ",'" & vDocNo & "',1 "
      gConnection.Execute vQuery
      Me.PGBUpdate.Value = Me.PGBUpdate + 1
      End If
      Next i
   End If
   
   If Me.CMBModule.ListIndex = 1 Then
      For i = 1 To Me.ListViewDocNo.ListItems.Count
      If Me.ListViewDocNo.ListItems(i).Checked = True Then
      vDocNo = Me.ListViewDocNo.ListItems(i).SubItems(1)
      vQuery = "exec dbo.USP_NP_UpdateConfirmModuleSale " & vType & ",'" & vDocNo & "',1 "
      gConnection.Execute vQuery
      Me.PGBUpdate.Value = Me.PGBUpdate + 1
      End If
      Next i
   End If
   
   If Me.CMBModule.ListIndex = 4 Then
      For i = 1 To Me.ListViewDocNo.ListItems.Count
      If Me.ListViewDocNo.ListItems(i).Checked = True Then
      vDocNo = Me.ListViewDocNo.ListItems(i).SubItems(1)
      vQuery = "exec dbo.USP_NP_UpdateConfirmModuleStock " & vType & ",'" & vDocNo & "',1 "
      gConnection.Execute vQuery
      Me.PGBUpdate.Value = Me.PGBUpdate + 1
      End If
      Next i
   End If
   
   
   Me.ListViewDocNo.ListItems.Clear
   Me.PGBUpdate.Value = 0
   Me.PGBUpdate.Visible = False
   MsgBox "͹��ѵ��͡��÷�����͡��� ���º�������� ��سҵ�Ǩ�ͺ", vbInformation, "Send Information Message"
   
   Me.CMBModule.ListIndex = 0
   Me.DTPDocDate.Value = Now
   Me.CHKAll.Value = 0
   Me.CMBDocType.SetFocus
   End If

End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub CMDSave_Click()
Dim vAnswer As Integer
Dim vDocNo As String
Dim i As Integer
Dim vQuery As String
Dim vCountSelect As Integer
Dim vType As Integer

On Error GoTo ErrDescription


If Me.ListViewDocNo.ListItems.Count > 0 Then
   vAnswer = MsgBox("�س��ͧ��� ¡��ԡ���͹��ѵԢͧ�͡��÷�����͡������������ ?", vbYesNo, "Send Message Question ?")
   
   If vAnswer = 6 Then
   
   For i = 1 To Me.ListViewDocNo.ListItems.Count
      If Me.ListViewDocNo.ListItems(i).Checked = True Then
         vCountSelect = vCountSelect + 1
      End If
   Next i
   
   If vCountSelect = 0 Then
      MsgBox "�ѧ��������͡�͡��÷��� ¡��ԡ���͹��ѵ� ��سҵ�Ǩ�ͺ", vbCritical, "Send Error Message"
      Me.ListViewDocNo.SetFocus
      Exit Sub
   End If
   
   Me.PGBUpdate.Visible = True
   Me.PGBUpdate.Min = 0
   Me.PGBUpdate.Max = vCountSelect
   
   vType = Me.CMBDocType.ListIndex
   
   If Me.CMBModule.ListIndex = 0 Then
      For i = 1 To Me.ListViewDocNo.ListItems.Count
      If Me.ListViewDocNo.ListItems(i).Checked = True Then
      vDocNo = Me.ListViewDocNo.ListItems(i).SubItems(1)
      vQuery = "exec dbo.USP_NP_UpdateConfirmModuleBuy " & vType & ",'" & vDocNo & "',0 "
      gConnection.Execute vQuery
      Me.PGBUpdate.Value = Me.PGBUpdate + 1
      End If
      Next i
   End If
   
   If Me.CMBModule.ListIndex = 1 Then
      For i = 1 To Me.ListViewDocNo.ListItems.Count
      If Me.ListViewDocNo.ListItems(i).Checked = True Then
      vDocNo = Me.ListViewDocNo.ListItems(i).SubItems(1)
      vQuery = "exec dbo.USP_NP_UpdateConfirmModuleSale " & vType & ",'" & vDocNo & "',0 "
      gConnection.Execute vQuery
      Me.PGBUpdate.Value = Me.PGBUpdate + 1
      End If
      Next i
   End If
   
   If Me.CMBModule.ListIndex = 4 Then
      For i = 1 To Me.ListViewDocNo.ListItems.Count
      If Me.ListViewDocNo.ListItems(i).Checked = True Then
      vDocNo = Me.ListViewDocNo.ListItems(i).SubItems(1)
      vQuery = "exec dbo.USP_NP_UpdateConfirmModuleStock " & vType & ",'" & vDocNo & "',0 "
      gConnection.Execute vQuery
      Me.PGBUpdate.Value = Me.PGBUpdate + 1
      End If
      Next i
   End If
   
   Me.ListViewDocNo.ListItems.Clear
   Me.PGBUpdate.Value = 0
   Me.PGBUpdate.Visible = False
   MsgBox "¡��ԡ ���͹��ѵ��͡��÷�����͡��� ���º�������� ��سҵ�Ǩ�ͺ", vbInformation, "Send Information Message"
   
   Me.CMBModule.ListIndex = 0
   Me.DTPDocDate.Value = Now
   Me.CHKAll.Value = 0
   Me.CMBDocType.SetFocus
   End If

End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub DTPDocDate_Change()
Call CheckData
End Sub

Private Sub Form_Load()
Me.DTPDocDate.Value = Now
Call SetListViewColor(ListViewDocNo, PicPoint, vbWhite, vbLightGreen)
Call CreateModule
End Sub

Public Sub CreateModule()
Me.CMBModule.AddItem ("1.�Ѵ����")
Me.CMBModule.AddItem ("2.�Ѵ���")
Me.CMBModule.AddItem ("3.���˹��")
Me.CMBModule.AddItem ("4.�١˹��")
Me.CMBModule.AddItem ("5.�Թ��Ҥ���ѧ")
Me.CMBModule.AddItem ("6.����и�Ҥ��")

Me.CMBModule.ListIndex = 0
End Sub


Public Sub Buy()
Me.CMBDocType.Clear
Me.CMBDocType.AddItem ("�͡��� 1.��ʹͫ����Թ���")
Me.CMBDocType.AddItem ("�͡��� 2.���觫����Թ���")
Me.CMBDocType.AddItem ("�͡��� 3.㺨����Թ�Ѵ��")
Me.CMBDocType.AddItem ("�͡��� 4.㺨����Թ��ǧ˹��")
Me.CMBDocType.AddItem ("�͡��� 5.��Ѻ�Թ��Ҩҡ��ë���")
Me.CMBDocType.AddItem ("�͡��� 6.㺵��˹��ҡ��ë���")
Me.CMBDocType.AddItem ("�͡��� 7.��觤׹�Թ���")
Me.CMBDocType.AddItem ("�͡��� 8.�Ŵ˹��")
Me.CMBDocType.AddItem ("�͡��� 9.㺫����Թ�����к�ԡ��")
Me.CMBDocType.AddItem ("�͡��� 10.��觤׹�Թ���/Ŵ˹��")
Me.CMBDocType.AddItem ("�͡��� 11.�����˹��/�����Թ������˹��")
End Sub

Public Sub Sale()
Me.CMBDocType.Clear
Me.CMBDocType.AddItem ("�͡��� 1.��ʹ��Ҥ�")
Me.CMBDocType.AddItem ("�͡��� 2.���觢�¤�ҧ��(BackOrder)")
Me.CMBDocType.AddItem ("�͡��� 3.���觨ͧ")
Me.CMBDocType.AddItem ("�͡��� 4.���觢��")
Me.CMBDocType.AddItem ("�͡��� 5.��Ѻ�Թ�Ѵ��")
Me.CMBDocType.AddItem ("�͡��� 6.��Ѻ�Թ��ǧ˹��")
Me.CMBDocType.AddItem ("�͡��� 7.㺤׹�Թ�Ѻ��ǧ˹��")
Me.CMBDocType.AddItem ("�͡��� 8.����Թ���,��ԡ��")
Me.CMBDocType.AddItem ("�͡��� 9.����Թ��� POS")
Me.CMBDocType.AddItem ("�͡��� 10.��Ѻ�׹�Թ���/Ŵ˹��")
Me.CMBDocType.AddItem ("�͡��� 11.�����˹��/�����Թ���(�١���)")
End Sub

Public Sub Vendor()
Me.CMBDocType.Clear
Me.CMBDocType.AddItem ("�͡��� 1.���˹��¡��")
Me.CMBDocType.AddItem ("�͡��� 2.������˹����� �")
Me.CMBDocType.AddItem ("�͡��� 3.��Ѻ�ҧ���")
Me.CMBDocType.AddItem ("�͡��� 4.㺨��ª���˹��")
Me.CMBDocType.AddItem ("�͡��� 5.�Ѵ˹���٭(���˹��)")
End Sub

Public Sub Customer()
Me.CMBDocType.Clear
Me.CMBDocType.AddItem ("�͡��� 1.�١˹��¡�ҵ鹻�")
Me.CMBDocType.AddItem ("�͡��� 2.����١˹����� �")
Me.CMBDocType.AddItem ("�͡��� 3.��ҧ���")
Me.CMBDocType.AddItem ("�͡��� 4.��ҧ����ѵ��ѵ�")
Me.CMBDocType.AddItem ("�͡��� 5.����稪��Ǥ���")
Me.CMBDocType.AddItem ("�͡��� 6.������Ѻ�Թ/�Ѻ����˹��")
Me.CMBDocType.AddItem ("�͡��� 7.�Ѵ˹���٭(�١˹��)")
End Sub

Public Sub ItemStock()
Me.CMBDocType.Clear
Me.CMBDocType.AddItem ("�͡��� 1.�Թ���¡��")
Me.CMBDocType.AddItem ("�͡��� 2.㺢��ԡ���Թ���,�ѵ�شԺ")
Me.CMBDocType.AddItem ("�͡��� 3.��ԡ���Թ���,�ѵ�شԺ")
Me.CMBDocType.AddItem ("�͡��� 4.��Ѻ�׹�Թ���,�ѵ�شԺ")
Me.CMBDocType.AddItem ("�͡��� 5.��Ѻ�Թ���������ٻ")
Me.CMBDocType.AddItem ("�͡��� 6.㺢��͹�Թ���")
Me.CMBDocType.AddItem ("�͡��� 7.��͹�Թ��������ҧ��ѧ")
Me.CMBDocType.AddItem ("�͡��� 8.㺻�Ѻ��ا�Թ�����ѧ��Ǩ�Ѻ")
End Sub

Public Sub CheckData()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset

Dim vListDoc As ListItem
Dim i As Integer
Dim vDocdate As String
Dim vType As Integer
Dim vMemIsConfirm As Integer

On Error GoTo ErrDescription

If Me.CMBModule.Text <> "" And Me.CMBDocType.Text <> "" Then

 If Me.CMBModule.ListIndex = 0 Then
   Me.ListViewDocNo.ListItems.Clear
   vDocdate = Me.DTPDocDate.Day & "/" & Me.DTPDocDate.Month & "/" & Me.DTPDocDate.Year
   vType = Me.CMBDocType.ListIndex
   
   vQuery = "exec dbo.USP_NP_CancelConfirmModuleBuy " & vType & ",'" & vDocdate & "' "
   If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
       vRecordset.MoveFirst
       For i = 1 To vRecordset.RecordCount
               vMemIsConfirm = Trim(vRecordset.Fields("isconfirm").Value)
               If vMemIsConfirm = 0 Then
               Set vListDoc = Me.ListViewDocNo.ListItems.Add(, , i, , 1)
               Else
               Set vListDoc = Me.ListViewDocNo.ListItems.Add(, , i, , 2)
               End If
               vListDoc.SubItems(1) = Trim(vRecordset.Fields("docno").Value)
               vListDoc.SubItems(2) = Trim(vRecordset.Fields("docdate").Value)
               vListDoc.SubItems(3) = Trim(vRecordset.Fields("code").Value) & "/" & Trim(vRecordset.Fields("name1").Value)
               vListDoc.SubItems(4) = Trim(vRecordset.Fields("mydescription").Value)
               vListDoc.SubItems(5) = Trim(vRecordset.Fields("creatorcode").Value)
       vRecordset.MoveNext
   Next i
   End If
   vRecordset.Close
   End If
   
   If Me.CMBModule.ListIndex = 1 Then
   Me.ListViewDocNo.ListItems.Clear
   vDocdate = Me.DTPDocDate.Day & "/" & Me.DTPDocDate.Month & "/" & Me.DTPDocDate.Year
   vType = Me.CMBDocType.ListIndex
   
   vQuery = "exec dbo.USP_NP_CancelConfirmModuleSale " & vType & ",'" & vDocdate & "' "
   If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
       vRecordset.MoveFirst
       For i = 1 To vRecordset.RecordCount
               vMemIsConfirm = Trim(vRecordset.Fields("isconfirm").Value)
               If vMemIsConfirm = 0 Then
               Set vListDoc = Me.ListViewDocNo.ListItems.Add(, , i, , 1)
               Else
               Set vListDoc = Me.ListViewDocNo.ListItems.Add(, , i, , 2)
               End If
               vListDoc.SubItems(1) = Trim(vRecordset.Fields("docno").Value)
               vListDoc.SubItems(2) = Trim(vRecordset.Fields("docdate").Value)
               vListDoc.SubItems(3) = Trim(vRecordset.Fields("code").Value) & "/" & Trim(vRecordset.Fields("name1").Value)
               vListDoc.SubItems(4) = Trim(vRecordset.Fields("mydescription").Value)
               vListDoc.SubItems(5) = Trim(vRecordset.Fields("creatorcode").Value)
       vRecordset.MoveNext
   Next i
   End If
   vRecordset.Close
   End If
   
   If Me.CMBModule.ListIndex = 2 Then
   
   End If
   
   If Me.CMBModule.ListIndex = 3 Then
   
   End If
   
   If Me.CMBModule.ListIndex = 4 Then
   Me.ListViewDocNo.ListItems.Clear
   vDocdate = Me.DTPDocDate.Day & "/" & Me.DTPDocDate.Month & "/" & Me.DTPDocDate.Year
   vType = Me.CMBDocType.ListIndex
   
   vQuery = "exec dbo.USP_NP_CancelConfirmModuleStock " & vType & ",'" & vDocdate & "' "
   If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
       vRecordset.MoveFirst
       For i = 1 To vRecordset.RecordCount
               vMemIsConfirm = Trim(vRecordset.Fields("isconfirm").Value)
               If vMemIsConfirm = 0 Then
               Set vListDoc = Me.ListViewDocNo.ListItems.Add(, , i, , 1)
               Else
               Set vListDoc = Me.ListViewDocNo.ListItems.Add(, , i, , 2)
               End If
               vListDoc.SubItems(1) = Trim(vRecordset.Fields("docno").Value)
               vListDoc.SubItems(2) = Trim(vRecordset.Fields("docdate").Value)
               vListDoc.SubItems(3) = Trim(vRecordset.Fields("code").Value) & "/" & Trim(vRecordset.Fields("name1").Value)
               vListDoc.SubItems(4) = Trim(vRecordset.Fields("mydescription").Value)
               vListDoc.SubItems(5) = Trim(vRecordset.Fields("creatorcode").Value)
       vRecordset.MoveNext
   Next i
   End If
   vRecordset.Close
   End If
   
Else
   Me.ListViewDocNo.ListItems.Clear
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub
