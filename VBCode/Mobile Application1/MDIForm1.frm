VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "Mobile Application Version1.6_Picking"
   ClientHeight    =   8805
   ClientLeft      =   2025
   ClientTop       =   750
   ClientWidth     =   14385
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIForm1.frx":0CFA
   Begin VB.Menu Order0 
      Caption         =   "�����Դ��Ͱҹ������"
      Begin VB.Menu Order001 
         Caption         =   "Connect Company"
      End
      Begin VB.Menu Order002 
         Caption         =   "Close Program"
      End
   End
   Begin VB.Menu Order1 
      Caption         =   "������÷ӧҹ"
      Begin VB.Menu Order101 
         Caption         =   "�示����鴾��������Ҥ�"
      End
      Begin VB.Menu Order102 
         Caption         =   "��Ǩ�ͺʵ�͡��Ш��ѹ"
      End
      Begin VB.Menu Order103 
         Caption         =   "�示����鴷�㺢��͹�Թ���"
      End
      Begin VB.Menu Order104 
         Caption         =   "�示��������Ҥ��Թ���"
      End
      Begin VB.Menu Order105 
         Caption         =   "�示������纷�����Թ���"
      End
      Begin VB.Menu Order106 
         Caption         =   "�ԧ�����鴷����Ժ"
      End
      Begin VB.Menu Order107 
         Caption         =   "��Ǩ�ͺ������Թ���"
      End
      Begin VB.Menu Order114 
         Caption         =   "��Ǩ�Ѻ�Թ��ҵ���к� Cycle-Count"
      End
      Begin VB.Menu Order111 
         Caption         =   "��Ǩ�Ѻʵ�͡��� Store"
      End
      Begin VB.Menu Order112 
         Caption         =   "¡��ԡ㺻�Ѻ��ا��ѧ��õ�Ǩ�Ѻ"
      End
      Begin VB.Menu Order110 
         Caption         =   "��Ǩ�Ѻʵ�͡��������"
      End
      Begin VB.Menu Order108 
         Caption         =   "�Ѻʵ�͡��Шӻ�"
      End
      Begin VB.Menu Order109 
         Caption         =   "��Ǩ�ͺ�������Թ��������§"
      End
      Begin VB.Menu Order113 
         Caption         =   "��§ҹ �Դ����������Թ��ҵԴź��Ш��ѹ"
      End
   End
   Begin VB.Menu Order2 
      Caption         =   "˹�ҵ�ҧ���������Դ���"
      WindowList      =   -1  'True
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MDIForm_Load()
Order1.Enabled = False
Order2.Enabled = False
Form001.Show
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
Dim vAnswer As Integer
Dim vQuery As String

vAnswer = MsgBox("�س��ͧ����͡�ҡ�����ҹ��������������", vbYesNo, "Message Question")
If vAnswer = 6 Then
    If gConnection.State = 1 Then
      vQuery = "exec dbo.USP_MB_UpdateUsedStockNo '" & vRemDocDate & "','" & vRemStockDocNo & "','" & vRemCountID & "','" & vRemWHCode & "','" & vRemStoreCode & "','" & vRemRowID & "','" & vUserID & "',0 "
      gConnection.Execute vQuery
      gConnection.Close
    End If
Else
    Cancel = True
End If
End Sub

Private Sub Order001_Click()
Form001.Show
End Sub

Private Sub Order002_Click()
Unload MDIForm1
End Sub

Private Sub Order101_Click()
Form101.Show
Form101.SetFocus
End Sub

Private Sub Order102_Click()
MsgBox "�����ҹ ����ͧ �ι����Ŵ� ���� ����ͧ MC3000", vbCritical, "Send Information Message"

'Form102.Show
'Form102.SetFocus
End Sub

Private Sub Order103_Click()
Form104.Show
Form104.SetFocus
End Sub

Private Sub Order104_Click()
Form103.Show
Form103.SetFocus
End Sub

Private Sub Order105_Click()
'MsgBox "�����ҹ ����ͧ �ι����Ŵ� ���� ����ͧ MC3000", vbCritical, "Send Information Message"

Form105.Show
Form105.SetFocus
End Sub

Private Sub Order106_Click()
'MsgBox "¡��ԡ��ҹ", vbCritical, "Send Information Message"
Form106.Show
Form106.SetFocus
End Sub

Private Sub Order107_Click()
'MsgBox "���ѧ����㹪�ǧ��Ѻ��ا����� �ա����ҳ 1 �ҷԵ��", vbCritical, "Send Error Message"
Form107.Show
Form107.SetFocus
End Sub

Private Sub Order108_Click()
'Form108.Show
'Form108.SetFocus
End Sub

Private Sub Order109_Click()
Form109.Show
Form109.SetFocus
End Sub

Private Sub Order110_Click()
'MsgBox "���ѧ����㹪�ǧ��Ѻ��ا����� �ա����ҳ 1 �ҷԵ��", vbCritical, "Send Error Message"
'Form110.Show
'Form110.SetFocus
End Sub

Private Sub Order111_Click()
'MsgBox "���ѧ����㹪�ǧ��Ѻ��ا����� �ա����ҳ 1 �ҷԵ��", vbCritical, "Send Error Message"
Form112.Show
Form112.SetFocus
End Sub

Private Sub Order112_Click()
FormCancelAdjustStock.Show
FormCancelAdjustStock.SetFocus
End Sub

Private Sub Order113_Click()
Form113.Show
Form113.SetFocus
End Sub

Private Sub Order114_Click()
Form114.Show
Form114.SetFocus
End Sub
