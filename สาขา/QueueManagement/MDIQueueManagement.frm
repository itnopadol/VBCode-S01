VERSION 5.00
Begin VB.MDIForm MDIQueueManagement 
   BackColor       =   &H8000000C&
   Caption         =   "����� �Ѵ��ä�ǨѴ�Թ���"
   ClientHeight    =   9315
   ClientLeft      =   4560
   ClientTop       =   3195
   ClientWidth     =   14880
   Icon            =   "MDIQueueManagement.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIQueueManagement.frx":27A2
   StartUpPosition =   2  'CenterScreen
   Begin VB.Menu OrderFiles 
      Caption         =   "     ���"
      NegotiatePosition=   2  'Middle
      Begin VB.Menu OrderLogIN 
         Caption         =   "���͡��ҷӧҹ����"
      End
      Begin VB.Menu Line1 
         Caption         =   "-"
      End
      Begin VB.Menu OrderExit 
         Caption         =   "�͡�����"
      End
   End
   Begin VB.Menu OrderProgram 
      Caption         =   "     �����"
      Begin VB.Menu OrderQueue 
         Caption         =   "�����èѴ����Թ���"
      End
      Begin VB.Menu OrderCustReceive 
         Caption         =   "��ѺʶҹФ�Ƿ���١����Ѻ�ͧ�ҡ㺨����Թ���"
      End
   End
   Begin VB.Menu OrderReport 
      Caption         =   "     ��§ҹ"
   End
   Begin VB.Menu OrderWindows 
      Caption         =   "     ˹�ҵ�ҧ"
      WindowList      =   -1  'True
   End
End
Attribute VB_Name = "MDIQueueManagement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MDIForm_Load()
OrderProgram.Enabled = False
OrderReport.Enabled = False
OrderWindows.Enabled = False
FrmLogIN.Show
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
Dim vQuestion As Integer
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String

vQuestion = MsgBox("�س��ͧ����͡�ҡ��������������", vbYesNo + vbCritical, "��ͤ����ͺ���")
If vQuestion = 6 Then
  If sConnection.State = 1 Then
  sConnection.Close
  End If
  If vConnection.State = 1 Then
      vConnection.Close
  End If
  If qConnection.State = 1 Then
      qConnection.Close
  End If
Else
  Cancel = True
End If
End Sub

Private Sub OrderCustReceive_Click()
Unload FrmQueue
FrmCheckCustReceive.Show
FrmCheckCustReceive.SetFocus
End Sub

Private Sub OrderExit_Click()
Dim vQuestion As Integer
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String

vQuestion = MsgBox("�س��ͧ����͡�ҡ��������������", vbYesNo + vbCritical, "��ͤ����ͺ���")
If vQuestion = 6 Then
  If sConnection.State = 1 Then
  sConnection.Close
  End If
  If vConnection.State = 1 Then
      vConnection.Close
  End If
  If qConnection.State = 1 Then
      qConnection.Close
  End If
Else
  Exit Sub
End If
End Sub

Private Sub OrderItemPayment_Click()
'Unload FrmQueue
'FrmPayItemCust.Show
'FrmPayItemCust.SetFocus
'FrmPrintReceiveSlip.Show
'FrmPrintReceiveSlip.SetFocus
End Sub

Private Sub OrderItemPaymentReserve_Click()
'FormPrintPayGoodsRes.Show
'FormPrintPayGoodsRes.SetFocus
End Sub

Private Sub OrderLogIN_Click()
MDIQueueManagement.Caption = Trim("����� �Ѵ��ä�ǨѴ�Թ���")
Unload FrmQueue
Unload FrmPicker
Unload FrmCheckQTY
OrderProgram.Enabled = False
OrderReport.Enabled = False
OrderWindows.Enabled = False
FrmLogIN.Show
FrmLogIN.SetFocus
End Sub

Private Sub OrderQueue_Click()
Unload FrmPayItemCust
FrmQueue.Show
FrmQueue.SetFocus
End Sub

Private Sub OrderReport_Click()
FrmReport101.Show
FrmReport101.SetFocus
End Sub
