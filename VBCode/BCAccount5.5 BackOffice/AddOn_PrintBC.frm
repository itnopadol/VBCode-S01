VERSION 5.00
Begin VB.MDIForm AddOn_PrintBC 
   BackColor       =   &H8000000C&
   Caption         =   "BCAccount5.5 BackOffice"
   ClientHeight    =   8985
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   12000
   LinkTopic       =   "MDIForm1"
   Picture         =   "AddOn_PrintBC.frx":0000
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu Order1 
      Caption         =   "���"
      Begin VB.Menu Order11 
         Caption         =   "���͡����ѷ����"
      End
      Begin VB.Menu Order12 
         Caption         =   "�͡�ҡ�����"
      End
   End
   Begin VB.Menu Order2 
      Caption         =   "�Ѵ����"
      Begin VB.Menu Order21 
         Caption         =   "¡��ԡ��è��ª����͡��ë���"
      End
   End
   Begin VB.Menu Order3 
      Caption         =   "�Ѵ���"
      Begin VB.Menu Order31 
         Caption         =   "������͡��â��"
      End
      Begin VB.Menu Order32 
         Caption         =   "�Ѿഷ ���������բ�¡�͹�͡��§ҹ"
      End
   End
   Begin VB.Menu Order4 
      Caption         =   "���˹��"
      Begin VB.Menu Order42 
         Caption         =   "�������Ӥѭ������˹����� � "
      End
      Begin VB.Menu Order41 
         Caption         =   "��§ҹ"
         Begin VB.Menu Order411 
            Caption         =   "��§ҹ ��ػ�������� �¡������˹��"
         End
         Begin VB.Menu Order412 
            Caption         =   "��§ҹ����͹������˹��"
         End
         Begin VB.Menu Order414 
            Caption         =   "��§ҹ ����͹������˹�� �����ǧ����"
         End
         Begin VB.Menu Order413 
            Caption         =   "��§ҹ�Ѵ�Ѵ�Ө������˹��"
         End
         Begin VB.Menu Order415 
            Caption         =   "��§ҹ�ʹ���˹���Ш���͹"
         End
      End
   End
   Begin VB.Menu Order5 
      Caption         =   "�١˹��"
      Begin VB.Menu Order52 
         Caption         =   "�������Ӥѭ����١˹����� �"
      End
      Begin VB.Menu Order51 
         Caption         =   "��§ҹ"
         Begin VB.Menu Order511 
            Caption         =   "��§ҹ ����͹����١˹�� ������"
         End
         Begin VB.Menu Order512 
            Caption         =   "��§ҹ ����͹����١˹�� �����ǧ����"
         End
         Begin VB.Menu Order513 
            Caption         =   "��§ҹ�ʹ�١˹���Ш���͹"
         End
      End
   End
   Begin VB.Menu Order6 
      Caption         =   "��/�ѵ�"
   End
   Begin VB.Menu Order7 
      Caption         =   "��Ҥ��/�Թʴ"
   End
   Begin VB.Menu Order8 
      Caption         =   "�Թ��Ҥ���ѧ"
   End
   Begin VB.Menu Order9 
      Caption         =   "�ѭ��"
      Begin VB.Menu Order91 
         Caption         =   "��§ҹ �¡������"
      End
      Begin VB.Menu Order92 
         Caption         =   "��§ҹ ��ش����ѹ"
      End
      Begin VB.Menu Order93 
         Caption         =   "��§ҹ �����ͧ"
      End
      Begin VB.Menu Order94 
         Caption         =   "��§ҹ ��ػ��������"
      End
   End
   Begin VB.Menu Order0 
      Caption         =   "���������"
   End
End
Attribute VB_Name = "AddOn_PrintBC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MDIForm_Load()
Order2.Enabled = False
Order3.Enabled = False
Order4.Enabled = False
Order5.Enabled = False
Order6.Enabled = False
Order7.Enabled = False
Order8.Enabled = False
Order9.Enabled = False
Order0.Enabled = False
FormLogIN.Show
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
Dim vQuestion As Integer
On Error GoTo Errdescription

vQuestion = MsgBox("�س��ͧ����͡�ҡ��������������", vbYesNo + vbCritical, "��ͤ����ͺ���")
If vQuestion = 6 Then
            If gConnection.State = 1 Then
                gConnection.Close
            End If
            If vConnection.State = 1 Then
                vConnection.Close
            End If
End
Else
Exit Sub
End If

Errdescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Order11_Click()
FormLogIN.Show
End Sub

Private Sub Order12_Click()
Dim vQuestion As Integer
On Error GoTo Errdescription

vQuestion = MsgBox("�س��ͧ����͡�ҡ��������������", vbYesNo + vbCritical, "��ͤ����ͺ���")
If vQuestion = 6 Then
            If gConnection.State = 1 Then
                gConnection.Close
            End If
            If vConnection.State = 1 Then
                vConnection.Close
            End If
End
Else
Exit Sub
End If

Errdescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Order21_Click()
Form21.Show
Form21.SetFocus
End Sub

Private Sub Order31_Click()
Form31.Show
Form31.SetFocus
End Sub

Private Sub Order32_Click()
Form32.Show
Form32.SetFocus
End Sub

Private Sub Order411_Click()
Form411.Show
Form411.SetFocus
End Sub

Private Sub Order412_Click()
Form412.Show
Form412.SetFocus
End Sub

Private Sub Order413_Click()
Form413.Show
Form413.SetFocus
End Sub

Private Sub Order414_Click()
Form414.Show
Form414.SetFocus
End Sub

Private Sub Order415_Click()
Form415.Show
Form415.SetFocus
End Sub

Private Sub Order42_Click()
Form42.Show
Form42.SetFocus
End Sub

Private Sub Order511_Click()
Form511.Show
Form511.SetFocus
End Sub

Private Sub Order512_Click()
Form512.Show
Form512.SetFocus
End Sub

Private Sub Order513_Click()
Form513.Show
Form513.SetFocus
End Sub

Private Sub Order52_Click()
Form52.Show
Form52.SetFocus
End Sub

Private Sub Order91_Click()
Form91.Show
Form91.SetFocus
End Sub

Private Sub Order92_Click()
Form92.Show
Form92.SetFocus
End Sub

Private Sub Order93_Click()
Form93.Show
Form93.SetFocus
End Sub

Private Sub Order94_Click()
Form94.Show
Form94.SetFocus
End Sub
