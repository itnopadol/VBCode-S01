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
      Caption         =   "โหมดติดต่อฐานข้อมูล"
      Begin VB.Menu Order001 
         Caption         =   "Connect Company"
      End
      Begin VB.Menu Order002 
         Caption         =   "Close Program"
      End
   End
   Begin VB.Menu Order1 
      Caption         =   "โหมดการทำงาน"
      Begin VB.Menu Order101 
         Caption         =   "เช็คบาร์โค้ดพิมพ์ป้ายราคา"
      End
      Begin VB.Menu Order102 
         Caption         =   "ตรวจสอบสต็อกประจำวัน"
      End
      Begin VB.Menu Order103 
         Caption         =   "เช็คบาร์โค้ดทำใบขอโอนสินค้า"
      End
      Begin VB.Menu Order104 
         Caption         =   "เช็คบาร์โค้ดเช็คราคาสินค้า"
      End
      Begin VB.Menu Order105 
         Caption         =   "เช็คบาร์โค้ดเก็บที่เก็บสินค้า"
      End
      Begin VB.Menu Order106 
         Caption         =   "ยิงบาร์โค้ดทำใบหยิบ"
      End
      Begin VB.Menu Order107 
         Caption         =   "ตรวจสอบที่เก็บสินค้า"
      End
      Begin VB.Menu Order114 
         Caption         =   "ตรวจนับสินค้าตามระบบ Cycle-Count"
      End
      Begin VB.Menu Order111 
         Caption         =   "ตรวจนับสต๊อกตาม Store"
      End
      Begin VB.Menu Order112 
         Caption         =   "ยกเลิกใบปรับปรุงหลังการตรวจนับ"
      End
      Begin VB.Menu Order110 
         Caption         =   "ตรวจนับสต๊อกตามที่เก็บ"
      End
      Begin VB.Menu Order108 
         Caption         =   "นับสต็อกประจำปี"
      End
      Begin VB.Menu Order109 
         Caption         =   "ตรวจสอบข้อมูลสินค้าใกล้เคียง"
      End
      Begin VB.Menu Order113 
         Caption         =   "รายงาน ติดตามการแก้ไขสินค้าติดลบประจำวัน"
      End
   End
   Begin VB.Menu Order2 
      Caption         =   "หน้าต่างโปรแกรมที่เปิดไว้"
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

vAnswer = MsgBox("คุณต้องการออกจากการใช้งานโปรแกรมใช่หรือไม่", vbYesNo, "Message Question")
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
MsgBox "ให้ใช้งาน เครื่อง แฮนด์เฮลด์ หรือ เครื่อง MC3000", vbCritical, "Send Information Message"

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
'MsgBox "ให้ใช้งาน เครื่อง แฮนด์เฮลด์ หรือ เครื่อง MC3000", vbCritical, "Send Information Message"

Form105.Show
Form105.SetFocus
End Sub

Private Sub Order106_Click()
'MsgBox "ยกเลิกใช้งาน", vbCritical, "Send Information Message"
Form106.Show
Form106.SetFocus
End Sub

Private Sub Order107_Click()
'MsgBox "กำลังอยู่ในช่วงปรับปรุงโปรแกรม อีกประมาณ 1 อาทิตย์", vbCritical, "Send Error Message"
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
'MsgBox "กำลังอยู่ในช่วงปรับปรุงโปรแกรม อีกประมาณ 1 อาทิตย์", vbCritical, "Send Error Message"
'Form110.Show
'Form110.SetFocus
End Sub

Private Sub Order111_Click()
'MsgBox "กำลังอยู่ในช่วงปรับปรุงโปรแกรม อีกประมาณ 1 อาทิตย์", vbCritical, "Send Error Message"
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
