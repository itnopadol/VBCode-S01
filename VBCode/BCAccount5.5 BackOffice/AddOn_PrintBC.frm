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
      Caption         =   "ไฟล์"
      Begin VB.Menu Order11 
         Caption         =   "เลือกบริษัทใหม่"
      End
      Begin VB.Menu Order12 
         Caption         =   "ออกจากโปรแกรม"
      End
   End
   Begin VB.Menu Order2 
      Caption         =   "จัดซื้อ"
      Begin VB.Menu Order21 
         Caption         =   "ยกเลิกการจ่ายชำระเอกสารซื้อ"
      End
   End
   Begin VB.Menu Order3 
      Caption         =   "จัดขาย"
      Begin VB.Menu Order31 
         Caption         =   "พิมพ์เอกสารขาย"
      End
      Begin VB.Menu Order32 
         Caption         =   "อัพเดท ข้อมูลภาษีขายก่อนออกรายงาน"
      End
   End
   Begin VB.Menu Order4 
      Caption         =   "เจ้าหนี้"
      Begin VB.Menu Order42 
         Caption         =   "พิมพ์ใบสำคัญตั้งเจ้าหนี้อื่น ๆ "
      End
      Begin VB.Menu Order41 
         Caption         =   "รายงาน"
         Begin VB.Menu Order411 
            Caption         =   "รายงาน สรุปค่าใช้จ่าย แยกตามเจ้าหนี้"
         End
         Begin VB.Menu Order412 
            Caption         =   "รายงานเคลื่อนไหวเจ้าหนี้"
         End
         Begin VB.Menu Order414 
            Caption         =   "รายงาน เคลื่อนไหวเจ้าหนี้ ตามช่วงเวลา"
         End
         Begin VB.Menu Order413 
            Caption         =   "รายงานตัดมัดจำจ่ายเจ้าหนี้"
         End
         Begin VB.Menu Order415 
            Caption         =   "รายงานยอดเจ้าหนี้ประจำเดือน"
         End
      End
   End
   Begin VB.Menu Order5 
      Caption         =   "ลูกหนี้"
      Begin VB.Menu Order52 
         Caption         =   "พิมพ์ใบสำคัญตั้งลูกหนี้อื่น ๆ"
      End
      Begin VB.Menu Order51 
         Caption         =   "รายงาน"
         Begin VB.Menu Order511 
            Caption         =   "รายงาน เคลื่อนไหวลูกหนี้ ทั้งหมด"
         End
         Begin VB.Menu Order512 
            Caption         =   "รายงาน เคลื่อนไหวลูกหนี้ ตามช่วงเวลา"
         End
         Begin VB.Menu Order513 
            Caption         =   "รายงานยอดลูกหนี้ประจำเดือน"
         End
      End
   End
   Begin VB.Menu Order6 
      Caption         =   "เช็ค/บัตร"
   End
   Begin VB.Menu Order7 
      Caption         =   "ธนาคาร/เงินสด"
   End
   Begin VB.Menu Order8 
      Caption         =   "สินค้าคงคลัง"
   End
   Begin VB.Menu Order9 
      Caption         =   "บัญชี"
      Begin VB.Menu Order91 
         Caption         =   "รายงาน แยกประเภท"
      End
      Begin VB.Menu Order92 
         Caption         =   "รายงาน สมุดรายวัน"
      End
      Begin VB.Menu Order93 
         Caption         =   "รายงาน งบทดลอง"
      End
      Begin VB.Menu Order94 
         Caption         =   "รายงาน สรุปค่าใช้จ่าย"
      End
   End
   Begin VB.Menu Order0 
      Caption         =   "ช่วยเหลือ"
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

vQuestion = MsgBox("คุณต้องการออกจากโปรแกรมใช่หรือไม่", vbYesNo + vbCritical, "ข้อความสอบถาม")
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

vQuestion = MsgBox("คุณต้องการออกจากโปรแกรมใช่หรือไม่", vbYesNo + vbCritical, "ข้อความสอบถาม")
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
