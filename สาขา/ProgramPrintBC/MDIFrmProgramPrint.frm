VERSION 5.00
Begin VB.MDIForm MDIFrmProgramPrint 
   BackColor       =   &H8000000C&
   Caption         =   "โปรแกรมพิมพ์เอกสาร BCAccount 5.5 Version 1.1"
   ClientHeight    =   10305
   ClientLeft      =   1740
   ClientTop       =   855
   ClientWidth     =   14880
   Icon            =   "MDIFrmProgramPrint.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIFrmProgramPrint.frx":08CA
   Begin VB.Menu Order1 
      Caption         =   "ไฟล์"
      Begin VB.Menu Order1_0 
         Caption         =   "เลือกบริษัททำงาน"
      End
      Begin VB.Menu Order1_2 
         Caption         =   "-"
      End
      Begin VB.Menu Order1_1 
         Caption         =   "ออกโปรแกรม"
      End
   End
   Begin VB.Menu Order2 
      Caption         =   "จัดซื้อ"
      Begin VB.Menu Order2_8 
         Caption         =   "พิมพ์ใบเสนอซื้อสินค้า"
      End
      Begin VB.Menu Order2_0 
         Caption         =   "พิมพ์ใบสั่งซื้อ"
      End
      Begin VB.Menu Order2_4 
         Caption         =   "พิมพ์ใบตรวจรับสินค้าเข้าคลัง"
      End
      Begin VB.Menu Order2_7 
         Caption         =   "ยกเลิกรายตัวสินค้าใบเสนอซื้อที่ไม่ได้อนุมัติ"
      End
      Begin VB.Menu Order2_6 
         Caption         =   "รวมสินค้าในใบ PR ที่อนุมัติแล้ว"
      End
      Begin VB.Menu Order2_5 
         Caption         =   "ดึงทะเบียนสินค้าที่ถูกยกเลิกกลับมาใช้ใหม่"
      End
      Begin VB.Menu Order2_9 
         Caption         =   "ลบเอกสาร การอนุมัติใบเสนอซื้อสินค้า"
      End
      Begin VB.Menu Order2_3 
         Caption         =   "-"
      End
      Begin VB.Menu Order2_1 
         Caption         =   "เปลี่ยนวันที่หมดอายุของใบสั่งซื้อ"
      End
      Begin VB.Menu Order2_2 
         Caption         =   "รายงาน ตรวจสอบใบสั่งซื้อ"
      End
      Begin VB.Menu Order2_52 
         Caption         =   "รายงาน ตรวจสอบสถานะใบเสนอซื้อสินค้า (PR)"
      End
   End
   Begin VB.Menu Order3 
      Caption         =   "จัดขาย"
      Begin VB.Menu Order3_0 
         Caption         =   "พิมพ์ใบเสนอราคา"
      End
      Begin VB.Menu Order3_1 
         Caption         =   "พิมพ์ใบBackOrder"
      End
      Begin VB.Menu Order3_14 
         Caption         =   "ยกเลิก Quotation (ใบเสนอราคา)"
      End
      Begin VB.Menu Order3_19 
         Caption         =   "ขอยกเลิกใบ Back Order"
      End
      Begin VB.Menu Order3_17 
         Caption         =   "เปลี่ยนวันที่หมดอายุใบเสนอราคาและBack Order"
      End
      Begin VB.Menu Order3_9 
         Caption         =   "-"
      End
      Begin VB.Menu OrderReserve 
         Caption         =   "พิมพ์ใบกำกับสินค้า"
      End
      Begin VB.Menu Order3_18 
         Caption         =   "พิมพ์ใบขออนุมัติขายสินค้าลูกค้าที่เกินวงเงิน"
      End
      Begin VB.Menu Order3_2 
         Caption         =   "พิมพ์ใบสั่งขาย/ใบจัดคิว/ใบจัดสินค้า"
      End
      Begin VB.Menu OrderPick 
         Caption         =   "PickingRequest"
      End
      Begin VB.Menu CheckOut 
         Caption         =   "CheckOutItem"
      End
      Begin VB.Menu Order3_15 
         Caption         =   "ยกเลิกเอกสารขาย"
      End
      Begin VB.Menu Order3_10 
         Caption         =   "-"
      End
      Begin VB.Menu Order3_3 
         Caption         =   "พิมพ์ใบมัดจำ"
      End
      Begin VB.Menu Order3_11 
         Caption         =   "-"
      End
      Begin VB.Menu Order3_4 
         Caption         =   "พิมพ์เอกสารขายสินค้า"
      End
      Begin VB.Menu Order3_20 
         Caption         =   "พิมพ์ใบกำกับภาษีเอกสาร POS"
      End
      Begin VB.Menu Order3_12 
         Caption         =   "-"
      End
      Begin VB.Menu Order3_5 
         Caption         =   "พิมพ์ใบรับคืนสินค้า/ลดหนี้"
      End
      Begin VB.Menu OrderEditReturn 
         Caption         =   "แก้ไขข้อมูลใบลดหนี้"
      End
      Begin VB.Menu Order3_13 
         Caption         =   "-"
      End
      Begin VB.Menu Order3_6 
         Caption         =   "พิมพ์ใบเพิ่มหนี้สินค้า(ลูกค้า)"
      End
      Begin VB.Menu Order3_7 
         Caption         =   "-"
      End
      Begin VB.Menu Order37 
         Caption         =   "แก้ไขรายละเอียดเอกสารขายที่ถูกอนุมัติแล้ว"
      End
      Begin VB.Menu Order391 
         Caption         =   "-"
      End
      Begin VB.Menu Order3_8 
         Caption         =   "รายงาน"
         Begin VB.Menu Order3_81 
            Caption         =   "รายงานยอดขายพนักงานขายประจำเดือน"
         End
         Begin VB.Menu Order3_82 
            Caption         =   "รายงาน การขายสินค้า ณ จุดขายต่าง ๆ"
         End
         Begin VB.Menu Order392 
            Caption         =   "รายงานใบมัดจำ แยกตามลูกค้า"
         End
         Begin VB.Menu Order393 
            Caption         =   "รายงาน ยอดขายก๊อก"
         End
         Begin VB.Menu Order394 
            Caption         =   "รายงาน สรุปบิลขายลดราคาสินค้า"
         End
         Begin VB.Menu Order395 
            Caption         =   "รายงาน เกี่ยวกับเอกสารขาย"
         End
         Begin VB.Menu Order396 
            Caption         =   "รายงาน Run Number เอกสาร"
         End
      End
   End
   Begin VB.Menu DO1 
      Caption         =   "จัดส่ง"
      Begin VB.Menu DO101 
         Caption         =   "จัดการระบบ"
         Begin VB.Menu DO102 
            Caption         =   "กำหนด ข้อมูลระดับความสำคัญ"
         End
         Begin VB.Menu DO104 
            Caption         =   "กำหนด ข้อมูลสถานที่ขนส่ง"
         End
         Begin VB.Menu DO105 
            Caption         =   "กำหนด ข้อมูลเส้นทางขนส่ง"
         End
         Begin VB.Menu DO106 
            Caption         =   "กำหนด ข้อมูลรถขนส่ง"
         End
         Begin VB.Menu DO107 
            Caption         =   "กำหนด ข้อมูลพนักงานขนส่ง"
         End
      End
      Begin VB.Menu DO108 
         Caption         =   "โปรแกรม"
         Begin VB.Menu DO109 
            Caption         =   "บันทึกใบขนส่ง"
         End
         Begin VB.Menu DO114 
            Caption         =   "กำหนดเวลาส่งสินค้าตามใบขอเข้าคิวขนส่ง"
         End
      End
      Begin VB.Menu DO110 
         Caption         =   "อนุมัติ ใบจัดคิวสินค้า"
      End
      Begin VB.Menu DO111 
         Caption         =   "คำนวณยอดสินค้าในใบจัดคิว"
      End
      Begin VB.Menu DO112 
         Caption         =   "รายงาน"
         Begin VB.Menu DO113 
            Caption         =   "รายงาน การคิดค่าเที่ยวพนักงานขนส่ง"
         End
      End
   End
   Begin VB.Menu Order4 
      Caption         =   "เจ้าหนี้"
      Begin VB.Menu Order41 
         Caption         =   "พิมพ์ใบสำคัญตั้งเจ้าหนี้อื่น ๆ"
      End
      Begin VB.Menu Order42 
         Caption         =   "รับวางบิลเจ้าหนี้ชั่วคราว"
      End
      Begin VB.Menu Order4_0 
         Caption         =   "รายงาน"
         Begin VB.Menu Order40_1 
            Caption         =   "รายงานยอดซื้อตามผู้จำหน่าย"
         End
         Begin VB.Menu Order40_2 
            Caption         =   "รายงานสรุปค่าใช้จ่ายประจำวัน"
         End
         Begin VB.Menu Order40_3 
            Caption         =   "รายงานเคลื่อนไหวเจ้าหนี้-ทั้งหมด"
         End
         Begin VB.Menu Order40_5 
            Caption         =   "รายงาน เคลื่อนไหวลูกหนี้ ตามช่วงวันที่"
         End
         Begin VB.Menu Order40_4 
            Caption         =   "รายงานมัดจำจ่าย เจ้าหนี้"
         End
         Begin VB.Menu Order40_6 
            Caption         =   "รายงาน ยอดเจ้าหนี้ ณ วันที่"
         End
      End
   End
   Begin VB.Menu Order5 
      Caption         =   "ลูกหนี้"
      Begin VB.Menu Order5_1 
         Caption         =   "พิมพ์ใบวางบิล"
      End
      Begin VB.Menu Order5_2 
         Caption         =   "พิมพ์ใบเสร็จรับเงิน"
      End
      Begin VB.Menu Order54_6 
         Caption         =   "พิมพ์จดหมายทวงหนี้"
      End
      Begin VB.Menu Order54_7 
         Caption         =   "พิมพ์ใบสำคัญตั้งลูกหนี้อื่น ๆ"
      End
      Begin VB.Menu Order55 
         Caption         =   "เพิ่มทะเบียนสถานที่ขนส่ง"
      End
      Begin VB.Menu Order55_1 
         Caption         =   "พิมพ์รายการหนี้ค้างชำระตามลูกค้า"
      End
      Begin VB.Menu Order5_6 
         Caption         =   "พิมพ์เอกสารตรวจสอบข้อมูลลูกค้าและระเบียนลูกค้า"
      End
      Begin VB.Menu Order5_7 
         Caption         =   "เพิ่มทะเบียนกลุ่มวงเงินลูกค้า"
      End
      Begin VB.Menu Order5_3 
         Caption         =   "-"
      End
      Begin VB.Menu Order5_4 
         Caption         =   "รายงาน"
         Begin VB.Menu Order54_1 
            Caption         =   "รายงานรับชำระหนี้ประจำวัน"
            Begin VB.Menu Order541_1 
               Caption         =   "รายงานรับชำระหนี้ประจำวัน"
            End
            Begin VB.Menu Order541_2 
               Caption         =   "รายงานยอดลูกหนี้ประจำวันแยกตามกลุ่มลูกค้า"
            End
         End
         Begin VB.Menu Order54_2 
            Caption         =   "รายงานยอดลูกหนี้ประจำเดือน"
            Begin VB.Menu Order542_2 
               Caption         =   "รายงานยอดลูกหนี้ประจำเดือนตามประเภทลูกหนี้"
            End
            Begin VB.Menu Order542_3 
               Caption         =   "รายงานยอดลูกหนี้ประจำเดือนตามกลุ่มลูกหนี้_สินเชื่อ"
            End
         End
         Begin VB.Menu Order54_3 
            Caption         =   "รายงานยอดเคลื่อนไหวลูกหนี้ เพื่อดูยอดคงค้าง"
         End
         Begin VB.Menu Order54_4 
            Caption         =   "รายงานเคลื่อนไหวลูกหนี้ ทั้งหมด"
         End
         Begin VB.Menu Order54_8 
            Caption         =   "รายงาน เคลื่อนไหวลูกหนี้ ตามช่วงวันที่"
         End
         Begin VB.Menu Order54_5 
            Caption         =   "รายงาน เช็คตามลูกค้า"
         End
         Begin VB.Menu Order54_9 
            Caption         =   "พิมพ์ Label จ่าหน้าซองจดหมาย"
         End
         Begin VB.Menu Order51_1 
            Caption         =   "พิมพ์รายงาน ใบรับวางบิลของพนักงานเก็บเงิน ตามช่วงวันที่"
         End
      End
   End
   Begin VB.Menu Order6 
      Caption         =   "เช็ค/บัตร"
      Begin VB.Menu Order6_0 
         Caption         =   "พิมพ์เช็ค"
      End
      Begin VB.Menu Order6_1 
         Caption         =   "พิมพ์ตั๋วแลกเงิน"
      End
      Begin VB.Menu Order6_2 
         Caption         =   "พิมพ์เอกสารยกเลิกเช็ครับ"
      End
      Begin VB.Menu Order6_3 
         Caption         =   "พิมพ์เอกสารแลกเปลี่ยนเช็ค"
      End
      Begin VB.Menu Order63 
         Caption         =   "รายงาน"
         Begin VB.Menu Order631 
            Caption         =   "รายงาน ประวัติเช็คคืน"
         End
      End
   End
   Begin VB.Menu Order7 
      Caption         =   "ธนาคาร/เงินสด"
      Begin VB.Menu Order7_0 
         Caption         =   "พิมพ์ใบนำฝาก"
      End
      Begin VB.Menu Order71 
         Caption         =   "พิมพ์ใบนำฝากเงินสด"
      End
      Begin VB.Menu Order072 
         Caption         =   "พิมพ์เช็คและใบนำฝากจากการโอนเงินระหว่างธนาคาร"
      End
   End
   Begin VB.Menu Order8 
      Caption         =   "สินค้าคงคลัง"
      Begin VB.Menu Order87 
         Caption         =   "บันทึกข้อมูลเวลาการจัดสินค้าใบหยิบ"
      End
      Begin VB.Menu Order871 
         Caption         =   "บันทึกเวลาจัดสินค้าในโฮมมาร์ท"
      End
      Begin VB.Menu Order814 
         Caption         =   "พิมพ์ใบขอเบิกสินค้า/วัตถุดิบ"
      End
      Begin VB.Menu Order86 
         Caption         =   "พิมพ์ใบขอโอนสินค้า"
      End
      Begin VB.Menu Order89 
         Caption         =   "คำนวณจำนวนสินค้าในใบขอโอนสินค้า"
      End
      Begin VB.Menu Order810 
         Caption         =   "ยกเลิกสินค้าในใบขอโอน"
      End
      Begin VB.Menu Order83 
         Caption         =   "พิมพ์ใบบันทึกโอนสินค้าระหว่างคลัง"
      End
      Begin VB.Menu Order84 
         Caption         =   "พิมพ์ใบบันทึกเบิกใช้สินค้า-วัตถุดิบ"
      End
      Begin VB.Menu OrderItemIssue 
         Caption         =   "รายงาน เอกสารขอเบิกสินค้าประจำวัน"
      End
      Begin VB.Menu Order85 
         Caption         =   "พิมพ์ใบบันทึกปรุบปรุงหลังการตรวจนับ"
      End
      Begin VB.Menu Order81 
         Caption         =   "โปรแกรม รวมรหัสสินค้า"
      End
      Begin VB.Menu Order82 
         Caption         =   "รายงาน สินค้าขายดี"
      End
      Begin VB.Menu OrderChangeItemandPrice 
         Caption         =   "รายงาน การเปลี่ยนรหัสและราคาสินค้า"
      End
      Begin VB.Menu StockCardBC 
         Caption         =   "StockCard BCAccount"
      End
      Begin VB.Menu Order88 
         Caption         =   "StockCard GP"
      End
      Begin VB.Menu OrderPrint 
         Caption         =   "รายงาน เกี่ยวกับการพิมพ์ใบจ่ายใบหยิบ"
      End
      Begin VB.Menu TransferPrint 
         Caption         =   "รายงาน ใบขอโอนไม่ได้ทำใบโอนสินค้า"
      End
      Begin VB.Menu OrderStockCountHMX 
         Caption         =   "รายงาน ข้อมูลตรวจนับสินค้า HMX ประจำวัน"
      End
   End
   Begin VB.Menu Order9 
      Caption         =   "บัญชี"
      Begin VB.Menu Order91 
         Caption         =   "ยกเลิกการผ่านบัญชี BCAccount"
      End
      Begin VB.Menu Order92 
         Caption         =   "รายงานสรุปค่าใช้จ่ายประจำปี"
      End
      Begin VB.Menu Order93 
         Caption         =   "รายงานงบทดลอง"
      End
      Begin VB.Menu Order94 
         Caption         =   "รายงาน แยกประเภท"
      End
      Begin VB.Menu Order95 
         Caption         =   "รายงาน สมุดรายวัน"
      End
      Begin VB.Menu Order96 
         Caption         =   "รายงานจัดลำดับ"
      End
      Begin VB.Menu Order99 
         Caption         =   "รายงาน การจ่ายเงินประจำวัน"
      End
      Begin VB.Menu OrderTaxPurchase 
         Caption         =   "บันทึกภาษีซื้อ"
      End
      Begin VB.Menu Order97 
         Caption         =   "เปลี่ยนเลขที่เอกสาร/เลขที่ภาษี"
      End
      Begin VB.Menu Order98 
         Caption         =   "ตรวจสอบความครบถ้วนของเอกสาร"
      End
      Begin VB.Menu Order991 
         Caption         =   "แก้ไขการอนุมัติเอกสารต่าง ๆ"
      End
      Begin VB.Menu OrderPOSItem 
         Caption         =   "เปิดติดลบสินค้าขาย POS"
      End
   End
   Begin VB.Menu Order0 
      Caption         =   "ช่วยเหลือ"
      Begin VB.Menu Order0_1 
         Caption         =   "พิมพ์ทดแทน"
      End
      Begin VB.Menu Order0_3 
         Caption         =   "พิมพ์ใบหยิบสินค้าทดแทน"
      End
      Begin VB.Menu Order0_2 
         Caption         =   "พิมพ์เอกสารขายสำคัญ"
      End
      Begin VB.Menu Order310 
         Caption         =   "พิมพ์ทดแทนใบจ่ายสินค้า (ใบเหลือง)"
      End
   End
   Begin VB.Menu nWindows 
      Caption         =   "หน้าต่างที่เปิดไว้แล้ว"
      WindowList      =   -1  'True
   End
End
Attribute VB_Name = "MDIFrmProgramPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub CheckOut_Click()
On Error GoTo ErrDescription

MsgBox ("ยกเลิกการใช้งานชั่วคราว")

'vFormActivate = "Form311"
'Call ChekAuthorityAccess
'If vAccess = 1 Then
'FormCheckOutHoldBill.Show
'FormCheckOutHoldBill.SetFocus
'Else
 '   MsgBox "User ID : " & vUserID & " ไม่สามารถเข้าใช้หน้าจอนี้ได้ เนื่องจากไม่มีสิทธิ์การใช้งาน ", vbInformation, "ข้อความแจ้งเตือน"
  '  Exit Sub
'End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub DO102_Click()
FrmOrder201.Show
FrmOrder201.SetFocus
End Sub

Private Sub DO103_Click()
FrmOrder202.Show
FrmOrder202.SetFocus
End Sub

Private Sub DO104_Click()
FrmOrder203.Show
FrmOrder203.SetFocus
End Sub

Private Sub DO105_Click()
FrmOrder204.Show
FrmOrder204.SetFocus
End Sub

Private Sub DO106_Click()
FrmOrder205.Show
FrmOrder205.SetFocus
End Sub

Private Sub DO107_Click()
FrmOrder206.Show
FrmOrder206.SetFocus
End Sub

Private Sub DO109_Click()
FormDelivery.Show
FormDelivery.SetFocus
End Sub

Private Sub DO110_Click()
FrmOrder401.Show
FrmOrder401.SetFocus
End Sub

Private Sub DO111_Click()
FrmOrder402.Show
FrmOrder402.Show
End Sub

Private Sub DO113_Click()
FrmOrder403.Show
FrmOrder403.SetFocus
End Sub

Private Sub DO114_Click()

On Error GoTo ErrDescription

FormQueueApprove.Show
FormQueueApprove.SetFocus

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

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
DO1.Enabled = False
nWindows.Enabled = False
FrmLogIN.Show
vChkFrmActivate = 0
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
Dim vQuestion As Integer
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String

vQuestion = MsgBox("คุณต้องการออกจากโปรแกรมใช่หรือไม่", vbYesNo + vbCritical, "ข้อความสอบถาม")
If vQuestion = 6 Then
            If gConnection.State = 1 Then
            gConnection.Close
            End If
            If vConnection.State = 1 Then
                vConnection.Close
            End If
            If vClosePR = 0 Then
            If vCheckDuplicate <> 1 Then
                On Error Resume Next
                vQuery = "delete npmaster.dbo.TB_CK_UserActivateProgram  where userid = '" & vUserID & "' and jobid = 1"
                gConnection.Execute (vQuery)
            End If
            End If
Else
Cancel = True
End If

End Sub

Private Sub Order0_1_Click()
On Error GoTo ErrDescription

vFormActivate = "Form01"
Call ChekAuthorityAccess
If vAccess = 1 Then
Form01.Show
Form01.SetFocus
Else
    MsgBox "User ID : " & vUserID & " ไม่สามารถเข้าใช้หน้าจอนี้ได้ เนื่องจากไม่มีสิทธิ์การใช้งาน ", vbInformation, "ข้อความแจ้งเตือน"
    Exit Sub
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Order0_2_Click()
On Error GoTo ErrDescription

vFormActivate = "Form02"
Call ChekAuthorityAccess
If vAccess = 1 Then
Form02.Show
Form02.SetFocus
Else
    MsgBox "User ID : " & vUserID & " ไม่สามารถเข้าใช้หน้าจอนี้ได้ เนื่องจากไม่มีสิทธิ์การใช้งาน ", vbInformation, "ข้อความแจ้งเตือน"
    Exit Sub
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Order0_3_Click()
On Error GoTo ErrDescription

'vFormActivate = "Form03"
'Call ChekAuthorityAccess
'If vAccess = 1 Then
'Form03.Show
'Form03.SetFocus
'Else
 '   MsgBox "User ID : " & vUserID & " ไม่สามารถเข้าใช้หน้าจอนี้ได้ เนื่องจากไม่มีสิทธิ์การใช้งาน ", vbInformation, "ข้อความแจ้งเตือน"
  '  Exit Sub
'End If

MsgBox "การทดแทน พิมพ์ใบหยิบสามารถทำได้ที่ หน้าพิมพ์ใบสั่งขายสั่งจองได้เลย ตรงปุ่มจัดสินค้า", vbInformation, "Send Information Message"

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Order072_Click()
On Error GoTo ErrDescription

vFormActivate = "Form72"
Call ChekAuthorityAccess
If vAccess = 1 Then
Form72.Show
Form72.SetFocus
Else
    MsgBox "User ID : " & vUserID & " ไม่สามารถเข้าใช้หน้าจอนี้ได้ เนื่องจากไม่มีสิทธิ์การใช้งาน ", vbInformation, "ข้อความแจ้งเตือน"
    Exit Sub
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Order1_0_Click()
Order2.Enabled = False
Order3.Enabled = False
Order4.Enabled = False
Order5.Enabled = False
Order6.Enabled = False
Order7.Enabled = False
Order8.Enabled = False
Order9.Enabled = False
Order0.Enabled = False
DO1.Enabled = False
nWindows.Enabled = False
MDIFrmProgramPrint.Caption = "โปรแกรมพิมพ์เอกสาร BCAccount 5.5"
FrmLogIN.Show
Unload Form01
Unload Form02
Unload Form2_51
Unload Form2_52
Unload Form2_6
Unload Form2_7
Unload Form20
Unload Form21
Unload Form22
Unload Form3_14
Unload Form3_15
Unload Form3_81
Unload Form3_82
Unload Form30
Unload Form31
Unload Form310
Unload Form32
Unload Form33
Unload Form34
Unload Form35
Unload Form36
Unload Form39
Unload Form392
Unload Form393
Unload Form395
Unload Form40_1
Unload Form40_2
Unload Form40_3
Unload Form40_4
Unload Form40_5
Unload Form41
Unload Form411
Unload Form42
Unload Form5_1
Unload Form5_2
Unload Form54_3
Unload Form54_4
Unload Form54_5
Unload Form54_6
Unload Form54_7
Unload Form54_8
Unload Form54_9
Unload Form541_1
Unload Form542_1
Unload Form542_2
Unload Form6_0
Unload Form6_1
Unload Form631
Unload Form7_0
Unload Form71
Unload Form72
Unload Form81
Unload Form810
Unload Form811
Unload Form812
Unload Form813
Unload Form82
Unload Form86
Unload Form87
Unload Form871
Unload Form88
Unload Form89
Unload Form91
Unload Form92
Unload Form93
Unload Form936
Unload Form94
Unload Form95
Unload Form96
Unload Form97
Unload Form98
Unload FormPrintReserve
Unload FrmOrder006
Unload Form312
Unload FormDelivery
Unload FrmOrder010
Unload FrmOrder013
Unload FrmOrder201
Unload FrmOrder202
Unload FrmOrder203
Unload FrmOrder204
Unload FrmOrder205
Unload FrmOrder206

End Sub

Private Sub Order1_1_Click()
Dim vQuestion As Integer
On Error GoTo ErrDescription

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

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Order2_0_Click()
On Error GoTo ErrDescription

vFormActivate = "Form20"
Call ChekAuthorityAccess
If vAccess = 1 Then
    Form20.Show
    Form20.SetFocus
Else
    MsgBox "User ID : " & vUserID & " ไม่สามารถเข้าใช้หน้าจอนี้ได้ เนื่องจากไม่มีสิทธิ์การใช้งาน ", vbInformation, "ข้อความแจ้งเตือน"
    Exit Sub
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Order2_1_Click()
On Error GoTo ErrDescription

vFormActivate = "Form21"
Call ChekAuthorityAccess
If vAccess = 1 Then
    Form21.Show
    Form21.SetFocus
Else
    MsgBox "User ID : " & vUserID & " ไม่สามารถเข้าใช้หน้าจอนี้ได้ เนื่องจากไม่มีสิทธิ์การใช้งาน ", vbInformation, "ข้อความแจ้งเตือน"
    Exit Sub
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Order2_2_Click()
On Error GoTo ErrDescription

vFormActivate = "Form22"
Call ChekAuthorityAccess
If vAccess = 1 Then
Form22.Show
Form22.SetFocus
Else
    MsgBox "User ID : " & vUserID & " ไม่สามารถเข้าใช้หน้าจอนี้ได้ เนื่องจากไม่มีสิทธิ์การใช้งาน ", vbInformation, "ข้อความแจ้งเตือน"
    Exit Sub
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Order2_51_Click()
'Form2_51.Show
'Form2_51.SetFocus
End Sub

Private Sub Order2_4_Click()
'On Error GoTo ErrDescription

vFormActivate = "Form2_4"
'Call ChekAuthorityAccess
'If vAccess = 1 Then
    Form2_4.Show
    Form2_4.SetFocus
'Else
 '   MsgBox "User ID : " & vUserID & " ไม่สามารถเข้าใช้หน้าจอนี้ได้ เนื่องจากไม่มีสิทธิ์การใช้งาน ", vbInformation, "ข้อความแจ้งเตือน"
  '  Exit Sub
'End If

'ErrDescription:
'if Err.Description <> "" Then
'MsgBox Err.Description
'Exit Sub
'End If
End Sub

Private Sub Order2_5_Click()
'On Error GoTo ErrDescription

'vFormActivate = "Form2_5"
'Call ChekAuthorityAccess
'If vAccess = 1 Then
    Form2_5.Show
    Form2_5.SetFocus
'Else
 '   MsgBox "User ID : " & vUserID & " ไม่สามารถเข้าใช้หน้าจอนี้ได้ เนื่องจากไม่มีสิทธิ์การใช้งาน ", vbInformation, "ข้อความแจ้งเตือน"
  '  Exit Sub
'End If

'ErrDescription:
'if Err.Description <> "" Then
'MsgBox Err.Description
'Exit Sub
'End If
End Sub

Private Sub Order2_52_Click()
On Error GoTo ErrDescription

vFormActivate = "Form2_52"
Call ChekAuthorityAccess
If vAccess = 1 Then
Form2_52.Show
Form2_52.SetFocus
Else
    MsgBox "User ID : " & vUserID & " ไม่สามารถเข้าใช้หน้าจอนี้ได้ เนื่องจากไม่มีสิทธิ์การใช้งาน ", vbInformation, "ข้อความแจ้งเตือน"
    Exit Sub
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Order2_6_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vCheckUserID As Integer

On Error GoTo ErrDescription

vFormActivate = "Form2_6"
Call ChekAuthorityAccess
If vAccess = 1 Then
    If vChkFrmActivate = 0 Then
    vQuery = "select userid from npmaster.dbo.TB_CK_UserActivateProgram where userid = '" & vUserID & "' and jobid = 1"
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vCheckUserID = 1
        vClosePR = 1
        MsgBox "มี UserID : " & vUserID & " เข้ามาใช้งานในหน้านี้แล้ว กรุณาตรวจสอบ ไม่งั้นจะไม่สามารถทำการรวม PR ได้ กรุณาติดต่อคอมฯนะครับ"
    Else
        vCheckUserID = 0
        vClosePR = 0
    End If
    vRecordset.Close
    If vCheckUserID = 0 Then
        Form2_6.Show
        Form2_6.SetFocus
        vQuery = "insert into npmaster.dbo.TB_CK_UserActivateProgram (UserID,ActiveDateTime,ActiveStatus,JobId) " _
                            & " values ('" & vUserID & "',getdate(),1,1)"
        gConnection.Execute (vQuery)
        vChkFrmActivate = 1
    End If
    Else
        Form2_6.SetFocus
    End If
Else
    MsgBox "User ID : " & vUserID & " ไม่สามารถเข้าใช้หน้าจอนี้ได้ เนื่องจากไม่มีสิทธิ์การใช้งาน ", vbInformation, "ข้อความแจ้งเตือน"
    Exit Sub
End If

ErrDescription:
If Err.Description <> "" Then
    If Err.Number = -2147217873 Then
        vCheckDuplicate = 1
        MsgBox "ผู้ใช้งาน  " & vUserID & "  ได้เข้าใช้งานหน้านี้อยู่แล้ว ไม่สามารถเข้าใช้งานซ้ำกันได้"
        Exit Sub
    Else
        MsgBox Err.Description
        Exit Sub
    End If
End If
End Sub

Private Sub Order2_7_Click()
On Error GoTo ErrDescription

vFormActivate = "Form2_7"
Call ChekAuthorityAccess
If vAccess = 1 Then
Form2_7.Show
Form2_7.SetFocus
Else
    MsgBox "User ID : " & vUserID & " ไม่สามารถเข้าใช้หน้าจอนี้ได้ เนื่องจากไม่มีสิทธิ์การใช้งาน ", vbInformation, "ข้อความแจ้งเตือน"
    Exit Sub
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Order2_8_Click()
Call ChekAuthorityAccess
Form2_8.Show
Form2_8.SetFocus
End Sub

Private Sub Order2_9_Click()
On Error GoTo ErrDescription

'vFormActivate = "Form6_2"
'Call ChekAuthorityAccess
'If vAccess = 1 Then
Form2_9.Show
Form2_9.SetFocus
'Else
 '   MsgBox "User ID : " & vUserID & " ไม่สามารถเข้าใช้หน้าจอนี้ได้ เนื่องจากไม่มีสิทธิ์การใช้งาน ", vbInformation, "ข้อความแจ้งเตือน"
  '  Exit Sub
'End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Order3_0_Click()
On Error GoTo ErrDescription

vFormActivate = "Form30"
Call ChekAuthorityAccess
If vAccess = 1 Then
Form30.Show
Form30.SetFocus
Else
    MsgBox "User ID : " & vUserID & " ไม่สามารถเข้าใช้หน้าจอนี้ได้ เนื่องจากไม่มีสิทธิ์การใช้งาน ", vbInformation, "ข้อความแจ้งเตือน"
    Exit Sub
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Order3_1_Click()
On Error GoTo ErrDescription

vFormActivate = "Form31"
Call ChekAuthorityAccess
If vAccess = 1 Then
Form31.Show
Form31.SetFocus
Else
    MsgBox "User ID : " & vUserID & " ไม่สามารถเข้าใช้หน้าจอนี้ได้ เนื่องจากไม่มีสิทธิ์การใช้งาน ", vbInformation, "ข้อความแจ้งเตือน"
    Exit Sub
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Order3_14_Click()
On Error GoTo ErrDescription

vFormActivate = "Form3_14"
Call ChekAuthorityAccess
If vAccess = 1 Then
Form3_14.Show
Form3_14.SetFocus
Else
    MsgBox "User ID : " & vUserID & " ไม่สามารถเข้าใช้หน้าจอนี้ได้ เนื่องจากไม่มีสิทธิ์การใช้งาน ", vbInformation, "ข้อความแจ้งเตือน"
    Exit Sub
End If
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Order3_15_Click()
On Error GoTo ErrDescription

vFormActivate = "Form3_15"
Call ChekAuthorityAccess
If vAccess = 1 Then
Form3_15.Show
Form3_15.SetFocus
Else
    MsgBox "User ID : " & vUserID & " ไม่สามารถเข้าใช้หน้าจอนี้ได้ เนื่องจากไม่มีสิทธิ์การใช้งาน ", vbInformation, "ข้อความแจ้งเตือน"
    Exit Sub
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Order3_16_Click()
'Form3_16.Show
'Form3_16.SetFocus
MsgBox "ยังไม่เปิดให้ใช้งาน", vbCritical, "Send Information Message"
End Sub

Private Sub Order3_17_Click()
Form3_17.Show
Form3_17.SetFocus
End Sub

Private Sub Order3_18_Click()
Form3_18.Show
Form3_18.SetFocus
End Sub

Private Sub Order3_19_Click()
On Error GoTo ErrDescription

MsgBox "ยังไม่เปิดให้ใช้งาน กำลังปรับปรุง", vbCritical, "Send Error Message"
'vFormActivate = "Form3_19"
'Call ChekAuthorityAccess
'If vAccess = 1 Then
'Form3_19.Show
'Form3_19.SetFocus
'Else
 '   MsgBox "User ID : " & vUserID & " ไม่สามารถเข้าใช้หน้าจอนี้ได้ เนื่องจากไม่มีสิทธิ์การใช้งาน ", vbInformation, "ข้อความแจ้งเตือน"
  '  Exit Sub
'End If
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Order3_2_Click()
On Error GoTo ErrDescription

vFormActivate = "Form311"
Call ChekAuthorityAccess
If vAccess = 1 Then
Form311.Show
Form311.SetFocus
Else
    MsgBox "User ID : " & vUserID & " ไม่สามารถเข้าใช้หน้าจอนี้ได้ เนื่องจากไม่มีสิทธิ์การใช้งาน ", vbInformation, "ข้อความแจ้งเตือน"
    Exit Sub
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Order3_20_Click()
'On Error GoTo ErrDescription

vFormActivate = "Form3_20"
'Call ChekAuthorityAccess
'If vAccess = 1 Then
Form3_20.Show
Form3_20.SetFocus
'Else
 '   MsgBox "User ID : " & vUserID & " ไม่สามารถเข้าใช้หน้าจอนี้ได้ เนื่องจากไม่มีสิทธิ์การใช้งาน ", vbInformation, "ข้อความแจ้งเตือน"
  '  Exit Sub
'End If

'ErrDescription:
'If Err.Description <> "" Then
'MsgBox Err.Description
'Exit Sub
'End If
End Sub

Private Sub Order3_3_Click()
On Error GoTo ErrDescription

vFormActivate = "Form33"
Call ChekAuthorityAccess
If vAccess = 1 Then
Form33.Show
Form33.SetFocus
Else
    MsgBox "User ID : " & vUserID & " ไม่สามารถเข้าใช้หน้าจอนี้ได้ เนื่องจากไม่มีสิทธิ์การใช้งาน ", vbInformation, "ข้อความแจ้งเตือน"
    Exit Sub
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Order3_4_Click()
On Error GoTo ErrDescription

vFormActivate = "Form34"
Call ChekAuthorityAccess
If vAccess = 1 Then
Form34.Show
Form34.SetFocus
Else
    MsgBox "User ID : " & vUserID & " ไม่สามารถเข้าใช้หน้าจอนี้ได้ เนื่องจากไม่มีสิทธิ์การใช้งาน ", vbInformation, "ข้อความแจ้งเตือน"
    Exit Sub
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Order3_5_Click()
On Error GoTo ErrDescription

vFormActivate = "Form35"
Call ChekAuthorityAccess
If vAccess = 1 Then
Form35.Show
Form35.SetFocus
Else
    MsgBox "User ID : " & vUserID & " ไม่สามารถเข้าใช้หน้าจอนี้ได้ เนื่องจากไม่มีสิทธิ์การใช้งาน ", vbInformation, "ข้อความแจ้งเตือน"
    Exit Sub
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Order3_6_Click()
On Error GoTo ErrDescription

vFormActivate = "Form36"
Call ChekAuthorityAccess
If vAccess = 1 Then
Form36.Show
Form36.SetFocus
Else
    MsgBox "User ID : " & vUserID & " ไม่สามารถเข้าใช้หน้าจอนี้ได้ เนื่องจากไม่มีสิทธิ์การใช้งาน ", vbInformation, "ข้อความแจ้งเตือน"
    Exit Sub
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Order3_81_Click()
On Error GoTo ErrDescription

vFormActivate = "Form3_81"
Call ChekAuthorityAccess
If vAccess = 1 Then
Form3_81.Show
Form3_81.SetFocus
Else
    MsgBox "User ID : " & vUserID & " ไม่สามารถเข้าใช้หน้าจอนี้ได้ เนื่องจากไม่มีสิทธิ์การใช้งาน ", vbInformation, "ข้อความแจ้งเตือน"
    Exit Sub
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Order3_82_Click()
On Error GoTo ErrDescription

vFormActivate = "Form3_82"
Call ChekAuthorityAccess
If vAccess = 1 Then
Form3_82.Show
Form3_82.SetFocus
Else
    MsgBox "User ID : " & vUserID & " ไม่สามารถเข้าใช้หน้าจอนี้ได้ เนื่องจากไม่มีสิทธิ์การใช้งาน ", vbInformation, "ข้อความแจ้งเตือน"
    Exit Sub
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Order310_Click()
On Error GoTo ErrDescription

vFormActivate = "Form310"
Call ChekAuthorityAccess
If vAccess = 1 Then
Form310.Show
Form310.SetFocus
Else
    MsgBox "User ID : " & vUserID & " ไม่สามารถเข้าใช้หน้าจอนี้ได้ เนื่องจากไม่มีสิทธิ์การใช้งาน ", vbInformation, "ข้อความแจ้งเตือน"
    Exit Sub
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
'Form310.Show
'Form310.SetFocus
End Sub

Private Sub Order311_Click()

End Sub

Private Sub Order37_Click()
On Error GoTo ErrDescription

vFormActivate = "Form39"
Call ChekAuthorityAccess
If vAccess = 1 Then
Form39.Show
Form39.SetFocus
Else
    MsgBox "User ID : " & vUserID & " ไม่สามารถเข้าใช้หน้าจอนี้ได้ เนื่องจากไม่มีสิทธิ์การใช้งาน ", vbInformation, "ข้อความแจ้งเตือน"
    Exit Sub
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Order392_Click()
On Error GoTo ErrDescription

vFormActivate = "Form392"
Call ChekAuthorityAccess
If vAccess = 1 Then
Form392.Show
Form392.SetFocus
Else
    MsgBox "User ID : " & vUserID & " ไม่สามารถเข้าใช้หน้าจอนี้ได้ เนื่องจากไม่มีสิทธิ์การใช้งาน ", vbInformation, "ข้อความแจ้งเตือน"
    Exit Sub
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Order393_Click()
On Error GoTo ErrDescription

vFormActivate = "Form32"
Call ChekAuthorityAccess
If vAccess = 1 Then
Form32.Show
Form32.SetFocus
Else
    MsgBox "User ID : " & vUserID & " ไม่สามารถเข้าใช้หน้าจอนี้ได้ เนื่องจากไม่มีสิทธิ์การใช้งาน ", vbInformation, "ข้อความแจ้งเตือน"
    Exit Sub
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Order394_Click()
On Error GoTo ErrDescription

vFormActivate = "Form393"
Call ChekAuthorityAccess
If vAccess = 1 Then
Form393.Show
Form393.SetFocus
Else
    MsgBox "User ID : " & vUserID & " ไม่สามารถเข้าใช้หน้าจอนี้ได้ เนื่องจากไม่มีสิทธิ์การใช้งาน ", vbInformation, "ข้อความแจ้งเตือน"
    Exit Sub
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Order395_Click()
On Error GoTo ErrDescription

vFormActivate = "Form395"
Call ChekAuthorityAccess
If vAccess = 1 Then
Form395.Show
Form395.SetFocus
Else
    MsgBox "User ID : " & vUserID & " ไม่สามารถเข้าใช้หน้าจอนี้ได้ เนื่องจากไม่มีสิทธิ์การใช้งาน ", vbInformation, "ข้อความแจ้งเตือน"
    Exit Sub
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Order396_Click()
On Error GoTo ErrDescription

vFormActivate = "Form936"
Call ChekAuthorityAccess
If vAccess = 1 Then
Form936.Show
Form936.SetFocus
Else
    MsgBox "User ID : " & vUserID & " ไม่สามารถเข้าใช้หน้าจอนี้ได้ เนื่องจากไม่มีสิทธิ์การใช้งาน ", vbInformation, "ข้อความแจ้งเตือน"
    Exit Sub
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Order40_1_Click()
On Error GoTo ErrDescription

vFormActivate = "Form40_1"
Call ChekAuthorityAccess
If vAccess = 1 Then
Form40_1.Show
Form40_1.SetFocus
Else
    MsgBox "User ID : " & vUserID & " ไม่สามารถเข้าใช้หน้าจอนี้ได้ เนื่องจากไม่มีสิทธิ์การใช้งาน ", vbInformation, "ข้อความแจ้งเตือน"
    Exit Sub
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Order40_2_Click()
On Error GoTo ErrDescription

vFormActivate = "Form40_2"
Call ChekAuthorityAccess
If vAccess = 1 Then
Form40_2.Show
Form40_2.SetFocus
Else
    MsgBox "User ID : " & vUserID & " ไม่สามารถเข้าใช้หน้าจอนี้ได้ เนื่องจากไม่มีสิทธิ์การใช้งาน ", vbInformation, "ข้อความแจ้งเตือน"
    Exit Sub
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Order40_3_Click()
On Error GoTo ErrDescription

vFormActivate = "Form40_3"
Call ChekAuthorityAccess
If vAccess = 1 Then
Form40_3.Show
Form40_3.SetFocus
Else
    MsgBox "User ID : " & vUserID & " ไม่สามารถเข้าใช้หน้าจอนี้ได้ เนื่องจากไม่มีสิทธิ์การใช้งาน ", vbInformation, "ข้อความแจ้งเตือน"
    Exit Sub
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Order40_4_Click()
On Error GoTo ErrDescription

vFormActivate = "Form40_4"
Call ChekAuthorityAccess
If vAccess = 1 Then
Form40_4.Show
Form40_4.SetFocus
Else
    MsgBox "User ID : " & vUserID & " ไม่สามารถเข้าใช้หน้าจอนี้ได้ เนื่องจากไม่มีสิทธิ์การใช้งาน ", vbInformation, "ข้อความแจ้งเตือน"
    Exit Sub
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Order40_5_Click()
On Error GoTo ErrDescription

vFormActivate = "Form40_5"
Call ChekAuthorityAccess
If vAccess = 1 Then
Form40_5.Show
Form40_5.SetFocus
Else
    MsgBox "User ID : " & vUserID & " ไม่สามารถเข้าใช้หน้าจอนี้ได้ เนื่องจากไม่มีสิทธิ์การใช้งาน ", vbInformation, "ข้อความแจ้งเตือน"
    Exit Sub
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Order40_6_Click()
On Error GoTo ErrDescription

vFormActivate = "Form40_6"
Call ChekAuthorityAccess
If vAccess = 1 Then
    Form40_6.Show
    Form40_6.SetFocus
Else
    MsgBox "User ID : " & vUserID & " ไม่สามารถเข้าใช้หน้าจอนี้ได้ เนื่องจากไม่มีสิทธิ์การใช้งาน ", vbInformation, "ข้อความแจ้งเตือน"
    Exit Sub
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Order41_Click()
On Error GoTo ErrDescription

vFormActivate = "Form41"
Call ChekAuthorityAccess
If vAccess = 1 Then
Form41.Show
Form41.SetFocus
Else
    MsgBox "User ID : " & vUserID & " ไม่สามารถเข้าใช้หน้าจอนี้ได้ เนื่องจากไม่มีสิทธิ์การใช้งาน ", vbInformation, "ข้อความแจ้งเตือน"
    Exit Sub
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Order42_Click()
On Error GoTo ErrDescription

vFormActivate = "Form42"
Call ChekAuthorityAccess
If vAccess = 1 Then
Form42.Show
Form42.SetFocus
Else
    MsgBox "User ID : " & vUserID & " ไม่สามารถเข้าใช้หน้าจอนี้ได้ เนื่องจากไม่มีสิทธิ์การใช้งาน ", vbInformation, "ข้อความแจ้งเตือน"
    Exit Sub
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Order5_1_Click()
On Error GoTo ErrDescription

vFormActivate = "Form5_1"
Call ChekAuthorityAccess
If vAccess = 1 Then
Form5_1.Show
Form5_1.SetFocus
Else
    MsgBox "User ID : " & vUserID & " ไม่สามารถเข้าใช้หน้าจอนี้ได้ เนื่องจากไม่มีสิทธิ์การใช้งาน ", vbInformation, "ข้อความแจ้งเตือน"
    Exit Sub
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Order5_2_Click()
On Error GoTo ErrDescription

vFormActivate = "Form5_2"
Call ChekAuthorityAccess
If vAccess = 1 Then
Form5_2.Show
Form5_2.SetFocus
Else
    MsgBox "User ID : " & vUserID & " ไม่สามารถเข้าใช้หน้าจอนี้ได้ เนื่องจากไม่มีสิทธิ์การใช้งาน ", vbInformation, "ข้อความแจ้งเตือน"
    Exit Sub
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Order5_6_Click()
On Error GoTo ErrDescription

vFormActivate = "Form5_6"
'Call ChekAuthorityAccess
'If vAccess = 1 Then
Form5_6.Show
Form5_6.SetFocus
'Else
 '   MsgBox "User ID : " & vUserID & " ไม่สามารถเข้าใช้หน้าจอนี้ได้ เนื่องจากไม่มีสิทธิ์การใช้งาน ", vbInformation, "ข้อความแจ้งเตือน"
  '  Exit Sub
'End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Order5_7_Click()
On Error GoTo ErrDescription

vFormActivate = "Form5_7"
Call ChekAuthorityAccess
If vAccess = 1 Then
Form5_7.Show
Form5_7.SetFocus
Else
    MsgBox "User ID : " & vUserID & " ไม่สามารถเข้าใช้หน้าจอนี้ได้ เนื่องจากไม่มีสิทธิ์การใช้งาน ", vbInformation, "ข้อความแจ้งเตือน"
    Exit Sub
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Order51_1_Click()
On Error GoTo ErrDescription

'vFormActivate = "Form51_1"
'Call ChekAuthorityAccess
'If vAccess = 1 Then
Form51_1.Show
Form51_1.SetFocus
'Else
 '   MsgBox "User ID : " & vUserID & " ไม่สามารถเข้าใช้หน้าจอนี้ได้ เนื่องจากไม่มีสิทธิ์การใช้งาน ", vbInformation, "ข้อความแจ้งเตือน"
  '  Exit Sub
'End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Order54_3_Click()
On Error GoTo ErrDescription

vFormActivate = "Form54_3"
Call ChekAuthorityAccess
If vAccess = 1 Then
Form54_3.Show
Form54_3.SetFocus
Else
    MsgBox "User ID : " & vUserID & " ไม่สามารถเข้าใช้หน้าจอนี้ได้ เนื่องจากไม่มีสิทธิ์การใช้งาน ", vbInformation, "ข้อความแจ้งเตือน"
    Exit Sub
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Order54_4_Click()
On Error GoTo ErrDescription

vFormActivate = "Form54_4"
Call ChekAuthorityAccess
If vAccess = 1 Then
Form54_4.Show
Form54_4.SetFocus
Else
    MsgBox "User ID : " & vUserID & " ไม่สามารถเข้าใช้หน้าจอนี้ได้ เนื่องจากไม่มีสิทธิ์การใช้งาน ", vbInformation, "ข้อความแจ้งเตือน"
    Exit Sub
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Order54_5_Click()
On Error GoTo ErrDescription

vFormActivate = "Form54_5"
Call ChekAuthorityAccess
If vAccess = 1 Then
Form54_5.Show
Form54_5.SetFocus
Else
    MsgBox "User ID : " & vUserID & " ไม่สามารถเข้าใช้หน้าจอนี้ได้ เนื่องจากไม่มีสิทธิ์การใช้งาน ", vbInformation, "ข้อความแจ้งเตือน"
    Exit Sub
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Order54_6_Click()
On Error GoTo ErrDescription

vFormActivate = "Form54_6"
Call ChekAuthorityAccess
If vAccess = 1 Then
Form54_6.Show
Form54_6.SetFocus
Else
    MsgBox "User ID : " & vUserID & " ไม่สามารถเข้าใช้หน้าจอนี้ได้ เนื่องจากไม่มีสิทธิ์การใช้งาน ", vbInformation, "ข้อความแจ้งเตือน"
    Exit Sub
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Order54_7_Click()
On Error GoTo ErrDescription

vFormActivate = "Form54_7"
Call ChekAuthorityAccess
If vAccess = 1 Then
Form54_7.Show
Form54_7.SetFocus
Else
    MsgBox "User ID : " & vUserID & " ไม่สามารถเข้าใช้หน้าจอนี้ได้ เนื่องจากไม่มีสิทธิ์การใช้งาน ", vbInformation, "ข้อความแจ้งเตือน"
    Exit Sub
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Order54_8_Click()
On Error GoTo ErrDescription

vFormActivate = "Form54_8"
Call ChekAuthorityAccess
If vAccess = 1 Then
Form54_8.Show
Form54_8.SetFocus
Else
    MsgBox "User ID : " & vUserID & " ไม่สามารถเข้าใช้หน้าจอนี้ได้ เนื่องจากไม่มีสิทธิ์การใช้งาน ", vbInformation, "ข้อความแจ้งเตือน"
    Exit Sub
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Order54_9_Click()
Form54_9.Show
Form54_9.SetFocus
End Sub

Private Sub Order541_1_Click()
On Error GoTo ErrDescription

vFormActivate = "Form541_1"
Call ChekAuthorityAccess
If vAccess = 1 Then
Form541_1.Show
Form541_1.SetFocus
Else
    MsgBox "User ID : " & vUserID & " ไม่สามารถเข้าใช้หน้าจอนี้ได้ เนื่องจากไม่มีสิทธิ์การใช้งาน ", vbInformation, "ข้อความแจ้งเตือน"
    Exit Sub
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Order541_2_Click()
'Form541_2.Show
End Sub

'Private Sub Order542_1_Click()
'On Error GoTo ErrDescription

'vFormActivate = "Form542_1"
'Call ChekAuthorityAccess
'If vAccess = 1 Then
'Form542_1.Show
'Form542_1.SetFocus
'Else
 '   MsgBox "User ID : " & vUserID & " ไม่สามารถเข้าใช้หน้าจอนี้ได้ เนื่องจากไม่มีสิทธิ์การใช้งาน ", vbInformation, "ข้อความแจ้งเตือน"
  '  Exit Sub
'End If

'ErrDescription:
'If Err.Description <> "" Then
'MsgBox Err.Description
'Exit Sub
'End If
'End Sub

Private Sub Order542_2_Click()
On Error GoTo ErrDescription

vFormActivate = "Form542_2"
Call ChekAuthorityAccess
If vAccess = 1 Then
Form542_2.Show
Form542_2.SetFocus
Else
    MsgBox "User ID : " & vUserID & " ไม่สามารถเข้าใช้หน้าจอนี้ได้ เนื่องจากไม่มีสิทธิ์การใช้งาน ", vbInformation, "ข้อความแจ้งเตือน"
    Exit Sub
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Order542_3_Click()
On Error GoTo ErrDescription

vFormActivate = "Form542_3"
Call ChekAuthorityAccess
If vAccess = 1 Then
Form542_3.Show
Form542_3.SetFocus
Else
    MsgBox "User ID : " & vUserID & " ไม่สามารถเข้าใช้หน้าจอนี้ได้ เนื่องจากไม่มีสิทธิ์การใช้งาน ", vbInformation, "ข้อความแจ้งเตือน"
    Exit Sub
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Order55_1_Click()
On Error GoTo ErrDescription

vFormActivate = "Form55_1"
Call ChekAuthorityAccess
If vAccess = 1 Then
Form55_1.Show
Form55_1.SetFocus
Else
    MsgBox "User ID : " & vUserID & " ไม่สามารถเข้าใช้หน้าจอนี้ได้ เนื่องจากไม่มีสิทธิ์การใช้งาน ", vbInformation, "ข้อความแจ้งเตือน"
    Exit Sub
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Order55_Click()
On Error GoTo ErrDescription

vFormActivate = "Form55"
Call ChekAuthorityAccess
If vAccess = 1 Then
Form55.Show
Form55.SetFocus
Else
    MsgBox "User ID : " & vUserID & " ไม่สามารถเข้าใช้หน้าจอนี้ได้ เนื่องจากไม่มีสิทธิ์การใช้งาน ", vbInformation, "ข้อความแจ้งเตือน"
    Exit Sub
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Order6_0_Click()
On Error GoTo ErrDescription

vFormActivate = "Form6_0"
Call ChekAuthorityAccess
If vAccess = 1 Then
Form6_0.Show
Form6_0.SetFocus
Else
    MsgBox "User ID : " & vUserID & " ไม่สามารถเข้าใช้หน้าจอนี้ได้ เนื่องจากไม่มีสิทธิ์การใช้งาน ", vbInformation, "ข้อความแจ้งเตือน"
    Exit Sub
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Order6_1_Click()
On Error GoTo ErrDescription

vFormActivate = "Form6_1"
Call ChekAuthorityAccess
If vAccess = 1 Then
Form6_1.Show
Form6_1.SetFocus
Else
    MsgBox "User ID : " & vUserID & " ไม่สามารถเข้าใช้หน้าจอนี้ได้ เนื่องจากไม่มีสิทธิ์การใช้งาน ", vbInformation, "ข้อความแจ้งเตือน"
    Exit Sub
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Order6_2_Click()
On Error GoTo ErrDescription

'vFormActivate = "Form6_2"
'Call ChekAuthorityAccess
'If vAccess = 1 Then
Form6_2.Show
Form6_2.SetFocus
'Else
 '   MsgBox "User ID : " & vUserID & " ไม่สามารถเข้าใช้หน้าจอนี้ได้ เนื่องจากไม่มีสิทธิ์การใช้งาน ", vbInformation, "ข้อความแจ้งเตือน"
  '  Exit Sub
'End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Order6_3_Click()
On Error GoTo ErrDescription

'vFormActivate = "Form6_2"
'Call ChekAuthorityAccess
'If vAccess = 1 Then
Form6_3.Show
Form6_3.SetFocus
'Else
 '   MsgBox "User ID : " & vUserID & " ไม่สามารถเข้าใช้หน้าจอนี้ได้ เนื่องจากไม่มีสิทธิ์การใช้งาน ", vbInformation, "ข้อความแจ้งเตือน"
  '  Exit Sub
'End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Order631_Click()
On Error GoTo ErrDescription

vFormActivate = "Form631"
Call ChekAuthorityAccess
If vAccess = 1 Then
Form631.Show
Form631.SetFocus
Else
    MsgBox "User ID : " & vUserID & " ไม่สามารถเข้าใช้หน้าจอนี้ได้ เนื่องจากไม่มีสิทธิ์การใช้งาน ", vbInformation, "ข้อความแจ้งเตือน"
    Exit Sub
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Order7_0_Click()
On Error GoTo ErrDescription

vFormActivate = "Form7_0"
Call ChekAuthorityAccess
If vAccess = 1 Then
Form7_0.Show
Form7_0.SetFocus
Else
    MsgBox "User ID : " & vUserID & " ไม่สามารถเข้าใช้หน้าจอนี้ได้ เนื่องจากไม่มีสิทธิ์การใช้งาน ", vbInformation, "ข้อความแจ้งเตือน"
    Exit Sub
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Order999_Click()
'MDIFrmProgramPrint.Arrange (1)
End Sub

Private Sub Order71_Click()
On Error GoTo ErrDescription

vFormActivate = "Form71"
Call ChekAuthorityAccess
If vAccess = 1 Then
Form71.Show
Form71.SetFocus
Else
    MsgBox "User ID : " & vUserID & " ไม่สามารถเข้าใช้หน้าจอนี้ได้ เนื่องจากไม่มีสิทธิ์การใช้งาน ", vbInformation, "ข้อความแจ้งเตือน"
    Exit Sub
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Order81_Click()
On Error GoTo ErrDescription

vFormActivate = "Form81"
Call ChekAuthorityAccess
If vAccess = 1 Then
Form81.Show
Form81.SetFocus
Else
    MsgBox "User ID : " & vUserID & " ไม่สามารถเข้าใช้หน้าจอนี้ได้ เนื่องจากไม่มีสิทธิ์การใช้งาน ", vbInformation, "ข้อความแจ้งเตือน"
    Exit Sub
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Order810_Click()
On Error GoTo ErrDescription

vFormActivate = "Form810"
Call ChekAuthorityAccess
If vAccess = 1 Then
Form810.Show
Form810.SetFocus
Else
    MsgBox "User ID : " & vUserID & " ไม่สามารถเข้าใช้หน้าจอนี้ได้ เนื่องจากไม่มีสิทธิ์การใช้งาน ", vbInformation, "ข้อความแจ้งเตือน"
    Exit Sub
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Order814_Click()
On Error GoTo ErrDescription

vFormActivate = "Form814"
Call ChekAuthorityAccess
If vAccess = 1 Then
Form814.Show
Form814.SetFocus
Else
    MsgBox "User ID : " & vUserID & " ไม่สามารถเข้าใช้หน้าจอนี้ได้ เนื่องจากไม่มีสิทธิ์การใช้งาน ", vbInformation, "ข้อความแจ้งเตือน"
    Exit Sub
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Order82_Click()
On Error GoTo ErrDescription

vFormActivate = "Form82"
Call ChekAuthorityAccess
If vAccess = 1 Then
Form82.Show
Form82.SetFocus
Else
    MsgBox "User ID : " & vUserID & " ไม่สามารถเข้าใช้หน้าจอนี้ได้ เนื่องจากไม่มีสิทธิ์การใช้งาน ", vbInformation, "ข้อความแจ้งเตือน"
    Exit Sub
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Order86_Click()
On Error GoTo ErrDescription

vFormActivate = "Form86"
Call ChekAuthorityAccess
If vAccess = 1 Then
Form86.Show
Form86.SetFocus
Else
    MsgBox "User ID : " & vUserID & " ไม่สามารถเข้าใช้หน้าจอนี้ได้ เนื่องจากไม่มีสิทธิ์การใช้งาน ", vbInformation, "ข้อความแจ้งเตือน"
    Exit Sub
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Order87_Click()
On Error GoTo ErrDescription

vFormActivate = "Form87"
Call ChekAuthorityAccess
If vAccess = 1 Then
Form87.Show
Form87.SetFocus
Else
    MsgBox "User ID : " & vUserID & " ไม่สามารถเข้าใช้หน้าจอนี้ได้ เนื่องจากไม่มีสิทธิ์การใช้งาน ", vbInformation, "ข้อความแจ้งเตือน"
    Exit Sub
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Order871_Click()
On Error GoTo ErrDescription

vFormActivate = "Form871"
Call ChekAuthorityAccess
If vAccess = 1 Then
Form871.Show
Form871.SetFocus
Else
    MsgBox "User ID : " & vUserID & " ไม่สามารถเข้าใช้หน้าจอนี้ได้ เนื่องจากไม่มีสิทธิ์การใช้งาน ", vbInformation, "ข้อความแจ้งเตือน"
    Exit Sub
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Order88_Click()
On Error GoTo ErrDescription

vFormActivate = "Form88"
Call ChekAuthorityAccess
If vAccess = 1 Then
Form88.Show
Form88.SetFocus
Else
    MsgBox "User ID : " & vUserID & " ไม่สามารถเข้าใช้หน้าจอนี้ได้ เนื่องจากไม่มีสิทธิ์การใช้งาน ", vbInformation, "ข้อความแจ้งเตือน"
    Exit Sub
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Order89_Click()
On Error GoTo ErrDescription

vFormActivate = "Form89"
Call ChekAuthorityAccess
If vAccess = 1 Then
Form89.Show
Form89.SetFocus
Else
    MsgBox "User ID : " & vUserID & " ไม่สามารถเข้าใช้หน้าจอนี้ได้ เนื่องจากไม่มีสิทธิ์การใช้งาน ", vbInformation, "ข้อความแจ้งเตือน"
    Exit Sub
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Order91_Click()
'On Error GoTo ErrDescription

'vFormActivate = "Form91"
'Call ChekAuthorityAccess
'If vAccess = 1 Then
'Form91.Show
'Form91.SetFocus
'Else
 '   MsgBox "User ID : " & vUserID & " ไม่สามารถเข้าใช้หน้าจอนี้ได้ เนื่องจากไม่มีสิทธิ์การใช้งาน ", vbInformation, "ข้อความแจ้งเตือน"
  '  Exit Sub
'End If

'ErrDescription:
'If Err.Description <> "" Then
'MsgBox Err.Description
'Exit Sub
'End If
End Sub


Private Sub Order92_Click()
'On Error GoTo ErrDescription

'vFormActivate = "Form92"
'Call ChekAuthorityAccess
'If vAccess = 1 Then
'Form92.Show
'Form92.SetFocus
'Else
 '   MsgBox "User ID : " & vUserID & " ไม่สามารถเข้าใช้หน้าจอนี้ได้ เนื่องจากไม่มีสิทธิ์การใช้งาน ", vbInformation, "ข้อความแจ้งเตือน"
  '  Exit Sub
'End If

'ErrDescription:
'If Err.Description <> "" Then
'MsgBox Err.Description
'Exit Sub
'End If
End Sub

Private Sub Order93_Click()
'On Error GoTo ErrDescription

'vFormActivate = "Form93"
'Call ChekAuthorityAccess
'If vAccess = 1 Then
'Form93.Show
'Form93.SetFocus
'Else
 '   MsgBox "User ID : " & vUserID & " ไม่สามารถเข้าใช้หน้าจอนี้ได้ เนื่องจากไม่มีสิทธิ์การใช้งาน ", vbInformation, "ข้อความแจ้งเตือน"
  '  Exit Sub
'End If

'ErrDescription:
'If Err.Description <> "" Then
'MsgBox Err.Description
'Exit Sub
'End If
End Sub

Private Sub Order94_Click()
'On Error GoTo ErrDescription

'vFormActivate = "Form94"
'Call ChekAuthorityAccess
'If vAccess = 1 Then
'Form94.Show
'Form94.SetFocus
'Else
 '   MsgBox "User ID : " & vUserID & " ไม่สามารถเข้าใช้หน้าจอนี้ได้ เนื่องจากไม่มีสิทธิ์การใช้งาน ", vbInformation, "ข้อความแจ้งเตือน"
  '  Exit Sub
'End If

'ErrDescription:
'If Err.Description <> "" Then
'MsgBox Err.Description
'Exit Sub
'End If
End Sub

Private Sub Order95_Click()
'On Error GoTo ErrDescription

'vFormActivate = "Form95"
'Call ChekAuthorityAccess
'If vAccess = 1 Then
'Form95.Show
'Form95.SetFocus
'Else
 '   MsgBox "User ID : " & vUserID & " ไม่สามารถเข้าใช้หน้าจอนี้ได้ เนื่องจากไม่มีสิทธิ์การใช้งาน ", vbInformation, "ข้อความแจ้งเตือน"
  '  Exit Sub
'End If

'ErrDescription:
'If Err.Description <> "" Then
'MsgBox Err.Description
'Exit Sub
'End If
End Sub

Private Sub Order96_Click()
'On Error GoTo ErrDescription

'vFormActivate = "Form96"
'Call ChekAuthorityAccess
'If vAccess = 1 Then
'Form96.Show
'Form96.SetFocus
'Else
 '   MsgBox "User ID : " & vUserID & " ไม่สามารถเข้าใช้หน้าจอนี้ได้ เนื่องจากไม่มีสิทธิ์การใช้งาน ", vbInformation, "ข้อความแจ้งเตือน"
  '  Exit Sub
'End If

'ErrDescription:
'If Err.Description <> "" Then
'MsgBox Err.Description
'Exit Sub
'End If
End Sub

Private Sub Order97_Click()
'On Error GoTo ErrDescription

'vFormActivate = "Form97"
'Call ChekAuthorityAccess
'If vAccess = 1 Then
'Form97.Show
'Form97.SetFocus
'Else
 '   MsgBox "User ID : " & vUserID & " ไม่สามารถเข้าใช้หน้าจอนี้ได้ เนื่องจากไม่มีสิทธิ์การใช้งาน ", vbInformation, "ข้อความแจ้งเตือน"
  '  Exit Sub
'End If

'ErrDescription:
'If Err.Description <> "" Then
'MsgBox Err.Description
'Exit Sub
'End If
End Sub

Private Sub Order98_Click()
'On Error GoTo ErrDescription

'vFormActivate = "Form98"
'Call ChekAuthorityAccess
'If vAccess = 1 Then
Form98.Show
Form98.SetFocus
'Else
 '   MsgBox "User ID : " & vUserID & " ไม่สามารถเข้าใช้หน้าจอนี้ได้ เนื่องจากไม่มีสิทธิ์การใช้งาน ", vbInformation, "ข้อความแจ้งเตือน"
  '  Exit Sub
'End If

'ErrDescription:
'If Err.Description <> "" Then
'MsgBox Err.Description
'Exit Sub
'End If
End Sub

Private Sub Order99_Click()
'Form99.Show
'Form99.SetFocus
End Sub

Private Sub Order991_Click()
Form991.Show
Form991.SetFocus
End Sub

Private Sub OrderChangeItemandPrice_Click()
On Error GoTo ErrDescription

vFormActivate = "Form811"
Call ChekAuthorityAccess
If vAccess = 1 Then
Form811.Show
Form811.SetFocus
Else
    MsgBox "User ID : " & vUserID & " ไม่สามารถเข้าใช้หน้าจอนี้ได้ เนื่องจากไม่มีสิทธิ์การใช้งาน ", vbInformation, "ข้อความแจ้งเตือน"
    Exit Sub
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub OrderEditReturn_Click()
FormEditReturn.Show
FormEditReturn.SetFocus
End Sub

Private Sub OrderItemIssue_Click()
On Error GoTo ErrDescription

'vFormActivate = "Form812"
'Call ChekAuthorityAccess
'If vAccess = 1 Then
FormItemIssue.Show
FormItemIssue.SetFocus
'Else
 '   MsgBox "User ID : " & vUserID & " ไม่สามารถเข้าใช้หน้าจอนี้ได้ เนื่องจากไม่มีสิทธิ์การใช้งาน ", vbInformation, "ข้อความแจ้งเตือน"
  '  Exit Sub
'End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub OrderPick_Click()
On Error Resume Next

MsgBox ("ยกเลิกการใช้งานชั่วคราว")

'On Error GoTo ErrDescription

'vFormActivate = "Form311"
'Call ChekAuthorityAccess
'If vAccess = 1 Then
'FormPickingRequest.Show
'FormPickingRequest.SetFocus
'Else
 '   MsgBox "User ID : " & vUserID & " ไม่สามารถเข้าใช้หน้าจอนี้ได้ เนื่องจากไม่มีสิทธิ์การใช้งาน ", vbInformation, "ข้อความแจ้งเตือน"
  '  Exit Sub
'End If

'ErrDescription:
'If Err.Description <> "" Then
'MsgBox Err.Description
'Exit Sub
'End If
End Sub

Private Sub OrderPOSItem_Click()
Dim vDateThai As String

vDateThai = Format(Now, "dddd")
If (UCase(vUserID) = Trim(UCase("boontiwa")) Or UCase(vUserID) = Trim(UCase("gittichai")) Or UCase(vUserID) = Trim(UCase("somchai")) Or UCase(vUserID) = Trim(UCase("jip"))) Or UCase(vUserID) = Trim(UCase("somrod")) Or UCase(vUserID) = Trim(UCase("porntip")) Then
    FormOpenItemPOS.Show
    FormOpenItemPOS.SetFocus
Else
    MsgBox "User นี้ไม่มีสิทธิ์ในการเปิดขายติดลบ POS", vbCritical, "Send Error"
    Exit Sub
End If

End Sub

Private Sub OrderPrint_Click()
On Error GoTo ErrDescription

vFormActivate = "Form812"
Call ChekAuthorityAccess
If vAccess = 1 Then
Form812.Show
Form812.SetFocus
Else
    MsgBox "User ID : " & vUserID & " ไม่สามารถเข้าใช้หน้าจอนี้ได้ เนื่องจากไม่มีสิทธิ์การใช้งาน ", vbInformation, "ข้อความแจ้งเตือน"
    Exit Sub
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub OrderReserve_Click()
On Error GoTo ErrDescription

vFormActivate = "FormPrintReserve"
Call ChekAuthorityAccess
If vAccess = 1 Then
FormPrintReserve.Show
FormPrintReserve.SetFocus
Else
    MsgBox "User ID : " & vUserID & " ไม่สามารถเข้าใช้หน้าจอนี้ได้ เนื่องจากไม่มีสิทธิ์การใช้งาน ", vbInformation, "ข้อความแจ้งเตือน"
    Exit Sub
End If
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If

End Sub

Private Sub OrderStockCountHMX_Click()
FormStockCount.Show
FormStockCount.SetFocus
End Sub

Private Sub OrderTaxPurchase_Click()
FormTaxPurchase.Show
FormTaxPurchase.SetFocus
End Sub

Private Sub StockCardBC_Click()
On Error GoTo ErrDescription

'vFormActivate = "FormStockCardBC"
'Call ChekAuthorityAccess
'If vAccess = 1 Then
FormStockCardBC.Show
FormStockCardBC.SetFocus
'Else
 '   MsgBox "User ID : " & vUserID & " ไม่สามารถเข้าใช้หน้าจอนี้ได้ เนื่องจากไม่มีสิทธิ์การใช้งาน ", vbInformation, "ข้อความแจ้งเตือน"
  '  Exit Sub
'End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub TransferPrint_Click()
On Error GoTo ErrDescription

vFormActivate = "Form813"
Call ChekAuthorityAccess
If vAccess = 1 Then
Form813.Show
Form813.SetFocus
Else
    MsgBox "User ID : " & vUserID & " ไม่สามารถเข้าใช้หน้าจอนี้ได้ เนื่องจากไม่มีสิทธิ์การใช้งาน ", vbInformation, "ข้อความแจ้งเตือน"
    Exit Sub
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub
