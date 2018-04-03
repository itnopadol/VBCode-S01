VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "โปรแกรมจัดการโปรโมชั่น เวอร์ชั่น 1.2"
   ClientHeight    =   9015
   ClientLeft      =   2550
   ClientTop       =   750
   ClientWidth     =   14385
   LinkTopic       =   "MDIForm1"
   Begin VB.Menu Order0 
      Caption         =   "โปรแกรม"
      Begin VB.Menu Order001 
         Caption         =   "เลือก User เข้าทำงานใหม่"
      End
      Begin VB.Menu Order002 
         Caption         =   "กำหนดสิทธิ์เข้าใช้งาน"
      End
   End
   Begin VB.Menu Order1 
      Caption         =   "สำหรับฝ่ายขาย"
      Begin VB.Menu Order101 
         Caption         =   "สร้างใบเสนอสินค้า"
      End
      Begin VB.Menu Order105 
         Caption         =   "สร้างใบขอยกเลิกสินค้าโปรโมชั่น"
      End
      Begin VB.Menu Order0001 
         Caption         =   "-"
      End
      Begin VB.Menu Order103 
         Caption         =   "พิมพ์ป้ายราคาโปรโมชั่น"
      End
      Begin VB.Menu Order104 
         Caption         =   "ลบ User ค้าง"
      End
   End
   Begin VB.Menu Order2 
      Caption         =   "สำหรับฝ่ายจัดซื้อ"
      Begin VB.Menu Order201 
         Caption         =   "ตรวจสอบใบเสนอสินค้า"
      End
   End
   Begin VB.Menu Order3 
      Caption         =   "สำหรับผู้อนุมัติ"
      Begin VB.Menu Order301 
         Caption         =   "อนุมัติใบเสนอสินค้า"
      End
      Begin VB.Menu Order302 
         Caption         =   "ยกเลิกโปรโมชั่น"
      End
   End
   Begin VB.Menu Order4 
      Caption         =   "สำหรับฝ่ายการตลาด"
      Begin VB.Menu Order401 
         Caption         =   "เพิ่มทะเบียนโปรโมชั่น"
      End
      Begin VB.Menu Order402 
         Caption         =   "เพิ่มและพิมพ์คูปอง"
      End
   End
   Begin VB.Menu Order9 
      Caption         =   "หน้าต่างที่เปิดไว้"
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
Form000.SetFocus
Order0.Enabled = False
Order1.Enabled = False
Order2.Enabled = False
Order3.Enabled = False
Order4.Enabled = False
Order9.Enabled = False
vMemCommand = 0
End Sub


Private Sub MDIForm_Unload(Cancel As Integer)
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vCheckDocno As String
Dim vCheckPromo As String
Dim vCheckExitProgram As Integer

On Error Resume Next

vCheckExitProgram = MsgBox("คุณต้องการออกจากโปรแกรมเสนอสินค้าโปรโมชั่นนี้แล้ว ใช่หรือไม่", vbYesNo, "ข้อความสอบถาม")
If vCheckExitProgram = 6 Then
    If gConnectionString <> "" Then
            vCheckDocno = Trim(Form201.Text101.Text)
            vCheckPromo = Left(Trim(Form201.Text102.Text), InStr(Trim(Form201.Text102.Text), "/") - 1)
    
            If vCheckUsedID = 1 Then
                vQuery = "delete npmaster.dbo.TB_CK_UserActivateProgram where userid = '" & vUserID & "' and jobid = 2"
                gConnection.Execute vQuery
                
                vQuery = "delete npmaster.dbo. TB_PM_TempCheckItemDuplicateLine where userid = '" & vUserID & "'  "
                gConnection.Execute vQuery
            End If
    
        If vCheckDocno <> "" Then
            vQuery = "USP_PM_DeleteCheckDuplicatItemLine '" & vCheckDocno & "','" & vCheckPromo & "','" & vUserID & "' "
            gConnection.Execute vQuery
            
            vQuery = "delete npmaster.dbo. TB_PM_TempCheckItemDuplicateLine where userid = '" & vUserID & "'  "
            gConnection.Execute vQuery
    
        End If
        If gConnection.State = 1 Then
            gConnection.Close
        End If
    End If
Else
    Cancel = True
End If
End Sub

Private Sub Order001_Click()
Form000.Show
Form000.SetFocus
MDIForm1.Caption = "โปรแกรมจัดการโปรโมชั่น เวอร์ชั่น 1.1"
Order1.Enabled = False
Order2.Enabled = False
Order3.Enabled = False
Order4.Enabled = False
Order9.Enabled = False
Form000.Text101.SetFocus
Form000.Text102.Text = ""
vMemCommand = 0
Unload Form101
Unload Form103
Unload Form201
Unload Form301
Unload Form401
End Sub

Private Sub Order002_Click()
    Form002.SetFocus
    Form002.Show
End Sub

Private Sub Order101_Click()
On Error GoTo ErrDescription

vFormActivate = "Form201"


Call ChekAuthorityAccess
If vAccess = 1 Then
    Form201.SetFocus
    Form201.Show
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

Private Sub Order102_Click()
'Form202.Show
'Form202.SetFocus
End Sub

Private Sub Order103_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vCheckUserID As Integer

On Error GoTo ErrDescription

vFormActivate = "Form103"
Call ChekAuthorityAccess
If vAccess = 1 Then
    If vCheckSetFocus = 0 Then
        vQuery = "select userid from npmaster.dbo.TB_CK_UserActivateProgram  where userid = '" & vUserID & "' and jobid = 2"
        If OpenDatabase(gConnection, vRecordset, vQuery) <> 0 Then
            vCheckUserID = 1
            MsgBox "มี UserID : " & vUserID & " เข้ามาใช้งานในหน้านี้แล้ว กรุณาตรวจสอบ ไม่งั้นจะไม่สามารถพิมพ์ป้ายราคาได้ กรุณาติดต่อคอมฯนะครับ"
        Else
            vCheckUserID = 0
        End If
        vRecordset.Close
        
        If vCheckUserID = 0 Then
            vQuery = "insert into npmaster.dbo.TB_CK_UserActivateProgram (UserID,ActiveDateTime,ActiveStatus,JobId) " _
                                & " values ('" & vUserID & "',getdate(),1,2)"
            gConnection.Execute (vQuery)
            vCheckUsedID = 1
            vCheckSetFocus = 1
            Form103.Show
            Form103.SetFocus
        End If
    Else
            Form103.SetFocus
    End If
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

Private Sub Order104_Click()
On Error GoTo ErrDescription

vFormActivate = "Form104"
Call ChekAuthorityAccess
If vAccess = 1 Then
    Form104.SetFocus
    Form104.Show
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

Private Sub Order105_Click()
On Error GoTo ErrDescription

'vFormActivate = "Form105"
'Call ChekAuthorityAccess
'If vAccess = 1 Then
    Form105.SetFocus
    Form105.Show
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

Private Sub Order201_Click()
'On Error GoTo ErrDescription
vFormActivate = "Form401"
Call ChekAuthorityAccess

If vAccess = 1 Then
    Form401.Show
    Form401.SetFocus
Else
    MsgBox "User ID : " & vUserID & " ไม่สามารถเข้าใช้หน้าจอนี้ได้ เนื่องจากไม่มีสิทธิ์การใช้งาน ", vbInformation, "ข้อความแจ้งเตือน"
    Exit Sub
End If

'ErrDescription:
'If Err.Description <> "" Then
'MsgBox Err.Description
'Exit Sub
'End If
End Sub

Private Sub Order301_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vListConfirmPromotion As ListItem
Dim vListConfirmItem As ListItem

On Error GoTo ErrDescription

vFormActivate = "Form301"
Call ChekAuthorityAccess
If vAccess = 1 Then
    Form301.ListView101.ListItems.Clear
    vQuery = "exec USP_PM_CheckOrConfirmSearch '1' "
    If OpenDatabaseBPlus(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    While Not vRecordset.EOF
        Set vListConfirmPromotion = Form301.ListView101.ListItems.Add(, , Trim(vRecordset.Fields("docno").Value))
        If IsNull(Trim(vRecordset.Fields("secmanname").Value)) Then
            vListConfirmPromotion.SubItems(1) = ""
         Else
            vListConfirmPromotion.SubItems(1) = Trim(vRecordset.Fields("secmanname").Value)
        End If
        If IsNull(Trim(vRecordset.Fields("createdate").Value)) Then
        vListConfirmPromotion.SubItems(2) = ""
        Else
            vListConfirmPromotion.SubItems(2) = Trim(vRecordset.Fields("createdate").Value)
        End If
        If IsNull(Trim(vRecordset.Fields("checkdate").Value)) Then
            vListConfirmPromotion.SubItems(3) = ""
        Else
            vListConfirmPromotion.SubItems(3) = Trim(vRecordset.Fields("checkdate").Value)
        End If
        If IsNull(Trim(vRecordset.Fields("checkername").Value)) Then
            vListConfirmPromotion.SubItems(4) = ""
        Else
            vListConfirmPromotion.SubItems(4) = Trim(vRecordset.Fields("checkername").Value)
        End If
        vListConfirmPromotion.SubItems(5) = Trim(vRecordset.Fields("pmname").Value)
    vRecordset.MoveNext
    Wend
    End If
    vRecordset.Close
    
    vQuery = "exec USP_PM_CheckOrConfirmSearch '2' "
        If OpenDatabaseBPlus(gConnection, vRecordset, vQuery) <> 0 Then
            vRecordset.MoveFirst
            While Not vRecordset.EOF
                Set vListConfirmItem = Form301.ListView102.ListItems.Add(, , Trim(vRecordset.Fields("docno").Value))
                If IsNull(Trim(vRecordset.Fields("secmanname").Value)) Then
                    vListConfirmItem.SubItems(1) = ""
                 Else
                    vListConfirmItem.SubItems(1) = Trim(vRecordset.Fields("secmanname").Value)
                End If
                If IsNull(Trim(vRecordset.Fields("createdate").Value)) Then
                vListConfirmItem.SubItems(2) = ""
                Else
                    vListConfirmItem.SubItems(2) = Trim(vRecordset.Fields("createdate").Value)
                End If
                If IsNull(Trim(vRecordset.Fields("checkdate").Value)) Then
                    vListConfirmItem.SubItems(3) = ""
                Else
                    vListConfirmItem.SubItems(3) = Trim(vRecordset.Fields("checkdate").Value)
                End If
                If IsNull(Trim(vRecordset.Fields("checkername").Value)) Then
                    vListConfirmItem.SubItems(4) = ""
                Else
                    vListConfirmItem.SubItems(4) = Trim(vRecordset.Fields("checkername").Value)
                End If
            vRecordset.MoveNext
            Wend
        End If
        vRecordset.Close
    Form301.Show
    Form301.SetFocus
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

Private Sub Order302_Click()
On Error GoTo ErrDescription

vFormActivate = "Form302"
Call ChekAuthorityAccess
If vAccess = 1 Then
    Form302.SetFocus
    Form302.Show
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

Private Sub Order401_Click()
On Error GoTo ErrDescription

vFormActivate = "Form101"
Call ChekAuthorityAccess
If vAccess = 1 Then
    Form101.SetFocus
    Form101.Show
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

Private Sub Order402_Click()
On Error GoTo ErrDescription

'vFormActivate = "Form102"
'Call ChekAuthorityAccess
'If vAccess = 1 Then
    'Form102.SetFocus
    'Form102.Show
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

