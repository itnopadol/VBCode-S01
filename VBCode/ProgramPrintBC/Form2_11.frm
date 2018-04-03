VERSION 5.00
Begin VB.Form Form2_11 
   Caption         =   "ยกเลิกอนุมัติ ใบสั่งซื้อสินค้า (สำหรับ CAT)"
   ClientHeight    =   8115
   ClientLeft      =   5130
   ClientTop       =   1980
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form2_11.frx":0000
   ScaleHeight     =   8115
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CMDPrint 
      Caption         =   "ยกเลิกอนุมัติ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5580
      TabIndex        =   2
      Top             =   2520
      Width           =   1815
   End
   Begin VB.TextBox TXTDocNo 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   4560
      TabIndex        =   1
      Top             =   1620
      Width           =   2835
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "เลขที่เอกสาร :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   465
      Left            =   2820
      TabIndex        =   0
      Top             =   1680
      Width           =   1635
   End
End
Attribute VB_Name = "Form2_11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMDPrint_Click()
Dim vQuery As String
Dim vRecordset As New Recordset
Dim vDocNo As String
Dim vAnswer As Integer
Dim vGetDocNo As String

On Error Resume Next

'If Me.TXTDocNo.Text <> "" And (UCase(vUserID) = "YURAPORN" Or UCase(vUserID) = "AMPAN" Or UCase(vUserID) = "KANYAPATCH" Or UCase(vUserID) = "SOMROD" Or UCase(vUserID) = "KHANNIKAR" Or UCase(vUserID) = "KAEWALIN" Or UCase(vUserID) = "RONNARONG" Or UCase(vUserID) = "WACHIPORN" Or UCase(vUserID) = "WARAMATSW" Or UCase(vUserID) = "PANITHI" Or UCase(vUserID) = "SIRIRANYA" Or UCase(vUserID) = "NAPAPANKK" Or UCase(vUserID) = "SOMROD") Then

Call ChekAuthorityAccess

If Me.TXTDocNo.Text <> "" And (vDepartment = "MC" Or vDepartment = "BY") Then
    vDocNo = Me.TXTDocNo.Text
        
    vQuery = "select docno from  dbo.bcpurchaseorder where docno = '" & vDocNo & "' and iscancel = 0 and isconfirm = 1 "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vGetDocNo = Trim(vRecordset.Fields("docno").Value)
    Else
    MsgBox "เอกสารใบสั่งซื้อสินค้าเลขที่ " & vDocNo & " ไม่สามารถยกเลิกอนุมัติได้ กรุณาตรวจสอบ"
    Me.TXTDocNo.SetFocus
    Exit Sub
    End If
    vRecordset.Close
        
    
    vAnswer = MsgBox("คุณต้องการยกเลิก อนุมัติใบสั่งซื้อสินค้าเลขที่ " & vDocNo & " นี้ใช่หรือไม่ ?", vbYesNo, "Send Question Message ?")
        
    If vAnswer = 6 Then
    vQuery = "Update dbo.bcpurchaseOrder Set isconfirm = 0 where docno = '" & vDocNo & "' "
    gConnection.Execute vQuery
    
    MsgBox ("ยกเลิก อนุมัติเอกสารใบสั่งซื้อสินค้าเรียบร้อยแล้ว แก้ไขเอกสารได้เลยครับ")
    End If
    
    Me.TXTDocNo.Text = ""
    Me.TXTDocNo.SetFocus
End If
End Sub

