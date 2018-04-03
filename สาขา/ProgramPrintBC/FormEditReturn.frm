VERSION 5.00
Begin VB.Form FormEditReturn 
   Caption         =   "แก้ไขส่วนลดในเอกสารลดหนี้"
   ClientHeight    =   8055
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "FormEditReturn.frx":0000
   ScaleHeight     =   8055
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CMD103 
      Caption         =   "ประมวลผล"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4410
      TabIndex        =   15
      Top             =   5760
      Width           =   1185
   End
   Begin VB.CommandButton CMD102 
      Caption         =   "คำนวณ"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5670
      TabIndex        =   10
      Top             =   1530
      Width           =   1185
   End
   Begin VB.TextBox Text102 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3960
      TabIndex        =   6
      Top             =   1530
      Width           =   1635
   End
   Begin VB.TextBox Text101 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3960
      TabIndex        =   5
      Top             =   990
      Width           =   1635
   End
   Begin VB.Label LBL106 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   3960
      TabIndex        =   17
      Top             =   2790
      Width           =   1635
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "มูลค่าสินค้าลดหนี้ :"
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   2295
      TabIndex        =   16
      Top             =   2835
      Width           =   1590
   End
   Begin VB.Label LBL105 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   3960
      TabIndex        =   14
      Top             =   5265
      Width           =   1635
   End
   Begin VB.Label LBL104 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   3960
      TabIndex        =   13
      Top             =   4770
      Width           =   1635
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ยอดภาษี :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   330
      Left            =   2655
      TabIndex        =   12
      Top             =   5265
      Width           =   1230
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ยอดลดหนี้ก่อนภาษี :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   285
      Left            =   2295
      TabIndex        =   11
      Top             =   4770
      Width           =   1590
   End
   Begin VB.Label LBL103 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   3960
      TabIndex        =   9
      Top             =   4275
      Width           =   1635
   End
   Begin VB.Label LBL102 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   3960
      TabIndex        =   8
      Top             =   3780
      Width           =   1635
   End
   Begin VB.Label LBL101 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   3960
      TabIndex        =   7
      Top             =   3285
      Width           =   1635
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "มูลค่าที่ถูกต้อง :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   2700
      TabIndex        =   4
      Top             =   4275
      Width           =   1185
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "มูลค่าเดิม :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   2745
      TabIndex        =   3
      Top             =   3780
      Width           =   1140
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "มูลค่าลดหนี้ :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   330
      Left            =   2835
      TabIndex        =   2
      Top             =   3285
      Width           =   1050
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ส่วนลดการค้า :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   330
      Left            =   2835
      TabIndex        =   1
      Top             =   1530
      Width           =   1050
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "เลขที่ใบลดหนี้ :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   330
      Left            =   2745
      TabIndex        =   0
      Top             =   990
      Width           =   1140
   End
End
Attribute VB_Name = "FormEditReturn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vDicountOld As Currency

Private Sub CMD102_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDiscountChange As Currency
Dim vSumOfItemAmount As Currency
Dim vSumOfTaxAmount As Currency
Dim vSumOfOldAmount As Currency
Dim vDocNo As String

On Error GoTo ErrDescription

If Text102.Text <> "" Then
    vDiscountChange = Text102.Text
    If vDiscountChange <> vDicountOld Then
    vDocNo = Trim(Text101.Text)
        vQuery = "exec dbo.USP_BC_SearchCreditnote '" & vDocNo & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vSumOfItemAmount = Trim(vRecordset.Fields("sumofitemamount").Value)
            vSumOfOldAmount = Trim(vRecordset.Fields("sumoldamount").Value)
            
        End If
        vRecordset.Close
        vDicountOld = CCur(Text102.Text)
        LBL101.Caption = Format(CCur(vSumOfItemAmount) - CCur(Text102.Text), "##,##0.00")
        LBL103.Caption = Format(CCur(vSumOfOldAmount) - CCur(LBL101.Caption), "##,##0.00")
        LBL104.Caption = Format(((CCur(LBL101.Caption) * 100) / 107), "##,##0.00")
        LBL105.Caption = Format(CCur(LBL101.Caption) - CCur(LBL104.Caption), "##,##0.00")
    End If
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Private Sub CMD103_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocNo As String
Dim vDiffAmount As Currency
Dim vSumTrueAmount As Currency
Dim vSumOfTotalTax As Currency
Dim vSumBeforeTax As Currency
Dim vSumTaxAmount As Currency
Dim vAnswer As Integer
Dim vDiscountAmount As Currency

On Error GoTo ErrDescription

If Text101.Text <> "" And Text102.Text <> "" Then
    vDocNo = Trim(Text101.Text)
    vDiscountAmount = CCur(Text102.Text)
    vAnswer = MsgBox("ต้องการเปลี่ยนส่วนลดการค้าของใบลดหนี้เลขที่ " & vDocNo & " นี้หรือไม่", vbYesNo, "Send Question ?")
    If vAnswer = 6 Then
        vDiffAmount = CCur(LBL106.Caption) - CCur(Text102.Text)
        vSumTrueAmount = CCur(LBL102.Caption) - vDiffAmount
        vSumOfTotalTax = vDiffAmount
        vSumBeforeTax = Format((vSumOfTotalTax * 100) / 107, "##,##0.00")
        vSumTaxAmount = vSumOfTotalTax - vSumBeforeTax
        
        vQuery = "begin tran"
        gConnection.Execute vQuery
        
        vQuery = "exec dbo.USP_BC_UpdateCreditNote '" & vDocNo & "'," & vDiscountAmount & ",'" & vDiscountAmount & "'," & vDiffAmount & "," & vSumTrueAmount & "," & vSumOfTotalTax & "," & vSumBeforeTax & "," & vSumTaxAmount & " "
        gConnection.Execute vQuery
        MsgBox "เปลี่ยนข้อมูลลดหนี้การค้าของใบลดหนี้เลขที่ " & vDocNo & " เรียบร้อยแล้วครับ", vbInformation, "Send Message"
        
        vQuery = "commit tran"
        gConnection.Execute vQuery
        
        Text101.Text = ""
        Text102.Text = ""
        LBL101.Caption = ""
        LBL102.Caption = ""
        LBL103.Caption = ""
        LBL104.Caption = ""
        LBL105.Caption = ""
        LBL106.Caption = ""
        Text101.SetFocus
    Else
        Exit Sub
    End If
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    vQuery = "rollback tran"
    gConnection.Execute vQuery
End If
End Sub

Private Sub Text101_KeyPress(KeyAscii As Integer)
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocNo As String

On Error GoTo ErrDescription

If KeyAscii = 13 Then
    If Text101.Text <> "" Then
        vDocNo = Trim(Text101.Text)
        vQuery = "exec dbo.USP_BC_SearchCreditnote '" & vDocNo & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vDicountOld = Trim(vRecordset.Fields("invoicedisc").Value)
            Text102.Text = Format(Trim(vRecordset.Fields("invoicedisc").Value), "##,##0.00")
            LBL101.Caption = Format(Trim(vRecordset.Fields("sumoftotaltax").Value), "##,##0.00")
            LBL102.Caption = Format(Trim(vRecordset.Fields("sumoldamount").Value), "##,##0.00")
            LBL103.Caption = Format(Trim(vRecordset.Fields("sumtrueamount").Value), "##,##0.00")
            LBL104.Caption = Format(Trim(vRecordset.Fields("sumofbeforetax").Value), "##,##0.00")
            LBL105.Caption = Format(Trim(vRecordset.Fields("sumoftaxamount").Value), "##,##0.00")
            LBL106.Caption = Format(Trim(vRecordset.Fields("sumofitemamount").Value), "##,##0.00")
            CMD102.Enabled = True
        Else
            Text102.Text = ""
            LBL101.Caption = ""
            LBL102.Caption = ""
            LBL103.Caption = ""
            LBL104.Caption = ""
            LBL105.Caption = ""
            CMD102.Enabled = False
        End If
        vRecordset.Close
    End If
End If


ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Private Sub Text101_LostFocus()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocNo As String

On Error GoTo ErrDescription

If Text101.Text <> "" Then
    vDocNo = Trim(Text101.Text)
    vQuery = "exec dbo.USP_BC_SearchCreditnote '" & vDocNo & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vDicountOld = Trim(vRecordset.Fields("invoicedisc").Value)
        Text102.Text = Format(Trim(vRecordset.Fields("invoicedisc").Value), "##,##0.00")
        LBL101.Caption = Format(Trim(vRecordset.Fields("sumoftotaltax").Value), "##,##0.00")
        LBL102.Caption = Format(Trim(vRecordset.Fields("sumoldamount").Value), "##,##0.00")
        LBL103.Caption = Format(Trim(vRecordset.Fields("sumtrueamount").Value), "##,##0.00")
        LBL104.Caption = Format(Trim(vRecordset.Fields("sumofbeforetax").Value), "##,##0.00")
        LBL105.Caption = Format(Trim(vRecordset.Fields("sumoftaxamount").Value), "##,##0.00")
        CMD102.Enabled = True
    Else
        Text102.Text = ""
        LBL101.Caption = ""
        LBL102.Caption = ""
        LBL103.Caption = ""
        LBL104.Caption = ""
        LBL105.Caption = ""
        CMD102.Enabled = False
    End If
    vRecordset.Close
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Private Sub Text102_LostFocus()
Dim vCheckText As String

On Error GoTo ErrDescription

vCheckText = Trim(Text102.Text)
Call CheckNumber(vCheckText)
If vCheckValueNumber = False Then
    Text102.Text = Format(vDicountOld, "##,##0.00")
    MsgBox "ต้องใส่ตัวเลขเท่านั้น", vbCritical, "Send Error"
Else
    Text102.Text = Format(Text102.Text, "##,##0.00")
    Call CMD102_Click
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub
