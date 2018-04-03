VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form3_17 
   Caption         =   "เปลี่ยนวันที่หมดอายุใบเสนอราคาและ Back Order"
   ClientHeight    =   8370
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form3_17.frx":0000
   ScaleHeight     =   8370
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CMD101 
      Caption         =   "เปลี่ยน"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5715
      TabIndex        =   3
      Top             =   3240
      Width           =   735
   End
   Begin MSComCtl2.DTPicker DTPicker101 
      Height          =   330
      Left            =   4860
      TabIndex        =   2
      Top             =   2475
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   66584577
      CurrentDate     =   38860
   End
   Begin VB.TextBox Text101 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   4860
      TabIndex        =   0
      Top             =   1350
      Width           =   1590
   End
   Begin VB.Label LBL101 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   4860
      TabIndex        =   1
      Top             =   1890
      Width           =   1590
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "เปลี่ยนเป็นวันที่ :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   3510
      TabIndex        =   6
      Top             =   2475
      Width           =   1275
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "วันที่หมดอายุ :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   330
      Left            =   3285
      TabIndex        =   5
      Top             =   1890
      Width           =   1500
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "เลขที่เอกสาร :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   330
      Left            =   3375
      TabIndex        =   4
      Top             =   1350
      Width           =   1410
   End
End
Attribute VB_Name = "Form3_17"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMD101_Click()
Dim vRecordset  As New ADODB.Recordset
Dim vQuery As String
Dim vDocNo As String
Dim vDocdate As Date
Dim vExpireDate As Date
Dim vAnswer As Integer

On Error GoTo ErrDescription

If Text101.Text <> "" And LBL101.Caption <> "" Then
    vDocNo = Trim(Text101.Text)
    vAnswer = MsgBox("คุณต้องการเปลี่ยนวันที่หมดอายุของเอกสารเลขที่ " & vDocNo & " นี้ใช่หรือไม่", vbYesNo, "Question ?")
    If vAnswer = 6 Then
        vExpireDate = DTPicker101.Day & "/" & DTPicker101.Month & "/" & DTPicker101.Year
        vQuery = "exec dbo.USP_BC_QuotationUpdateExpireDate '" & vDocNo & "','" & vExpireDate & "' "
        gConnection.Execute vQuery
    Else
        Exit Sub
    End If
    MsgBox "เปลี่ยนวันที่หมดอายุของเลขที่เอกสาร " & vDocNo & " เรียบร้อยแล้ว", vbInformation, "Send Message"
    Text101.Text = ""
    LBL101.Caption = ""
    DTPicker101 = Now
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub

Private Sub Form_Load()
DTPicker101 = Now
End Sub

Private Sub Text101_KeyPress(KeyAscii As Integer)
Dim vRecordset  As New ADODB.Recordset
Dim vQuery As String
Dim vDocNo As String
Dim vDocdate As Date
Dim vExpireDate As Date

On Error GoTo ErrDescription

If KeyAscii = 13 Then
    If Text101.Text <> "" Then
        vDocNo = Trim(Text101.Text)
        vQuery = "exec dbo.USP_BC_QuotationChangeExpireDate  '" & vDocNo & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vExpireDate = Trim(vRecordset.Fields("expiredate").Value)
            vDocdate = Trim(vRecordset.Fields("docdate").Value)
        Else
            MsgBox "ไม่มีข้อมูลของเอกสารเลขที่ " & vDocNo & " กรุณาตรวจสอบ", vbCritical, "Send Message"
            Exit Sub
        End If
        vRecordset.Close
        LBL101.Caption = vExpireDate
        
    End If
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub
