VERSION 5.00
Begin VB.Form Form21 
   Caption         =   "ยกเลิกจ่ายชำระเอกสารซื้อ"
   ClientHeight    =   8985
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11970
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form21.frx":0000
   ScaleHeight     =   8985
   ScaleWidth      =   11970
   WindowState     =   2  'Maximized
   Begin VB.OptionButton Opt213 
      BackColor       =   &H00FFFFFF&
      Caption         =   "เอกสารส่งคืน/ลดหนี้"
      Height          =   315
      Left            =   7800
      TabIndex        =   6
      Top             =   1200
      Width           =   2265
   End
   Begin VB.OptionButton Opt212 
      BackColor       =   &H00FFFFFF&
      Caption         =   "เอกสารตั้งหนี้อื่น ๆ "
      Height          =   315
      Left            =   5175
      TabIndex        =   5
      Top             =   1200
      Width           =   2265
   End
   Begin VB.OptionButton Opt211 
      BackColor       =   &H00FFFFFF&
      Caption         =   "เอกสารตั้งหนี้จากการซื้อ"
      Height          =   315
      Left            =   2550
      TabIndex        =   4
      Top             =   1200
      Value           =   -1  'True
      Width           =   2265
   End
   Begin VB.CommandButton CMD211 
      Caption         =   "ยกเลิกการอนุมัติ"
      Height          =   615
      Left            =   4575
      TabIndex        =   3
      Top             =   3450
      Width           =   1440
   End
   Begin VB.TextBox TXT211 
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
      Height          =   390
      Left            =   3600
      TabIndex        =   2
      Top             =   2475
      Width           =   2415
   End
   Begin VB.Label LBL212 
      BackStyle       =   0  'Transparent
      Caption         =   "เลขที่เอกสาร"
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
      Height          =   390
      Left            =   2550
      TabIndex        =   1
      Top             =   2475
      Width           =   990
   End
   Begin VB.Label LBL211 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ยกเลิกจ่ายชำระเอกสารซื้อ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   14.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   540
      Left            =   2550
      TabIndex        =   0
      Top             =   225
      Width           =   7515
   End
End
Attribute VB_Name = "Form21"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMD211_Click()
Dim vQuery As String
Dim vDocNo As String
Dim vRecordset As New ADODB.Recordset
Dim vConfirm As Integer
Dim vQuestion As Integer

On Error GoTo Errdescription

vDocNo = Trim(TXT211.Text)

If vDocNo <> "" Then
        If Opt211.Value = True Then
                vQuery = "select docno,StatementState from BCAPInvoice where docno = '" & vDocNo & "' "
        ElseIf Opt212.Value = True Then
                vQuery = "select docno,StatementState from BCAPOTHERDEBT where docno = '" & vDocNo & "' "
        ElseIf Opt213.Value = True Then
                vQuery = "select docno,StatementState from BCStkRefund where docno = '" & vDocNo & "' "
        End If
                                If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                                    vConfirm = Trim(vRecordset.Fields("StatementState").Value)
                                Else
                                MsgBox "เอกสารเลขที่  " & vDocNo & "  ไม่มีในระบบครับ กรุณาใส่เงื่อนไขใหม่ครับ", vbInformation + vbCritical, "ข้อความเตือน"
                                Exit Sub
                                End If
        vRecordset.Close
                                                If vConfirm = 1 Then
                                                    vQuestion = MsgBox("คุณต้องการยกเลิกการจ่ายชำระเลขที่เอกสาร  " & vDocNo & "   นี้ใช่หรือไม่", vbYesNo, "ข้อความสอบถาม")
                                                            If vQuestion = 6 Then
                                                                        If Opt211.Value = True Then
                                                                                vQuery = "Update BCAPInvoice  Set StatementState = 0 where docno = '" & vDocNo & "' "
                                                                        ElseIf Opt212.Value = True Then
                                                                                vQuery = "Update BCAPOTHERDEBT Set StatementState = 0 where docno = '" & vDocNo & "' "
                                                                        ElseIf Opt213.Value = True Then
                                                                                vQuery = "Update BCStkRefund Set StatementState = 0 where docno = '" & vDocNo & "' "
                                                                        End If
                                                                gConnection.Execute vQuery
                                                                MsgBox "เอกสารเลขที่  " & vDocNo & " ได้ทำการยกเลิกการจ่ายชำระเรียบร้อยแล้ว ", vbInformation, "ข้อความแจ้งให้ทราบ"
                                                            Else
                                                                Exit Sub
                                                            End If
                                            Else
                                                MsgBox "เอกสารเลขที่  " & vDocNo & "  ยังไม่ได้จ่ายชำระครับ", vbInformation, "ข้อความแจ้งให้ทราบ"
                                            End If
End If

Errdescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

