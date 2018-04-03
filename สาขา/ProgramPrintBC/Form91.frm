VERSION 5.00
Begin VB.Form Form91 
   Caption         =   "ยกเลิกการผ่านบัญชี"
   ClientHeight    =   8370
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12000
   Icon            =   "Form91.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form91.frx":08CA
   ScaleHeight     =   8370
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CMD911 
      Caption         =   "กดยกเลิกการผ่านบัญชี"
      Height          =   615
      Left            =   4275
      TabIndex        =   3
      Top             =   2100
      Width           =   1890
   End
   Begin VB.TextBox TXT911 
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
      Height          =   465
      Left            =   3450
      TabIndex        =   2
      Top             =   1275
      Width           =   2715
   End
   Begin VB.Label LBL912 
      BackStyle       =   0  'Transparent
      Caption         =   "เลขที่เอกสารที่ต้องการยกเลิกการผ่านบัญชี"
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
      Height          =   465
      Left            =   300
      TabIndex        =   1
      Top             =   1275
      Width           =   2865
   End
   Begin VB.Label LBL911 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ยกเลิกการผ่านบัญชี"
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
      Top             =   300
      Width           =   7665
   End
End
Attribute VB_Name = "Form91"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMD911_Click()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vDocno As String
Dim vQuestion As Integer
Dim vSource As Integer
Dim vTable As String
Dim vFields As String
Dim vIspostGL As Integer

On Error GoTo ErrDescription

vDocno = Trim(TXT911.Text)
vQuery = "select source from bctrans where docno = '" & vDocno & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vSource = Trim(vRecordset.Fields("source").Value)
End If
vRecordset.Close

vQuery = "select * from np_gl_source where source = " & vSource & " "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vTable = Trim(vRecordset.Fields("tablename").Value)
    vFields = Trim(vRecordset.Fields("glfields").Value)
End If
vRecordset.Close


If vTable <> "" Then
    vQuery = "select " & vFields & " from " & vTable & " where docno = '" & vDocno & "' "
            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                vIspostGL = Trim(vRecordset.Fields("" & vFields & "").Value)
            End If
    vRecordset.Close

        If vIspostGL = 0 Then
            MsgBox "เอกสารเลขที่   " & vDocno & "  ยังไม่ได้ผ่านบัญชี คุณไม่สามารถยกเลิกการผ่านบัญชีได้ครับ ", vbInformation, "  ข้อความแจ้งให้ทราบ"
        ElseIf vIspostGL = 1 Then
            MsgBox "เอกสารเลขที่   " & vDocno & "  ได้ผ่านบัญชีไปแล้ว ", vbInformation, "  ข้อความแจ้งให้ทราบ"
            vQuestion = MsgBox("คุณต้องการยกเลิกการผ่านบัญชีเลขที่เอกสารนี้ใช่หรือไม่", vbCritical + vbYesNo, "ข้อความสอบถาม")
                If vQuestion = 6 Then
                    vQuery = "execute Usp_GL_Cancel '" & vDocno & "' "
                    gConnection.Execute vQuery
                    MsgBox "เลขที่เอกสาร       " & vDocno & "     ได้ถูกยกเลิกการผ่านบัญชีเรียบร้อยแล้ว", vbInformation, "ข้อความแจ้งให้ทราบ"
                Else
                    Exit Sub
                End If
            End If
Else
     MsgBox "เอกสารเลขที่   " & vDocno & "  ไม่มีในระบบ", vbInformation, "  ข้อความแจ้งให้ทราบ"

End If


TXT911.Text = ""

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub
