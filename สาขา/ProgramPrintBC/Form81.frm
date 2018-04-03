VERSION 5.00
Begin VB.Form Form81 
   Caption         =   "หน้ารวมรหัสสินค้า"
   ClientHeight    =   8340
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form81.frx":0000
   ScaleHeight     =   8340
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CMD811 
      Caption         =   "รวมรหัส"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3825
      TabIndex        =   2
      Top             =   3600
      Width           =   1440
   End
   Begin VB.TextBox TXT812 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   2850
      TabIndex        =   1
      Top             =   2775
      Width           =   2415
   End
   Begin VB.TextBox TXT811 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   2850
      TabIndex        =   0
      Top             =   1950
      Width           =   2415
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "รหัสนี้จะคงอยู่ในระบบ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2880
      TabIndex        =   9
      Top             =   2475
      Width           =   3840
   End
   Begin VB.Label LBLItemName2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   5355
      TabIndex        =   8
      Top             =   2790
      Width           =   4650
   End
   Begin VB.Label LBLItemName1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   5355
      TabIndex        =   7
      Top             =   1950
      Width           =   4695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "รหัสนี้จะไปเป็นรหัสข้างล่างแทน"
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
      Height          =   300
      Left            =   2880
      TabIndex        =   6
      Top             =   1620
      Width           =   3315
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "หน้ารวมรหัสสินค้า"
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
      TabIndex        =   5
      Top             =   300
      Width           =   7515
   End
   Begin VB.Label LBL812 
      BackStyle       =   0  'Transparent
      Caption         =   "รหัสสินค้าที่ 2 :"
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
      Height          =   315
      Left            =   1710
      TabIndex        =   4
      Top             =   2790
      Width           =   1065
   End
   Begin VB.Label LBL811 
      BackStyle       =   0  'Transparent
      Caption         =   "รหัสสินค้าที่ 1 :"
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
      Height          =   315
      Left            =   1710
      TabIndex        =   3
      Top             =   1980
      Width           =   1140
   End
End
Attribute VB_Name = "Form81"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMD811_Click()
Dim vItem1 As String, vItem2 As String
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vQuestion As Integer

On Error GoTo ErrDescription

vItem1 = Trim(TXT811.Text)
vItem2 = Trim(TXT812.Text)

If Me.LBLItemName1.Caption = "" Or Me.LBLItemName2.Caption = "" Then
MsgBox "กรุณาตรวจสอบสินค้าที่จะรวมรหัส ", vbCritical, "Send Error Message"
Exit Sub
End If

If vItem1 <> "" And vItem2 <> "" Then
                    vQuery = "select code from bcitem where code = '" & vItem1 & "' "
                    If OpenDataBase(gConnection, vRecordset, vQuery) = 0 Then
                        MsgBox "รหัสสินค้ารหัส   " & vItem1 & " ไม่มีในระบบ กรุณาตรวจสอบด้วยครับ ", vbInformation + vbCritical, "ข้อความเตือน"
                        Exit Sub
                    End If
                    vRecordset.Close
                    
                    vQuery = "select code from bcitem where code = '" & vItem2 & "' "
                    If OpenDataBase(gConnection, vRecordset, vQuery) = 0 Then
                        MsgBox "รหัสสินค้ารหัส   " & vItem2 & " ไม่มีในระบบ กรุณาตรวจสอบด้วยครับ ", vbInformation + vbCritical, "ข้อความเตือน"
                        Exit Sub
                    End If
                    vRecordset.Close
                    
                    If vItem1 = vItem2 Then
                        MsgBox "สินค้ารหัสเดียวกันไม่สามารถใช้งานโปรแกรมนี้ได้ กรุณาตรวจสอบด้วยครับ ", vbInformation + vbCritical, "ข้อความเตือน"
                        Exit Sub
                    End If
                    
                    vQuestion = MsgBox(" การรวมรหัสสินค้า 2 ตัวนี้ รหัสตัวที่1 จะเปลี่ยนมาใช้รหัสที่ 2 แทน คุณต้องการรวมรหัสสินค้าสองตัวนี้ใช่หรือไม่", vbYesNo, "ข้อความสอบถาม")
                    If vQuestion = 6 Then
                    vQuery = "execute USP_IV_ItmMerge '" & vItem1 & "','" & vItem2 & "' "
                    gConnection.Execute vQuery
                    MsgBox "โปรแกรม รวมรหัสสินค้าได้รวมรหัสสินค้าเรียบร้อยแล้วครับ"
                    Else
                    Exit Sub
                    End If
Else
    MsgBox "กรุณากรอกข้อมูลให้ครบด้วยครับ", vbInformation, "ข้อความแจ้งเตือน"
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub TXT811_Change()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vItemCode As String

If Me.TXT811.Text <> "" Then
vItemCode = Me.TXT811.Text
vQuery = "exec dbo.USP_NP_ItemDataDetails1 '" & vItemCode & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    Me.LBLItemName1.Caption = vRecordset.Fields("name1").Value
Else
    Me.LBLItemName1.Caption = ""
End If
vRecordset.Close
Else
    Me.LBLItemName1.Caption = ""
End If

End Sub

Private Sub TXT812_Change()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vItemCode As String

If Me.TXT812.Text <> "" Then
vItemCode = Me.TXT812.Text
vQuery = "exec dbo.USP_NP_ItemDataDetails1 '" & vItemCode & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    Me.LBLItemName2.Caption = vRecordset.Fields("name1").Value
Else
    Me.LBLItemName2.Caption = ""
End If
vRecordset.Close
Else
    Me.LBLItemName2.Caption = ""
End If

End Sub
