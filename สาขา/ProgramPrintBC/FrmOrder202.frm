VERSION 5.00
Begin VB.Form FrmOrder202 
   Caption         =   "Form202 กำหนด ผู้รับสินค้า"
   ClientHeight    =   8055
   ClientLeft      =   1350
   ClientTop       =   645
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "FrmOrder202.frx":0000
   ScaleHeight     =   8055
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text105 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   3150
      TabIndex        =   5
      Top             =   3375
      Width           =   2625
   End
   Begin VB.TextBox Text104 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   3150
      TabIndex        =   4
      Top             =   2925
      Width           =   2625
   End
   Begin VB.CommandButton CMD105 
      Height          =   330
      Left            =   4950
      Picture         =   "FrmOrder202.frx":72FB
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "ออก"
      Top             =   5850
      Width           =   330
   End
   Begin VB.CommandButton CMD104 
      Height          =   330
      Left            =   4500
      Picture         =   "FrmOrder202.frx":7665
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "ยกเลิกข้อมูล"
      Top             =   5850
      Width           =   330
   End
   Begin VB.CommandButton CMD103 
      Height          =   330
      Left            =   4050
      Picture         =   "FrmOrder202.frx":9803
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "ค้นหาข้อมูล"
      Top             =   5850
      Width           =   330
   End
   Begin VB.CommandButton CMD102 
      Height          =   330
      Left            =   3600
      Picture         =   "FrmOrder202.frx":9BD0
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "บันทึกและปรับปรุงข้อมูล"
      Top             =   5850
      Width           =   330
   End
   Begin VB.CommandButton CMD101 
      Height          =   330
      Left            =   3150
      Picture         =   "FrmOrder202.frx":9EF7
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "เคลียร์ หน้าจอ"
      Top             =   5850
      Width           =   330
   End
   Begin VB.TextBox Text106 
      Appearance      =   0  'Flat
      Height          =   1545
      Left            =   3150
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   3825
      Width           =   5910
   End
   Begin VB.TextBox Text103 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   3150
      TabIndex        =   3
      Top             =   2475
      Width           =   2625
   End
   Begin VB.ComboBox CMB101 
      Height          =   315
      Left            =   3150
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   1575
      Width           =   1365
   End
   Begin VB.TextBox Text102 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   3150
      TabIndex        =   2
      Top             =   2025
      Width           =   2625
   End
   Begin VB.TextBox Text101 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   330
      Left            =   3150
      TabIndex        =   0
      Top             =   1125
      Width           =   1365
   End
   Begin VB.Image Image102 
      Height          =   300
      Left            =   3150
      Picture         =   "FrmOrder202.frx":A2DC
      Top             =   630
      Width           =   570
   End
   Begin VB.Image Image101 
      Height          =   300
      Left            =   3150
      Picture         =   "FrmOrder202.frx":A818
      Top             =   630
      Width           =   570
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "หมายเหตุ :"
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
      Left            =   945
      TabIndex        =   18
      Top             =   3825
      Width           =   2130
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "โทรศัพท์มือถือ :"
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
      Left            =   1035
      TabIndex        =   17
      Top             =   3375
      Width           =   2040
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "โทรศัพท์บ้าน :"
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
      Height          =   285
      Left            =   1035
      TabIndex        =   16
      Top             =   2925
      Width           =   1995
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "นามสกุล :"
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
      Height          =   285
      Left            =   1125
      TabIndex        =   15
      Top             =   2475
      Width           =   1905
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ชื่อ :"
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
      Left            =   1215
      TabIndex        =   14
      Top             =   2025
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "คำนำหน้าชื่อ :"
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
      Left            =   1260
      TabIndex        =   13
      Top             =   1575
      Width           =   1770
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ID :"
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
      Height          =   285
      Left            =   1305
      TabIndex        =   12
      Top             =   1125
      Width           =   1725
   End
End
Attribute VB_Name = "FrmOrder202"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMB101_Change()
'vCheckChageDataReceive = 1
End Sub

Private Sub CMD101_Click()
On Error GoTo ErrDescription

'Text101.Text = ""
'Text102.Text = ""
'Text103.Text = ""
'Text104.Text = ""
'MaskEdBox101.Text = "000-000000"
'MaskEdBox102.Text = "00-000-0000"
'Image101.Visible = True
'image102.Visible = False
'vCheckReceiveOpen = 0

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Private Sub CMD102_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vPriority As String
Dim vMydescription As String
Dim vIsCancel As String
Dim vID As String
Dim vTitleName As String
Dim vFirstName As String
Dim vSurName As String
Dim vHomePhone As String
Dim vMobilePhone As String
Dim vActiveStatus As String
    
On Error GoTo ErrDescription

'If Text102.Text <> "" And Text103.Text <> "" Then
    'vTitleName = Trim(CMB101.Text)
    'vFirstName = Trim(Text102.Text)
    'vSurName = Trim(Text103.Text)
    'vHomePhone = Trim(MaskEdBox101.Text)
    'vMobilePhone = Trim(MaskEdBox102.Text)
    'vActiveStatus = 1
    'vMydescription = Trim(Text104.Text)
    
    'If vCheckReceiveOpen = 0 Then
     '   Call CheckIsReceive
      '  If vCheckIsReceive = 1 Then
       '     MsgBox "ชื่อผู้รับสินค้า  " & vFirstName & "  " & vSurName & "   มีอยู่แล้วกรุณาตรวจสอบ", vbCritical, "Send Error"
        '    Exit Sub
        'End If
        'vID = "Null"
    'Else
     '   vID = Trim(Text101.Text)
    'End If
    'vQuery = "exec bcnp.dbo.USP_DO_ReceiveUpdate " & vID & ",'" & vTitleName & "','" & vFirstName & "'," _
                        & " '" & vSurName & "','" & vHomePhone & "','" & vMobilePhone & "','" & vActiveStatus & "','" & vUserID & "' ,'" & vMydescription & "' "
    'gConnection.Execute vQuery
    'If vCheckReceiveOpen = 0 Then
     '   MsgBox "บันทึกข้อมูลผู้รับสินค้า " & vTitleName & " " & vFirstName & "  " & vSurName & "    เรียบร้อยแล้ว", vbInformation, "Send Message"
      '  If vCheckAddReceiver = 1 Then
            'FrmOrder202.Hide
            'FrmOrder007.SetFocus
            'vCheckAddReceiver = 0
        'End If
    'ElseIf vCheckReceiveOpen = 1 Then
     '   MsgBox "บันทึกการแก้ไขข้อมูลผู้รับสินค้า " & vTitleName & " " & vFirstName & "  " & vSurName & "    เรียบร้อยแล้ว", vbInformation, "Send Message"
    'End If

     '   Text101.Text = ""
      '  Text102.Text = ""
       ' Text103.Text = ""
        'Text104.Text = ""
        'MaskEdBox101.Text = "000-000000"
        'MaskEdBox102.Text = "00-000-0000"
        'Image101.Visible = True
        'Image102.Visible = False
        'vCheckReceiveOpen = 0
        
        'MsgBox "บันทึกข้อมูลเรียบร้อยแล้วครับ", vbInformation, "Send Information"
            
'Else
 '   MsgBox "ต้องใส่ข้อมูล ชื่อและนามสกุล ให้ครบถึงจะบันทึกข้อมูลได้", vbInformation, "Send Information"
'End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Private Sub CMD103_Click()
On Error GoTo ErrDescription

'MDIFrmProgramPrint.Enabled = False
'FrmOrder007.Show
'vReceiveModule = 2

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If

End Sub

Private Sub CMD104_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vID As Integer
Dim vAnswer As Integer

On Error GoTo ErrDescription

'If Text101.Text <> "" Then
 '   If vCheckReceiveOpen = 1 Then
  '  vAnswer = MsgBox("คุณต้องการยกเลิก ระดับความสำคัญนี้ใช่หรือไม่", vbYesNo, "Question Respond")
   '     If vAnswer = 6 Then
    '        vID = Trim(Text101.Text)
     '       vQuery = "Update npmaster.dbo.TB_DO_Receive set activestatus = 0 where id = " & vID & " "
      '      gConnection.Execute vQuery
       '     Text101.Text = ""
        '    Text102.Text = ""
         '   Text103.Text = ""
          '  Text104.Text = ""
            'MaskEdBox101.Text = "000-000000"
            'MaskEdBox102.Text = "00-000-0000"
            'Image101.Visible = True
            'Image102.Visible = False
            'vCheckReceiveOpen = 0
        'Else
         '   Exit Sub
        'End If
    'End If
'End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Private Sub CMD105_Click()
'Unload FrmOrder202
End Sub

Private Sub Form_Load()
'CMB101.AddItem Trim("คุณ")
'CMB101.AddItem Trim("นาย")
'CMB101.AddItem Trim("น.ส.")
'CMB101.AddItem Trim("นาง")
'Image101.Visible = True
'Image102.Visible = False
End Sub

Public Sub CheckIsCancel()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vPriority As String
Dim vID As Integer

On Error GoTo ErrDescription

'vID = Trim(Text101.Text)
'vQuery = "select id,iscancel from npmaster.dbo.TB_DO_Receive Where id = '" & vID & "' "
'If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
 '   vCheckIsCancel1 = Trim(vRecordset.Fields("iscancel").Value)
'Else
 '   vCheckIsCancel1 = 0
'End If
'vRecordset.Close

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Public Sub CheckIsReceive()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vFirstName  As String
Dim vSurName As String

On Error GoTo ErrDescription

'vFirstName = Trim(Text102.Text)
'vSurName = Trim(Text103.Text)
'vQuery = "select id from npmaster.dbo.TB_DO_Receive Where firstname = '" & vFirstName & "' and surname = '" & vSurName & "' "
'If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
 '   vCheckIsReceive = 1
'Else
 '   vCheckIsReceive = 0
'End If
'vRecordset.Close

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Public Sub CheckIsReceiveUpDate()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vFirstName  As String
Dim vSurName As String

On Error GoTo ErrDescription

'vFirstName = Trim(Text102.Text)
'vSurName = Trim(Text103.Text)
'vQuery = "select id from npmaster.dbo.TB_DO_Receive Where firstname = '" & vFirstName & "' and surname = '" & vSurName & "' "
'If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
 '   vCheckIsReceive = 1
'Else
 '   vCheckIsReceive = 0
'End If
'vRecordset.Close

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Private Sub MaskEdBox101_Change()
'vCheckChageDataReceive = 1
End Sub

Private Sub MaskEdBox102_Change()
'vCheckChageDataReceive = 1
End Sub

Private Sub Text102_Change()
'vCheckChageDataReceive = 1
End Sub

Private Sub Text103_Change()
'vCheckChageDataReceive = 1
End Sub

Private Sub Text104_Change()
'vCheckChageDataReceive = 1
End Sub


