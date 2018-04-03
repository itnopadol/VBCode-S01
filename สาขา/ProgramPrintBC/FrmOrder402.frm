VERSION 5.00
Begin VB.Form FrmOrder402 
   Caption         =   "คำนวณยอดขนส่งใบจัดคิวสินค้า"
   ClientHeight    =   8040
   ClientLeft      =   3225
   ClientTop       =   900
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "FrmOrder402.frx":0000
   ScaleHeight     =   8040
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CMD102 
      Caption         =   "ออก"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7200
      TabIndex        =   5
      Top             =   3060
      Width           =   870
   End
   Begin VB.CommandButton CMD101 
      Caption         =   "คำนวณ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6075
      TabIndex        =   4
      Top             =   3060
      Width           =   870
   End
   Begin VB.TextBox Text101 
      Appearance      =   0  'Flat
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
      Left            =   3960
      TabIndex        =   3
      Top             =   2115
      Width           =   1815
   End
   Begin VB.ComboBox CMB101 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3960
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1530
      Width           =   4065
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "เลขที่อ้างอิง"
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
      Left            =   2700
      TabIndex        =   2
      Top             =   2115
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ประเภทเอกสาร"
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
      Left            =   2700
      TabIndex        =   0
      Top             =   1530
      Width           =   1230
   End
End
Attribute VB_Name = "FrmOrder402"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMB101_Change()
Text101.SetFocus
End Sub

Private Sub CMB101_Click()
Text101.SetFocus
End Sub

Private Sub CMD101_Click()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vDocument As String
Dim vCheckDocNo As String
Dim vFix As Integer

On Error GoTo ErrDescription

If Text101.Text <> "" Then
    vDocument = Trim(Text101.Text)
        If CMB101.ListIndex = 0 Then
            vCheckDocNo = UCase(Left(vDocument, 2))
            If vCheckDocNo <> Trim("QR") And vCheckDocNo <> Trim("DR") Then
                MsgBox "เลือกประเภทการคำนวณไม่ตรงกับเลขที่เอกสาร", vbInformation, "Send Message"
                Exit Sub
            End If
            vQuery = "exec bcnp.dbo.USP_DO_QueueRemainCal '" & vDocument & "' "
            gConnection.Execute vQuery
        ElseIf CMB101.ListIndex = 1 Then
            vFix = 1
            vQuery = "exec bcnp.dbo.USP_DO_QueueRemainCal '" & vDocument & "' ," & vFix & " "
            gConnection.Execute vQuery
        End If
        MsgBox "คำนวณข้อมูลเสร็จแล้วครับ", vbInformation, "Send Information"
        Text101.Text = ""
        Text101.SetFocus
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Private Sub CMD102_Click()
Unload FrmOrder402
End Sub

Private Sub Form_Load()
CMB101.AddItem Trim("คำนวณ จากเลขที่ใบจัดคิว")
CMB101.AddItem Trim("คำนวณ จากเลขที่อ้างอิงรายตัวสินค้า")
End Sub

