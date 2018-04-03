VERSION 5.00
Begin VB.Form Form102_1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "แก้ไขจำนวน"
   ClientHeight    =   4725
   ClientLeft      =   5970
   ClientTop       =   2670
   ClientWidth     =   6840
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form102_1.frx":0000
   ScaleHeight     =   4725
   ScaleWidth      =   6840
   Begin VB.CommandButton Command102 
      Caption         =   "ลบรายการ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   5310
      TabIndex        =   7
      Top             =   3195
      Width           =   1215
   End
   Begin VB.CommandButton Command101 
      Caption         =   "แก้ไข"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   4005
      TabIndex        =   4
      Top             =   3195
      Width           =   1215
   End
   Begin VB.TextBox Text101 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   29.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   810
      Left            =   3435
      MaxLength       =   5
      TabIndex        =   3
      Top             =   2250
      Width           =   3075
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "จำนวนที่นับได้"
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
      Height          =   315
      Left            =   0
      TabIndex        =   10
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label103 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000E&
      Height          =   315
      Left            =   1275
      TabIndex        =   9
      Top             =   1800
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "(กรณี ที่ไม่ใส่ยอดแก้ไข กด Enter ได้เลย)"
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
      Height          =   315
      Left            =   135
      TabIndex        =   8
      Top             =   2565
      Width           =   3165
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ชื่อสินค้า"
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
      Height          =   315
      Left            =   525
      TabIndex        =   6
      Top             =   1350
      Width           =   690
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "รหัสสินค้า"
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
      Height          =   315
      Left            =   450
      TabIndex        =   5
      Top             =   900
      Width           =   765
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "จำนวนที่ต้องการแก้ไข"
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
      Height          =   315
      Left            =   1530
      TabIndex        =   2
      Top             =   2250
      Width           =   1740
   End
   Begin VB.Label Label102 
      BackColor       =   &H80000009&
      Height          =   315
      Left            =   1275
      TabIndex        =   1
      Top             =   1350
      Width           =   5265
   End
   Begin VB.Label Label101 
      BackColor       =   &H80000009&
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
      Left            =   1275
      TabIndex        =   0
      Top             =   900
      Width           =   2190
   End
End
Attribute VB_Name = "Form102_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command101_Click()
On Error Resume Next

    Form102.Enabled = True
    If Text101.Text <> "" Then
    Form102.ListView101.ListItems(vItemClick).ListSubItems(2).Text = Format(Trim(Form102_1.Text101.Text), "##,##0.00")
    Form102.ListView101.ListItems(vItemClick).ListSubItems(3).Text = Format(Trim(Form102_1.Text101.Text) - Form102.ListView101.ListItems(vItemClick).ListSubItems(1).Text, "##,##0.00")
    End If
    Label101.Caption = ""
    Label102.Caption = ""
    Text101.Text = ""
    Form102_1.Hide
    Form102.Text101.SetFocus
    Unload Form102_1

End Sub

Private Sub Command102_Click()
Form102.Enabled = True
Form102.Text101.SetFocus
Unload Form102_1
Form102.ListView101.ListItems.Remove (vItemClick)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Unload Form102_1
Form102.Enabled = True
Form102.Text101.SetFocus
End Sub


Private Sub Text101_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Len(Text101.Text) < 7 Then
        Call Command101_Click
    Else
        MsgBox "จำนวนสินค้ามากเกินกว่าความเป็นจริง กรุณาตรวจสอบ"
        Exit Sub
    End If
End If
End Sub
