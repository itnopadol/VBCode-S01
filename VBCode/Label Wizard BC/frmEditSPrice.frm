VERSION 5.00
Begin VB.Form frmEditSPrice 
   Caption         =   "แก้ไขระดับราคาปกติ"
   ClientHeight    =   3480
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5220
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   5220
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ยกเลิก"
      Height          =   375
      Left            =   3720
      TabIndex        =   12
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "บันทึก"
      Height          =   375
      Left            =   2400
      TabIndex        =   11
      Top             =   3000
      Width           =   1095
   End
   Begin VB.TextBox txtIndex 
      Height          =   375
      Left            =   600
      TabIndex        =   6
      Text            =   "txtIndex"
      Top             =   3000
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   2535
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4695
      Begin VB.TextBox txtPrice 
         Height          =   420
         Left            =   2040
         TabIndex        =   10
         Text            =   "txtPrice"
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Frame Frame2 
         Height          =   2535
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   1815
         Begin VB.Label Label4 
            Caption         =   "ราคาปกติ :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   5
            Top             =   1850
            Width           =   1455
         End
         Begin VB.Label Label3 
            Caption         =   "ราคาพิเศษ :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   1440
            Width           =   1455
         End
         Begin VB.Label Label2 
            Caption         =   "ชื่อสินค้า :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   3
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "Item Number :"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   2
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.Label lbSPrice 
         Caption         =   "lbSPrice"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   9
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label lbItemDesc 
         Caption         =   "lbItemDesc"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2040
         TabIndex        =   8
         Top             =   840
         Width           =   2415
      End
      Begin VB.Label lbItemNumber 
         Caption         =   "lbItemNumber"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   7
         Top             =   360
         Width           =   2415
      End
   End
End
Attribute VB_Name = "frmEditSPrice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
        Unload Me
End Sub

Private Sub cmdSave_Click()
        'Dim ListX As ListItem
        If IsNumeric(txtPrice) = False Then
                MsgBox "กรุณาใส่จำนวนที่เป็นตัวเลข", vbInformation + vbOKOnly, "คำแนะนำ"
                Exit Sub
        End If
        
        ' Add ระดับราคาปกติกลับคืน
        'Set ListX = frmSPrice.LV_SPrice.ListItems(4).SubItems
        frmSPrice.LV_SPrice.ListItems(Int(txtIndex.Text)).SubItems(4) = Trim(txtPrice.Text)
        Unload Me
End Sub

Private Sub txtPrice_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
                Call cmdSave_Click
        End If
End Sub
