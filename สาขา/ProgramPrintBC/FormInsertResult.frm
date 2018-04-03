VERSION 5.00
Begin VB.Form FormInsertResult 
   Caption         =   "กรอกข้อมูล การขนส่งสินค้า"
   ClientHeight    =   8055
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "FormInsertResult.frx":0000
   ScaleHeight     =   8055
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin VB.OptionButton Option101 
      BackColor       =   &H8000000E&
      Caption         =   "1.สมบูรณ์"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2790
      TabIndex        =   13
      Top             =   4185
      Width           =   2715
   End
   Begin VB.OptionButton Option102 
      BackColor       =   &H8000000E&
      Caption         =   "2.ไม่ต้องเก็บเงิน"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2790
      TabIndex        =   12
      Top             =   4635
      Width           =   2715
   End
   Begin VB.OptionButton Option103 
      BackColor       =   &H8000000E&
      Caption         =   "3.พนักงานขับรถไม่ได้รับแจ้ง"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2790
      TabIndex        =   11
      Top             =   5085
      Width           =   2715
   End
   Begin VB.OptionButton Option104 
      BackColor       =   &H8000000E&
      Caption         =   "4.หาคนชำระเงินไม่ได้"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2790
      TabIndex        =   10
      Top             =   5535
      Width           =   2715
   End
   Begin VB.OptionButton Option105 
      BackColor       =   &H8000000E&
      Caption         =   "5.อื่น ๆ "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2790
      TabIndex        =   9
      Top             =   5985
      Width           =   2715
   End
   Begin VB.CommandButton CMD101 
      Caption         =   "ตกลง"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4410
      TabIndex        =   8
      Top             =   6435
      Width           =   1095
   End
   Begin VB.TextBox Text104 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   2790
      TabIndex        =   7
      Top             =   2655
      Width           =   1815
   End
   Begin VB.TextBox Text103 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
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
      Left            =   2790
      TabIndex        =   5
      Top             =   2160
      Width           =   1815
   End
   Begin VB.TextBox Text102 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
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
      Left            =   2790
      TabIndex        =   3
      Top             =   1665
      Width           =   1815
   End
   Begin VB.TextBox Text101 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
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
      Left            =   2790
      TabIndex        =   0
      Top             =   1170
      Width           =   1815
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ผลการขนส่ง :"
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
      TabIndex        =   14
      Top             =   3870
      Width           =   1635
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "จำนวนเงินที่เก็บได้ :"
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
      Left            =   900
      TabIndex        =   6
      Top             =   2655
      Width           =   1770
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "จำนวนเงินที่ต้องเก็บ :"
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
      Left            =   900
      TabIndex        =   4
      Top             =   2160
      Width           =   1770
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "เลขที่ใบจัดคิว :"
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
      Left            =   1440
      TabIndex        =   2
      Top             =   1665
      Width           =   1230
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "เลขที่ใบขนส่ง :"
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
      Left            =   990
      TabIndex        =   1
      Top             =   1170
      Width           =   1680
   End
End
Attribute VB_Name = "FormInsertResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMD101_Click()
Dim i As Integer
Dim vHoardDate As Date

On Error GoTo ErrDescription

If Text104.Text <> "" Then
    If Option101.Value <> False Or Option102.Value <> False Or Option103.Value <> False Or Option104.Value <> False Or Option105.Value <> False Then
        'If Option201.Value <> False Or Option202.Value <> False Or Option203.Value <> False Or Option204.Value <> False Or Option205.Value <> False Or Option206.Value <> False Then
            i = FormSendResult.ListView101.SelectedItem.Index
            'vHoardDate = DTPicker101.Day & "/" & DTPicker101.Month & "/" & DTPicker101.Year
            FormSendResult.ListView101.ListItems.Item(i).SubItems(4) = Trim(Text104.Text)
            'FormSendResult.ListView101.ListItems.Item(i).SubItems(5) = vHoardDate
            If Option101.Value = True Then
                FormSendResult.ListView101.ListItems.Item(i).SubItems(4) = 0
            ElseIf Option102.Value = True Then
                FormSendResult.ListView101.ListItems.Item(i).SubItems(4) = 1
            ElseIf Option103.Value = True Then
                FormSendResult.ListView101.ListItems.Item(i).SubItems(4) = 2
            ElseIf Option104.Value = True Then
                FormSendResult.ListView101.ListItems.Item(i).SubItems(4) = 3
            ElseIf Option105.Value = True Then
                FormSendResult.ListView101.ListItems.Item(i).SubItems(4) = 4
            End If
            'If Option201.Value = True Then
             '   FormSendResult.ListView101.ListItems.Item(i).SubItems(5) = 0
            'ElseIf Option202.Value = True Then
             '   FormSendResult.ListView101.ListItems.Item(i).SubItems(5) = 1
            'ElseIf Option203.Value = True Then
             '   FormSendResult.ListView101.ListItems.Item(i).SubItems(5) = 2
            'ElseIf Option204.Value = True Then
             '   FormSendResult.ListView101.ListItems.Item(i).SubItems(5) = 3
            'ElseIf Option205.Value = True Then
             '   FormSendResult.ListView101.ListItems.Item(i).SubItems(5) = 4
            'ElseIf Option206.Value = True Then
             '   FormSendResult.ListView101.ListItems.Item(i).SubItems(5) = 5
            'End If
            
            Unload FormInsertResult
            FormSendResult.Enabled = True
        'Else
         '   MsgBox "กรุณา เลือกผลการขนส่งด้วย", vbInformation, "Send Information"
          '  Exit Sub
        'End If
    Else
        MsgBox "กรุณา เลือกผลการเก็บเงินด้วย", vbInformation, "Send Information"
        Exit Sub
    End If
Else
    MsgBox "กรุณา กรอกจำนวนเงินที่เก็บมาด้วย", vbInformation, "Send Information"
    Text104.SetFocus
    Exit Sub
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Private Sub Form_Load()
'DTPicker101 = FormDelivery.DTPicker102.Value
End Sub
