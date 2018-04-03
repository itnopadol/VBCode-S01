VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form42 
   Caption         =   "ทำใบวางบิลเจ้าหนี้ชั่วคราว"
   ClientHeight    =   8415
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10425
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form42.frx":0000
   ScaleHeight     =   8415
   ScaleWidth      =   10425
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CMD101 
      Caption         =   "เลือกเอกสารทำใบวางบิล"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   750
      TabIndex        =   17
      Top             =   3000
      Width           =   1965
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3015
      Left            =   750
      TabIndex        =   16
      Top             =   3600
      Width           =   8940
      _ExtentX        =   15769
      _ExtentY        =   5318
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "เลขที่เอกสาร"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "วันที่เอกสาร"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "วันที่ครบกำหนด"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "หมายเหตุ"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "ยอดคงเหลือ"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "มูลค่าวางบิล"
         Object.Width           =   2646
      EndProperty
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   315
      Left            =   7950
      TabIndex        =   2
      Top             =   975
      Width           =   1740
      _ExtentX        =   3069
      _ExtentY        =   556
      _Version        =   393216
      Format          =   20643841
      CurrentDate     =   38594
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   5025
      TabIndex        =   1
      Top             =   975
      Width           =   1740
      _ExtentX        =   3069
      _ExtentY        =   556
      _Version        =   393216
      Format          =   20643841
      CurrentDate     =   38594
   End
   Begin VB.TextBox Text6 
      Height          =   315
      Left            =   1800
      TabIndex        =   3
      Top             =   1425
      Width           =   1740
   End
   Begin VB.TextBox Text5 
      Height          =   315
      Left            =   1800
      TabIndex        =   7
      Top             =   2325
      Width           =   1740
   End
   Begin VB.TextBox Text4 
      Height          =   315
      Left            =   5025
      TabIndex        =   6
      Top             =   1875
      Width           =   1740
   End
   Begin VB.TextBox Text3 
      Height          =   315
      Left            =   1800
      TabIndex        =   5
      Top             =   1875
      Width           =   1740
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   5025
      TabIndex        =   4
      Top             =   1425
      Width           =   1740
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   1800
      TabIndex        =   0
      Top             =   975
      Width           =   1740
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "หมายเหตุ"
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
      Left            =   750
      TabIndex        =   15
      Top             =   2325
      Width           =   840
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "รหัสแผนก"
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
      Left            =   4050
      TabIndex        =   14
      Top             =   1875
      Width           =   915
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "วันที่นัดชำระ"
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
      Left            =   6900
      TabIndex        =   13
      Top             =   975
      Width           =   1065
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "เครดิต"
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
      Left            =   750
      TabIndex        =   12
      Top             =   1875
      Width           =   690
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "ชื่อเจ้าหนี้"
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
      Left            =   4050
      TabIndex        =   11
      Top             =   1425
      Width           =   840
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "รหัสเจ้าหนี้"
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
      Left            =   750
      TabIndex        =   10
      Top             =   1425
      Width           =   990
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "วันที่วางบิล"
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
      Left            =   4050
      TabIndex        =   9
      Top             =   975
      Width           =   915
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "เลขที่เอกสาร"
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
      Left            =   750
      TabIndex        =   8
      Top             =   975
      Width           =   1365
   End
End
Attribute VB_Name = "Form42"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub TextLine_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    
End If
End Sub

'TextLine(0).Left = 1500
'TextLine(0).Top = 3600
'TextLine(0).Width = 1665
'TextLine(0).Height = 315
'TextLine(0).TabIndex = 8

'For i = 1 To 4
 '   If i < 5 Then
  '      Load TextLine(i)
   '     Set TextLine(i).Container = Form42
    '    TextLine(i).Visible = True
     '   TextLine(i).Left = TextLine(i - 1).Left + 1590
      ''  TextLine(i).Top = TextLine(0).Top
        'TextLine(i).Width = TextLine(0).Width
        'TextLine(i).Height = TextLine(0).Height
        'TextLine(i).Appearance = 0
        'TextLine(i).BackColor = "&H80000018"
        'TextLine(i).TabIndex = TextLine(0).TabIndex + i
    'Else
     '   Load TextLine(i)
      '  Set TextLine(i).Container = Form42
       ' TextLine(i).Visible = True
        ''TextLine(i).Left = TextLine(i - 1).Left + 1590
        'TextLine(i).Top = TextLine(0).Top
        'TextLine(i).Width = TextLine(0).Width
        'TextLine(i).Height = TextLine(0).Height
        'TextLine(i).Appearance = 0
        'TextLine(i).BackColor = "&H80000018"
        'TextLine(i).TabIndex = TextLine(0).TabIndex + i
    'End If
'Next i
