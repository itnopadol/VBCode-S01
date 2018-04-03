VERSION 5.00
Begin VB.Form Form39 
   Caption         =   "แก้ไขข้อมูลเอกสารขายที่ถูกอ้างอิงไปแล้ว"
   ClientHeight    =   8340
   ClientLeft      =   3210
   ClientTop       =   1320
   ClientWidth     =   12000
   Icon            =   "Form39.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form39.frx":08CA
   ScaleHeight     =   8340
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin VB.TextBox TextSend 
      Alignment       =   2  'Center
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
      Left            =   4545
      TabIndex        =   38
      Top             =   4815
      Width           =   1500
   End
   Begin VB.TextBox TextExpire 
      Alignment       =   2  'Center
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
      Left            =   4545
      TabIndex        =   15
      Top             =   4365
      Width           =   1500
   End
   Begin VB.TextBox TextValidaty 
      Alignment       =   2  'Center
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
      Left            =   4545
      TabIndex        =   16
      Top             =   3960
      Width           =   1500
   End
   Begin VB.TextBox TextSaleCode 
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
      Left            =   4545
      TabIndex        =   18
      Top             =   5715
      Width           =   1995
   End
   Begin VB.TextBox TextARCode 
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
      Left            =   4545
      TabIndex        =   17
      Top             =   5265
      Width           =   1995
   End
   Begin VB.ComboBox CMBAssert 
      Enabled         =   0   'False
      Height          =   315
      Left            =   9090
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   3960
      Width           =   1995
   End
   Begin VB.TextBox TextCredit 
      Alignment       =   2  'Center
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
      Left            =   4545
      TabIndex        =   14
      Top             =   3555
      Width           =   1500
   End
   Begin VB.ComboBox CMBDelivery 
      Enabled         =   0   'False
      Height          =   315
      Left            =   9090
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   3150
      Width           =   1995
   End
   Begin VB.ComboBox CMBSaleType 
      Enabled         =   0   'False
      Height          =   315
      Left            =   4545
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   3150
      Width           =   1995
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   7620
      Left            =   0
      ScaleHeight     =   7590
      ScaleWidth      =   1875
      TabIndex        =   22
      Top             =   1170
      Width           =   1905
   End
   Begin VB.CheckBox CHKSend 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "7.เปลี่ยนวันที่ส่งของภายใน"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4275
      TabIndex        =   7
      Top             =   4860
      Width           =   195
   End
   Begin VB.CheckBox CHKCredit 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "4.เปลี่ยนเครดิต"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4275
      TabIndex        =   4
      Top             =   3600
      Width           =   195
   End
   Begin VB.CheckBox CHKExpire 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "5.เปลี่ยนวันที่หมดอายุ"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4275
      TabIndex        =   5
      Top             =   4410
      Width           =   195
   End
   Begin VB.CheckBox CHKValidate 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "6.เปลี่ยนวันที่ยืนราคาของเอกสาร"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4275
      TabIndex        =   6
      Top             =   4005
      Width           =   195
   End
   Begin VB.CheckBox CHKAssert 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "8.เปลี่ยนสถานะการตอบกลับของลูกค้า"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   8820
      TabIndex        =   8
      Top             =   4005
      Width           =   195
   End
   Begin VB.CheckBox CHKSaleCode 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "9.เปลี่ยนรหัสพนักงานขาย"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4275
      TabIndex        =   9
      Top             =   5760
      Width           =   195
   End
   Begin VB.CheckBox CHKARCode 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "10.เปลี่ยนรหัสลูกค้า"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4275
      TabIndex        =   10
      Top             =   5310
      Width           =   195
   End
   Begin VB.CheckBox CHKDelivery 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "3.เปลี่ยนเงื่อนไขการขนส่ง"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   8820
      TabIndex        =   3
      Top             =   3195
      Width           =   195
   End
   Begin VB.CheckBox CHKSaleType 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "2.เปลี่ยนประเภทการขาย"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4275
      TabIndex        =   2
      Top             =   3195
      Width           =   195
   End
   Begin VB.CheckBox CHKConfirm 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "1.ยกเลิกการอนุมัติเอกสาร"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   4275
      TabIndex        =   1
      Top             =   2745
      Width           =   2940
   End
   Begin VB.CommandButton CMDProcess 
      Caption         =   "ประมวลผล"
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
      Height          =   435
      Left            =   10170
      TabIndex        =   19
      Top             =   6300
      Width           =   1365
   End
   Begin VB.TextBox TXTDocNo 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4545
      TabIndex        =   0
      Top             =   1890
      Width           =   1530
   End
   Begin VB.Label LBLSendDate 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   9090
      TabIndex        =   46
      Top             =   4815
      Width           =   1500
   End
   Begin VB.Label LBLExpiredate 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   9090
      TabIndex        =   45
      Top             =   4410
      Width           =   1500
   End
   Begin VB.Label LBLDuedate 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   9090
      TabIndex        =   44
      Top             =   3555
      Width           =   1500
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "วัน"
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
      Height          =   240
      Left            =   6120
      TabIndex        =   43
      Top             =   4815
      Width           =   330
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "วัน"
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
      Height          =   240
      Left            =   6120
      TabIndex        =   42
      Top             =   4365
      Width           =   330
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "วัน"
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
      Height          =   240
      Left            =   6120
      TabIndex        =   41
      Top             =   3960
      Width           =   375
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "วัน"
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
      Height          =   195
      Left            =   6120
      TabIndex        =   40
      Top             =   3555
      Width           =   375
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ส่งภายใน :"
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
      Height          =   240
      Left            =   3060
      TabIndex        =   39
      Top             =   4815
      Width           =   1140
   End
   Begin VB.Label LBLDocdate 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   4545
      TabIndex        =   37
      Top             =   2250
      Width           =   1530
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "วันที่เอกสาร :"
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
      Left            =   3150
      TabIndex        =   36
      Top             =   2250
      Width           =   1365
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "วันที่ครบกำหนด :"
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
      Left            =   7335
      TabIndex        =   35
      Top             =   3600
      Width           =   1410
   End
   Begin VB.Label Label7 
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
      Height          =   285
      Left            =   7065
      TabIndex        =   34
      Top             =   4410
      Width           =   1680
   End
   Begin VB.Label LBLARName 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   6930
      TabIndex        =   33
      Top             =   5265
      Width           =   4605
   End
   Begin VB.Label LBLSaleName 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   6930
      TabIndex        =   32
      Top             =   5715
      Width           =   4605
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ลูกค้า :"
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
      Left            =   3195
      TabIndex        =   31
      Top             =   5265
      Width           =   1005
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "พนักงานขาย :"
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
      Left            =   2070
      TabIndex        =   30
      Top             =   5715
      Width           =   2130
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "สถานะการตอบกลับ :"
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
      Left            =   6795
      TabIndex        =   29
      Top             =   4005
      Width           =   1950
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "วันที่นัดส่งของ :"
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
      Left            =   7110
      TabIndex        =   28
      Top             =   4815
      Width           =   1635
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ยืนราคา :"
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
      Left            =   2565
      TabIndex        =   27
      Top             =   3960
      Width           =   1635
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "หมดอายุภายใน :"
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
      Left            =   2565
      TabIndex        =   26
      Top             =   4365
      Width           =   1635
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "เครดิต :"
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
      Left            =   2565
      TabIndex        =   25
      Top             =   3555
      Width           =   1635
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "เงื่อนไขการขนส่ง :"
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
      Left            =   7110
      TabIndex        =   24
      Top             =   3150
      Width           =   1635
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ประเภทการขาย :"
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
      Left            =   2565
      TabIndex        =   23
      Top             =   3150
      Width           =   1635
   End
   Begin VB.Image IMG102 
      Height          =   300
      Left            =   11025
      Picture         =   "Form39.frx":7BC5
      Top             =   1305
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Image IMG101 
      Height          =   300
      Left            =   11025
      Picture         =   "Form39.frx":806E
      Top             =   1305
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Line Line1 
      BorderWidth     =   5
      Index           =   0
      X1              =   0
      X2              =   12060
      Y1              =   1125
      Y2              =   1125
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "แก้ไขข้อมูลเอกสารที่ถูกอ้างอิงไปแล้ว"
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
      Left            =   2700
      TabIndex        =   21
      Top             =   300
      Width           =   8910
   End
   Begin VB.Label LBL391 
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
      Height          =   300
      Left            =   3285
      TabIndex        =   20
      Top             =   1890
      Width           =   1215
   End
End
Attribute VB_Name = "Form39"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim vDocNo As String
Dim vQuery As String
Dim vCheckType As Integer
Dim vMemSaleType As Integer
Dim vMemDelivery As Integer
Dim vMemAssert As Integer
Dim vMemCredit As Integer
Dim vMemDueDate As Date
Dim vMemExpire As Integer
Dim vMemExpireDate As Date
Dim vMemValidaty As Integer
Dim vMemArCode As String
Dim vMemSaleCode As String
Dim vMemSend As Integer
Dim vMemSendDate As Date
Dim vMemArName As String
Dim vMemSaleName As String
Dim vMemDocDate As Date

Private Sub CHKARCode_Click()
On Error GoTo ErrDescription

If Me.IMG102.Visible = True And Me.LBLARName.Caption <> "" Then
   If Me.CHKARCode.Value = 1 Then
      Me.TextARCode.Enabled = True
      Me.TextARCode.SetFocus
   Else
      Me.TextARCode.Enabled = False
      Me.TextARCode.Text = vMemArCode
      Me.LBLARName.Caption = vMemArName
   End If
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub CHKAssert_Click()
On Error GoTo ErrDescription

If Me.IMG102.Visible = True And Me.LBLARName.Caption <> "" Then
   If Me.CHKAssert.Value = 1 Then
      Me.CMBAssert.Enabled = True
      Me.CMBAssert.SetFocus
   Else
      Me.CMBAssert.Enabled = False
      Me.CMBAssert.ListIndex = vMemAssert
   End If
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub CHKConfirm_Click()
On Error GoTo ErrDescription

If Me.IMG102.Visible = True And Me.LBLARName.Caption <> "" Then
   If vCheckType = 3 Then
      If Me.CHKConfirm.Value = 1 Then
         MsgBox "ใบสั่งจองสินค้า เมื่อยกเลิกการอนุมัติ เอกสารดังกล่าวจะถูกยกเลิกโดยอัตโนมัติ เนื่องจากเอกสารดังกล่าวเกี่ยวเนื่องกับใบมัดจำและใบโอนสินค้า", vbCritical, "Send Error Message"
      End If
   End If
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub CHKCredit_Click()
On Error GoTo ErrDescription

If Me.IMG102.Visible = True And Me.LBLARName.Caption <> "" Then
   If Me.CHKCredit.Value = 1 Then
      Me.TextCredit.Enabled = True
      Me.TextCredit.SetFocus
   Else
      Me.TextCredit.Enabled = False
      Me.TextCredit.Text = vMemCredit
   End If
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub CHKDelivery_Click()
On Error GoTo ErrDescription

If Me.IMG102.Visible = True And Me.LBLARName.Caption <> "" Then
   If Me.CHKDelivery.Value = 1 Then
      Me.CMBDelivery.Enabled = True
      Me.CMBDelivery.SetFocus
   Else
      Me.CMBDelivery.Enabled = False
      Me.CMBDelivery.ListIndex = vMemDelivery
   End If
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub CHKExpire_Click()
On Error GoTo ErrDescription

If Me.IMG102.Visible = True And Me.LBLARName.Caption <> "" Then
   If Me.CHKExpire.Value = 1 Then
      Me.TextExpire.Enabled = True
      Me.TextExpire.SetFocus
   Else
      Me.TextExpire.Enabled = False
      Me.TextExpire.Text = vMemExpire
   End If
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub CHKSaleCode_Click()
On Error GoTo ErrDescription

If Me.IMG102.Visible = True And Me.LBLARName.Caption <> "" Then
   If Me.CHKSaleCode.Value = 1 Then
      Me.TextSaleCode.Enabled = True
      Me.TextSaleCode.SetFocus
   Else
      Me.TextSaleCode.Enabled = False
      Me.TextSaleCode.Text = vMemSaleCode
      Me.LBLSaleName.Caption = vMemSaleName
   End If
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub CHKSaleType_Click()
On Error GoTo ErrDescription

If Me.IMG102.Visible = True And Me.LBLARName.Caption <> "" Then
   If Me.CHKSaleType.Value = 1 Then
      Me.CMBSaleType.Enabled = True
      Me.CMBSaleType.SetFocus
   Else
      Me.CMBSaleType.Enabled = False
      Me.CMBSaleType.ListIndex = vMemSaleType
   End If
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub CHKSend_Click()
On Error GoTo ErrDescription

If Me.IMG102.Visible = True And Me.LBLARName.Caption <> "" Then
   If Me.CHKSend.Value = 1 Then
      Me.TextSend.Enabled = True
      Me.TextSend.SetFocus
   Else
      Me.TextSend.Enabled = False
      Me.TextSend.Text = vMemSend
   End If
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub CHKValidate_Click()
On Error GoTo ErrDescription

If Me.IMG102.Visible = True And Me.LBLARName.Caption <> "" Then
   If Me.CHKValidate.Value = 1 Then
      Me.TextValidaty.Enabled = True
      Me.TextValidaty.SetFocus
   ElseIf Me.CHKValidate.Value = 0 And Me.CHKExpire.Value = 0 Then
      Me.TextValidaty.Enabled = False
      Me.TextValidaty.Text = vMemValidaty
      Me.TextExpire.Text = vMemExpire
   Else
      Me.TextValidaty.Enabled = False
      Me.TextValidaty.Text = vMemValidaty
   End If
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub CMD391_Click()
Dim vQuery As String
Dim vDocNo As String
Dim vRecordset As New ADODB.Recordset
Dim vConfirm As Integer
Dim vQuestion As Integer

On Error GoTo ErrDescription

'vDocno = Trim(TXT391.Text)

If vDocNo <> "" Then
        vQuery = "select isconfirm from bcnp.dbo.bcsaleorder where docno = '" & vDocNo & "' and sostatus = 1 "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vConfirm = Trim(vRecordset.Fields("isconfirm").Value)
        Else
        MsgBox "เอกสารเลขที่  " & vDocNo & "  ไม่มีในระบบครับ กรุณาใส่เงื่อนไขใหม่ครับ", vbInformation + vbCritical, "ข้อความเตือน"
        Exit Sub
        End If
        vRecordset.Close
    
    
                    If vConfirm = 1 Then
                        vQuestion = MsgBox("คุณต้องการยกเลิกการอนุมัติเลขที่เอกสาร  " & vDocNo & "   นี้ใช่หรือไม่", vbYesNo, "ข้อความสอบถาม")
                        
                        If vQuestion = 6 Then
                            vQuery = "Exec dbo.SP_PRG_00001 '" & vDocNo & "'  "
                            gConnection.Execute vQuery
                            MsgBox "เอกสารเลขที่  " & vDocNo & " ได้ทำการยกเลิกการอนุมัติเรียบร้อยแล้ว ", vbInformation, "ข้อความแจ้งให้ทราบ"
                        Else
                            Exit Sub
                        End If
                    Else
                        MsgBox "เอกสารเลขที่  " & vDocNo & "  ยังไม่ได้อนุมัติครับ", vbInformation, "ข้อความแจ้งให้ทราบ"
                    End If
                
Else
    MsgBox "กรุณาใส่เงื่อนไขให้ครบด้วยครับ", vbInformation + vbCritical, "ข้อความเตือน"
End If

ErrDescription:
If (Err.Number = -2147217911) Then
MsgBox "คุณไม่มีสิทธิ์หน้าจอนี้ กรุณาติดต่อแผนกคอมพิวเตอร์ ", vbInformation + vbCritical, "ข้อความเตือน"
Exit Sub
ElseIf Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Public Sub GetSaleType()
Me.CMBSaleType.Clear
Me.CMBSaleType.AddItem ("ขายสินค้าเงินสด")
Me.CMBSaleType.AddItem ("ขายสินค้าเงินเชื่อ")
Me.CMBSaleType.ListIndex = 0
End Sub

Public Sub GetDeliveryCondition()
Me.CMBDelivery.Clear
Me.CMBDelivery.AddItem ("รับเอง")
Me.CMBDelivery.AddItem ("ส่งให้")
Me.CMBDelivery.ListIndex = 0
End Sub

Public Sub GetCustomerAssert()
Me.CMBAssert.Clear
Me.CMBAssert.AddItem ("รอตอบกลับ")
Me.CMBAssert.AddItem ("ตอบกลับแล้ว")
Me.CMBAssert.AddItem ("ไม่รับในราคา")
Me.CMBAssert.ListIndex = 0
End Sub

Public Sub Load()
Me.CHKARCode.Enabled = False
Me.CHKAssert.Enabled = False
Me.CHKConfirm.Enabled = False
Me.CHKCredit.Enabled = False
Me.CHKDelivery.Enabled = False
Me.CHKExpire.Enabled = False
Me.CHKSaleCode.Enabled = False
Me.CHKSaleType.Enabled = False
Me.CHKSend.Enabled = False
Me.CHKValidate.Enabled = False

Me.CHKARCode.Value = 0
Me.CHKAssert.Value = 0
Me.CHKConfirm.Value = 0
Me.CHKCredit.Value = 0
Me.CHKDelivery.Value = 0
Me.CHKExpire.Value = 0
Me.CHKSaleCode.Value = 0
Me.CHKSaleType.Value = 0
Me.CHKSend.Value = 0
Me.CHKValidate.Value = 0

Me.CMBSaleType.ListIndex = 0
Me.CMBDelivery.ListIndex = 0
Me.CMBAssert.ListIndex = 0
Me.LBLExpiredate.Caption = ""
Me.LBLSendDate.Caption = ""
Me.LBLDuedate.Caption = ""
Me.TextCredit.Text = ""
Me.TextExpire.Text = ""
Me.TextValidaty.Text = ""
Me.TextSend.Text = ""
Me.TextARCode.Text = ""
Me.TextSaleCode.Text = ""
Me.LBLARName.Caption = ""
Me.LBLSaleName.Caption = ""
Me.LBLDocdate.Caption = ""
End Sub

Private Sub CMDProcess_Click()
Dim vAnswer As Integer
Dim vGetSaleType As Integer
Dim vGetDelivery As Integer
Dim vGetAssert As Integer
Dim vGetCredit As Integer
Dim vGetExpire As Integer
Dim vGetValidaty As Integer
Dim vGetArCode As String
Dim vGetSaleCode As String
Dim vGetSend As Integer
Dim vGetIsConfirm As Integer
Dim vGetDuedate As String
Dim vGetExpireDate As String
Dim vGetSendDate As String

On Error GoTo ErrDescription

If Me.CHKARCode.Value = 1 Or Me.CHKAssert.Value = 1 Or Me.CHKConfirm.Value = 1 Or Me.CHKCredit.Value = 1 Or Me.CHKDelivery.Value = 1 Or Me.CHKExpire.Value = 1 Or Me.CHKSaleCode.Value = 1 Or Me.CHKSaleType.Value = 1 Or Me.CHKSend.Value = 1 Or Me.CHKValidate.Value = 1 Then
   If Me.TXTDocNo.Text <> "" And Me.LBLARName.Caption <> "" And Me.LBLSaleName.Caption <> "" Then
   vDocNo = UCase(vDocNo)
      vAnswer = MsgBox("คุณต้องการแก้ไขข้อมูลของเลขที่เอกสาร " & vDocNo & " นี้ใช่หรือไม่ ", vbYesNo, "Send Information Question")
      If vAnswer = 6 Then
         vGetSaleType = Me.CMBSaleType.ListIndex
         If Me.TextSend.Text <> "" Then
            vGetSend = Me.TextSend.Text
         Else
            vGetSend = 0
         End If
         If Me.TextCredit.Text <> "" Then
            vGetCredit = Me.TextCredit.Text
         Else
            vGetCredit = 0
         End If
         If Me.TextExpire.Text <> "" Then
            vGetExpire = Me.TextExpire.Text
         Else
            vGetExpire = 0
         End If
         If Me.TextValidaty.Text <> "" Then
            vGetValidaty = Me.TextValidaty.Text
         Else
            vGetValidaty = 0
         End If
         vGetDelivery = Me.CMBDelivery.ListIndex
         vGetAssert = Me.CMBAssert.ListIndex
         vGetArCode = Me.TextARCode.Text
         vGetSaleCode = Me.TextSaleCode.Text
         vGetIsConfirm = 1
         
          vGetDuedate = Me.LBLDuedate.Caption
          vGetExpireDate = Me.LBLExpiredate.Caption
          vGetSendDate = Me.LBLSendDate.Caption
          
          On Error GoTo ErrRollBack
         
         'vQuery = "begin tran"
         'gConnection.Execute (vQuery)
         
         If Me.CHKConfirm.Value = 1 Then
            vQuery = "exec dbo.USP_SO_UpdateDataDetails 1,'" & vDocNo & "'," & vCheckType & "," & vGetSaleType & "," & vGetSend & "," & vGetCredit & "," & vGetExpire & "," & vGetValidaty & "," & vGetDelivery & "," & vGetAssert & ",'" & vGetArCode & "','" & vGetSaleCode & "' "
            gConnection.Execute (vQuery)
            vGetIsConfirm = 0
         End If
         
         If Me.CHKSaleType.Value = 1 And vMemSaleType <> vGetSaleType Then
            vQuery = "exec dbo.USP_SO_UpdateDataDetails 2,'" & vDocNo & "'," & vCheckType & "," & vGetSaleType & "," & vGetSend & "," & vGetCredit & "," & vGetExpire & "," & vGetValidaty & "," & vGetDelivery & "," & vGetAssert & ",'" & vGetArCode & "','" & vGetSaleCode & "' "
            gConnection.Execute (vQuery)
         End If
      
         If Me.CHKDelivery.Value = 1 And vMemDelivery <> vGetDelivery Then
            vQuery = "exec dbo.USP_SO_UpdateDataDetails 3,'" & vDocNo & "'," & vCheckType & "," & vGetSaleType & "," & vGetSend & "," & vGetCredit & "," & vGetExpire & "," & vGetValidaty & "," & vGetDelivery & "," & vGetAssert & ",'" & vGetArCode & "','" & vGetSaleCode & "' "
            gConnection.Execute (vQuery)
         End If
         
         If Me.CHKCredit.Value = 1 And vMemCredit <> vGetCredit Then
            vQuery = "exec dbo.USP_SO_UpdateDataDetails 4,'" & vDocNo & "'," & vCheckType & "," & vGetSaleType & "," & vGetSend & "," & vGetCredit & "," & vGetExpire & "," & vGetValidaty & "," & vGetDelivery & "," & vGetAssert & ",'" & vGetArCode & "','" & vGetSaleCode & "' "
            gConnection.Execute (vQuery)
         End If
         
         If Me.CHKExpire.Value = 1 And vMemExpire <> vGetExpire Then
            vQuery = "exec dbo.USP_SO_UpdateDataDetails 5,'" & vDocNo & "'," & vCheckType & "," & vGetSaleType & "," & vGetSend & "," & vGetCredit & "," & vGetExpire & "," & vGetValidaty & "," & vGetDelivery & "," & vGetAssert & ",'" & vGetArCode & "','" & vGetSaleCode & "' "
            gConnection.Execute (vQuery)
         End If
         
         If Me.CHKValidate.Value = 1 And vMemValidaty <> vGetValidaty Then
            vQuery = "exec dbo.USP_SO_UpdateDataDetails 6,'" & vDocNo & "'," & vCheckType & "," & vGetSaleType & "," & vGetSend & "," & vGetCredit & "," & vGetExpire & "," & vGetValidaty & "," & vGetDelivery & "," & vGetAssert & ",'" & vGetArCode & "','" & vGetSaleCode & "' "
            gConnection.Execute (vQuery)
         End If
         
         If Me.CHKSend.Value = 1 And vMemSend <> vGetSend Then
            vQuery = "exec dbo.USP_SO_UpdateDataDetails 7,'" & vDocNo & "'," & vCheckType & "," & vGetSaleType & "," & vGetSend & "," & vGetCredit & "," & vGetExpire & "," & vGetValidaty & "," & vGetDelivery & "," & vGetAssert & ",'" & vGetArCode & "','" & vGetSaleCode & "' "
            gConnection.Execute (vQuery)
         End If
         
         If Me.CHKAssert.Value = 1 Then
            vQuery = "exec dbo.USP_SO_UpdateDataDetails 8,'" & vDocNo & "'," & vCheckType & "," & vGetSaleType & "," & vGetSend & "," & vGetCredit & "," & vGetExpire & "," & vGetValidaty & "," & vGetDelivery & "," & vGetAssert & ",'" & vGetArCode & "','" & vGetSaleCode & "' "
            gConnection.Execute (vQuery)
         End If
         
         If Me.CHKSaleCode.Value = 1 And vMemSaleCode <> vGetSaleCode Then
            vQuery = "exec dbo.USP_SO_UpdateDataDetails 9,'" & vDocNo & "'," & vCheckType & "," & vGetSaleType & "," & vGetSend & "," & vGetCredit & "," & vGetExpire & "," & vGetValidaty & "," & vGetDelivery & "," & vGetAssert & ",'" & vGetArCode & "','" & vGetSaleCode & "' "
            gConnection.Execute (vQuery)
         End If
         
         If Me.CHKARCode.Value = 1 And vMemArCode <> vGetArCode Then
            vQuery = "exec dbo.USP_SO_UpdateDataDetails 10,'" & vDocNo & "'," & vCheckType & "," & vGetSaleType & "," & vGetSend & "," & vGetCredit & "," & vGetExpire & "," & vGetValidaty & "," & vGetDelivery & "," & vGetAssert & ",'" & vGetArCode & "','" & vGetSaleCode & "' "
            gConnection.Execute (vQuery)
         End If
         
         'vQuery = "commit tran"
         'gConnection.Execute (vQuery)
         
ErrRollBack:
If Err.Description <> "" Then
   vQuery = "rollback tran"
   gConnection.Execute (vQuery)
   MsgBox Err.Description, vbCritical, "Send Error Message"
   Exit Sub
End If
         
         vQuery = "exec dbo.USP_SO_HistoryUpdateDataDetails '" & vDocNo & "','" & vMemDocDate & "'," & vGetIsConfirm & "," & vMemSaleType & "," & vMemDelivery & ",'" & vMemDueDate & "','" & vMemExpireDate & "','" & vMemSendDate & "','" & vUserID & "'," & vMemAssert & " "
         gConnection.Execute (vQuery)
         
         MsgBox "เลขที่เอกสาร " & vDocNo & " ได้ทำการปรับปรุงข้อมูลที่ต้องการแก้ไขให้เรียบร้อยแล้ว กรุณาตรวจสอบข้อมูลใน BCAccount ", vbInformation, "Send Information Message"
         Call Form_Load
         Me.TXTDocNo.Text = ""
         Me.TXTDocNo.SetFocus
      End If
   Else
      MsgBox "กรุณากรอกข้อมูลให้ครบถ้วนและถูกต้อง", vbCritical, "Send Error Message"
   End If
Else
   MsgBox "ไม่มีการปรับปรุงข้อมูลของเอกสาร เนื่องจากไม่มีการแก้ไขข้อมูลเปลี่ยนไปจากเดิม", vbInformation, "Send Information Message"
End If




ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If

End Sub

Private Sub Form_Load()
Call GetSaleType
Call GetDeliveryCondition
Call GetCustomerAssert
Call Load
Me.IMG101.Visible = True
Me.IMG102.Visible = False
Me.CMDProcess.Enabled = False
End Sub

Private Sub TextARCode_Change()
Dim vARCode As String
Dim vRecordset As New ADODB.Recordset

On Error GoTo ErrDescription

If Me.TextARCode.Text <> "" Then
   vARCode = Me.TextARCode.Text
   vQuery = "select code,isnull(name1,'') as arname from dbo.bcar where code = '" & vARCode & "' "
   If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
      Me.LBLARName.Caption = vRecordset.Fields("arname").Value
   Else
      Me.LBLARName.Caption = ""
   End If
   vRecordset.Close
Else
   Me.LBLARName.Caption = ""
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub TextCredit_Change()
Dim vChangeCredit As Integer
Dim vDiffCredit As Integer

On Error GoTo ErrDescription

If Me.TextCredit.Text = "" Then
   vChangeCredit = 0
End If

If Me.TextCredit.Text <> "" And Me.TextCredit.Text <> "0" Then
   Call CheckNumber(Me.TextCredit.Text)
   If vCheckValueNumber = True Then
     If vMemCredit > 0 Then
         vChangeCredit = Me.TextCredit.Text
         vDiffCredit = vChangeCredit - vMemCredit
         Me.LBLDuedate.Caption = DateAdd("d", vDiffCredit, vMemDueDate)
      ElseIf vMemCredit = 0 And vChangeCredit > 0 Then
         vDiffCredit = vChangeCredit
         Me.LBLDuedate.Caption = DateAdd("d", vDiffCredit, vMemDocDate)
      Else
         Me.LBLDuedate.Caption = ""
      End If
   Else
      Me.LBLDuedate.Caption = vMemDueDate
      Me.TextCredit.Text = vMemCredit
      MsgBox "กรอกข้อมูลได้เฉพาะตัวเลขเท่านั้น", vbCritical, "Send Error Message"
   End If
Else
   Me.LBLDuedate.Caption = ""
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Public Sub CheckNumber(vData As String)
Dim vDocNo As String
Dim vText As String
Dim i As Integer

On Error GoTo ErrDescription

vDocNo = Trim(vData)

For i = 1 To Len(vData)
    If Mid(vDocNo, i, 1) = 0 Or Mid(vDocNo, i, 1) = 1 Or Mid(vDocNo, i, 1) = 2 Or Mid(vDocNo, i, 1) = 3 Or Mid(vDocNo, i, 1) = 4 Or Mid(vDocNo, i, 1) = 5 Or Mid(vDocNo, i, 1) = 6 Or Mid(vDocNo, i, 1) = 7 Or Mid(vDocNo, i, 1) = 8 Or Mid(vDocNo, i, 1) = 9 Then
        vCheckValueNumber = True
    Else
        vCheckValueNumber = False
        Exit Sub
    End If
Next i

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub


Private Sub TextExpire_Change()
Dim vChangeExpire As Integer
Dim vDiffExpire As Integer
Dim vChangeValidaty As Integer

On Error GoTo ErrDescription

   If Me.TextValidaty.Text = "" Then
      vChangeValidaty = 0
   End If
   
   If Me.TextExpire.Text = "" Then
      vChangeExpire = 0
   End If

If Me.TextExpire.Text <> "" And Me.TextExpire.Text <> "0" Then
   Call CheckNumber(Me.TextExpire.Text)
   If vCheckValueNumber = True Then
      If vMemExpire > 0 Then
         vChangeExpire = Me.TextExpire.Text
         If Me.TextValidaty.Text = "" Then
            vChangeValidaty = 0
         Else
            vChangeValidaty = Me.TextValidaty.Text
         End If
         vDiffExpire = vChangeExpire - vMemExpire
         Me.LBLExpiredate.Caption = DateAdd("d", vDiffExpire, vMemExpireDate)
      ElseIf vMemExpire = 0 And vChangeValidaty <> vMemValidaty Then
         vChangeExpire = Me.TextExpire.Text
         vDiffExpire = vChangeExpire - vMemExpire
         Me.LBLExpiredate.Caption = DateAdd("d", vDiffExpire, vMemDocDate)
      ElseIf vMemExpire = 0 And vChangeExpire > 0 Then
         vChangeExpire = Me.TextExpire.Text
         vDiffExpire = vChangeExpire - vMemExpire
         Me.LBLExpiredate.Caption = DateAdd("d", vDiffExpire, vMemDocDate)
      Else
         Me.LBLExpiredate.Caption = ""
      End If
   Else
      Me.LBLExpiredate.Caption = vMemExpireDate
      Me.TextExpire.Text = vMemExpire
      MsgBox "กรอกข้อมูลได้เฉพาะตัวเลขเท่านั้น", vbCritical, "Send Error Message"
   End If
Else
   Me.LBLExpiredate.Caption = ""
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub TextSaleCode_Change()
Dim vSaleCode As String
Dim vRecordset As New ADODB.Recordset

On Error GoTo ErrDescription

If Me.TextSaleCode.Text <> "" Then
   vSaleCode = Me.TextSaleCode.Text
   vQuery = "select code,isnull(name,'') as salename from dbo.bcsale where code = '" & vSaleCode & "' "
   If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
      Me.LBLSaleName.Caption = vRecordset.Fields("salename").Value
   Else
      Me.LBLSaleName.Caption = ""
   End If
   vRecordset.Close
Else
   Me.LBLSaleName.Caption = ""
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub TextSend_Change()
Dim vChangeSend As Integer
Dim vDiffSend As Integer

On Error GoTo ErrDescription

If Me.TextSend.Text = "" Then
   vChangeSend = 0
End If

If Me.TextSend.Text <> "" And Me.TextSend.Text <> "0" Then
   Call CheckNumber(Me.TextSend.Text)
   If vCheckValueNumber = True Then
      If vMemSend > 0 Then
         vChangeSend = Me.TextSend.Text
         vDiffSend = vChangeSend - vMemSend
         Me.LBLSendDate.Caption = DateAdd("d", vDiffSend, vMemSendDate)
      Else
         vChangeSend = Me.TextSend.Text
         vDiffSend = vChangeSend
         Me.LBLSendDate.Caption = DateAdd("d", vDiffSend, vMemDocDate)
      End If
   Else
      Me.LBLSendDate.Caption = vMemSendDate
      Me.TextSend.Text = vMemSend
      MsgBox "กรอกข้อมูลได้เฉพาะตัวเลขเท่านั้น", vbCritical, "Send Error Message"
   End If
Else
   Me.LBLSendDate.Caption = ""
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub TextValidaty_Change()
Dim vChangeValidaty As Integer
Dim vExpire As Integer

On Error GoTo ErrDescription

If Me.TextValidaty.Text = "" Then
   vChangeValidaty = 0
End If

If Me.TextValidaty.Text <> "" And Me.TextValidaty.Text <> "0" Then
   Call CheckNumber(Me.TextValidaty.Text)
   If vCheckValueNumber = True Then
      vChangeValidaty = Me.TextValidaty.Text
      If Me.CHKValidate.Value = 1 Then
      Me.TextExpire.Text = Me.TextValidaty.Text
      End If
   Else
      Me.TextValidaty.Text = vMemValidaty
      MsgBox "กรอกข้อมูลได้เฉพาะตัวเลขเท่านั้น", vbCritical, "Send Error Message"
   End If
ElseIf vMemExpire = 0 Then
   Me.TextExpire.Text = Me.TextValidaty.Text
ElseIf vChangeValidaty = 0 Then
  Me.TextExpire.Text = Me.TextValidaty.Text
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub TXTDocNo_Change()
If UCase(vDocNo) <> UCase(Me.TXTDocNo.Text) Then
   Call Form_Load
End If
End Sub

Private Sub TXTDocNo_LostFocus()
Me.TXTDocNo.Text = UCase(Me.TXTDocNo.Text)
End Sub

Private Sub TXTDocNo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 8 Then
   Call Form_Load
End If
End Sub

Private Sub TXTDocNo_KeyPress(KeyAscii As Integer)
Dim vRecordset As New ADODB.Recordset

On Error GoTo ErrDescription

If KeyAscii = 13 Then
   vDocNo = TXTDocNo.Text
   Call Form_Load
   vQuery = "exec dbo.USP_SO_CheckDocDetails '" & vDocNo & "' "
   If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
      vCheckType = vRecordset.Fields("type").Value
      Me.LBLDocdate.Caption = vRecordset.Fields("docdate").Value
      If vRecordset.Fields("isconfirm").Value = 0 Then
         Me.IMG101.Visible = True
         Me.IMG102.Visible = False
         MsgBox "เอกสารยังไม่ได้ถูกอ้างอิง แก้ไขที่โปรแกรม BCAccount 5.5 ได้ปกติ", vbCritical, "Send Error Message"
         vRecordset.Close
         Me.TXTDocNo.SetFocus
         Exit Sub
      ElseIf vRecordset.Fields("isconfirm").Value = 1 Then
         Me.IMG101.Visible = False
         Me.IMG102.Visible = True
      End If
      
      Me.CMDProcess.Enabled = True
      
      vMemDocDate = vRecordset.Fields("docdate").Value
      vMemSaleType = vRecordset.Fields("billtype").Value
      vMemDelivery = vRecordset.Fields("isconditionsend").Value
      vMemAssert = vRecordset.Fields("assertstatus").Value
      vMemCredit = vRecordset.Fields("creditday").Value
      vMemDueDate = vRecordset.Fields("duedate").Value
      vMemExpire = vRecordset.Fields("expirecredit").Value
      vMemValidaty = vRecordset.Fields("validity").Value
      vMemArCode = vRecordset.Fields("arcode").Value
      vMemSaleCode = vRecordset.Fields("salecode").Value
      vMemExpireDate = vRecordset.Fields("expiredate").Value
      vMemSend = vRecordset.Fields("deliveryday").Value
      vMemSendDate = vRecordset.Fields("deliverydate").Value
      vMemArName = vRecordset.Fields("arname").Value
      vMemSaleName = vRecordset.Fields("salename").Value

      If vRecordset.Fields("type").Value = 0 Then
      
         Me.CMBSaleType.ListIndex = vRecordset.Fields("billtype").Value
         Me.CMBDelivery.ListIndex = vRecordset.Fields("isconditionsend").Value
         Me.CMBAssert.ListIndex = vRecordset.Fields("assertstatus").Value
         Me.LBLExpiredate.Caption = vRecordset.Fields("expiredate").Value
         Me.LBLSendDate.Caption = vRecordset.Fields("deliverydate").Value
         Me.LBLDuedate.Caption = vRecordset.Fields("duedate").Value
         Me.TextCredit.Text = vRecordset.Fields("creditday").Value
         Me.TextExpire.Text = vRecordset.Fields("expirecredit").Value
         Me.TextValidaty.Text = vRecordset.Fields("validity").Value
         Me.TextSend.Text = vRecordset.Fields("deliveryday").Value
         Me.TextARCode.Text = vRecordset.Fields("arcode").Value
         Me.TextSaleCode.Text = vRecordset.Fields("salecode").Value
         Me.LBLARName.Caption = vRecordset.Fields("arname").Value
         Me.LBLSaleName.Caption = vRecordset.Fields("salename").Value
         
         Me.CHKARCode.Enabled = True
         Me.CHKAssert.Enabled = True
         Me.CHKConfirm.Enabled = True
         Me.CHKCredit.Enabled = True
         Me.CHKDelivery.Enabled = False
         Me.CHKExpire.Enabled = True
         Me.CHKSaleCode.Enabled = True
         Me.CHKSaleType.Enabled = True
         Me.CHKSend.Enabled = True
         Me.CHKValidate.Enabled = True
      ElseIf vRecordset.Fields("type").Value = 1 Then
      
         Me.CMBSaleType.ListIndex = vRecordset.Fields("billtype").Value
         Me.CMBDelivery.ListIndex = vRecordset.Fields("isconditionsend").Value
         Me.CMBAssert.ListIndex = vRecordset.Fields("assertstatus").Value
         Me.LBLExpiredate.Caption = vRecordset.Fields("expiredate").Value
         Me.LBLSendDate.Caption = vRecordset.Fields("deliverydate").Value
         Me.LBLDuedate.Caption = vRecordset.Fields("duedate").Value
         Me.TextCredit.Text = vRecordset.Fields("creditday").Value
         Me.TextExpire.Text = vRecordset.Fields("expirecredit").Value
         Me.TextValidaty.Text = vRecordset.Fields("validity").Value
         Me.TextSend.Text = vRecordset.Fields("deliveryday").Value
         Me.TextARCode.Text = vRecordset.Fields("arcode").Value
         Me.TextSaleCode.Text = vRecordset.Fields("salecode").Value
         Me.LBLARName.Caption = vRecordset.Fields("arname").Value
         Me.LBLSaleName.Caption = vRecordset.Fields("salename").Value
         
         Me.CHKARCode.Enabled = True
         Me.CHKAssert.Enabled = True
         Me.CHKConfirm.Enabled = True
         Me.CHKCredit.Enabled = True
         Me.CHKDelivery.Enabled = False
         Me.CHKExpire.Enabled = True
         Me.CHKSaleCode.Enabled = True
         Me.CHKSaleType.Enabled = True
         Me.CHKSend.Enabled = True
         Me.CHKValidate.Enabled = True
      ElseIf vRecordset.Fields("type").Value = 2 Then
      
         Me.CMBSaleType.ListIndex = vRecordset.Fields("billtype").Value
         Me.CMBDelivery.ListIndex = vRecordset.Fields("isconditionsend").Value
         Me.CMBAssert.ListIndex = 0
         Me.LBLExpiredate.Caption = vRecordset.Fields("expiredate").Value
         Me.LBLSendDate.Caption = vRecordset.Fields("deliverydate").Value
         Me.LBLDuedate.Caption = vRecordset.Fields("duedate").Value
         Me.TextCredit.Text = vRecordset.Fields("creditday").Value
         Me.TextExpire.Text = vRecordset.Fields("expirecredit").Value
         Me.TextValidaty.Text = ""
         Me.TextSend.Text = vRecordset.Fields("deliveryday").Value
         Me.TextARCode.Text = vRecordset.Fields("arcode").Value
         Me.TextSaleCode.Text = vRecordset.Fields("salecode").Value
         Me.LBLARName.Caption = vRecordset.Fields("arname").Value
         Me.LBLSaleName.Caption = vRecordset.Fields("salename").Value
         
         Me.CHKARCode.Enabled = True
         Me.CHKAssert.Enabled = False
         Me.CHKConfirm.Enabled = True
         Me.CHKCredit.Enabled = True
         Me.CHKDelivery.Enabled = True
         Me.CHKExpire.Enabled = False
         Me.CHKSaleCode.Enabled = True
         Me.CHKSaleType.Enabled = True
         Me.CHKSend.Enabled = True
         Me.CHKValidate.Enabled = False
      ElseIf vRecordset.Fields("type").Value = 3 Then
      
         Me.CMBSaleType.ListIndex = vRecordset.Fields("billtype").Value
         Me.CMBDelivery.ListIndex = vRecordset.Fields("isconditionsend").Value
         Me.CMBAssert.ListIndex = 0
         Me.LBLExpiredate.Caption = vRecordset.Fields("expiredate").Value
         Me.LBLSendDate.Caption = vRecordset.Fields("deliverydate").Value
         Me.LBLDuedate.Caption = vRecordset.Fields("duedate").Value
         Me.TextCredit.Text = vRecordset.Fields("creditday").Value
         Me.TextExpire.Text = vRecordset.Fields("expirecredit").Value
         Me.TextSend.Text = vRecordset.Fields("deliveryday").Value
         Me.TextARCode.Text = vRecordset.Fields("arcode").Value
         Me.TextSaleCode.Text = vRecordset.Fields("salecode").Value
         Me.LBLARName.Caption = vRecordset.Fields("arname").Value
         Me.LBLSaleName.Caption = vRecordset.Fields("salename").Value
         
         Me.CHKARCode.Enabled = True
         Me.CHKAssert.Enabled = False
         Me.CHKConfirm.Enabled = True
         Me.CHKCredit.Enabled = True
         Me.CHKDelivery.Enabled = True
         Me.CHKExpire.Enabled = False
         Me.CHKSaleCode.Enabled = True
         Me.CHKSaleType.Enabled = True
         Me.CHKSend.Enabled = True
         Me.CHKValidate.Enabled = False
      End If

   Else
      MsgBox "ไม่มีข้อมูลของเลขที่เอกสารที่ต้องการแก้ไขข้อมูล หรือ เอกสารดังกล่าวถูกยกเลิก  หรือ ถูกอ้างทำเอกสารอื่นไปแล้ว", vbCritical, "Send Error Message"
      Call Form_Load
   End If
   vRecordset.Close
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub
