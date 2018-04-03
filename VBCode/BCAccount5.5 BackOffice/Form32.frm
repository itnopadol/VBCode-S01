VERSION 5.00
Begin VB.Form Form32 
   Caption         =   "Update SaleTax"
   ClientHeight    =   8970
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form32.frx":0000
   ScaleHeight     =   8970
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CMD101 
      Caption         =   "UpDate SaleTax"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   4095
      TabIndex        =   0
      Top             =   1530
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "กดปุ่มนี้เพื่อทำการอัพเดท คำอธิบายของภาษีเพื่อออกรายงาน ภาษีขาย"
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
      Height          =   840
      Left            =   2205
      TabIndex        =   1
      Top             =   1035
      Width           =   6135
   End
End
Attribute VB_Name = "Form32"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMD101_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String

vQuery = "Update bcvat.dbo.bcoutputtax set shorttaxdesc = 'ขายสินค้า'  where shorttaxdesc like '%N/A%' "
gConnection.Execute vQuery
MsgBox "Update ข้อมูลที่จะออกรายงานภาษีขายเรียบร้อยแล้ว"
End Sub

