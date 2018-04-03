VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form21 
   Caption         =   "หน้าแก้ไขวันที่หมดอายุใบสั่งซื้อ"
   ClientHeight    =   8340
   ClientLeft      =   2655
   ClientTop       =   1260
   ClientWidth     =   12000
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   Picture         =   "Form2.frx":08CA
   ScaleHeight     =   8340
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CMD211 
      Caption         =   "UpDate"
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
      Height          =   465
      Left            =   7875
      TabIndex        =   4
      Top             =   3450
      Width           =   1140
   End
   Begin MSComCtl2.DTPicker DTP211 
      Height          =   390
      Left            =   1575
      TabIndex        =   3
      Top             =   3000
      Width           =   2265
      _ExtentX        =   3995
      _ExtentY        =   688
      _Version        =   393216
      Format          =   66584577
      CurrentDate     =   38022
   End
   Begin VB.TextBox TXT213 
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
      Height          =   390
      Left            =   6750
      TabIndex        =   2
      Top             =   2325
      Width           =   2265
   End
   Begin VB.TextBox TXT212 
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
      Height          =   390
      Left            =   1575
      TabIndex        =   1
      Top             =   2325
      Width           =   2265
   End
   Begin VB.TextBox TXT211 
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
      Height          =   390
      Left            =   1575
      TabIndex        =   0
      Top             =   1650
      Width           =   2265
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "แก้วันที่หมดอายุของใบสั่งซื้อ"
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
      TabIndex        =   9
      Top             =   300
      Width           =   7365
   End
   Begin VB.Label LBL214 
      BackStyle       =   0  'Transparent
      Caption         =   "วันที่ใบสั่งซื้อ"
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
      Height          =   390
      Left            =   5700
      TabIndex        =   8
      Top             =   2325
      Width           =   1065
   End
   Begin VB.Label LBL213 
      BackStyle       =   0  'Transparent
      Caption         =   "วันที่แก้ไข"
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
      Height          =   390
      Left            =   525
      TabIndex        =   7
      Top             =   3000
      Width           =   840
   End
   Begin VB.Label LBL212 
      BackStyle       =   0  'Transparent
      Caption         =   "วันที่หมดอายุ"
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
      Top             =   2325
      Width           =   990
   End
   Begin VB.Label LBL211 
      BackStyle       =   0  'Transparent
      Caption         =   "เลขที่ใบสั่งซื้อ"
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
      TabIndex        =   5
      Top             =   1725
      Width           =   990
   End
End
Attribute VB_Name = "Form21"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMD211_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocNo  As String
Dim vExpireDate As Date, vChangeDate As Date, vDocdate As Date

On Error GoTo ErrDescription

vDocNo = Trim(TXT211.Text)
vExpireDate = DTP211.Day & "/" & DTP211.Month & "/" & DTP211.Year
vQuery = "select docdate from bcpurchaseorder where docno = '" & vDocNo & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vDocdate = Trim(vRecordset.Fields("docdate").Value)
End If
vRecordset.Close
'-------------------------------------------------------------------------------------------
If vDocdate < vExpireDate Then
        vQuery = "set dateformat dmy  Update BCPurchaseOrder set expiredate = '" & vExpireDate & "' where DocNo = '" & vDocNo & "' "
        gConnection.Execute vQuery
'----------------------------------------------------------------------------------------------
        vQuery = "select Expiredate from bcpurchaseorder where docno = '" & vDocNo & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vChangeDate = Trim(vRecordset.Fields("expiredate").Value)
        End If
        vRecordset.Close
'-------------------------------------------------------------------------------------------
        MsgBox "วันที่หมดอายุได้ถูกแก้ไขเป็นวันที่ " & vChangeDate & "เรียบร้อยแล้ว", vbInformation, "ข้อความแจ้ง"
Else
    MsgBox "ไม่สามารถแก้ไขวันที่น้อยกว่าวันที่ทำเอกสารได้", vbCritical, "ข้อความเตือน"
End If


ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Form_Load()
DTP211.Value = Now
End Sub

Private Sub TXT211_KeyPress(KeyAscii As Integer)
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocNo As String
Dim vExpireDate As Date, vDocdate As Date

On Error GoTo ErrDescription
If KeyAscii = 13 Then
vDocNo = Trim(TXT211.Text)
vQuery = "select Docno ,Docdate , expiredate from BCPurchaseOrder where Docno = '" & vDocNo & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vExpireDate = Trim(vRecordset.Fields("expiredate").Value)
    vDocdate = Trim(vRecordset.Fields("Docdate").Value)
    CMD211.Enabled = True
Else
MsgBox "ไม่มีเลขที่ใบสั่งซื้อดังกล่าว ไม่สามารถแก้ไขวันที่หมดอายุได้ ", vbCritical, "ข้อความเตือน"
CMD211.Enabled = False
TXT211.SetFocus
End If
vRecordset.Close
'-----------------------------------------------------------------------------------------------------------------

End If
TXT212.Text = vExpireDate
TXT213.Text = vDocdate

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub
