VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProductDetail 
   Caption         =   "Form Product Detail"
   ClientHeight    =   7440
   ClientLeft      =   2250
   ClientTop       =   1725
   ClientWidth     =   11025
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7440
   ScaleWidth      =   11025
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   9495
      TabIndex        =   2
      Top             =   6210
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   8145
      TabIndex        =   1
      Top             =   6210
      Width           =   1215
   End
   Begin MSComctlLib.ListView LV_ProductDetail 
      Height          =   5730
      Left            =   315
      TabIndex        =   0
      Top             =   225
      Width           =   10380
      _ExtentX        =   18309
      _ExtentY        =   10107
      View            =   3
      LabelEdit       =   1
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
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "รหัสสินค้า"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "คลัง"
         Object.Width           =   1323
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "บาร์โค้ด"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "รายชื่อสินค้า"
         Object.Width           =   4939
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "หน่วย"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "ราคา"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "ราคาปกติ"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "ที่เก็บ"
         Object.Width           =   1764
      EndProperty
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "ช่องราคาตั้ง ต้องไม่เท่ากับ 0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   14.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   315
      TabIndex        =   4
      Top             =   6570
      Width           =   7530
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "รายการสินค้าที่จะพิมพ์ป้ายได้ ต้องมีรายการครบทุกช่อง"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   14.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   315
      TabIndex        =   3
      Top             =   6120
      Width           =   7485
   End
End
Attribute VB_Name = "frmProductDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
        Unload Me
        frmWizard.Enabled = True
        frmWizard.Text1.Text = ""
        frmWizard.SetFocus
End Sub

Private Sub cmdOK_Click()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset

        If LV_ProductDetail.ListItems.count <> 0 Then
        tmpItemNumber = Trim(LV_ProductDetail.SelectedItem.Text)
        
        vQuery = "exec dbo.USP_NP_CheckItemInRecProduct2 '" & tmpItemNumber & "' "
        If OpenDatabase(vConnection, vRecordset, vQuery) <> 0 Then
          vCheckRecProduct = vRecordset.Fields("vcount").Value
        End If
        vRecordset.Close
        If vCheckRecProduct > 0 Then
        tmpBarcod = Trim(LV_ProductDetail.SelectedItem.SubItems(2))
        tmpItemDesc = Trim(LV_ProductDetail.SelectedItem.SubItems(3))
        frmWizard.Label6.Caption = Trim(LV_ProductDetail.SelectedItem.SubItems(3))
        tmpUOFM = Trim(LV_ProductDetail.SelectedItem.SubItems(4))
        tmpPrice = Trim(LV_ProductDetail.SelectedItem.SubItems(5))
        tmpSPrice = Trim(LV_ProductDetail.SelectedItem.SubItems(6))
        tmpWHCode = Trim(LV_ProductDetail.SelectedItem.SubItems(1))
        tmpShelfCode = Trim(LV_ProductDetail.SelectedItem.SubItems(7))
        Else
        tmpBarcod = ""
        tmpItemDesc = ""
        frmWizard.Label6.Caption = ""
        tmpUOFM = ""
        tmpPrice = ""
        tmpSPrice = ""
        tmpWHCode = ""
        MsgBox "สินค้าดังกล่าวยังไม่ได้ระบุที่อยู่ของสินค้า ต้องทำการระบุที่อยู่สินค้าถึงจะพิมพ์ป้ายต่าง ๆ ได้", vbCritical, "Send Error Message"
        End If
        
        Unload Me
        frmWizard.Enabled = True
        frmWizard.Text2.Enabled = True
        frmWizard.Text2.SetFocus
        End If
End Sub

Private Sub LV_ProductDetail_DblClick()
    If LV_ProductDetail.ListItems.count <> 0 Then
    Call cmdOK_Click
    End If
End Sub

Private Sub LV_ProductDetail_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
                Call cmdOK_Click
        End If
End Sub
