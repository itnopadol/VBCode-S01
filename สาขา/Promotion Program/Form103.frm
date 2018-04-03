VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form Form103 
   Caption         =   "พิมพ์ป้ายราคาโปรโมชั่น"
   ClientHeight    =   9000
   ClientLeft      =   1545
   ClientTop       =   630
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport Crystal102 
      Left            =   765
      Top             =   8280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin Crystal.CrystalReport Crystal101 
      Left            =   270
      Top             =   8280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.CommandButton CMD106 
      Caption         =   "เคลียร์ตะกร้า"
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
      Left            =   270
      TabIndex        =   25
      Top             =   7785
      Width           =   1215
   End
   Begin VB.CommandButton CMD104 
      Height          =   315
      Left            =   11025
      Picture         =   "Form103.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7785
      Width           =   315
   End
   Begin VB.TextBox Text103 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   315
      Left            =   7125
      TabIndex        =   15
      Top             =   7785
      Width           =   3840
   End
   Begin VB.CommandButton CMD101 
      Caption         =   "พิมพ์ป้ายราคา"
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
      Left            =   9525
      TabIndex        =   7
      Top             =   8190
      Width           =   1440
   End
   Begin VB.Frame Frame2 
      Caption         =   "เลือกรายการสินค้าพิมพ์ป้าย"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5985
      Left            =   45
      TabIndex        =   9
      Top             =   1755
      Width           =   11895
      Begin VB.CommandButton CMD105 
         Caption         =   "ลงตะกร้า"
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
         Left            =   225
         TabIndex        =   4
         Top             =   3375
         Width           =   1215
      End
      Begin MSComctlLib.ListView ListView102 
         Height          =   2100
         Left            =   225
         TabIndex        =   5
         Top             =   3810
         Width           =   11490
         _ExtentX        =   20267
         _ExtentY        =   3704
         View            =   3
         LabelEdit       =   1
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
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "รหัสสินค้า"
            Object.Width           =   2734
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "บาร์โค้ด"
            Object.Width           =   2734
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "ราคาตั้ง"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "ราคาปกติ"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "ราคาโปรโมชั่น"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "หน่วยนับ"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "ชื่อสินค้า"
            Object.Width           =   6262
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "ส่วนลด"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "จำนวน"
            Object.Width           =   2117
         EndProperty
      End
      Begin VB.CheckBox Check101 
         Caption         =   "เลือกทั้งหมด"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   300
         TabIndex        =   20
         Top             =   300
         Width           =   1965
      End
      Begin MSComctlLib.ListView ListView101 
         Height          =   2745
         Left            =   225
         TabIndex        =   3
         Top             =   600
         Width           =   11490
         _ExtentX        =   20267
         _ExtentY        =   4842
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   10
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "รหัสสินค้า"
            Object.Width           =   2734
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "บาร์โค้ด"
            Object.Width           =   2734
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "ราคาตั้ง"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "ราคาปกติ"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "ราคาโปรโมชั่น"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "หน่วยนับ"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "ชื่อสินค้า"
            Object.Width           =   6262
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "ส่วนลด"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "สมาชิก"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "เลขที่เอกสาร"
            Object.Width           =   2646
         EndProperty
      End
      Begin VB.Label Label104 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   10725
         TabIndex        =   24
         Top             =   3555
         Width           =   915
      End
      Begin VB.Label Label9 
         Caption         =   "จำนวนสินค้าในตะกร้า :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   8925
         TabIndex        =   23
         Top             =   3555
         Width           =   1815
      End
      Begin VB.Label Label103 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   11025
         TabIndex        =   22
         Top             =   315
         Width           =   690
      End
      Begin VB.Label Label6 
         Caption         =   "จำนวนสินค้าในโปรโมชั่น :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   8925
         TabIndex        =   21
         Top             =   315
         Width           =   2040
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "เลือกเงื่อนไขการพิมพ์ป้าย"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1440
      Left            =   45
      TabIndex        =   8
      Top             =   135
      Width           =   11895
      Begin VB.TextBox Text104 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1485
         TabIndex        =   2
         ToolTipText     =   "กรอกแล้วกดปุ่ม Enter"
         Top             =   945
         Width           =   2715
      End
      Begin VB.CommandButton CMD103 
         Height          =   315
         Left            =   10950
         Picture         =   "Form103.frx":03CD
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   375
         Width           =   315
      End
      Begin VB.CommandButton CMD102 
         Height          =   315
         Left            =   5250
         Picture         =   "Form103.frx":079A
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   375
         Width           =   315
      End
      Begin VB.TextBox Text102 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   8100
         TabIndex        =   13
         Top             =   360
         Width           =   2790
      End
      Begin VB.TextBox Text101 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1500
         TabIndex        =   12
         Top             =   375
         Width           =   3690
      End
      Begin VB.Label Label7 
         Caption         =   "เลขที่ใบเสนอ"
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
         Left            =   270
         TabIndex        =   26
         Top             =   990
         Width           =   1140
      End
      Begin VB.Label Label102 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   10050
         TabIndex        =   19
         Top             =   1005
         Width           =   1740
      End
      Begin VB.Label Label101 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   7125
         TabIndex        =   18
         Top             =   1005
         Width           =   1365
      End
      Begin VB.Label Label5 
         Caption         =   "วันที่หมดโปรโมชั่น"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   8550
         TabIndex        =   17
         Top             =   1005
         Width           =   1515
      End
      Begin VB.Label Label4 
         Caption         =   "วันที่เริ่มโปรโมชั่น"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5700
         TabIndex        =   16
         Top             =   1005
         Width           =   1440
      End
      Begin VB.Label Label2 
         Caption         =   "เลือก Section Manager"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6075
         TabIndex        =   11
         Top             =   450
         Width           =   1890
      End
      Begin VB.Label Label1 
         Caption         =   "เลือกโปรโมชั่น"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   300
         TabIndex        =   10
         Top             =   450
         Width           =   1140
      End
   End
   Begin VB.Label Label3 
      Caption         =   "เลือกฟอร์มป้ายราคา"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5475
      TabIndex        =   14
      Top             =   7830
      Width           =   1590
   End
End
Attribute VB_Name = "Form103"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vCountBin As Integer
Dim vSortResult As Integer
Dim vClickKeyQty As Integer

Private Sub Check101_Click()
Dim i As Integer

On Error GoTo ErrDescription

For i = ListView101.ListItems.Count To 1 Step -1
    If Check101.Value = 1 Then
        If ListView101.ListItems.Item(i).ForeColor <> "&H000000FF" Then
            If ListView101.ListItems.Item(i).SubItems(2) <> "0" Then
                ListView101.ListItems.Item(i).Checked = True
            End If
        Else
            ListView101.ListItems.Item(i).Checked = False
        End If
    Else
        ListView101.ListItems.Item(i).Checked = False
    End If
Next i

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub CMD101_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim i As Integer
Dim vItemCode As String
Dim vItemName As String
Dim vBarCode As String
Dim vPrice As Currency
Dim vPromoPrice As Currency
Dim vPriceErect As Currency
Dim vUnitCode As String
Dim vDateStart As Date
Dim vDateEnd As Date
Dim vCountPrice As Integer
Dim vLabelID As String
Dim vLabRandom As Integer

Dim vCountPrintError As Integer
Dim n As Integer
Dim vPrintQty As Integer

Dim vPromotion As String
Dim vPromotionName As String
Dim vSecMan As String
Dim vPromoDocNo As String

On Error Resume Next

If ListView102.ListItems.Count <> 0 And Text103.Text <> "" Then


vPromotion = Left(Trim(Form103.Text101.Text), InStr(Trim(Form103.Text101.Text), "/") - 1)
vPromotionName = Right(Trim(Form103.Text101.Text), Len(Form103.Text101.Text) - InStr(Trim(Form103.Text101.Text), "/"))
vSecMan = Left(Trim(Form103.Text102.Text), InStr(Trim(Form103.Text102.Text), "/") - 1)
vPromoDocNo = Trim(Form103.Text104)

vQuery = "execute USP_PM_SelectItemPrintLabel '" & vPromotion & "','" & vSecMan & "','" & vPromoDocNo & "' "
If OpenDatabase(gConnection, vRecordset, vQuery) <> 0 Then
vRecordset.MoveFirst
vMemIsExpire = Trim(vRecordset.Fields("expire").Value)
End If
vRecordset.Close
            
    For i = ListView102.ListItems.Count To 1 Step -1
            vItemCode = Trim(ListView102.ListItems.Item(i).Text)
            vBarCode = Trim(ListView102.ListItems.Item(i).SubItems(1))
            vItemName = Trim(ListView102.ListItems.Item(i).SubItems(6))
            vPrice = Trim(ListView102.ListItems.Item(i).SubItems(3))
            vPromoPrice = Trim(ListView102.ListItems.Item(i).SubItems(4))
            vPriceErect = Trim(ListView102.ListItems.Item(i).SubItems(2))
            vUnitCode = Trim(ListView102.ListItems.Item(i).SubItems(5))
            vDateStart = Trim(Label101.Caption)
            vDateEnd = Trim(Label102.Caption)
            If InStr(Trim(ListView102.ListItems.Item(i).SubItems(4)), ".") <> 0 Then
                vCountPrice = InStr(Trim(ListView102.ListItems.Item(i).SubItems(4)), ".") - 1   'Len(Trim(ListView102.ListItems.Item(i).SubItems(4)))
            Else
                vCountPrice = Len(Trim(ListView102.ListItems.Item(i).SubItems(4)))
            End If
            
            vPrintQty = Trim(ListView102.ListItems.Item(i).SubItems(8))
            
            vQuery = "set dateformat dmy"
            gConnection.Execute vQuery
            
            If vPrintQty > 1 Then
            For n = 1 To vPrintQty
            vQuery = "execute USP_PM_InsertTempPrintLabel '" & vItemCode & "','" & vBarCode & "','" & vItemName & "'," & vPrice & "," & vPromoPrice & "," & vPriceErect & ",'" & vUnitCode & "'," & vCountPrice & ",'" & vDateStart & "','" & vDateEnd & "','" & vUserID & "','" & vPromotionName & "'  "
            gConnection.Execute vQuery
            Next n
            Else
            vQuery = "execute USP_PM_InsertTempPrintLabel '" & vItemCode & "','" & vBarCode & "','" & vItemName & "'," & vPrice & "," & vPromoPrice & "," & vPriceErect & ",'" & vUnitCode & "'," & vCountPrice & ",'" & vDateStart & "','" & vDateEnd & "','" & vUserID & "','" & vPromotionName & "' "
            gConnection.Execute vQuery
            End If
            
            'ListView102.ListItems.Remove (i)
            If vFormName = Trim("PL_NM_P1001") Or vFormName = Trim("PL_NM_P2001") Or vFormName = Trim("PL_NM_P3001") Or vFormName = Trim("PL_NM_P4001") Then
                vQuery = "exec usp_IV_UpdatePrintUpdateChangePrice '" & vItemCode & "','" & vUnitCode & "','" & vUserID & "','" & vFormName & "' "
                gConnection.Execute vQuery
            End If

    Next i
    ListView101.ListItems.Clear
    vLabelID = vFormName
    vQuery = "select labrandom from npmaster.dbo.tb_pm_label where labid = '" & vLabelID & "' and labrandom = 1 "
    If OpenDatabase(gConnection, vRecordset, vQuery) <> 0 Then
        vLabRandom = 1
    Else
        vLabRandom = 0
    End If
    vRecordset.Close
    
    If vLabRandom = 1 Then
        If vMemIsExpire >= 0 Then
        Call ProcessPrintRandom
        Else
        MsgBox "เอกสารหมดอายุโปรโมชั่น ไม่สามารถพิมพ์ฟอร์ม ป้ายราคาพิเศษได้ กรุณาตรวจสอบ", vbCritical, "Send Error Message"
        Exit Sub
        End If
    Else
        Call ProcessPrint
    End If
    
    Text103.Text = ""
    vQuery = "execute USP_PM_DeleteTempLabel '" & vUserID & "' "
    gConnection.Execute vQuery
    vMemIsExpire = 0
    
Else
    If ListView102.ListItems.Count = 0 Then
        MsgBox "ยังไม่มีรายการสินค้าที่จะพิมพ์ป้ายราคา"
    ElseIf Text103.Text = "" Then
        MsgBox "กรุณาเลือกฟอร์มป้ายราคาที่จะพิมพ์ด้วย"
    End If
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub CMD102_Click()
MDIForm1.Enabled = False
vMemCommand = 2
FormSearchMainPromotion.Show

End Sub

Private Sub CMD103_Click()
MDIForm1.Enabled = False
vMemCommand = 2
FormSearchSecMan.Show
End Sub

Private Sub CMD104_Click()
MDIForm1.Enabled = False
FormSearchLabel.Show
End Sub

Private Sub CMD105_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim i As Integer
Dim vBinItem As ListItem

On Error GoTo ErrDescription

If ListView101.ListItems.Count <> 0 Then
    For i = ListView101.ListItems.Count To 1 Step -1
        If ListView101.ListItems.Item(i).Checked = True Then
            Set vBinItem = ListView102.ListItems.Add(, , Trim(ListView101.ListItems.Item(i).Text))
            vBinItem.SubItems(1) = Trim(ListView101.ListItems.Item(i).SubItems(1))
            vBinItem.SubItems(2) = CheckDegit(Trim(ListView101.ListItems.Item(i).SubItems(2)))
            vBinItem.SubItems(3) = Trim(ListView101.ListItems.Item(i).SubItems(3))
            vBinItem.SubItems(4) = Trim(ListView101.ListItems.Item(i).SubItems(4))
            vBinItem.SubItems(5) = Trim(ListView101.ListItems.Item(i).SubItems(5))
            vBinItem.SubItems(6) = Trim(ListView101.ListItems.Item(i).SubItems(6))
            vBinItem.SubItems(7) = Trim(ListView101.ListItems.Item(i).SubItems(7))
            vBinItem.SubItems(8) = 1
            Label103.Caption = Label103.Caption - 1
            vCountBin = vCountBin + 1
            Label104.Caption = vCountBin
            ListView101.ListItems.Remove (i)
        End If
    Next i
    Check101.Value = 0
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub CMD106_Click()
ListView102.ListItems.Clear
End Sub

Private Sub Form_Load()
vCountBin = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String

On Error Resume Next

If vCheckUsedID = 1 Then
    vCheckSetFocus = 0
    vQuery = "delete npmaster.dbo.TB_CK_UserActivateProgram where userid = '" & vUserID & "' and jobid = 2"
    gConnection.Execute (vQuery)
End If
End Sub

Private Sub ListView101_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo ErrDescription

ListView101.Sorted = True
ListView101.SortKey = ColumnHeader.Index - 1
If vSortResult = 0 Then
    ListView101.SortOrder = lvwAscending
    vSortResult = 1
Else
    ListView101.SortOrder = lvwDescending
    vSortResult = 0
End If


ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub

Private Sub ListView101_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim i As Integer
Dim vPrice As Double
Dim vPriceErect As Double

If ListView101.ListItems.Item(Item.Index).ForeColor = "&H000000FF" Then
    MsgBox "ไม่สามารถเลือกสินค้าที่มีปัญหา มาพิมพ์ป้ายราคาได้ ต้องไปแก้ไขก่อน"
    ListView101.ListItems.Item(Item.Index).Checked = False
End If

vPrice = ListView101.ListItems.Item(Item.Index).SubItems(4)
vPriceErect = ListView101.ListItems.Item(Item.Index).SubItems(2)

If vPrice >= vPriceErect Then
    MsgBox "ไม่สามารถเลือกรายการสินค้านี้ได้ เพราะราคาตั้งน้อยกว่าราคาโปรโมชั่น กรุณาแก้ไขข้อมูลราคาตั้งก่อนนะครับ"
    ListView101.ListItems.Item(Item.Index).Checked = False
End If

If ListView101.ListItems.Item(Item.Index).SubItems(2) = "0" Then
    MsgBox "ไม่สามารถเลือกรายการสินค้านี้ได้ เพราะราคาตั้งเป็น 0 กรุณาแก้ไขข้อมูลราคาตั้งก่อนนะครับ"
    ListView101.ListItems.Item(Item.Index).Checked = False
End If
End Sub

Private Sub ListView101_KeyPress(KeyAscii As Integer)
Dim i As Integer

On Error GoTo ErrDescription

If KeyAscii = 1 Then
For i = ListView101.ListItems.Count To 1 Step -1
    If ListView101.ListItems.Item(i).ForeColor <> "&H000000FF" Then
        If ListView101.ListItems.Item(i).SubItems(2) <> 0 Then
            ListView101.ListItems.Item(i).Selected = True
            ListView101.ListItems.Item(i).Checked = True
            Check101.Value = 1
        End If
    End If
Next i
End If

If KeyAscii = 13 Then
    For i = ListView101.ListItems.Count To 1 Step -1
        If ListView101.ListItems.Item(i).ForeColor <> "&H000000FF" Then
            If ListView101.ListItems.Item(i).SubItems(2) <> 0 Then
                If ListView101.ListItems.Item(i).Selected = True Then
                    ListView101.ListItems.Item(i).Checked = True
                End If
            End If
        End If
    Next i
End If

If KeyAscii = 32 Then
    If vCheckSelect3 = 0 Then
        For i = ListView101.ListItems.Count To 1 Step -1
            If ListView101.ListItems(i).Selected = True Then
                If ListView101.ListItems.Item(i).SubItems(2) <> 0 Then
                    ListView101.ListItems(i).Checked = True
                Else
                    ListView101.ListItems(i).Checked = False
                End If
            End If
        Next i
        vCheckSelect3 = 1
    Else
        For i = ListView101.ListItems.Count To 1 Step -1
            If ListView101.ListItems(i).Selected = True Then
            ListView101.ListItems(i).Checked = False
            End If
        Next i
        vCheckSelect3 = 0
    End If
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub ListView102_DblClick()
Dim vQty As Integer
Dim vQtyWord As String

On Error GoTo ErrDescription

If Me.ListView102.ListItems.Count > 0 Then
vClickKeyQty = Me.ListView102.SelectedItem.Index
vQtyWord = InputBox("กรุณา กรอกจำนวนที่ต้องการพิมพ์", "จำนวนที่ต้องการพิมพ์", 2)

If vQtyWord <> "" Then
vQty = vQtyWord
Else
vQty = 0
End If

If vQty <> 0 Then
Me.ListView102.ListItems(vClickKeyQty).SubItems(8) = vQty
End If

End If


ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub ListView102_KeyDown(KeyCode As Integer, Shift As Integer)
Dim vItemSelect As Integer
Dim i As Integer

On Error GoTo ErrDescription

If KeyCode = 46 Then
    For i = ListView102.ListItems.Count To 1 Step -1
        If ListView102.ListItems.Item(i).Selected = True Then
            ListView102.ListItems.Remove (i)
        End If
    Next i
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub


Public Sub ProcessPrintRandom()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vLabelID As String
Dim vCountLabel As Integer
Dim vArray(4) As Integer
Dim i As Integer
Dim j As Integer
Dim vReportName(5) As String
Dim vLabelName As String

On Error GoTo ErrDescription

i = 0
vLabelID = vFormName
vQuery = "execute USP_PM_LabelRandom '" & vUserID & "' "
If OpenDatabase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    While Not vRecordset.EOF
    i = i + 1
    vArray(i) = Trim(vRecordset.Fields("countprice").Value)
    vRecordset.MoveNext
    Wend
End If
vRecordset.Close

For j = 1 To i Step 1

    If vArray(j) = 2 Then
        vReportName(j) = vLabelID & "1"
    ElseIf vArray(j) = 3 Then
        vReportName(j) = vLabelID & "2"
    ElseIf vArray(j) = 4 Then
        vReportName(j) = vLabelID & "3"
    ElseIf vArray(j) = 5 Then
        vReportName(j) = vLabelID & "4"
    End If
    
        vQuery = "execute USP_PM_LabelNameRandom '" & vLabelID & "','" & vReportName(j) & "' "
        If OpenDatabase(gConnection, vRecordset, vQuery) <> 0 Then
            vLabelName = Trim(vRecordset.Fields("labsubname").Value)
        End If
        vRecordset.Close
        
        With Crystal102
        .ReportFileName = vLabelName & ".rpt"
        .ParameterFields(0) = "@vUserID;" & vUserID & ";true"
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .Action = 1
        End With
        
Next j

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub ListView102_KeyPress(KeyAscii As Integer)
Dim vQty As Integer
Dim vQtyWord As String

On Error GoTo ErrDescription

If KeyAscii = 13 Then
If Me.ListView102.ListItems.Count > 0 Then
vClickKeyQty = Me.ListView102.SelectedItem.Index
vQtyWord = InputBox("กรุณา กรอกจำนวนที่ต้องการพิมพ์", "จำนวนที่ต้องการพิมพ์", 2)

If vQtyWord <> "" Then
vQty = vQtyWord
Else
vQty = 0
End If

If vQty <> 0 Then
Me.ListView102.ListItems(vClickKeyQty).SubItems(8) = vQty
End If

End If
End If


ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Text101_Change()
ListView101.ListItems.Clear
Text102.Text = ""
End Sub

Public Sub ProcessPrint()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vLabelID As String
Dim vCountLabel As Integer
Dim vReportName As String

On Error GoTo ErrDescription


vLabelID = vFormName
If vMemIsExpire <= 0 And vLabelID <> "PL_NM_P1001" Then
    MsgBox "เอกสารหมดอายุโปรโมชั่น ไม่สามารถพิมพ์ฟอร์ม ป้ายราคาพิเศษได้ กรุณาตรวจสอบ", vbCritical, "Send Error Message"
    Exit Sub
End If

vQuery = "select labpath from npmaster.dbo.tb_pm_label where labid = '" & vLabelID & "' "
If OpenDatabase(gConnection, vRecordset, vQuery) <> 0 Then
    vReportName = Trim(vRecordset.Fields("labpath").Value)
End If
vRecordset.Close

With Crystal102
.ReportFileName = vReportName & ".rpt"
.ParameterFields(0) = "@vUserID;" & vUserID & ";true"
.Destination = crptToWindow
.WindowState = crptMaximized
.Action = 1
End With

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub


Private Sub Text104_KeyPress(KeyAscii As Integer)
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocno As String
Dim vCheckIsConfirmStatus As Integer
Dim vCheckDocnoExist As Integer
Dim vPMCode As String

If KeyAscii = 13 And Text104.Text <> "" Then
    vDocno = Text104.Text
    vQuery = "select isnull(count(docno),0) as vCount from npmaster.dbo. TB_PM_Request  where docno = '" & vDocno & "' "
    If OpenDatabase(gConnection, vRecordset, vQuery) <> 0 Then
      vCheckDocnoExist = vRecordset.Fields("vCount").Value
    End If
    vRecordset.Close
    
    If vCheckDocnoExist > 0 Then
    vQuery = "select isnull(isconfirm,0) as isconfirm,isnull(pmcode,'') as pmcode from npmaster.dbo. TB_PM_Request  where docno = '" & vDocno & "' "
    If OpenDatabase(gConnection, vRecordset, vQuery) <> 0 Then
      vCheckIsConfirmStatus = vRecordset.Fields("isconfirm").Value
      vPMCode = vRecordset.Fields("pmcode").Value
    End If
    vRecordset.Close
    
      If vCheckIsConfirmStatus = 0 Then
        MsgBox "เอกสารยังไม่ได้ตรวจสอบและอนุมัติ พิมพ์ป้ายราคาไม่ได้", vbInformation, "Send Message"
        Exit Sub
      ElseIf vCheckIsConfirmStatus = 1 Then
        MsgBox "เอกสารตรวจสอบแล้วแต่ยังไม่ได้อนุมัติ พิมพ์ป้ายราคาไม่ได้", vbInformation, "Send Message"
        Exit Sub
      ElseIf vCheckIsConfirmStatus = 2 Then
        Call SelectItemPromoPrintLabel(vPMCode)
      End If
  Else
    MsgBox "ไม่มีเอกสารนี้ ในระบบโปรโมชั่น กรุณาตรวจสอบ", vbInformation, "Send Message"
    Exit Sub
  End If
End If
End Sub
