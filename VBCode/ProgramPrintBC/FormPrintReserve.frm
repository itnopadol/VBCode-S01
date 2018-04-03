VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "CRYSTL32.OCX"
Begin VB.Form FormPrintReserve 
   Caption         =   "พิมพ์ใบกำกับสินค้า"
   ClientHeight    =   8175
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12000
   Icon            =   "FormPrintReserve.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "FormPrintReserve.frx":1272
   ScaleHeight     =   8175
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin VB.OptionButton OPTA4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "ออกกระดาษ A4"
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   6165
      TabIndex        =   11
      Top             =   1935
      Width           =   2805
   End
   Begin VB.OptionButton OPTA5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "ออกกระดาษครึ่งหน้า"
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   6165
      TabIndex        =   10
      Top             =   1485
      Value           =   -1  'True
      Width           =   2805
   End
   Begin VB.TextBox Text102 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2250
      TabIndex        =   9
      Top             =   1440
      Width           =   2190
   End
   Begin VB.ComboBox CMBForm 
      Height          =   315
      Left            =   6165
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1035
      Width           =   2805
   End
   Begin Crystal.CrystalReport Crystal101 
      Left            =   360
      Top             =   7320
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
   Begin VB.CommandButton CMD102 
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
      Height          =   315
      Left            =   720
      TabIndex        =   2
      Top             =   2025
      Width           =   1440
   End
   Begin VB.CheckBox Check101 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "กำหนดจำนวนพิมพ์ = 1"
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
      Left            =   2250
      TabIndex        =   1
      Top             =   2025
      Width           =   2190
   End
   Begin MSComctlLib.ListView ListView101 
      Height          =   4260
      Left            =   720
      TabIndex        =   3
      Top             =   2430
      Width           =   10470
      _ExtentX        =   18468
      _ExtentY        =   7514
      View            =   3
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
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "รหัสสินค้า"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ชื่อสินค้า"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "จำนวน"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "หน่วยนับ"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "จำนวนพิมพ์"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "การขนส่ง"
         Object.Width           =   2117
      EndProperty
   End
   Begin VB.CommandButton CMD101 
      Caption         =   "พิมพ์เอกสาร"
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
      Left            =   9990
      TabIndex        =   4
      Top             =   6795
      Width           =   1215
   End
   Begin VB.TextBox Text101 
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
      Height          =   315
      Left            =   2250
      TabIndex        =   0
      Top             =   1035
      Width           =   2190
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "รหัสสินค้า :"
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
      Left            =   1080
      TabIndex        =   8
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ฟอร์มที่พิมพ์ :"
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
      Left            =   4545
      TabIndex        =   7
      Top             =   1035
      Width           =   1545
   End
   Begin VB.Label Label2 
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
      Height          =   315
      Left            =   1125
      TabIndex        =   5
      Top             =   1035
      Width           =   1065
   End
End
Attribute VB_Name = "FormPrintReserve"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMD101_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocNo As String
Dim vReportType As String
Dim vReportName As String
Dim vItemCode As String
Dim i As Integer
Dim vRepID As Integer
Dim vRepType As String
Dim vFormID As Integer

On Error GoTo ErrDescription

If Text101.Text <> "" Then
    vDocNo = Trim(Text101.Text)
    vQuery = "exec dbo.usp_SO_ItemReserveSearch '" & vDocNo & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vReportType = Trim(vRecordset.Fields("doctype").Value)
    End If
    vRecordset.Close

    vRepType = "SO"
    If Me.OPTA5.Value = True Then
       vRepID = 405
    ElseIf Me.OPTA4.Value = True Then
       vRepID = 259
    End If

    vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vReportName = Trim(vRecordset.Fields("reportname").Value)
    End If
    vRecordset.Close
        
    For i = 1 To ListView101.ListItems.Count
        If ListView101.ListItems.Item(i).Checked = True Then
                vItemCode = Trim(ListView101.ListItems.Item(i).Text)
                With Crystal101
                .ReportFileName = vReportName & ".rpt"
                .ParameterFields(0) = "@vDocNo;" & vDocNo & ";true"
                .ParameterFields(1) = "@vItemCode;" & vItemCode & ";true"
                .ParameterFields(2) = "@vDocType6;" & vReportType & ";true"
                .DetailCopies = Trim(ListView101.ListItems.Item(i).SubItems(4))
                .WindowState = crptMaximized
                .Destination = crptToWindow
                .Action = 1
                End With
        End If
    Next i
End If
Text101.Text = ""
ListView101.ListItems.Clear

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub CMD102_Click()
Dim i As Integer

On Error GoTo ErrDescription

Check101.Value = 1
For i = 1 To ListView101.ListItems.Count
    ListView101.ListItems.Item(i).Checked = True
Next i

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Form_Load()
Call InsertReportType
End Sub

Public Sub InsertReportType()
Me.CMBForm.AddItem ("สั่งจอง")
Me.CMBForm.AddItem ("รอส่งคืน")
Me.CMBForm.AddItem ("แตกเสียหาย")
Me.CMBForm.ListIndex = 0
End Sub
Private Sub ListView101_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim vPrintCount As String
Dim vPrintCount1 As Integer

On Error GoTo ErrDescription


If Check101.Value = 0 Then
        If ListView101.ListItems.Item(Item.Index).Checked = True Then
            vPrintCount = InputBox("กรอกจำนวนแผ่นที่ต้องการพิมพ์ :", "Data Print", 1)
            If vPrintCount = "" Then
                MsgBox "จำนวนต้องมากกว่า 0 หรือไม่ใช่ค่าว่าง", vbCritical, "Send Error Message"
                ListView101.ListItems.Item(Item.Index).Checked = False
                Me.ListView101.SetFocus
                Exit Sub
            End If

                Call CheckNumber(vPrintCount)
                If vCheckValueNumber = True Then
                   vPrintCount1 = vPrintCount
                Else
                MsgBox "จำนวนต้องมากกว่า 0 หรือไม่ใช่ค่าว่าง", vbCritical, "Send Error Message"
                ListView101.ListItems.Item(Item.Index).Checked = False
                Me.ListView101.SetFocus
                Exit Sub
                End If
                
             If vPrintCount <> "" And vPrintCount <> 0 Then
               ListView101.ListItems.Item(Item.Index).SubItems(4) = vPrintCount
            Else
               MsgBox "จำนวนต้องมากกว่า 0 หรือไม่ใช่ค่าว่าง", vbCritical, "Send Error Message"
               ListView101.ListItems.Item(Item.Index).Checked = False
               Me.ListView101.SetFocus
               Exit Sub
            End If
            ListView101.ListItems.Item(Item.Index).Checked = True
        End If
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Text101_KeyPress(KeyAscii As Integer)
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocNo As String
Dim vListItem As ListItem

On Error GoTo ErrDescription

If KeyAscii = 13 Then
    ListView101.ListItems.Clear
   If Text101.Text <> "" Then
        vDocNo = Trim(Text101.Text)
        vQuery = "exec dbo.usp_SO_ItemReserveSearch '" & vDocNo & "'"
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vRecordset.MoveFirst
            While Not vRecordset.EOF
                Set vListItem = ListView101.ListItems.Add(, , Trim(vRecordset.Fields("itemcode").Value))
                vListItem.SubItems(1) = Trim(vRecordset.Fields("itemname").Value)
                vListItem.SubItems(2) = Trim(vRecordset.Fields("qty").Value)
                vListItem.SubItems(3) = Trim(vRecordset.Fields("unitcode").Value)
                vListItem.SubItems(4) = 1
                vListItem.SubItems(5) = Trim(vRecordset.Fields("isconditionsend").Value)
                vRecordset.MoveNext
            Wend
        Else
            MsgBox "ไม่สามารถพิมพ์ใบกำกับสินค้าได้ เนื่องจากวันครบกำหนดเอกสารเกิน 15 วัน ", vbCritical, "Send Error"
        End If
        Me.ListView101.SetFocus
        vRecordset.Close

   End If
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub
