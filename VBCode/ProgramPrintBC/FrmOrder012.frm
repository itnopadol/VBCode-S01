VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmOrder012 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form012 ค้นหาพนักงานขนส่ง"
   ClientHeight    =   5010
   ClientLeft      =   3870
   ClientTop       =   2280
   ClientWidth     =   8325
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmOrder012.frx":0000
   ScaleHeight     =   5010
   ScaleWidth      =   8325
   Begin VB.CommandButton CMD103 
      Caption         =   "ออก"
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
      Left            =   7065
      TabIndex        =   5
      Top             =   4455
      Width           =   1005
   End
   Begin VB.CommandButton CMD102 
      Caption         =   "เลือก"
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
      Left            =   5850
      TabIndex        =   3
      Top             =   4455
      Width           =   1005
   End
   Begin MSComctlLib.ListView ListView101 
      Height          =   2940
      Left            =   225
      TabIndex        =   2
      Top             =   1395
      Width           =   7845
      _ExtentX        =   13838
      _ExtentY        =   5186
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
      NumItems        =   11
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ลำดับ"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "ตำแหน่ง"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "รหัสพนักงาน"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "ชื่อพนักงาน"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "EngName"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "คำอธิบาย"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "รหัสตำแหน่ง"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "เลขที่ใบขับขี่"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "วันที่หมดอายุใบขับขี่"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Activate"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton CMD101 
      Height          =   330
      Left            =   4365
      Picture         =   "FrmOrder012.frx":72FB
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "ค้นหา ตามเงื่อนไขที่กรอก"
      Top             =   900
      Width           =   330
   End
   Begin VB.TextBox Text101 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   1845
      TabIndex        =   0
      Top             =   900
      Width           =   2490
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ค้นหาพนักงานขนส่ง"
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
      Left            =   270
      TabIndex        =   4
      Top             =   945
      Width           =   1680
   End
End
Attribute VB_Name = "FrmOrder012"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMD101_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vSearch As String
Dim vSearchList As ListItem
Dim i As Integer

On Error Resume Next

ListView101.ListItems.Clear
vSearch = Trim(Text101.Text)
vQuery = "exec bcnp.dbo.USP_DO_EmpBPlusSearch '" & vSearch & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    i = 1
    While Not vRecordset.EOF
    Set vSearchList = ListView101.ListItems.Add(, , Trim(i))
    vSearchList.SubItems(1) = Trim(vRecordset.Fields("id").Value)
    vSearchList.SubItems(2) = Trim(vRecordset.Fields("positionname").Value)
    vSearchList.SubItems(3) = Trim(vRecordset.Fields("code").Value)
    vSearchList.SubItems(4) = Trim(vRecordset.Fields("name1").Value)
    vSearchList.SubItems(5) = Trim(vRecordset.Fields("name2").Value)
    vSearchList.SubItems(6) = Trim(vRecordset.Fields("MyDescription").Value)
    vSearchList.SubItems(7) = Trim(vRecordset.Fields("position").Value)
    vSearchList.SubItems(8) = Trim(vRecordset.Fields("LicenceNumber").Value)
    vSearchList.SubItems(9) = Trim(vRecordset.Fields("LicenceExpired").Value)
    vSearchList.SubItems(10) = Trim(vRecordset.Fields("ActiveStatus").Value)
    vRecordset.MoveNext
    i = i + 1
    Wend
Else
    MsgBox "ไม่มีข้อมูล ตามคำที่ใช้ค้นหา", vbInformation, "Send Information"
End If
vRecordset.Close
End Sub

Private Sub CMD102_Click()
Dim vListSelectItem As ListItem
Dim i As Integer
Dim j As Integer
Dim vCheckEmp As String
Dim vSelectEmp As String
Dim m As Integer
Dim n As Integer
Dim vSelectOK As Integer
Dim vCount As Integer

On Error GoTo ErrDescription

If ListView101.ListItems.Count <> 0 Then
    If vEmpModule = 1 Then
    
        Call CheckPosition
        
        If ListView101.ListItems.Count > 0 Then
            For i = 1 To ListView101.ListItems.Count
                If ListView101.ListItems.Item(i).Checked = True Then
                    If ListView101.ListItems.Item(i).SubItems(7) = 1 Then
                        vCheckSale = 1
                        vCount = vCount + 1
                        If vCount > 1 Then
                        MsgBox "ไม่สามารถเลือก คนขับรถได้มากกว่า1คน", vbCritical, "Send Error"
                        Exit Sub
                        End If
                    End If
                End If
            Next i
        End If
        
        If vCheckPosition = vCount Then
            MsgBox "ไม่สามารถเลือก คนขับรถได้มากกว่า1คน", vbCritical, "Send Error"
            Exit Sub
        End If
    For m = 1 To ListView101.ListItems.Count
            If FormDelivery.ListView102.ListItems.Count > 0 Then
                j = FormDelivery.ListView102.ListItems.Count
            Else
                j = 0
            End If
        If ListView101.ListItems.Item(m).Checked = True Then
            vCheckEmp = Trim(ListView101.ListItems.Item(m).SubItems(3))
            For n = 1 To FormDelivery.ListView102.ListItems.Count
                vSelectEmp = Trim(FormDelivery.ListView102.ListItems.Item(n).SubItems(2))
            If vCheckEmp = vSelectEmp Then
                vSelectOK = 0
                GoTo Line1
            Else
                vSelectOK = 1
            End If
            Next n
            If FormDelivery.ListView102.ListItems.Count = 0 Then
                vSelectOK = 1
            End If
Line1:
            If vSelectOK = 1 Then
                j = j + 1
                Set vListSelectItem = FormDelivery.ListView102.ListItems.Add(, , Trim(j))
                vListSelectItem.SubItems(1) = Trim(ListView101.ListItems.Item(m).SubItems(1))
                vListSelectItem.SubItems(2) = Trim(ListView101.ListItems.Item(m).SubItems(3))
                vListSelectItem.SubItems(3) = Trim(ListView101.ListItems.Item(m).SubItems(4))
                vListSelectItem.SubItems(4) = Trim(ListView101.ListItems.Item(m).SubItems(2))
                vListSelectItem.SubItems(5) = Format(0, "##,##0.00")
            End If
        End If
    Next m
    End If
    Unload FrmOrder012
End If
ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Private Sub CMD103_Click()
Unload FrmOrder012
End Sub

Private Sub Form_Unload(Cancel As Integer)
MDIFrmProgramPrint.Enabled = True
End Sub

Private Sub ListView101_DblClick()
On Error GoTo ErrDescription

If ListView101.ListItems.Count <> 0 Then
    If vEmpModule = 2 Then
        FrmOrder206.Text101 = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(1))
        FrmOrder206.Text102 = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(3))
        FrmOrder206.Text103 = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(4))
        FrmOrder206.Text104 = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(5))
        If Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(7)) = 1 Then
            FrmOrder206.Text105.Text = Trim("พนักงานขับรถ")
        Else
            FrmOrder206.Text105.Text = Trim("พนักงานติดตามรถ")
        End If
        FrmOrder206.Text106 = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(8))
        FrmOrder206.Text107 = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(6))
        FrmOrder206.DTPicker101 = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(9))
        If Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(10)) = 1 Then
            FrmOrder206.Image101.Visible = True
            FrmOrder206.Image102.Visible = False
        Else
            FrmOrder206.Image101.Visible = False
            FrmOrder206.Image102.Visible = True
        End If
        vCheckEmpOpen = 1
        Unload FrmOrder012
    End If
End If
ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Private Sub ListView101_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim i As Integer

End Sub

Private Sub Text101_KeyPress(KeyAscii As Integer)
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vSearch As String
Dim vSearchList As ListItem
Dim i As Integer

On Error Resume Next

If KeyAscii = 13 Then
    ListView101.ListItems.Clear
    vSearch = Trim(Text101.Text)
    vQuery = "exec bcnp.dbo.USP_DO_EmpBPlusSearch '" & vSearch & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vRecordset.MoveFirst
        i = 1
        While Not vRecordset.EOF
        Set vSearchList = ListView101.ListItems.Add(, , Trim(i))
        vSearchList.SubItems(1) = Trim(vRecordset.Fields("id").Value)
        vSearchList.SubItems(2) = Trim(vRecordset.Fields("positionname").Value)
        vSearchList.SubItems(3) = Trim(vRecordset.Fields("code").Value)
        vSearchList.SubItems(4) = Trim(vRecordset.Fields("name1").Value)
        vSearchList.SubItems(5) = Trim(vRecordset.Fields("name2").Value)
        vSearchList.SubItems(6) = Trim(vRecordset.Fields("MyDescription").Value)
        vSearchList.SubItems(7) = Trim(vRecordset.Fields("position").Value)
        vSearchList.SubItems(8) = Trim(vRecordset.Fields("LicenceNumber").Value)
        vSearchList.SubItems(9) = Trim(vRecordset.Fields("LicenceExpired").Value)
        vSearchList.SubItems(10) = Trim(vRecordset.Fields("ActiveStatus").Value)
        vRecordset.MoveNext
        i = i + 1
        Wend
    Else
        MsgBox "ไม่มีข้อมูล ตามคำที่ใช้ค้นหา", vbInformation, "Send Information"
    End If
    vRecordset.Close
End If
End Sub

Public Sub CheckPosition()
Dim i As Integer

If FormDelivery.ListView102.ListItems.Count > 0 Then
    For i = 1 To FormDelivery.ListView102.ListItems.Count
        If FormDelivery.ListView102.ListItems.Item(i).SubItems(4) = Trim("ขับรถ") Then
            vCheckPosition = 1
            Exit Sub
        Else
            vCheckPosition = 0
        End If
    Next i
Else
    vCheckPosition = 0
End If
End Sub

