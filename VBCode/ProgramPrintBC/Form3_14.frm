VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form3_14 
   Caption         =   "หน้ายกเลิกเอกสาร Quotation และ BackOrder"
   ClientHeight    =   8385
   ClientLeft      =   3000
   ClientTop       =   675
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form3_14.frx":0000
   ScaleHeight     =   8385
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin VB.OptionButton Option2 
      BackColor       =   &H80000009&
      Caption         =   "ยกเลิกการอนุมัติเอกสาร"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1575
      TabIndex        =   6
      Top             =   1725
      Width           =   2940
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H80000009&
      Caption         =   "ยกเลิกเอกสาร"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1575
      TabIndex        =   5
      Top             =   1275
      Value           =   -1  'True
      Width           =   2940
   End
   Begin VB.CommandButton CMDCancel 
      Caption         =   "ยกเลิกเอกสาร"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1575
      TabIndex        =   2
      Top             =   5850
      Width           =   1365
   End
   Begin MSComctlLib.ListView ListViewCanCelDocs 
      Height          =   2940
      Left            =   1575
      TabIndex        =   1
      Top             =   2700
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   5186
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
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
         Text            =   "รหัสสินค้า"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "ชื่อสินค้า"
         Object.Width           =   6703
      EndProperty
   End
   Begin VB.TextBox TXTCancel 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1575
      TabIndex        =   0
      Top             =   2100
      Width           =   2940
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ยกเลิกเอกสารใบเสนอราคา และ ใบ BackOrder ตามรายตัวสินค้า"
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
      Height          =   465
      Left            =   2925
      TabIndex        =   4
      Top             =   300
      Width           =   7065
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
      Height          =   390
      Left            =   600
      TabIndex        =   3
      Top             =   2100
      Width           =   990
   End
End
Attribute VB_Name = "Form3_14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMDCancel_Click()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vItemCode As String
Dim i As Integer, vCount As Integer
Dim vDocNo As String, vDiffQty As String
Dim vCheck As Boolean
Dim vListItems As ListItem
Dim vCheckQty As Integer
Dim vCountCancel As Integer
Dim vAmount As Double

On Error GoTo ErrDescription

For i = 1 To ListViewCanCelDocs.ListItems.Count
vCheck = ListViewCanCelDocs.ListItems(i).Checked
If vCheck = True Then
                    vDocNo = ListViewCanCelDocs.ListItems(i).Text
                    vItemCode = ListViewCanCelDocs.ListItems(i).ListSubItems.Item(2).Text
                    vQuery = "select qty-remainqty as diffQty from bcquotationsub where docno = '" & vDocNo & "' and itemcode = '" & vItemCode & "' "
                    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                        vDiffQty = vRecordset.Fields("diffqty").Value
                    End If
                    vRecordset.Close
                    If vDiffQty = 0 Then
                        vQuery = "update dbo.bcquotationsub  set iscancel = 1 ,remainqty = 0 where docno = '" & vDocNo & "' and itemcode = '" & vItemCode & "' "
                        gConnection.Execute vQuery
                    Else
                    vQuery = "Update bcquotationsub  set qty = " & vDiffQty & ",remainqty = 0 , amount =( " & vDiffQty & "*price ), " _
                                    & "     netamount = (" & vDiffQty & "*price)-(((" & vDiffQty & "*price)*100)/107) ,homeamount = (" & vDiffQty & "*price)-(((" & vDiffQty & "*price)*100)/107) ,iscancel = 1  " _
                                    & "     where docno = '" & vDocNo & "' and itemcode = '" & vItemCode & "' "
                    gConnection.Execute vQuery
                    End If
End If
Next i
vQuery = "select count(itemcode) as itemcode from bcquotationsub where docno = '" & vDocNo & "'"
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vCount = vRecordset.Fields("itemcode").Value
Else
    vCount = 0
End If
vRecordset.Close

If vCount <> 0 Then
            vQuery = "select sum(amount) as amount from bcquotationsub where docno = '" & vDocNo & "' "
            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                If IsNull(vRecordset.Fields("amount").Value) Then
                vAmount = 0
                Else
                vAmount = vRecordset.Fields("amount").Value
                End If
            End If
            vRecordset.Close
        
            vQuery = "update bcquotation set sumofitemamount = " & vAmount & "," _
                            & "     afterdiscount = " & vAmount & ", " _
                            & "     beforetaxamount = " & vAmount & "-(((" & vAmount & ")*100)/107) , " _
                            & "      taxamount = ((" & vAmount & ")*100)/107 , " _
                            & "      totalamount = " & vAmount & ",  " _
                            & "      netamount = " & vAmount & " " _
                            & "     where docno = '" & vDocNo & "' "
            gConnection.Execute vQuery
Else
            vQuery = "update dbo.bcquotation set iscancel = 1 ,isconfirm = 0,billstatus = 1where docno = '" & vDocNo & "' "
            GetConnect
            conn.Execute vQuery
End If
ListViewCanCelDocs.ListItems.Clear
vQuery = "select docno,docdate,itemcode,itemname from bcnp.dbo.bcquotationsub where docno = '" & vDocNo & "' and remainqty <> 0"
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
   Do Until vRecordset.EOF
    Set vListItems = ListViewCanCelDocs.ListItems.Add(, , vRecordset.Fields("Docno").Value)
    vListItems.SubItems(1) = vRecordset.Fields("docdate").Value
    vListItems.SubItems(2) = vRecordset.Fields("itemcode").Value
    vListItems.SubItems(3) = vRecordset.Fields("itemname").Value
    vRecordset.MoveNext
    Loop
End If
vRecordset.Close

vQuery = "select count(itemcode) as vSumQTY  from dbo.bcquotationsub where docno = '" & vDocNo & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vCheckQty = vRecordset.Fields("vSumQTY").Value
End If
vRecordset.Close
            
vQuery = "select count(itemcode) as vSumQTY  from dbo.bcquotationsub where docno = '" & vDocNo & "' and iscancel = 1"
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vCountCancel = vRecordset.Fields("vSumQTY").Value
End If
vRecordset.Close

If vCheckQty = vCountCancel Then
        vQuery = "update dbo.bcquotation set iscancel = 1 ,isconfirm = 0,billstatus = 1 where docno = '" & vDocNo & "' "
        gConnection.Execute vQuery
End If
MsgBox "ยกเลิกเอกสารเลขที่ " & vDocNo & " เรียบร้อยแล้ว "

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub TXTCancel_KeyPress(KeyAscii As Integer)
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocNo As String
Dim vItemList As ListItem
Dim vAnswer As Integer
Dim vCheckExist As Integer
Dim vCheckRef As Integer
Dim vRecordset1 As New ADODB.Recordset
Dim vRecordset2 As New ADODB.Recordset
Dim vItemCode As String
Dim vCheckLineRef As Integer
Dim vCheckStkRequest As Integer


On Error GoTo ErrDescription
If KeyAscii = 13 And Me.TXTCancel.Text <> "" Then
  If Option1.Value = True Then
    
    vDocNo = Trim(TXTCancel.Text)
    vQuery = "select isnull(count(backorderno),0) as vCount  from npmaster.dbo.tb_pr_generate where backorderno = '" & vDocNo & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
       vCheckStkRequest = vRecordset.Fields("vCount").Value
    End If
    vRecordset.Close
    
    If vCheckStkRequest > 0 Then
       MsgBox "เอกสาร เลขที่ '" & vDocNo & "' ได้ทำเอกสารเสนอซื้อสินค้าไปแล้วไม่สามารถยกเลิกเอกสารได้ โปรดติดต่อ แผนกจัดซื้อ กรณีต้องการยกเลิกการขายสินค้าในเอกสารดังกล่าว"
       Exit Sub
    End If
    
    ListViewCanCelDocs.ListItems.Clear
    vQuery = "select  isnull(count(docno),0) as vCount  from dbo.bcquotation where docno = '" & vDocNo & "'  and iscancel = 0"
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
      vCheckExist = vRecordset.Fields("vCount").Value
    End If
    vRecordset.Close
    
    If vCheckExist = 1 Then
      vQuery = "select isnull(count(stkreserveno),0) as vCount  from dbo.bcsaleordersub where stkreserveno  = '" & vDocNo & "' and iscancel = 0"
      If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vCheckRef = vRecordset.Fields("vCount").Value
      End If
      vRecordset.Close
      
      If vCheckRef = 0 Then
        vQuery = "select docno,docdate,itemcode,itemname from bcnp.dbo.bcquotationsub where docno = '" & vDocNo & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vRecordset.MoveFirst
          Do Until vRecordset.EOF
           Set vItemList = ListViewCanCelDocs.ListItems.Add(, , vRecordset.Fields("Docno").Value)
           vItemList.SubItems(1) = vRecordset.Fields("docdate").Value
           vItemList.SubItems(2) = vRecordset.Fields("itemcode").Value
           vItemList.SubItems(3) = vRecordset.Fields("itemname").Value
           vRecordset.MoveNext
           Loop
        End If
        vRecordset.Close
      Else
        vQuery = "select docno ,itemcode from dbo.bcquotationsub where docno = '" & vDocNo & "' "
        If OpenDataBase(gConnection, vRecordset1, vQuery) <> 0 Then
        vRecordset1.MoveFirst
        While Not vRecordset1.EOF
        vItemCode = Trim(vRecordset1.Fields("itemcode").Value)
        
        vQuery = "select isnull(count(stkreserveno),0) as vCount  from dbo.bcsaleordersub where stkreserveno  = '" & vDocNo & "' and itemcode = '" & vItemCode & "' and iscancel = 0"
        If OpenDataBase(gConnection, vRecordset2, vQuery) <> 0 Then
            vCheckLineRef = vRecordset2.Fields("vCount").Value
        End If
        vRecordset2.Close
        
        If vCheckLineRef = 0 Then
        vQuery = "select docno,docdate,itemcode,itemname from bcnp.dbo.bcquotationsub where docno = '" & vDocNo & "' and itemcode = '" & vItemCode & "' and iscancel = 0"
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vRecordset.MoveFirst
         Do Until vRecordset.EOF
          Set vItemList = ListViewCanCelDocs.ListItems.Add(, , vRecordset.Fields("Docno").Value)
          vItemList.SubItems(1) = vRecordset.Fields("docdate").Value
          vItemList.SubItems(2) = vRecordset.Fields("itemcode").Value
          vItemList.SubItems(3) = vRecordset.Fields("itemname").Value
          vRecordset.MoveNext
          Loop
        End If
        vRecordset.Close
        
        End If
        
        vRecordset1.MoveNext
        Wend
      End If
      vRecordset1.Close
      End If
    End If
    
  ElseIf Option2.Value = True Then
    vDocNo = Trim(TXTCancel.Text)
    vAnswer = MsgBox("คุณต้องการยกเลิกการอนุมัติเอกสาร เลขที่ " & vDocNo & " ใช่หรือไม่", vbOKCancel, "ข้อความสอบถาม")
    If vAnswer = 1 Then
      vQuery = "Update BCQuotation set IsConfirm = 0 , BillStatus = 0 where docno = '" & vDocNo & "' "
      gConnection.Execute vQuery
      TXTCancel.Text = ""
    Else
    Exit Sub
    End If
  End If
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub
