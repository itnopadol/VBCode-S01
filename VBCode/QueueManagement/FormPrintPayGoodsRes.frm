VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FormPrintPayGoodsRes 
   Caption         =   "พิมพ์ทดแทนใบจ่ายสินค้า"
   ClientHeight    =   8070
   ClientLeft      =   5130
   ClientTop       =   1605
   ClientWidth     =   12090
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8070
   ScaleWidth      =   12090
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport Crystal101 
      Left            =   1305
      Top             =   4905
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin VB.CommandButton CMD102 
      Height          =   285
      Left            =   6165
      Picture         =   "FormPrintPayGoodsRes.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1395
      Width           =   330
   End
   Begin VB.CommandButton CMD101 
      Caption         =   "พิมพ์"
      Height          =   375
      Left            =   5310
      TabIndex        =   6
      Top             =   3915
      Width           =   870
   End
   Begin VB.TextBox Text102 
      Appearance      =   0  'Flat
      Height          =   690
      Left            =   4500
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   3150
      Width           =   4965
   End
   Begin MSComctlLib.ListView ListView101 
      Height          =   1140
      Left            =   4500
      TabIndex        =   3
      Top             =   1845
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   2011
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "คลัง"
         Object.Width           =   2822
      EndProperty
   End
   Begin VB.TextBox Text101 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4500
      TabIndex        =   1
      Top             =   1395
      Width           =   1635
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "เหตุผลการพิมพ์ทดแทน :"
      Height          =   330
      Left            =   2250
      TabIndex        =   5
      Top             =   3150
      Width           =   2130
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "คลัง :"
      Height          =   285
      Left            =   3060
      TabIndex        =   2
      Top             =   1845
      Width           =   1320
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "เลขที่เอกสาร :"
      Height          =   240
      Left            =   2610
      TabIndex        =   0
      Top             =   1395
      Width           =   1770
   End
End
Attribute VB_Name = "FormPrintPayGoodsRes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vQuery As String
Dim vInvoiceNo As String
Dim vPayNumber As String
Dim vReasonDesc  As String


Private Sub CMD101_Click()
Dim vRecordset As New ADODB.Recordset
Dim vPrintDate As Date
Dim vReasonDesc As String
Dim vReserveCode As String
Dim vWHCode As String

If Text102.Text <> "" Then
  vWHCode = Trim(ListView101.SelectedItem.Text)
  vInvoiceNo = Trim(Text101.Text)
  If vWHCode <> "" Then
    If vWHCode = Trim("010") Then
      Call Print010
    ElseIf vWHCode = Trim("012") Then
      Call Print012
    ElseIf vWHCode = Trim("014") Then
      Call Print014
    ElseIf vWHCode = Trim("015") Then
      Call Print015
    ElseIf vWHCode = Trim("016") Then
      Call Print016
    ElseIf vWHCode = Trim("020") Then
      Call Print020
    ElseIf vWHCode = Trim("070") Then
      Call Print070
    ElseIf vWHCode = Trim("097") Then
      Call Print097
    End If
    ListView101.ListItems.Clear
    Text102.Text = ""
    Text101.SetFocus
  Else
    MsgBox "คลิ๊กเลือกคลังที่ต้องการพิมพ์ด้วย", vbCritical, "Send Error"
  End If
Else
  MsgBox "ต้องใส่เหตุผลในการพิมพ์ทดแทนใบจ่ายสินค้าด้วย", vbCritical, "Send Error"
End If
End Sub

Public Sub Print010()
Dim vReportName As String
Dim vDocno As String
Dim vRepID As Integer
Dim vRepType As String
Dim vRecordset As New ADODB.Recordset
Dim vDocno1 As String
Dim vDocGroup1 As String
Dim vCheck As Integer
Dim vWHCode As String
Dim vPrintDate As Date

vDocno = Trim(vInvoiceNo)
vWHCode = "010"
vPrintDate = CDate(Day(Now) & "/" & Month(Now) & "/" & Year(Now))
vQuery = "exec dbo.USP_NP_SearchPayGoodsPrintReserve '" & vInvoiceNo & "','" & vWHCode & "' "
If OpenDataBase(vConnection, vRecordset, vQuery) <> 0 Then
    vPayNumber = Trim(vRecordset.Fields("paynumber").Value)
Else
  vPayNumber = ""
End If
vRecordset.Close

If vPayNumber = "" Then
  MsgBox "ไม่สามารถพิมพ์ใบจ่ายสินค้าน้อยกว่าวันที่ทำงานได้", vbCritical, "Send Error"
  Exit Sub
End If
vReasonDesc = Trim(Text102.Text)

vQuery = "exec dbo.USP_NP_InsertPayGoodsReserve '" & vInvoiceNo & "','" & vPayNumber & "','" & vWHCode & "','" & vPrintDate & "','" & vReasonDesc & "','" & vUserID & "' "
vConnection.Execute vQuery
  
  
vDocno1 = UCase(Left(vDocno, 3))
vDocGroup1 = UCase(vDocno1)

If vDocGroup1 = "ICV" Or vDocGroup1 = "ICN" Or vDocGroup1 = "IVD" Or vDocGroup1 = "IVN" Or vDocGroup1 = "IVM" Then
vRepID = 125
vRepType = "INV"
Else
vRepID = 89
vRepType = "INV"
End If
vCheck = 1
vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
If OpenDataBase(vConnection, vRecordset, vQuery) <> 0 Then
With Crystal101
.ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
.ParameterFields(0) = "@DocNo;" & vDocno & ";true"
.ParameterFields(1) = "@vCheck;" & vCheck & ";true"
.Destination = crptToWindow
.WindowState = crptMaximized
.Action = 1
End With
End If
vRecordset.Close

End Sub

Public Sub Print012()
Dim vReportName As String
Dim vDocno As String
Dim vRepID As Integer
Dim vRepType As String
Dim vRecordset As New ADODB.Recordset
Dim vDocno1 As String
Dim vDocGroup1 As String
Dim vCheck As Integer
Dim vWHCode As String
Dim vPrintDate As Date

vDocno = Trim(vInvoiceNo)
vWHCode = "012"
vPrintDate = CDate(Day(Now) & "/" & Month(Now) & "/" & Year(Now))
vQuery = "exec dbo.USP_NP_SearchPayGoodsPrintReserve '" & vInvoiceNo & "','" & vWHCode & "' "
If OpenDataBase(vConnection, vRecordset, vQuery) <> 0 Then
    vPayNumber = Trim(vRecordset.Fields("paynumber").Value)
Else
  vPayNumber = ""
End If
vRecordset.Close

If vPayNumber = "" Then
  MsgBox "ไม่สามารถพิมพ์ใบจ่ายสินค้าน้อยกว่าวันที่ทำงานได้", vbCritical, "Send Error"
  Exit Sub
End If
vReasonDesc = Trim(Text102.Text)

vQuery = "exec dbo.USP_NP_InsertPayGoodsReserve '" & vInvoiceNo & "','" & vPayNumber & "','" & vWHCode & "','" & vPrintDate & "','" & vReasonDesc & "','" & vUserID & "' "
vConnection.Execute vQuery
  
  
vDocno1 = UCase(Left(vDocno, 3))
vDocGroup1 = UCase(vDocno1)

If vDocGroup1 = "ICV" Or vDocGroup1 = "ICN" Or vDocGroup1 = "IVD" Or vDocGroup1 = "IVN" Or vDocGroup1 = "IVM" Then
vRepID = 129
vRepType = "INV"
Else
vRepID = 91
vRepType = "INV"
End If
vCheck = 1
vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
If OpenDataBase(vConnection, vRecordset, vQuery) <> 0 Then
With Crystal101
.ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
.ParameterFields(0) = "@DocNo;" & vDocno & ";true"
.ParameterFields(1) = "@vCheck;" & vCheck & ";true"
.Destination = crptToWindow
.WindowState = crptMaximized
.Action = 1
End With
End If
vRecordset.Close

End Sub

Public Sub Print014()
Dim vReportName As String
Dim vDocno As String
Dim vRepID As Integer
Dim vRepType As String
Dim vRecordset As New ADODB.Recordset
Dim vDocno1 As String
Dim vDocGroup1 As String
Dim vCheck As Integer
Dim vWHCode As String
Dim vPrintDate As Date

vDocno = Trim(vInvoiceNo)
vWHCode = "014"
vPrintDate = CDate(Day(Now) & "/" & Month(Now) & "/" & Year(Now))
vQuery = "exec dbo.USP_NP_SearchPayGoodsPrintReserve '" & vInvoiceNo & "','" & vWHCode & "' "
If OpenDataBase(vConnection, vRecordset, vQuery) <> 0 Then
    vPayNumber = Trim(vRecordset.Fields("paynumber").Value)
Else
  vPayNumber = ""
End If
vRecordset.Close

If vPayNumber = "" Then
  MsgBox "ไม่สามารถพิมพ์ใบจ่ายสินค้าน้อยกว่าวันที่ทำงานได้", vbCritical, "Send Error"
  Exit Sub
End If
vReasonDesc = Trim(Text102.Text)

vQuery = "exec dbo.USP_NP_InsertPayGoodsReserve '" & vInvoiceNo & "','" & vPayNumber & "','" & vWHCode & "','" & vPrintDate & "','" & vReasonDesc & "','" & vUserID & "' "
vConnection.Execute vQuery
  
  
vDocno1 = UCase(Left(vDocno, 3))
vDocGroup1 = UCase(vDocno1)

If vDocGroup1 = "ICV" Or vDocGroup1 = "ICN" Or vDocGroup1 = "IVD" Or vDocGroup1 = "IVN" Or vDocGroup1 = "IVM" Then
vRepID = 170
vRepType = "INV"
Else
vRepID = 69
vRepType = "INV"
End If
vCheck = 1
vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
If OpenDataBase(vConnection, vRecordset, vQuery) <> 0 Then
With Crystal101
.ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
.ParameterFields(0) = "@DocNo;" & vDocno & ";true"
.ParameterFields(1) = "@vCheck;" & vCheck & ";true"
.Destination = crptToWindow
.WindowState = crptMaximized
.Action = 1
End With
End If
vRecordset.Close

End Sub

Public Sub Print015()
Dim vReportName As String
Dim vDocno As String
Dim vRepID As Integer
Dim vRepType As String
Dim vRecordset As New ADODB.Recordset
Dim vDocno1 As String
Dim vDocGroup1 As String
Dim vCheck As Integer
Dim vWHCode As String
Dim vPrintDate As Date

vDocno = Trim(vInvoiceNo)
vWHCode = "015"
vPrintDate = CDate(Day(Now) & "/" & Month(Now) & "/" & Year(Now))
vQuery = "exec dbo.USP_NP_SearchPayGoodsPrintReserve '" & vInvoiceNo & "','" & vWHCode & "' "
If OpenDataBase(vConnection, vRecordset, vQuery) <> 0 Then
    vPayNumber = Trim(vRecordset.Fields("paynumber").Value)
Else
  vPayNumber = ""
End If
vRecordset.Close

If vPayNumber = "" Then
  MsgBox "ไม่สามารถพิมพ์ใบจ่ายสินค้าน้อยกว่าวันที่ทำงานได้", vbCritical, "Send Error"
  Exit Sub
End If
vReasonDesc = Trim(Text102.Text)

vQuery = "exec dbo.USP_NP_InsertPayGoodsReserve '" & vInvoiceNo & "','" & vPayNumber & "','" & vWHCode & "','" & vPrintDate & "','" & vReasonDesc & "','" & vUserID & "' "
vConnection.Execute vQuery
  
  
vDocno1 = UCase(Left(vDocno, 3))
vDocGroup1 = UCase(vDocno1)

If vDocGroup1 = "ICV" Or vDocGroup1 = "ICN" Or vDocGroup1 = "IVD" Or vDocGroup1 = "IVN" Or vDocGroup1 = "IVM" Then
vRepID = 131
vRepType = "INV"
Else
vRepID = 92
vRepType = "INV"
End If
vCheck = 1
vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
If OpenDataBase(vConnection, vRecordset, vQuery) <> 0 Then
With Crystal101
.ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
.ParameterFields(0) = "@DocNo;" & vDocno & ";true"
.ParameterFields(1) = "@vCheck;" & vCheck & ";true"
.Destination = crptToWindow
.WindowState = crptMaximized
.Action = 1
End With
End If
vRecordset.Close

End Sub

Public Sub Print016()
Dim vReportName As String
Dim vDocno As String
Dim vRepID As Integer
Dim vRepType As String
Dim vRecordset As New ADODB.Recordset
Dim vDocno1 As String
Dim vDocGroup1 As String
Dim vCheck As Integer
Dim vWHCode As String
Dim vPrintDate As Date

vDocno = Trim(vInvoiceNo)
vWHCode = "016"
vPrintDate = CDate(Day(Now) & "/" & Month(Now) & "/" & Year(Now))
vQuery = "exec dbo.USP_NP_SearchPayGoodsPrintReserve '" & vInvoiceNo & "','" & vWHCode & "' "
If OpenDataBase(vConnection, vRecordset, vQuery) <> 0 Then
    vPayNumber = Trim(vRecordset.Fields("paynumber").Value)
Else
  vPayNumber = ""
End If
vRecordset.Close

If vPayNumber = "" Then
  MsgBox "ไม่สามารถพิมพ์ใบจ่ายสินค้าน้อยกว่าวันที่ทำงานได้", vbCritical, "Send Error"
  Exit Sub
End If
vReasonDesc = Trim(Text102.Text)

vQuery = "exec dbo.USP_NP_InsertPayGoodsReserve '" & vInvoiceNo & "','" & vPayNumber & "','" & vWHCode & "','" & vPrintDate & "','" & vReasonDesc & "','" & vUserID & "' "
vConnection.Execute vQuery
  
  
vDocno1 = UCase(Left(vDocno, 3))
vDocGroup1 = UCase(vDocno1)

If vDocGroup1 = "ICV" Or vDocGroup1 = "ICN" Or vDocGroup1 = "IVD" Or vDocGroup1 = "IVN" Or vDocGroup1 = "IVM" Then
vRepID = 217
vRepType = "INV"
Else
vRepID = 216
vRepType = "INV"
End If
vCheck = 1
vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
If OpenDataBase(vConnection, vRecordset, vQuery) <> 0 Then
With Crystal101
.ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
.ParameterFields(0) = "@DocNo;" & vDocno & ";true"
.ParameterFields(1) = "@vCheck;" & vCheck & ";true"
.Destination = crptToWindow
.WindowState = crptMaximized
.Action = 1
End With
End If
vRecordset.Close

End Sub

Public Sub Print020()
Dim vReportName As String
Dim vDocno As String
Dim vRepID As Integer
Dim vRepType As String
Dim vRecordset As New ADODB.Recordset
Dim vDocno1 As String
Dim vDocGroup1 As String
Dim vCheck As Integer
Dim vWHCode As String
Dim vPrintDate As Date

vDocno = Trim(vInvoiceNo)
vWHCode = "020"
vPrintDate = CDate(Day(Now) & "/" & Month(Now) & "/" & Year(Now))
vQuery = "exec dbo.USP_NP_SearchPayGoodsPrintReserve '" & vInvoiceNo & "','" & vWHCode & "' "
If OpenDataBase(vConnection, vRecordset, vQuery) <> 0 Then
    vPayNumber = Trim(vRecordset.Fields("paynumber").Value)
Else
  vPayNumber = ""
End If
vRecordset.Close

If vPayNumber = "" Then
  MsgBox "ไม่สามารถพิมพ์ใบจ่ายสินค้าน้อยกว่าวันที่ทำงานได้", vbCritical, "Send Error"
  Exit Sub
End If
vReasonDesc = Trim(Text102.Text)

vQuery = "exec dbo.USP_NP_InsertPayGoodsReserve '" & vInvoiceNo & "','" & vPayNumber & "','" & vWHCode & "','" & vPrintDate & "','" & vReasonDesc & "','" & vUserID & "' "
vConnection.Execute vQuery
  
  
vDocno1 = UCase(Left(vDocno, 3))
vDocGroup1 = UCase(vDocno1)

If vDocGroup1 = "ICV" Or vDocGroup1 = "ICN" Or vDocGroup1 = "IVD" Or vDocGroup1 = "IVN" Or vDocGroup1 = "IVM" Then
vRepID = 194
vRepType = "INV"
Else
vRepID = 193
vRepType = "INV"
End If
vCheck = 1
vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
If OpenDataBase(vConnection, vRecordset, vQuery) <> 0 Then
With Crystal101
.ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
.ParameterFields(0) = "@DocNo;" & vDocno & ";true"
.ParameterFields(1) = "@vCheck;" & vCheck & ";true"
.Destination = crptToWindow
.WindowState = crptMaximized
.Action = 1
End With
End If
vRecordset.Close

End Sub

Public Sub Print070()
Dim vReportName As String
Dim vDocno As String
Dim vRepID As Integer
Dim vRepType As String
Dim vRecordset As New ADODB.Recordset
Dim vDocno1 As String
Dim vDocGroup1 As String
Dim vCheck As Integer
Dim vWHCode As String
Dim vPrintDate As Date

vDocno = Trim(vInvoiceNo)
vWHCode = "070"
vPrintDate = CDate(Day(Now) & "/" & Month(Now) & "/" & Year(Now))
vQuery = "exec dbo.USP_NP_SearchPayGoodsPrintReserve '" & vInvoiceNo & "','" & vWHCode & "' "
If OpenDataBase(vConnection, vRecordset, vQuery) <> 0 Then
    vPayNumber = Trim(vRecordset.Fields("paynumber").Value)
Else
  vPayNumber = ""
End If
vRecordset.Close

If vPayNumber = "" Then
  MsgBox "ไม่สามารถพิมพ์ใบจ่ายสินค้าน้อยกว่าวันที่ทำงานได้", vbCritical, "Send Error"
  Exit Sub
End If
vReasonDesc = Trim(Text102.Text)

vQuery = "exec dbo.USP_NP_InsertPayGoodsReserve '" & vInvoiceNo & "','" & vPayNumber & "','" & vWHCode & "','" & vPrintDate & "','" & vReasonDesc & "','" & vUserID & "' "
vConnection.Execute vQuery
  
  
vDocno1 = UCase(Left(vDocno, 3))
vDocGroup1 = UCase(vDocno1)

If vDocGroup1 = "ICV" Or vDocGroup1 = "ICN" Or vDocGroup1 = "IVD" Or vDocGroup1 = "IVN" Or vDocGroup1 = "IVM" Then
vRepID = 298
vRepType = "INV"
Else
vRepID = 296
vRepType = "INV"
End If
vCheck = 1
vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
If OpenDataBase(vConnection, vRecordset, vQuery) <> 0 Then
With Crystal101
.ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
.ParameterFields(0) = "@DocNo;" & vDocno & ";true"
.ParameterFields(1) = "@vCheck;" & vCheck & ";true"
.Destination = crptToWindow
.WindowState = crptMaximized
.Action = 1
End With
End If
vRecordset.Close

End Sub

Public Sub Print097()
Dim vReportName As String
Dim vDocno As String
Dim vRepID As Integer
Dim vRepType As String
Dim vRecordset As New ADODB.Recordset
Dim vDocno1 As String
Dim vDocGroup1 As String
Dim vCheck As Integer
Dim vWHCode As String
Dim vPrintDate As Date

vDocno = Trim(vInvoiceNo)
vWHCode = "097"
vPrintDate = CDate(Day(Now) & "/" & Month(Now) & "/" & Year(Now))
vQuery = "exec dbo.USP_NP_SearchPayGoodsPrintReserve '" & vInvoiceNo & "','" & vWHCode & "' "
If OpenDataBase(vConnection, vRecordset, vQuery) <> 0 Then
    vPayNumber = Trim(vRecordset.Fields("paynumber").Value)
Else
  vPayNumber = ""
End If
vRecordset.Close

If vPayNumber = "" Then
  MsgBox "ไม่สามารถพิมพ์ใบจ่ายสินค้าน้อยกว่าวันที่ทำงานได้", vbCritical, "Send Error"
  Exit Sub
End If
vReasonDesc = Trim(Text102.Text)

vQuery = "exec dbo.USP_NP_InsertPayGoodsReserve '" & vInvoiceNo & "','" & vPayNumber & "','" & vWHCode & "','" & vPrintDate & "','" & vReasonDesc & "','" & vUserID & "' "
vConnection.Execute vQuery
  
  
vDocno1 = UCase(Left(vDocno, 3))
vDocGroup1 = UCase(vDocno1)

If vDocGroup1 = "ICV" Or vDocGroup1 = "ICN" Or vDocGroup1 = "IVD" Or vDocGroup1 = "IVN" Or vDocGroup1 = "IVM" Then
vRepID = 133
vRepType = "INV"
Else
vRepID = 97
vRepType = "INV"
End If
vCheck = 1
vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
If OpenDataBase(vConnection, vRecordset, vQuery) <> 0 Then
With Crystal101
.ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
.ParameterFields(0) = "@DocNo;" & vDocno & ";true"
.ParameterFields(1) = "@vCheck;" & vCheck & ";true"
.Destination = crptToWindow
.WindowState = crptMaximized
.Action = 1
End With
End If
vRecordset.Close

End Sub

Private Sub CMD102_Click()
Dim vRecordset As New ADODB.Recordset
Dim vDocno As String
Dim vListItem As ListItem

vDocno = Trim(Text101.Text)
If vDocno <> "" Then
  vQuery = "exec dbo.USP_NP_InvoiceGroupWareHouseRes '" & vDocno & "'," & vSelectZoneID & " "
  If OpenDataBase(vConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    ListView101.ListItems.Clear
    While Not vRecordset.EOF
      Set vListItem = ListView101.ListItems.Add(, , Trim(vRecordset.Fields("whcode").Value))
    vRecordset.MoveNext
    Wend
  Else
    MsgBox "ไม่มีข้อมูลเลขที่ใบจ่ายดังกล่าว  อาจยังไม่ได้พิมพ์ที่หน้าคิวพิมพ์ใบจ่าย หรือ อาจถูกยกเลิกไปแล้ว  กรุณาตรวจสอบอีกครั้ง", vbCritical, "Send Error"
    Exit Sub
  End If
  vRecordset.Close
Else
  MsgBox "ไม่มีข้อมูลเลขที่เอกสารขายที่ต้องการพิมพ์", vbCritical, "Send Error"
End If
End Sub

