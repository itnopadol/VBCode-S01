VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form Form310 
   Caption         =   "หน้าพิมพ์ใบจ่ายสินค้า"
   ClientHeight    =   8340
   ClientLeft      =   3675
   ClientTop       =   1560
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form310.frx":0000
   ScaleHeight     =   8340
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin VB.ComboBox CMBFamily 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2475
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   4095
      Width           =   7125
   End
   Begin VB.ComboBox CMBZone 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2475
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   3420
      Width           =   1320
   End
   Begin VB.ComboBox CMBWHCode 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2475
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   2115
      Width           =   1320
   End
   Begin Crystal.CrystalReport CrystalReport310 
      Left            =   1320
      Top             =   6255
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
   Begin VB.CommandButton CMD3101 
      Caption         =   "พิมพ์ทดแทน"
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
      Left            =   2385
      TabIndex        =   5
      Top             =   5220
      Width           =   1350
   End
   Begin VB.ComboBox CMB3101 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2475
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   2760
      Width           =   1320
   End
   Begin VB.TextBox TXT3101 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2475
      TabIndex        =   0
      Top             =   1425
      Width           =   2940
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "กลุ่มสินค้า :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1125
      TabIndex        =   11
      Top             =   4095
      Width           =   1185
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "โซน :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1125
      TabIndex        =   10
      Top             =   3465
      Width           =   1140
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "คลัง :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   990
      TabIndex        =   9
      Top             =   2115
      Width           =   1275
   End
   Begin VB.Label LBL3103 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ชั้นเก็บที่พิมพ์ :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   300
      Left            =   450
      TabIndex        =   8
      Top             =   2790
      Width           =   1830
   End
   Begin VB.Label LBL3102 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "เลขที่ใบเหลือง :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   300
      Left            =   540
      TabIndex        =   7
      Top             =   1440
      Width           =   1725
   End
   Begin VB.Label LBL3101 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "หน้าพิมพ์ใบจ่ายสินค้า (ใบเหลือง)"
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
      Left            =   3015
      TabIndex        =   6
      Top             =   225
      Width           =   8490
   End
End
Attribute VB_Name = "Form310"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMB3101_Change()
'Dim vQuery As String
'Dim vRecordset As New ADODB.Recordset
'Dim vDocNo As String
'Dim vWHCode As String
'Dim vShelfCode As String
'Dim n As Integer


'If Me.TXT3101.Text <> "" And Me.CMBWHCode.Text <> "" Then
 ' vDocNo = Me.TXT3101.Text
  'vWHCode = Me.CMBWHCode.Text
  'vShelfCode = Me.CMB3101.Text
      
   
   'vQuery = "exec dbo.USP_INV_SearchShelfPrintSlip2 '" & vDocNo & "' "
   'If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
   'Me.CMBZone.Clear
   'vRecordset.MoveFirst
   'While Not vRecordset.EOF
    '  CMBZone.AddItem Trim(vRecordset.Fields("pickzone").Value)
   'vRecordset.MoveNext
   'Wend
   'vRecordset.Close
   'Else
   'Me.CMBZone.Clear
   'End If
   
'End If
End Sub

Private Sub CMBWHCode_Change()
'Dim vQuery As String
'Dim vRecordset As New ADODB.Recordset
'Dim vDocNo As String
'Dim vWHCode As String
'Dim n As Integer

'If Me.TXT3101.Text <> "" Then
 '  vDocNo = Me.TXT3101.Text
  ' vWHCode = Me.CMBWHCode.Text
   
   'vQuery = "exec dbo.USP_INV_SearchWHCodePrintSlip '" & vDocNo & "','" & vWHCode & "' "
   'If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
   'Me.CMB3101.Clear
   'vRecordset.MoveFirst
   'While Not vRecordset.EOF
   'CMB3101.AddItem Trim(vRecordset.Fields("shelfcode").Value)
   'vRecordset.MoveNext
   'Wend
   'vRecordset.Close
   'Else
   'Me.CMB3101.Clear
   'End If
   
'End If
End Sub

Private Sub CMD3101_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vShelfCode As String
Dim vWHCode As String
Dim vDocNo As String
Dim vIsCompleteSave As Integer
Dim vTableName As String
Dim vPickZone As String
Dim vZoneID As String
Dim vFamilyCode As String



On Error GoTo ErrDescription

If Me.TXT3101.Text = "" Then
Me.TXT3101.SetFocus
Exit Sub
End If

If Me.CMBWHCode.Text = "" Then
MsgBox "กรุณาเลือกคลัง", vbCritical, "Send Error Message"
Me.CMBWHCode.SetFocus
Exit Sub
End If

If Me.CMB3101.Text = "" Then
MsgBox "กรุณาเลือกชั้นเก็บ", vbCritical, "Send Error Message"
Me.CMB3101.SetFocus
Exit Sub
End If

If Me.CMBZone.Text = "" Then
MsgBox "กรุณาเลือกโซน", vbCritical, "Send Error Message"
Me.CMBZone.SetFocus
Exit Sub
End If

If Me.CMBFamily.Text = "" Then
MsgBox "กรุณาเลือกกลุ่มสินค้า", vbCritical, "Send Error Message"
Me.CMBFamily.SetFocus
Exit Sub
End If

vWHCode = Me.CMBWHCode.Text
vShelfCode = Trim(CMB3101.Text)
vDocNo = Trim(TXT3101.Text)
vTableName = "BCARInvoice"
vZoneID = Me.CMBZone.Text
vFamilyCode = Me.CMBFamily.Text

vQuery = "exec dbo.USP_NP_SearchIsCompleteSave '" & vTableName & "' ,'" & vDocNo & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vIsCompleteSave = vRecordset.Fields("iscompletesave").Value
End If
vRecordset.Close

If vIsCompleteSave = 0 Then
MsgBox "เอกสารยังบันทึกไม่สมบูรณ์ ยังไม่สามารถพิมพ์ได้ กรุณารอสักครู่", vbCritical, "Send Error Message"
Me.CMD3101.SetFocus
Exit Sub
End If


 'vQuery = "exec dbo.USP_INV_SearchBillPayItemZone '" & vDocNo & "','" & vWHCode & "','" & vShelfCode & "','" & vPickZone & "' "
 vQuery = "exec dbo.USP_INV_NewSearchInvoicePrintSlip '" & vDocNo & "','" & vWHCode & "','" & vShelfCode & "','" & vZoneID & "','" & vFamilyCode & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vRecordset.MoveFirst
  vPickZone = Trim(vRecordset.Fields("pickzone").Value)
End If
vRecordset.Close


If vDocNo <> "" Then
Call PrintBillPayItemZone(vDocNo, vWHCode, vShelfCode, vZoneID, vPickZone, vFamilyCode)
Else
   MsgBox "กรุณาใส่เลขที่เอกสารด้วยนะครับ", vbInformation + vbCritical, "ข้อความเตือน"
End If



ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
End If

End Sub

Public Sub PrintItem_AVL01(WHCode As String, PickZone As String, ZoneID As String)
Dim vReportName As String
Dim vDocNo As String
Dim vRepID As Integer
Dim vRepType As String
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vAutoNumber As String, vDocNo1 As String, vDocGroup1 As String
Dim vCheck As Integer
Dim vGenerateNumber, vHeader As String
Dim vWHCode As String
Dim vShelfCode As String
Dim vMydescription As String
Dim vCheckNumber As String
Dim vCount  As Integer

On Error GoTo ErrDescription

vDocNo = Trim(Me.TXT3101.Text)
vWHCode = WHCode
vShelfCode = Trim(CMB3101.Text)
vQuery = "select * from npmaster.dbo.NP_PayGoods where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vCheckNumber = Trim(vRecordset.Fields("invoiceno").Value)
End If
vRecordset.Close

vQuery = "Update npmaster.dbo.NP_PayGoods set PrintedTimeStamp = getdate() where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
gConnection.Execute vQuery

If vCheckNumber = "" Then
vQuery = "select header,autonumber  from npmaster.dbo.NP_Generate_DocNo where headertype = 8 "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vAutoNumber = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
vHeader = Trim(vRecordset.Fields("header").Value)
End If
vRecordset.Close
vGenerateNumber = vHeader & "-" & vAutoNumber
vMydescription = "ทดแทนการเปลี่ยนคลังของเอกสาร"

vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1 where headertype = 8 "
gConnection.Execute vQuery

vQuery = "exec dbo.USP_INV_InsertPrintSlip1 '" & vDocNo & "','" & vGenerateNumber & "','" & vWHCode & "','" & vShelfCode & "','" & ZoneID & "','" & PickZone & "','" & vMydescription & "','" & vUserID & "' "
gConnection.Execute vQuery

End If
        
'If UCase(Left(vDocNo, 3)) = "S01" Or UCase(Left(vDocNo, 3)) = "S02" Then
 '  vDocGroup1 = Left(Right(TXT011.Text, Len(TXT011.Text) - InStr(TXT011.Text, "-")), 3) 'Left(Trim(TXT011.Text), 3)
'Else
 '  vDocGroup1 = Left(Trim(vDocNo), 3)
'End If

If UCase(Left(vDocNo, 3)) = "S01" Or UCase(Left(vDocNo, 3)) = "S02" Or UCase(Left(vDocNo, 3)) = "W01" Or UCase(Left(vDocNo, 3)) = "W02" Then
   vDocNo1 = UCase(Left(Right(vDocNo, Len(vDocNo) - InStr(vDocNo, "-")), 3)) 'UCase(Left(vDocNo, 3))
Else
   vDocNo1 = UCase(Left(vDocNo, 3))
End If

vDocGroup1 = vDocNo1
If vDocGroup1 = "ICV" Or vDocGroup1 = "ICN" Or vDocGroup1 = "IVD" Or vDocGroup1 = "IVN" Or vDocGroup1 = "IVM" Then
vRepID = 426
vRepType = "INV"
Else
vRepID = 424
vRepType = "INV"
End If
vCheck = 1

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
With CrystalReport310
.ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
.ParameterFields(0) = "@vDocNo;" & vDocNo & ";true"
.ParameterFields(1) = "@vPickZone;" & PickZone & ";true"
.Destination = crptToWindow
.WindowState = crptMaximized
.Action = 1
End With
End If
vRecordset.Close
                            
vQuery = "select lastprintcount,whcode  from npmaster.dbo.np_paygoods where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vCount = Trim(vRecordset.Fields("lastprintcount").Value)
End If
vRecordset.Close

vCount = vCount + 1
vQuery = "Update npmaster.dbo.np_paygoods  set lastprintcount = " & vCount & " ,lastuserprint = '" & vUserID & "' where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
gConnection.Execute vQuery
        
                
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If
End Sub

Public Sub PrintBillPayItemZone(vDocNo As String, vWHCode As String, vShelfCode As String, vZoneID As String, vPickZone As String, vFamilyCode As String)
Dim vReportName As String
Dim vRepID As Integer
Dim vRepType As String
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vRecordset1 As New ADODB.Recordset
Dim vAutoNumber As String, vDocNo1 As String, vDocGroup1 As String
Dim vCheck As Integer
Dim vGenerateNumber, vHeader As String
Dim vMydescription As String
Dim vCheckNumber As String
Dim vCount  As Integer
Dim vNextZoneNumber As Integer
Dim vExQuery As String


On Error GoTo ErrDescription


vQuery = "select * from npmaster.dbo.NP_PayGoods where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' and zoneid = '" & vZoneID & "' and pickzone ='" & vPickZone & "' and familycode = '" & vFamilyCode & "'"
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vCheckNumber = Trim(vRecordset.Fields("invoiceno").Value)
End If
vRecordset.Close

vQuery = "Update npmaster.dbo.NP_PayGoods set PrintedTimeStamp = getdate() where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' and zoneid = '" & vZoneID & "' and pickzone = '" & vPickZone & "' and familycode = '" & vFamilyCode & "'"
gConnection.Execute vQuery

If vCheckNumber = "" Then
    vQuery = "select header,autonumber  from npmaster.dbo.NP_Generate_DocNo where headertype = 8 "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vAutoNumber = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
    vHeader = Trim(vRecordset.Fields("header").Value)
    End If
    vRecordset.Close
    vGenerateNumber = vHeader & "-" & vAutoNumber
    vMydescription = "ทดแทนการเปลี่ยนคลังของเอกสาร"
    
    vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1 where headertype = 8 "
    gConnection.Execute vQuery
    
    'vQuery = "exec dbo.USP_INV_InsertPrintSlip1 '" & vDocNo & "','" & vGenerateNumber & "','" & vWHCode & "','" & vShelfCode & "','" & vZoneID & "','" & vPickZone & "','" & vMydescription & "','" & vUserID & "' "
    'gConnection.Execute vQuery
    
    vQuery = "exec dbo.USP_INV_InsertInvoicePrintSlip '" & vDocNo & "','" & vGenerateNumber & "','" & vWHCode & "','" & vShelfCode & "','" & vFamilyCode & "','" & vZoneID & "','" & vPickZone & "','','" & vUserID & "' "
    gConnection.Execute vQuery
    
    
    vQuery = "exec dbo.USP_NP_SearchMaxZoneNumber  '" & vPickZone & "' "
    If OpenDataBase(gConnection, vRecordset1, vQuery) <> 0 Then
        vNextZoneNumber = Trim(vRecordset1.Fields("zonenumber").Value)
        
        vExQuery = "exec dbo.USP_NP_UpdatePayGoodRunningZone '" & vDocNo & "','" & vWHCode & "','" & vShelfCode & "','" & vPickZone & "','" & vPickZone & "','" & vGenerateNumber & "'," & vNextZoneNumber & " "
        gConnection.Execute vExQuery
            
    
    End If
    vRecordset1.Close


End If
        

If UCase(Left(vDocNo, 3)) = "S01" Or UCase(Left(vDocNo, 3)) = "S02" Or UCase(Left(vDocNo, 3)) = "W01" Or UCase(Left(vDocNo, 3)) = "W02" Then
   vDocNo1 = UCase(Left(Right(vDocNo, Len(vDocNo) - InStr(vDocNo, "-")), 3)) 'UCase(Left(vDocNo, 3))
Else
   vDocNo1 = UCase(Left(vDocNo, 3))
End If

vDocGroup1 = vDocNo1
If vDocGroup1 = "ICV" Or vDocGroup1 = "ICN" Or vDocGroup1 = "IVD" Or vDocGroup1 = "IVN" Or vDocGroup1 = "IVM" Then
vRepID = 526
vRepType = "INV"
Else
vRepID = 524
vRepType = "INV"
End If
vCheck = 0

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
With CrystalReport310
.ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
.ParameterFields(0) = "@vDocNo;" & vDocNo & ";true"
.ParameterFields(1) = "@vWHCode;" & vWHCode & ";true"
.ParameterFields(2) = "@vShelfCode;" & vShelfCode & ";true"
.ParameterFields(3) = "@vZoneID;" & vZoneID & ";true"
.ParameterFields(4) = "@vCheck;" & vCheck & ";true"
.ParameterFields(5) = "@vFamilyCode;" & vFamilyCode & ";true"
.Destination = crptToWindow
.WindowState = crptMaximized
.Action = 1
End With
End If
vRecordset.Close
                            
vQuery = "select lastprintcount,whcode  from npmaster.dbo.np_paygoods where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vCount = Trim(vRecordset.Fields("lastprintcount").Value)
End If
vRecordset.Close

vCount = vCount + 1
vQuery = "Update npmaster.dbo.np_paygoods  set lastprintcount = " & vCount & " ,lastuserprint = '" & vUserID & "' where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
gConnection.Execute vQuery
        
                
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If
End Sub



Public Sub PrintItem_AVL02(WHCode As String, PickZone As String, ZoneID As String)
Dim vReportName As String
Dim vDocNo As String
Dim vRepID As Integer
Dim vRepType As String
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vAutoNumber As String, vDocNo1 As String, vDocGroup1 As String
Dim vCheck As Integer
Dim vGenerateNumber, vHeader As String
Dim vWHCode As String
Dim vShelfCode As String
Dim vMydescription As String
Dim vCheckNumber As String
Dim vCount  As Integer

On Error GoTo ErrDescription

vDocNo = Trim(Me.TXT3101.Text)
vWHCode = WHCode
vShelfCode = Trim(CMB3101.Text)
vQuery = "select * from npmaster.dbo.NP_PayGoods where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vCheckNumber = Trim(vRecordset.Fields("invoiceno").Value)
End If
vRecordset.Close

vQuery = "Update npmaster.dbo.NP_PayGoods set PrintedTimeStamp = getdate() where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
gConnection.Execute vQuery

If vCheckNumber = "" Then
vQuery = "select header,autonumber  from npmaster.dbo.NP_Generate_DocNo where headertype = 8 "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vAutoNumber = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
vHeader = Trim(vRecordset.Fields("header").Value)
End If
vRecordset.Close
vGenerateNumber = vHeader & "-" & vAutoNumber
vMydescription = "ทดแทนการเปลี่ยนคลังของเอกสาร"

vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1 where headertype = 8 "
gConnection.Execute vQuery

vQuery = "exec dbo.USP_INV_InsertPrintSlip1 '" & vDocNo & "','" & vGenerateNumber & "','" & vWHCode & "','" & vShelfCode & "','" & ZoneID & "','" & PickZone & "','" & vMydescription & "','" & vUserID & "' "
gConnection.Execute vQuery

End If
        
'If UCase(Left(vDocNo, 3)) = "S01" Or UCase(Left(vDocNo, 3)) = "S02" Then
 '  vDocGroup1 = Left(Right(TXT011.Text, Len(TXT011.Text) - InStr(TXT011.Text, "-")), 3) 'Left(Trim(TXT011.Text), 3)
'Else
 '  vDocGroup1 = Left(Trim(vDocNo), 3)
'End If

If UCase(Left(vDocNo, 3)) = "S01" Or UCase(Left(vDocNo, 3)) = "S02" Or UCase(Left(vDocNo, 3)) = "W01" Or UCase(Left(vDocNo, 3)) = "W02" Then
   vDocNo1 = UCase(Left(Right(vDocNo, Len(vDocNo) - InStr(vDocNo, "-")), 3)) 'UCase(Left(vDocNo, 3))
Else
   vDocNo1 = UCase(Left(vDocNo, 3))
End If

vDocGroup1 = vDocNo1
If vDocGroup1 = "ICV" Or vDocGroup1 = "ICN" Or vDocGroup1 = "IVD" Or vDocGroup1 = "IVN" Or vDocGroup1 = "IVM" Then
vRepID = 426
vRepType = "INV"
Else
vRepID = 424
vRepType = "INV"
End If
vCheck = 1

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
With CrystalReport310
.ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
.ParameterFields(0) = "@vDocNo;" & vDocNo & ";true"
.ParameterFields(1) = "@vPickZone;" & PickZone & ";true"
.Destination = crptToWindow
.WindowState = crptMaximized
.Action = 1
End With
End If
vRecordset.Close
                            
vQuery = "select lastprintcount,whcode  from npmaster.dbo.np_paygoods where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vCount = Trim(vRecordset.Fields("lastprintcount").Value)
End If
vRecordset.Close

vCount = vCount + 1
vQuery = "Update npmaster.dbo.np_paygoods  set lastprintcount = " & vCount & " ,lastuserprint = '" & vUserID & "' where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
gConnection.Execute vQuery
        
                
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If
End Sub

Public Sub PrintItem_BAK01(WHCode As String, PickZone As String, ZoneID As String)
Dim vReportName As String
Dim vDocNo As String
Dim vRepID As Integer
Dim vRepType As String
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vAutoNumber As String, vDocNo1 As String, vDocGroup1 As String
Dim vCheck As Integer
Dim vGenerateNumber, vHeader As String
Dim vWHCode As String
Dim vShelfCode As String
Dim vMydescription As String
Dim vCheckNumber As String
Dim vCount  As Integer

On Error GoTo ErrDescription

vDocNo = Trim(Me.TXT3101.Text)
vWHCode = WHCode
vShelfCode = Trim(CMB3101.Text)
vQuery = "select * from npmaster.dbo.NP_PayGoods where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vCheckNumber = Trim(vRecordset.Fields("invoiceno").Value)
End If
vRecordset.Close

vQuery = "Update npmaster.dbo.NP_PayGoods set PrintedTimeStamp = getdate() where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
gConnection.Execute vQuery

If vCheckNumber = "" Then
vQuery = "select header,autonumber  from npmaster.dbo.NP_Generate_DocNo where headertype = 8 "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vAutoNumber = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
vHeader = Trim(vRecordset.Fields("header").Value)
End If
vRecordset.Close
vGenerateNumber = vHeader & "-" & vAutoNumber
vMydescription = "ทดแทนการเปลี่ยนคลังของเอกสาร"

vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1 where headertype = 8 "
gConnection.Execute vQuery

vQuery = "exec dbo.USP_INV_InsertPrintSlip1 '" & vDocNo & "','" & vGenerateNumber & "','" & vWHCode & "','" & vShelfCode & "','" & ZoneID & "','" & PickZone & "','" & vMydescription & "','" & vUserID & "' "
gConnection.Execute vQuery

End If
        

'vDocno1 = UCase(Left(Right(vDocNo, Len(vDocNo) - InStr(vDocNo, "-")), 3)) 'UCase(Left(vDocNo, 3))

If UCase(Left(vDocNo, 3)) = "S01" Or UCase(Left(vDocNo, 3)) = "S02" Or UCase(Left(vDocNo, 3)) = "W01" Or UCase(Left(vDocNo, 3)) = "W02" Then
   vDocNo1 = UCase(Left(Right(vDocNo, Len(vDocNo) - InStr(vDocNo, "-")), 3)) 'UCase(Left(vDocNo, 3))
Else
   vDocNo1 = UCase(Left(vDocNo, 3))
End If

vDocGroup1 = vDocNo1
If vDocGroup1 = "ICV" Or vDocGroup1 = "ICN" Or vDocGroup1 = "IVD" Or vDocGroup1 = "IVN" Or vDocGroup1 = "IVM" Then
vRepID = 485
vRepType = "INV"
Else
vRepID = 484
vRepType = "INV"
End If
vCheck = 1

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
With CrystalReport310
.ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
.ParameterFields(0) = "@vDocNo;" & vDocNo & ";true"
.ParameterFields(1) = "@vPickZone;" & PickZone & ";true"
.Destination = crptToWindow
.WindowState = crptMaximized
.Action = 1
End With
End If
vRecordset.Close
                            
vQuery = "select lastprintcount,whcode  from npmaster.dbo.np_paygoods where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vCount = Trim(vRecordset.Fields("lastprintcount").Value)
End If
vRecordset.Close

vCount = vCount + 1
vQuery = "Update npmaster.dbo.np_paygoods  set lastprintcount = " & vCount & " ,lastuserprint = '" & vUserID & "' where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
gConnection.Execute vQuery
        
                
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If
End Sub


Public Sub PrintItem_BAK02(WHCode As String, PickZone As String, ZoneID As String)
Dim vReportName As String
Dim vDocNo As String
Dim vRepID As Integer
Dim vRepType As String
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vAutoNumber As String, vDocNo1 As String, vDocGroup1 As String
Dim vCheck As Integer
Dim vGenerateNumber, vHeader As String
Dim vWHCode As String
Dim vShelfCode As String
Dim vMydescription As String
Dim vCheckNumber As String
Dim vCount  As Integer

On Error GoTo ErrDescription

vDocNo = Trim(Me.TXT3101.Text)
vWHCode = WHCode
vShelfCode = Trim(CMB3101.Text)
vQuery = "select * from npmaster.dbo.NP_PayGoods where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vCheckNumber = Trim(vRecordset.Fields("invoiceno").Value)
End If
vRecordset.Close

vQuery = "Update npmaster.dbo.NP_PayGoods set PrintedTimeStamp = getdate() where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
gConnection.Execute vQuery

If vCheckNumber = "" Then
vQuery = "select header,autonumber  from npmaster.dbo.NP_Generate_DocNo where headertype = 8 "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vAutoNumber = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
vHeader = Trim(vRecordset.Fields("header").Value)
End If
vRecordset.Close
vGenerateNumber = vHeader & "-" & vAutoNumber
vMydescription = "ทดแทนการเปลี่ยนคลังของเอกสาร"

vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1 where headertype = 8 "
gConnection.Execute vQuery

vQuery = "exec dbo.USP_INV_InsertPrintSlip1 '" & vDocNo & "','" & vGenerateNumber & "','" & vWHCode & "','" & vShelfCode & "','" & ZoneID & "','" & PickZone & "','" & vMydescription & "','" & vUserID & "' "
gConnection.Execute vQuery

End If
        

'vDocno1 = UCase(Left(Right(vDocNo, Len(vDocNo) - InStr(vDocNo, "-")), 3)) 'UCase(Left(vDocNo, 3))

If UCase(Left(vDocNo, 3)) = "S01" Or UCase(Left(vDocNo, 3)) = "S02" Or UCase(Left(vDocNo, 3)) = "W01" Or UCase(Left(vDocNo, 3)) = "W02" Then
   vDocNo1 = UCase(Left(Right(vDocNo, Len(vDocNo) - InStr(vDocNo, "-")), 3)) 'UCase(Left(vDocNo, 3))
Else
   vDocNo1 = UCase(Left(vDocNo, 3))
End If

vDocGroup1 = vDocNo1
If vDocGroup1 = "ICV" Or vDocGroup1 = "ICN" Or vDocGroup1 = "IVD" Or vDocGroup1 = "IVN" Or vDocGroup1 = "IVM" Then
vRepID = 485
vRepType = "INV"
Else
vRepID = 484
vRepType = "INV"
End If
vCheck = 1

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
With CrystalReport310
.ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
.ParameterFields(0) = "@vDocNo;" & vDocNo & ";true"
.ParameterFields(1) = "@vPickZone;" & PickZone & ";true"
.Destination = crptToWindow
.WindowState = crptMaximized
.Action = 1
End With
End If
vRecordset.Close
                            
vQuery = "select lastprintcount,whcode  from npmaster.dbo.np_paygoods where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vCount = Trim(vRecordset.Fields("lastprintcount").Value)
End If
vRecordset.Close

vCount = vCount + 1
vQuery = "Update npmaster.dbo.np_paygoods  set lastprintcount = " & vCount & " ,lastuserprint = '" & vUserID & "' where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
gConnection.Execute vQuery
        
                
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If
End Sub

Public Sub PrintItem_BAK04(WHCode As String, PickZone As String, ZoneID As String)
Dim vReportName As String
Dim vDocNo As String
Dim vRepID As Integer
Dim vRepType As String
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vAutoNumber As String, vDocNo1 As String, vDocGroup1 As String
Dim vCheck As Integer
Dim vGenerateNumber, vHeader As String
Dim vWHCode As String
Dim vShelfCode As String
Dim vMydescription As String
Dim vCheckNumber As String
Dim vCount  As Integer

On Error GoTo ErrDescription

vDocNo = Trim(Me.TXT3101.Text)
vWHCode = WHCode
vShelfCode = Trim(CMB3101.Text)
vQuery = "select * from npmaster.dbo.NP_PayGoods where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vCheckNumber = Trim(vRecordset.Fields("invoiceno").Value)
End If
vRecordset.Close

vQuery = "Update npmaster.dbo.NP_PayGoods set PrintedTimeStamp = getdate() where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
gConnection.Execute vQuery

If vCheckNumber = "" Then
vQuery = "select header,autonumber  from npmaster.dbo.NP_Generate_DocNo where headertype = 8 "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vAutoNumber = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
vHeader = Trim(vRecordset.Fields("header").Value)
End If
vRecordset.Close
vGenerateNumber = vHeader & "-" & vAutoNumber
vMydescription = "ทดแทนการเปลี่ยนคลังของเอกสาร"

vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1 where headertype = 8 "
gConnection.Execute vQuery

vQuery = "exec dbo.USP_INV_InsertPrintSlip1 '" & vDocNo & "','" & vGenerateNumber & "','" & vWHCode & "','" & vShelfCode & "','" & ZoneID & "','" & PickZone & "','" & vMydescription & "','" & vUserID & "' "
gConnection.Execute vQuery

End If
        

'vDocno1 = UCase(Left(Right(vDocNo, Len(vDocNo) - InStr(vDocNo, "-")), 3)) 'UCase(Left(vDocNo, 3))

If UCase(Left(vDocNo, 3)) = "S01" Or UCase(Left(vDocNo, 3)) = "S02" Or UCase(Left(vDocNo, 3)) = "W01" Or UCase(Left(vDocNo, 3)) = "W02" Then
   vDocNo1 = UCase(Left(Right(vDocNo, Len(vDocNo) - InStr(vDocNo, "-")), 3)) 'UCase(Left(vDocNo, 3))
Else
   vDocNo1 = UCase(Left(vDocNo, 3))
End If

vDocGroup1 = vDocNo1
If vDocGroup1 = "ICV" Or vDocGroup1 = "ICN" Or vDocGroup1 = "IVD" Or vDocGroup1 = "IVN" Or vDocGroup1 = "IVM" Then
vRepID = 485
vRepType = "INV"
Else
vRepID = 484
vRepType = "INV"
End If
vCheck = 1

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
With CrystalReport310
.ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
.ParameterFields(0) = "@vDocNo;" & vDocNo & ";true"
.ParameterFields(1) = "@vPickZone;" & PickZone & ";true"
.Destination = crptToWindow
.WindowState = crptMaximized
.Action = 1
End With
End If
vRecordset.Close
                            
vQuery = "select lastprintcount,whcode  from npmaster.dbo.np_paygoods where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vCount = Trim(vRecordset.Fields("lastprintcount").Value)
End If
vRecordset.Close

vCount = vCount + 1
vQuery = "Update npmaster.dbo.np_paygoods  set lastprintcount = " & vCount & " ,lastuserprint = '" & vUserID & "' where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
gConnection.Execute vQuery
        
                
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If
End Sub

Public Sub PrintItem_BAK05(WHCode As String, PickZone As String, ZoneID As String)
Dim vReportName As String
Dim vDocNo As String
Dim vRepID As Integer
Dim vRepType As String
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vAutoNumber As String, vDocNo1 As String, vDocGroup1 As String
Dim vCheck As Integer
Dim vGenerateNumber, vHeader As String
Dim vWHCode As String
Dim vShelfCode As String
Dim vMydescription As String
Dim vCheckNumber As String
Dim vCount  As Integer

On Error GoTo ErrDescription

vDocNo = Trim(Me.TXT3101.Text)
vWHCode = WHCode
vShelfCode = Trim(CMB3101.Text)
vQuery = "select * from npmaster.dbo.NP_PayGoods where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vCheckNumber = Trim(vRecordset.Fields("invoiceno").Value)
End If
vRecordset.Close

vQuery = "Update npmaster.dbo.NP_PayGoods set PrintedTimeStamp = getdate() where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
gConnection.Execute vQuery

If vCheckNumber = "" Then
vQuery = "select header,autonumber  from npmaster.dbo.NP_Generate_DocNo where headertype = 8 "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vAutoNumber = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
vHeader = Trim(vRecordset.Fields("header").Value)
End If
vRecordset.Close
vGenerateNumber = vHeader & "-" & vAutoNumber
vMydescription = "ทดแทนการเปลี่ยนคลังของเอกสาร"

vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1 where headertype = 8 "
gConnection.Execute vQuery

vQuery = "exec dbo.USP_INV_InsertPrintSlip1 '" & vDocNo & "','" & vGenerateNumber & "','" & vWHCode & "','" & vShelfCode & "','" & ZoneID & "','" & PickZone & "','" & vMydescription & "','" & vUserID & "' "
gConnection.Execute vQuery

End If
        

'vDocno1 = UCase(Left(Right(vDocNo, Len(vDocNo) - InStr(vDocNo, "-")), 3)) 'UCase(Left(vDocNo, 3))

If UCase(Left(vDocNo, 3)) = "S01" Or UCase(Left(vDocNo, 3)) = "S02" Or UCase(Left(vDocNo, 3)) = "W01" Or UCase(Left(vDocNo, 3)) = "W02" Then
   vDocNo1 = UCase(Left(Right(vDocNo, Len(vDocNo) - InStr(vDocNo, "-")), 3)) 'UCase(Left(vDocNo, 3))
Else
   vDocNo1 = UCase(Left(vDocNo, 3))
End If

vDocGroup1 = vDocNo1
If vDocGroup1 = "ICV" Or vDocGroup1 = "ICN" Or vDocGroup1 = "IVD" Or vDocGroup1 = "IVN" Or vDocGroup1 = "IVM" Then
vRepID = 485
vRepType = "INV"
Else
vRepID = 484
vRepType = "INV"
End If
vCheck = 1

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
With CrystalReport310
.ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
.ParameterFields(0) = "@vDocNo;" & vDocNo & ";true"
.ParameterFields(1) = "@vPickZone;" & PickZone & ";true"
.Destination = crptToWindow
.WindowState = crptMaximized
.Action = 1
End With
End If
vRecordset.Close
                            
vQuery = "select lastprintcount,whcode  from npmaster.dbo.np_paygoods where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vCount = Trim(vRecordset.Fields("lastprintcount").Value)
End If
vRecordset.Close

vCount = vCount + 1
vQuery = "Update npmaster.dbo.np_paygoods  set lastprintcount = " & vCount & " ,lastuserprint = '" & vUserID & "' where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
gConnection.Execute vQuery
        
                
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If
End Sub

Public Sub PrintItem_BAK06(WHCode As String, PickZone As String, ZoneID As String)
Dim vReportName As String
Dim vDocNo As String
Dim vRepID As Integer
Dim vRepType As String
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vAutoNumber As String, vDocNo1 As String, vDocGroup1 As String
Dim vCheck As Integer
Dim vGenerateNumber, vHeader As String
Dim vWHCode As String
Dim vShelfCode As String
Dim vMydescription As String
Dim vCheckNumber As String
Dim vCount  As Integer

On Error GoTo ErrDescription

vDocNo = Trim(Me.TXT3101.Text)
vWHCode = WHCode
vShelfCode = Trim(CMB3101.Text)
vQuery = "select * from npmaster.dbo.NP_PayGoods where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vCheckNumber = Trim(vRecordset.Fields("invoiceno").Value)
End If
vRecordset.Close

vQuery = "Update npmaster.dbo.NP_PayGoods set PrintedTimeStamp = getdate() where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
gConnection.Execute vQuery

If vCheckNumber = "" Then
vQuery = "select header,autonumber  from npmaster.dbo.NP_Generate_DocNo where headertype = 8 "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vAutoNumber = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
vHeader = Trim(vRecordset.Fields("header").Value)
End If
vRecordset.Close
vGenerateNumber = vHeader & "-" & vAutoNumber
vMydescription = "ทดแทนการเปลี่ยนคลังของเอกสาร"

vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1 where headertype = 8 "
gConnection.Execute vQuery

vQuery = "exec dbo.USP_INV_InsertPrintSlip1 '" & vDocNo & "','" & vGenerateNumber & "','" & vWHCode & "','" & vShelfCode & "','" & ZoneID & "','" & PickZone & "','" & vMydescription & "','" & vUserID & "' "
gConnection.Execute vQuery

End If
        

'vDocno1 = UCase(Left(Right(vDocNo, Len(vDocNo) - InStr(vDocNo, "-")), 3)) 'UCase(Left(vDocNo, 3))

If UCase(Left(vDocNo, 3)) = "S01" Or UCase(Left(vDocNo, 3)) = "S02" Or UCase(Left(vDocNo, 3)) = "W01" Or UCase(Left(vDocNo, 3)) = "W02" Then
   vDocNo1 = UCase(Left(Right(vDocNo, Len(vDocNo) - InStr(vDocNo, "-")), 3)) 'UCase(Left(vDocNo, 3))
Else
   vDocNo1 = UCase(Left(vDocNo, 3))
End If

vDocGroup1 = vDocNo1
If vDocGroup1 = "ICV" Or vDocGroup1 = "ICN" Or vDocGroup1 = "IVD" Or vDocGroup1 = "IVN" Or vDocGroup1 = "IVM" Then
vRepID = 485
vRepType = "INV"
Else
vRepID = 484
vRepType = "INV"
End If
vCheck = 1

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
With CrystalReport310
.ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
.ParameterFields(0) = "@vDocNo;" & vDocNo & ";true"
.ParameterFields(1) = "@vPickZone;" & PickZone & ";true"
.Destination = crptToWindow
.WindowState = crptMaximized
.Action = 1
End With
End If
vRecordset.Close
                            
vQuery = "select lastprintcount,whcode  from npmaster.dbo.np_paygoods where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vCount = Trim(vRecordset.Fields("lastprintcount").Value)
End If
vRecordset.Close

vCount = vCount + 1
vQuery = "Update npmaster.dbo.np_paygoods  set lastprintcount = " & vCount & " ,lastuserprint = '" & vUserID & "' where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
gConnection.Execute vQuery
        
                
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If
End Sub


Public Sub PrintItem_BK1(WHCode As String)
Dim vReportName As String
Dim vDocNo As String
Dim vRepID As Integer
Dim vRepType As String
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vAutoNumber As String, vDocNo1 As String, vDocGroup1 As String
Dim vCheck As Integer
Dim vGenerateNumber, vHeader As String
Dim vWHCode As String
Dim vShelfCode As String
Dim vMydescription As String
Dim vCheckNumber As String
Dim vCount  As Integer

On Error GoTo ErrDescription

vDocNo = Trim(Me.TXT3101.Text)
vWHCode = WHCode
vShelfCode = Trim(CMB3101.Text)
vQuery = "select * from npmaster.dbo.NP_PayGoods where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vCheckNumber = Trim(vRecordset.Fields("invoiceno").Value)
End If
vRecordset.Close

vQuery = "Update npmaster.dbo.NP_PayGoods set PrintedTimeStamp = getdate() where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
gConnection.Execute vQuery

If vCheckNumber = "" Then
vQuery = "select header,autonumber  from npmaster.dbo.NP_Generate_DocNo where headertype = 8 "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vAutoNumber = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
vHeader = Trim(vRecordset.Fields("header").Value)
End If
vRecordset.Close
vGenerateNumber = vHeader & "-" & vAutoNumber
vMydescription = "ทดแทนการเปลี่ยนคลังของเอกสาร"

vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1 where headertype = 8 "
gConnection.Execute vQuery

vQuery = "exec dbo.USP_INV_InsertPrintSlip '" & vDocNo & "','" & vGenerateNumber & "','" & vWHCode & "','" & vShelfCode & "','" & vMydescription & "','" & vUserID & "' "
gConnection.Execute vQuery

End If
        

'vDocno1 = UCase(Left(Right(vDocNo, Len(vDocNo) - InStr(vDocNo, "-")), 3)) 'UCase(Left(vDocNo, 3))

If UCase(Left(vDocNo, 3)) = "S01" Or UCase(Left(vDocNo, 3)) = "S02" Or UCase(Left(vDocNo, 3)) = "W01" Or UCase(Left(vDocNo, 3)) = "W02" Then
   vDocNo1 = UCase(Left(Right(vDocNo, Len(vDocNo) - InStr(vDocNo, "-")), 3)) 'UCase(Left(vDocNo, 3))
Else
   vDocNo1 = UCase(Left(vDocNo, 3))
End If

vDocGroup1 = vDocNo1
If vDocGroup1 = "ICV" Or vDocGroup1 = "ICN" Or vDocGroup1 = "IVD" Or vDocGroup1 = "IVN" Or vDocGroup1 = "IVM" Then
vRepID = 430
vRepType = "INV"
Else
vRepID = 428
vRepType = "INV"
End If
vCheck = 1

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
With CrystalReport310
.ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
.ParameterFields(0) = "@vDocNo;" & vDocNo & ";true"
.ParameterFields(1) = "@vCheck;" & vCheck & ";true"
.Destination = crptToWindow
.WindowState = crptMaximized
.Action = 1
End With
End If
vRecordset.Close
                            
vQuery = "select lastprintcount,whcode  from npmaster.dbo.np_paygoods where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vCount = Trim(vRecordset.Fields("lastprintcount").Value)
End If
vRecordset.Close

vCount = vCount + 1
vQuery = "Update npmaster.dbo.np_paygoods  set lastprintcount = " & vCount & " ,lastuserprint = '" & vUserID & "' where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
gConnection.Execute vQuery
        
                
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If
End Sub


Public Sub PrintItem_BK2(WHCode As String)
Dim vReportName As String
Dim vDocNo As String
Dim vRepID As Integer
Dim vRepType As String
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vAutoNumber As String, vDocNo1 As String, vDocGroup1 As String
Dim vCheck As Integer
Dim vGenerateNumber, vHeader As String
Dim vWHCode As String
Dim vShelfCode As String
Dim vMydescription As String
Dim vCheckNumber As String
Dim vCount  As Integer

On Error GoTo ErrDescription

vDocNo = Trim(Me.TXT3101.Text)
vWHCode = WHCode
vShelfCode = Trim(CMB3101.Text)
vQuery = "select * from npmaster.dbo.NP_PayGoods where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vCheckNumber = Trim(vRecordset.Fields("invoiceno").Value)
End If
vRecordset.Close

vQuery = "Update npmaster.dbo.NP_PayGoods set PrintedTimeStamp = getdate() where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
gConnection.Execute vQuery

If vCheckNumber = "" Then
vQuery = "select header,autonumber  from npmaster.dbo.NP_Generate_DocNo where headertype = 8 "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vAutoNumber = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
vHeader = Trim(vRecordset.Fields("header").Value)
End If
vRecordset.Close
vGenerateNumber = vHeader & "-" & vAutoNumber
vMydescription = "ทดแทนการเปลี่ยนคลังของเอกสาร"

vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1 where headertype = 8 "
gConnection.Execute vQuery

vQuery = "exec dbo.USP_INV_InsertPrintSlip '" & vDocNo & "','" & vGenerateNumber & "','" & vWHCode & "','" & vShelfCode & "','" & vMydescription & "','" & vUserID & "' "
gConnection.Execute vQuery

End If
        

'vDocno1 = UCase(Left(Right(vDocNo, Len(vDocNo) - InStr(vDocNo, "-")), 3)) 'UCase(Left(vDocNo, 3))

If UCase(Left(vDocNo, 3)) = "S01" Or UCase(Left(vDocNo, 3)) = "S02" Or UCase(Left(vDocNo, 3)) = "W01" Or UCase(Left(vDocNo, 3)) = "W02" Then
   vDocNo1 = UCase(Left(Right(vDocNo, Len(vDocNo) - InStr(vDocNo, "-")), 3)) 'UCase(Left(vDocNo, 3))
Else
   vDocNo1 = UCase(Left(vDocNo, 3))
End If

vDocGroup1 = vDocNo1
If vDocGroup1 = "ICV" Or vDocGroup1 = "ICN" Or vDocGroup1 = "IVD" Or vDocGroup1 = "IVN" Or vDocGroup1 = "IVM" Then
vRepID = 434
vRepType = "INV"
Else
vRepID = 432
vRepType = "INV"
End If
vCheck = 1

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
With CrystalReport310
.ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
.ParameterFields(0) = "@vDocNo;" & vDocNo & ";true"
.ParameterFields(1) = "@vCheck;" & vCheck & ";true"
.Destination = crptToWindow
.WindowState = crptMaximized
.Action = 1
End With
End If
vRecordset.Close
                            
vQuery = "select lastprintcount,whcode  from npmaster.dbo.np_paygoods where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vCount = Trim(vRecordset.Fields("lastprintcount").Value)
End If
vRecordset.Close

vCount = vCount + 1
vQuery = "Update npmaster.dbo.np_paygoods  set lastprintcount = " & vCount & " ,lastuserprint = '" & vUserID & "' where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
gConnection.Execute vQuery
        
                
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If
End Sub

Public Sub PrintItem_BK3(WHCode As String)
Dim vReportName As String
Dim vDocNo As String
Dim vRepID As Integer
Dim vRepType As String
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vAutoNumber As String, vDocNo1 As String, vDocGroup1 As String
Dim vCheck As Integer
Dim vGenerateNumber, vHeader As String
Dim vWHCode As String
Dim vShelfCode As String
Dim vMydescription As String
Dim vCheckNumber As String
Dim vCount  As Integer

On Error GoTo ErrDescription

vDocNo = Trim(Me.TXT3101.Text)
vWHCode = WHCode
vShelfCode = Trim(CMB3101.Text)
vQuery = "select * from npmaster.dbo.NP_PayGoods where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vCheckNumber = Trim(vRecordset.Fields("invoiceno").Value)
End If
vRecordset.Close

vQuery = "Update npmaster.dbo.NP_PayGoods set PrintedTimeStamp = getdate() where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
gConnection.Execute vQuery

If vCheckNumber = "" Then
vQuery = "select header,autonumber  from npmaster.dbo.NP_Generate_DocNo where headertype = 8 "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vAutoNumber = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
vHeader = Trim(vRecordset.Fields("header").Value)
End If
vRecordset.Close
vGenerateNumber = vHeader & "-" & vAutoNumber
vMydescription = "ทดแทนการเปลี่ยนคลังของเอกสาร"

vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1 where headertype = 8 "
gConnection.Execute vQuery

vQuery = "exec dbo.USP_INV_InsertPrintSlip '" & vDocNo & "','" & vGenerateNumber & "','" & vWHCode & "','" & vShelfCode & "','" & vMydescription & "','" & vUserID & "' "
gConnection.Execute vQuery

End If
        

'vDocno1 = UCase(Left(Right(vDocNo, Len(vDocNo) - InStr(vDocNo, "-")), 3)) 'UCase(Left(vDocNo, 3))

If UCase(Left(vDocNo, 3)) = "S01" Or UCase(Left(vDocNo, 3)) = "S02" Or UCase(Left(vDocNo, 3)) = "W01" Or UCase(Left(vDocNo, 3)) = "W02" Then
   vDocNo1 = UCase(Left(Right(vDocNo, Len(vDocNo) - InStr(vDocNo, "-")), 3)) 'UCase(Left(vDocNo, 3))
Else
   vDocNo1 = UCase(Left(vDocNo, 3))
End If
vDocGroup1 = vDocNo1
If vDocGroup1 = "ICV" Or vDocGroup1 = "ICN" Or vDocGroup1 = "IVD" Or vDocGroup1 = "IVN" Or vDocGroup1 = "IVM" Then
vRepID = 438
vRepType = "INV"
Else
vRepID = 436
vRepType = "INV"
End If
vCheck = 1

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
With CrystalReport310
.ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
.ParameterFields(0) = "@vDocNo;" & vDocNo & ";true"
.ParameterFields(1) = "@vCheck;" & vCheck & ";true"
.Destination = crptToWindow
.WindowState = crptMaximized
.Action = 1
End With
End If
vRecordset.Close
                            
vQuery = "select lastprintcount,whcode  from npmaster.dbo.np_paygoods where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vCount = Trim(vRecordset.Fields("lastprintcount").Value)
End If
vRecordset.Close

vCount = vCount + 1
vQuery = "Update npmaster.dbo.np_paygoods  set lastprintcount = " & vCount & " ,lastuserprint = '" & vUserID & "' where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
gConnection.Execute vQuery
        
                
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If
End Sub


Public Sub PrintItem_SPO(WHCode As String, PickZone As String, ZoneID As String)
Dim vReportName As String
Dim vDocNo As String
Dim vRepID As Integer
Dim vRepType As String
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vAutoNumber As String, vDocNo1 As String, vDocGroup1 As String
Dim vCheck As Integer
Dim vGenerateNumber, vHeader As String
Dim vWHCode As String
Dim vShelfCode As String
Dim vMydescription As String
Dim vCheckNumber As String
Dim vCount  As Integer

On Error GoTo ErrDescription

vDocNo = Trim(Me.TXT3101.Text)
vWHCode = WHCode
vShelfCode = Trim(CMB3101.Text)
vQuery = "select * from npmaster.dbo.NP_PayGoods where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vCheckNumber = Trim(vRecordset.Fields("invoiceno").Value)
End If
vRecordset.Close

vQuery = "Update npmaster.dbo.NP_PayGoods set PrintedTimeStamp = getdate() where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
gConnection.Execute vQuery

If vCheckNumber = "" Then
vQuery = "select header,autonumber  from npmaster.dbo.NP_Generate_DocNo where headertype = 8 "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vAutoNumber = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
vHeader = Trim(vRecordset.Fields("header").Value)
End If
vRecordset.Close
vGenerateNumber = vHeader & "-" & vAutoNumber
vMydescription = "ทดแทนการเปลี่ยนคลังของเอกสาร"

vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1 where headertype = 8 "
gConnection.Execute vQuery

vQuery = "exec dbo.USP_INV_InsertPrintSlip1 '" & vDocNo & "','" & vGenerateNumber & "','" & vWHCode & "','" & vShelfCode & "','" & ZoneID & "','" & PickZone & "','" & vMydescription & "','" & vUserID & "' "
gConnection.Execute vQuery

End If
        

'vDocno1 = UCase(Left(Right(vDocNo, Len(vDocNo) - InStr(vDocNo, "-")), 3)) 'UCase(Left(vDocNo, 3))

If UCase(Left(vDocNo, 3)) = "S01" Or UCase(Left(vDocNo, 3)) = "S02" Or UCase(Left(vDocNo, 3)) = "W01" Or UCase(Left(vDocNo, 3)) = "W02" Then
   vDocNo1 = UCase(Left(Right(vDocNo, Len(vDocNo) - InStr(vDocNo, "-")), 3)) 'UCase(Left(vDocNo, 3))
Else
   vDocNo1 = UCase(Left(vDocNo, 3))
End If

vDocGroup1 = vDocNo1
If vDocGroup1 = "ICV" Or vDocGroup1 = "ICN" Or vDocGroup1 = "IVD" Or vDocGroup1 = "IVN" Or vDocGroup1 = "IVM" Then
vRepID = 442
vRepType = "INV"
Else
vRepID = 440
vRepType = "INV"
End If
vCheck = 1

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
With CrystalReport310
.ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
.ParameterFields(0) = "@vDocNo;" & vDocNo & ";true"
.ParameterFields(1) = "@vPickZone;" & PickZone & ";true"
.Destination = crptToWindow
.WindowState = crptMaximized
.Action = 1
End With
End If
vRecordset.Close
                            
vQuery = "select lastprintcount,whcode  from npmaster.dbo.np_paygoods where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vCount = Trim(vRecordset.Fields("lastprintcount").Value)
End If
vRecordset.Close

vCount = vCount + 1
vQuery = "Update npmaster.dbo.np_paygoods  set lastprintcount = " & vCount & " ,lastuserprint = '" & vUserID & "' where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
gConnection.Execute vQuery
        
                
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If
End Sub

Public Sub PrintItem_DMG(WHCode As String, PickZone As String, ZoneID As String)
Dim vReportName As String
Dim vDocNo As String
Dim vRepID As Integer
Dim vRepType As String
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vAutoNumber As String, vDocNo1 As String, vDocGroup1 As String
Dim vCheck As Integer
Dim vGenerateNumber, vHeader As String
Dim vWHCode As String
Dim vShelfCode As String
Dim vMydescription As String
Dim vCheckNumber As String
Dim vCount  As Integer

On Error GoTo ErrDescription

vDocNo = Trim(Me.TXT3101.Text)
vWHCode = WHCode
vShelfCode = Trim(CMB3101.Text)
vQuery = "select * from npmaster.dbo.NP_PayGoods where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vCheckNumber = Trim(vRecordset.Fields("invoiceno").Value)
End If
vRecordset.Close

vQuery = "Update npmaster.dbo.NP_PayGoods set PrintedTimeStamp = getdate() where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
gConnection.Execute vQuery

If vCheckNumber = "" Then
vQuery = "select header,autonumber  from npmaster.dbo.NP_Generate_DocNo where headertype = 8 "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vAutoNumber = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
vHeader = Trim(vRecordset.Fields("header").Value)
End If
vRecordset.Close
vGenerateNumber = vHeader & "-" & vAutoNumber
vMydescription = "ทดแทนการเปลี่ยนคลังของเอกสาร"

vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1 where headertype = 8 "
gConnection.Execute vQuery

vQuery = "exec dbo.USP_INV_InsertPrintSlip1 '" & vDocNo & "','" & vGenerateNumber & "','" & vWHCode & "','" & vShelfCode & "','" & ZoneID & "','" & PickZone & "','" & vMydescription & "','" & vUserID & "' "
gConnection.Execute vQuery

End If
        
'vDocno1 = UCase(Left(Right(vDocNo, Len(vDocNo) - InStr(vDocNo, "-")), 3)) 'UCase(Left(vDocNo, 3))

If UCase(Left(vDocNo, 3)) = "S01" Or UCase(Left(vDocNo, 3)) = "S02" Or UCase(Left(vDocNo, 3)) = "W01" Or UCase(Left(vDocNo, 3)) = "W02" Then
   vDocNo1 = UCase(Left(Right(vDocNo, Len(vDocNo) - InStr(vDocNo, "-")), 3)) 'UCase(Left(vDocNo, 3))
Else
   vDocNo1 = UCase(Left(vDocNo, 3))
End If

vDocGroup1 = vDocNo1
If vDocGroup1 = "ICV" Or vDocGroup1 = "ICN" Or vDocGroup1 = "IVD" Or vDocGroup1 = "IVN" Or vDocGroup1 = "IVM" Then
vRepID = 458
vRepType = "INV"
Else
vRepID = 457
vRepType = "INV"
End If
vCheck = 1

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
With CrystalReport310
.ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
.ParameterFields(0) = "@vDocNo;" & vDocNo & ";true"
.ParameterFields(1) = "@vPickZone;" & PickZone & ";true"
.Destination = crptToWindow
.WindowState = crptMaximized
.Action = 1
End With
End If
vRecordset.Close
                            
vQuery = "select lastprintcount,whcode  from npmaster.dbo.np_paygoods where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vCount = Trim(vRecordset.Fields("lastprintcount").Value)
End If
vRecordset.Close

vCount = vCount + 1
vQuery = "Update npmaster.dbo.np_paygoods  set lastprintcount = " & vCount & " ,lastuserprint = '" & vUserID & "' where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
gConnection.Execute vQuery
        
                
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If
End Sub

Public Sub PrintItem_VND(WHCode As String, PickZone As String, ZoneID As String)
Dim vReportName As String
Dim vDocNo As String
Dim vRepID As Integer
Dim vRepType As String
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vAutoNumber As String, vDocNo1 As String, vDocGroup1 As String
Dim vCheck As Integer
Dim vGenerateNumber, vHeader As String
Dim vWHCode As String
Dim vShelfCode As String
Dim vMydescription As String
Dim vCheckNumber As String
Dim vCount  As Integer

On Error GoTo ErrDescription

vDocNo = Trim(Me.TXT3101.Text)
vWHCode = WHCode
vShelfCode = Trim(CMB3101.Text)
vQuery = "select * from npmaster.dbo.NP_PayGoods where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vCheckNumber = Trim(vRecordset.Fields("invoiceno").Value)
End If
vRecordset.Close

vQuery = "Update npmaster.dbo.NP_PayGoods set PrintedTimeStamp = getdate() where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
gConnection.Execute vQuery

If vCheckNumber = "" Then
vQuery = "select header,autonumber  from npmaster.dbo.NP_Generate_DocNo where headertype = 8 "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vAutoNumber = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
vHeader = Trim(vRecordset.Fields("header").Value)
End If
vRecordset.Close
vGenerateNumber = vHeader & "-" & vAutoNumber
vMydescription = "ทดแทนการเปลี่ยนคลังของเอกสาร"

vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1 where headertype = 8 "
gConnection.Execute vQuery

vQuery = "exec dbo.USP_INV_InsertPrintSlip1 '" & vDocNo & "','" & vGenerateNumber & "','" & vWHCode & "','" & vShelfCode & "','" & ZoneID & "','" & PickZone & "','" & vMydescription & "','" & vUserID & "' "
gConnection.Execute vQuery

End If
        
'vDocno1 = UCase(Left(Right(vDocNo, Len(vDocNo) - InStr(vDocNo, "-")), 3)) 'UCase(Left(vDocNo, 3))

If UCase(Left(vDocNo, 3)) = "S01" Or UCase(Left(vDocNo, 3)) = "S02" Or UCase(Left(vDocNo, 3)) = "W01" Or UCase(Left(vDocNo, 3)) = "W02" Then
   vDocNo1 = UCase(Left(Right(vDocNo, Len(vDocNo) - InStr(vDocNo, "-")), 3)) 'UCase(Left(vDocNo, 3))
Else
   vDocNo1 = UCase(Left(vDocNo, 3))
End If

vDocGroup1 = vDocNo1
If vDocGroup1 = "ICV" Or vDocGroup1 = "ICN" Or vDocGroup1 = "IVD" Or vDocGroup1 = "IVN" Or vDocGroup1 = "IVM" Then
vRepID = 462
vRepType = "INV"
Else
vRepID = 461
vRepType = "INV"
End If
vCheck = 1

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
With CrystalReport310
.ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
.ParameterFields(0) = "@vDocNo;" & vDocNo & ";true"
.ParameterFields(1) = "@vPickZone;" & PickZone & ";true"
.Destination = crptToWindow
.WindowState = crptMaximized
.Action = 1
End With
End If
vRecordset.Close
                            
vQuery = "select lastprintcount,whcode  from npmaster.dbo.np_paygoods where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vCount = Trim(vRecordset.Fields("lastprintcount").Value)
End If
vRecordset.Close

vCount = vCount + 1
vQuery = "Update npmaster.dbo.np_paygoods  set lastprintcount = " & vCount & " ,lastuserprint = '" & vUserID & "' where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
gConnection.Execute vQuery
        
                
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If
End Sub

Public Sub PrintItem_OFS(WHCode As String, PickZone As String, ZoneID As String)
Dim vReportName As String
Dim vDocNo As String
Dim vRepID As Integer
Dim vRepType As String
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vAutoNumber As String, vDocNo1 As String, vDocGroup1 As String
Dim vCheck As Integer
Dim vGenerateNumber, vHeader As String
Dim vWHCode As String
Dim vShelfCode As String
Dim vMydescription As String
Dim vCheckNumber As String
Dim vCount  As Integer

On Error GoTo ErrDescription

vDocNo = Trim(Me.TXT3101.Text)
vWHCode = WHCode
vShelfCode = Trim(CMB3101.Text)
vQuery = "select * from npmaster.dbo.NP_PayGoods where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vCheckNumber = Trim(vRecordset.Fields("invoiceno").Value)
End If
vRecordset.Close

vQuery = "Update npmaster.dbo.NP_PayGoods set PrintedTimeStamp = getdate() where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
gConnection.Execute vQuery

If vCheckNumber = "" Then
vQuery = "select header,autonumber  from npmaster.dbo.NP_Generate_DocNo where headertype = 8 "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vAutoNumber = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
vHeader = Trim(vRecordset.Fields("header").Value)
End If
vRecordset.Close
vGenerateNumber = vHeader & "-" & vAutoNumber
vMydescription = "ทดแทนการเปลี่ยนคลังของเอกสาร"

vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1 where headertype = 8 "
gConnection.Execute vQuery

vQuery = "exec dbo.USP_INV_InsertPrintSlip1 '" & vDocNo & "','" & vGenerateNumber & "','" & vWHCode & "','" & vShelfCode & "','" & ZoneID & "','" & PickZone & "','" & vMydescription & "','" & vUserID & "' "
gConnection.Execute vQuery

End If
        

'vDocno1 = UCase(Left(Right(vDocNo, Len(vDocNo) - InStr(vDocNo, "-")), 3)) 'UCase(Left(vDocNo, 3))

If UCase(Left(vDocNo, 3)) = "S01" Or UCase(Left(vDocNo, 3)) = "S02" Or UCase(Left(vDocNo, 3)) = "W01" Or UCase(Left(vDocNo, 3)) = "W02" Then
   vDocNo1 = UCase(Left(Right(vDocNo, Len(vDocNo) - InStr(vDocNo, "-")), 3)) 'UCase(Left(vDocNo, 3))
Else
   vDocNo1 = UCase(Left(vDocNo, 3))
End If

vDocGroup1 = vDocNo1
If vDocGroup1 = "ICV" Or vDocGroup1 = "ICN" Or vDocGroup1 = "IVD" Or vDocGroup1 = "IVN" Or vDocGroup1 = "IVM" Then
vRepID = 446
vRepType = "INV"
Else
vRepID = 445
vRepType = "INV"
End If
vCheck = 1

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
With CrystalReport310
.ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
.ParameterFields(0) = "@vDocNo;" & vDocNo & ";true"
.ParameterFields(1) = "@vPickZone;" & PickZone & ";true"
.Destination = crptToWindow
.WindowState = crptMaximized
.Action = 1
End With
End If
vRecordset.Close
                            
vQuery = "select lastprintcount,whcode  from npmaster.dbo.np_paygoods where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vCount = Trim(vRecordset.Fields("lastprintcount").Value)
End If
vRecordset.Close

vCount = vCount + 1
vQuery = "Update npmaster.dbo.np_paygoods  set lastprintcount = " & vCount & " ,lastuserprint = '" & vUserID & "' where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
gConnection.Execute vQuery
        
                
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If
End Sub

Public Sub PrintItem_SHW(WHCode As String, PickZone As String, ZoneID As String)
Dim vReportName As String
Dim vDocNo As String
Dim vRepID As Integer
Dim vRepType As String
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vAutoNumber As String, vDocNo1 As String, vDocGroup1 As String
Dim vCheck As Integer
Dim vGenerateNumber, vHeader As String
Dim vWHCode As String
Dim vShelfCode As String
Dim vMydescription As String
Dim vCheckNumber As String
Dim vCount  As Integer

On Error GoTo ErrDescription

vDocNo = Trim(Me.TXT3101.Text)
vWHCode = WHCode
vShelfCode = Trim(CMB3101.Text)
vQuery = "select * from npmaster.dbo.NP_PayGoods where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vCheckNumber = Trim(vRecordset.Fields("invoiceno").Value)
End If
vRecordset.Close

vQuery = "Update npmaster.dbo.NP_PayGoods set PrintedTimeStamp = getdate() where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
gConnection.Execute vQuery

If vCheckNumber = "" Then
vQuery = "select header,autonumber  from npmaster.dbo.NP_Generate_DocNo where headertype = 8 "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vAutoNumber = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
vHeader = Trim(vRecordset.Fields("header").Value)
End If
vRecordset.Close
vGenerateNumber = vHeader & "-" & vAutoNumber
vMydescription = "ทดแทนการเปลี่ยนคลังของเอกสาร"

vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1 where headertype = 8 "
gConnection.Execute vQuery

vQuery = "exec dbo.USP_INV_InsertPrintSlip1 '" & vDocNo & "','" & vGenerateNumber & "','" & vWHCode & "','" & vShelfCode & "','" & ZoneID & "','" & PickZone & "','" & vMydescription & "','" & vUserID & "' "
gConnection.Execute vQuery

End If
        

'vDocno1 = UCase(Left(Right(vDocNo, Len(vDocNo) - InStr(vDocNo, "-")), 3)) 'UCase(Left(vDocNo, 3))

If UCase(Left(vDocNo, 3)) = "S01" Or UCase(Left(vDocNo, 3)) = "S02" Or UCase(Left(vDocNo, 3)) = "W01" Or UCase(Left(vDocNo, 3)) = "W02" Then
   vDocNo1 = UCase(Left(Right(vDocNo, Len(vDocNo) - InStr(vDocNo, "-")), 3)) 'UCase(Left(vDocNo, 3))
Else
   vDocNo1 = UCase(Left(vDocNo, 3))
End If

vDocGroup1 = vDocNo1
If vDocGroup1 = "ICV" Or vDocGroup1 = "ICN" Or vDocGroup1 = "IVD" Or vDocGroup1 = "IVN" Or vDocGroup1 = "IVM" Then
vRepID = 450
vRepType = "INV"
Else
vRepID = 449
vRepType = "INV"
End If
vCheck = 1

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
With CrystalReport310
.ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
.ParameterFields(0) = "@vDocNo;" & vDocNo & ";true"
.ParameterFields(1) = "@vPickZone;" & PickZone & ";true"
.Destination = crptToWindow
.WindowState = crptMaximized
.Action = 1
End With
End If
vRecordset.Close
                            
vQuery = "select lastprintcount,whcode  from npmaster.dbo.np_paygoods where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vCount = Trim(vRecordset.Fields("lastprintcount").Value)
End If
vRecordset.Close

vCount = vCount + 1
vQuery = "Update npmaster.dbo.np_paygoods  set lastprintcount = " & vCount & " ,lastuserprint = '" & vUserID & "' where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
gConnection.Execute vQuery
        
                
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If
End Sub

Public Sub PrintItem_RSV(WHCode As String, PickZone As String, ZoneID As String)
Dim vReportName As String
Dim vDocNo As String
Dim vRepID As Integer
Dim vRepType As String
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vAutoNumber As String, vDocNo1 As String, vDocGroup1 As String
Dim vCheck As Integer
Dim vGenerateNumber, vHeader As String
Dim vWHCode As String
Dim vShelfCode As String
Dim vMydescription As String
Dim vCheckNumber As String
Dim vCount  As Integer

On Error GoTo ErrDescription

vDocNo = Trim(Me.TXT3101.Text)
vWHCode = WHCode
vShelfCode = Trim(CMB3101.Text)
vQuery = "select * from npmaster.dbo.NP_PayGoods where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vCheckNumber = Trim(vRecordset.Fields("invoiceno").Value)
End If
vRecordset.Close

vQuery = "Update npmaster.dbo.NP_PayGoods set PrintedTimeStamp = getdate() where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
gConnection.Execute vQuery

If vCheckNumber = "" Then
vQuery = "select header,autonumber  from npmaster.dbo.NP_Generate_DocNo where headertype = 8 "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vAutoNumber = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
vHeader = Trim(vRecordset.Fields("header").Value)
End If
vRecordset.Close
vGenerateNumber = vHeader & "-" & vAutoNumber
vMydescription = "ทดแทนการเปลี่ยนคลังของเอกสาร"

vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1 where headertype = 8 "
gConnection.Execute vQuery

vQuery = "exec dbo.USP_INV_InsertPrintSlip1 '" & vDocNo & "','" & vGenerateNumber & "','" & vWHCode & "','" & vShelfCode & "','" & ZoneID & "','" & PickZone & "','" & vMydescription & "','" & vUserID & "' "
gConnection.Execute vQuery

End If
        

'vDocno1 = UCase(Left(Right(vDocNo, Len(vDocNo) - InStr(vDocNo, "-")), 3)) 'UCase(Left(vDocNo, 3))

If UCase(Left(vDocNo, 3)) = "S01" Or UCase(Left(vDocNo, 3)) = "S02" Or UCase(Left(vDocNo, 3)) = "W01" Or UCase(Left(vDocNo, 3)) = "W02" Then
   vDocNo1 = UCase(Left(Right(vDocNo, Len(vDocNo) - InStr(vDocNo, "-")), 3)) 'UCase(Left(vDocNo, 3))
Else
   vDocNo1 = UCase(Left(vDocNo, 3))
End If

vDocGroup1 = vDocNo1
If vDocGroup1 = "ICV" Or vDocGroup1 = "ICN" Or vDocGroup1 = "IVD" Or vDocGroup1 = "IVN" Or vDocGroup1 = "IVM" Then
vRepID = 454
vRepType = "INV"
Else
vRepID = 453
vRepType = "INV"
End If
vCheck = 1

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
With CrystalReport310
.ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
.ParameterFields(0) = "@vDocNo;" & vDocNo & ";true"
.ParameterFields(1) = "@vPickZone;" & PickZone & ";true"
.Destination = crptToWindow
.WindowState = crptMaximized
.Action = 1
End With
End If
vRecordset.Close
                            
vQuery = "select lastprintcount,whcode  from npmaster.dbo.np_paygoods where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vCount = Trim(vRecordset.Fields("lastprintcount").Value)
End If
vRecordset.Close

vCount = vCount + 1
vQuery = "Update npmaster.dbo.np_paygoods  set lastprintcount = " & vCount & " ,lastuserprint = '" & vUserID & "' where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "' and shelfcode = '" & vShelfCode & "' "
gConnection.Execute vQuery
        
                
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If
End Sub


Private Sub TXT3101_KeyPress(KeyAscii As Integer)
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocNo As String, vDocNo1 As String
Dim i  As Integer, a As Integer, b As Integer, d As Integer
Dim vCount As Integer
Dim n As Integer
Dim vWHCode As String
Dim vShelfCode As String
Dim vCheckWH As String
Dim vPickZone As String
Dim vFamilyCode As String
Dim vCheckZone As String
Dim vCheckShelfCode As String
Dim vCheckFamilyCode As String

On Error GoTo ErrDescription

Me.CMBWHCode.Clear
Me.CMB3101.Clear
Me.CMBZone.Clear
Me.CMBFamily.Clear


If KeyAscii = 13 Then
      vDocNo = Trim(TXT3101.Text)
      vQuery = "exec dbo.USP_INV_SearchShelfInvoicePrintSlip '" & vDocNo & "' "
      If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
         vRecordset.MoveFirst
         While Not vRecordset.EOF
         
         vDocNo1 = Trim(vRecordset.Fields("docno").Value)
         vShelfCode = Trim(vRecordset.Fields("shelfcode").Value)
         vWHCode = Trim(vRecordset.Fields("whcode").Value)
         vPickZone = Trim(vRecordset.Fields("zoneid").Value)
         vFamilyCode = Trim(vRecordset.Fields("familycode").Value)
         
         If Me.CMBWHCode.ListCount > 0 Then
            For n = 0 To Me.CMBWHCode.ListCount - 1
               vCheckWH = Me.CMBWHCode.List(n)
               If vCheckWH <> vWHCode Then
               Me.CMBWHCode.AddItem vWHCode
               End If
            Next n
         Else
            Me.CMBWHCode.AddItem vWHCode
         End If
         
        If Me.CMB3101.ListCount > 0 Then
            For b = 0 To Me.CMB3101.ListCount - 1
               vCheckShelfCode = Me.CMB3101.List(b)
               If vCheckShelfCode <> vShelfCode Then
               Me.CMB3101.AddItem vShelfCode
               End If
            Next b
         Else
            Me.CMB3101.AddItem vShelfCode
         End If
         
         If Me.CMBZone.ListCount > 0 Then
         
            For a = 0 To Me.CMBZone.ListCount - 1
               vCheckZone = Me.CMBZone.List(a)
               If vCheckZone <> vPickZone Then
               Me.CMBZone.AddItem vPickZone
               End If
            Next a
            
         Else
         Me.CMBZone.AddItem Trim(vRecordset.Fields("zoneid").Value)
         End If
         

         If Me.CMBFamily.ListCount > 0 Then
         
            For d = 0 To Me.CMBFamily.ListCount - 1
               vCheckFamilyCode = Me.CMBFamily.List(d)
               If vCheckFamilyCode <> vFamilyCode Then
               Me.CMBFamily.AddItem vFamilyCode
               End If
            Next d
            
         Else
         Me.CMBFamily.AddItem Trim(vRecordset.Fields("familycode").Value)
         End If
         
         
         vRecordset.MoveNext
         Wend
      Else
        MsgBox "ไม่มีเอกสารเลขที่  " & vDocNo & " ในระบบ", vbInformation + vbCritical, "ข้อความเตือน"
        Exit Sub
      End If
      vRecordset.Close
      
      If Me.CMB3101.ListCount > 0 Then
         Me.CMB3101.Text = Me.CMB3101.List(0)
      End If
      
      If Me.CMBWHCode.ListCount > 0 Then
         Me.CMBWHCode.Text = Me.CMBWHCode.List(0)
      End If
                    
      MsgBox "เลือกคลังที่จะพิมพ์ใบจ่ายสินค้า (ใบเหลือง)", vbInformation + vbCritical, "ข้อความเตือน"
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
End If
End Sub


Public Sub PrintItem010()
Dim vDocNo As String
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vCheckItemNotLocation015 As Integer
Dim vCheckItemLocation015 As String

On Error GoTo ErrDescription

vDocNo = UCase(Trim(TXT3101.Text))

vQuery = "exec dbo.USP_INV_CheckItemNotLocation015 '" & vDocNo & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
  vCheckItemNotLocation015 = vRecordset.Fields("vcount").Value
End If
vRecordset.Close

vQuery = "exec dbo.USP_INV_CheckItemLocation015 '" & vDocNo & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
  vCheckItemLocation015 = vRecordset.Fields("vcount").Value
End If
vRecordset.Close


If vCheckItemNotLocation015 > 0 Then
  Call PrintItem010_Zone010
End If
If vCheckItemLocation015 > 0 Then
  Call PrintItem010_Zone015
End If
        
        
        'vWHCode = Trim(CMB3101.Text)
        'vQuery = "select * from npmaster.dbo.NP_PayGoods where invoiceno = '" & vDocno & "'  and whcode = '" & vWHCode & "' "
        'If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
         '   vCheckNumber = Trim(vRecordset.Fields("invoiceno").Value)
        'End If
        'vRecordset.Close
        
        'vQuery = "Update npmaster.dbo.NP_PayGoods set PrintedTimeStamp = getdate() where invoiceno = '" & vDocno & "'  and whcode = '" & vWHCode & "' "
        'gConnection.Execute vQuery
        
        'If vCheckNumber = "" Then
               '  vQuery = "select header,autonumber  from npmaster.dbo.NP_Generate_DocNo where headertype = 8 "
              '  If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
             '       vAutoNumber = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
            '        vHeader = Trim(vRecordset.Fields("header").Value)
           '     End If
          '      vRecordset.Close
         '       vGenerateNumber = vHeader & "-" & vAutoNumber
        '
              '  vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1 where headertype = 8 "
             '   gConnection.Execute vQuery
            '
           '     vQuery = "insert npmaster.dbo.np_paygoods (invoiceno,paynumber,paydatetime,whcode,UserPrint,LastPrintCount,Checked,description) " _
          '      & " values ('" & vDocno & "','" & vGenerateNumber & "',getdate(),'010','" & vUserID & "',1,0,'ทดแทนการเปลี่ยนคลังของเอกสาร')"
         '       gConnection.Execute vQuery
        
        'End If
        
        'vDocno1 = UCase(Left(vDocno, 3))
        'vDocGroup1 = vDocno1
       ' If vDocGroup1 = "ICV" Or vDocGroup1 = "ICN" Or vDocGroup1 = "IVD" Or vDocGroup1 = "IVN" Then
      '  vRepID = 126
     '   vRepType = "INV"
    '    Else
   '     vRepID = 93
  '      vRepType = "INV"
 '       End If
'vCheck = 1
 'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
                   '         If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                  '              With CrystalReport310
                 '                   .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
                '                    .ParameterFields(0) = "@DocNo;" & vDocno & ";true"
               '                     .ParameterFields(1) = "@vCheck;" & vCheck & ";true"
              '                      .Destination = crptToWindow
             '                       .WindowState = crptMaximized
            '                        .Action = 1
           '                     End With
          '                  End If
         '                   vRecordset.Close
                            
        'vQuery = "select lastprintcount,whcode  from npmaster.dbo.np_paygoods where invoiceno = '" & vDocno & "' and whcode = '" & vWHCode & "' "
        'If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
         '   vCount = Trim(vRecordset.Fields("lastprintcount").Value)
        'End If
        'vRecordset.Close
        
        'vCount = vCount + 1
        'vQuery = "Update npmaster.dbo.np_paygoods  set lastprintcount = " & vCount & " ,lastuserprint = '" & vUserID & "' where invoiceno = '" & vDocno & "' and whcode = '" & vWHCode & "'"
        'gConnection.Execute vQuery
                
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If
End Sub

Public Sub PrintItem010_Zone010()
Dim vReportName As String
Dim vDocNo As String
Dim vRepID As Integer
Dim vRepType As String
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vCount As Integer
Dim vWHCode As String
Dim vWHCode1 As Integer
Dim vCheckNumber As String, vDocGroup1 As String
Dim vAutoNumber As String, vDocNo1, vHeader As String
Dim vGenerateNumber  As String
Dim vCheck As Integer

On Error GoTo ErrDescription

vDocNo = UCase(Trim(TXT3101.Text))
vWHCode = Trim(CMB3101.Text)
vQuery = "select * from npmaster.dbo.NP_PayGoods where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vCheckNumber = Trim(vRecordset.Fields("invoiceno").Value)
End If
vRecordset.Close

vQuery = "Update npmaster.dbo.NP_PayGoods set PrintedTimeStamp = getdate() where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' "
gConnection.Execute vQuery

If vCheckNumber = "" Then
   vQuery = "select header,autonumber  from npmaster.dbo.NP_Generate_DocNo where headertype = 8 "
  If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
      vAutoNumber = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
      vHeader = Trim(vRecordset.Fields("header").Value)
  End If
  vRecordset.Close
  vGenerateNumber = vHeader & "-" & vAutoNumber

  vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1 where headertype = 8 "
  gConnection.Execute vQuery
  
  vQuery = "insert npmaster.dbo.np_paygoods (invoiceno,paynumber,paydatetime,whcode,UserPrint,LastPrintCount,Checked,description) " _
  & " values ('" & vDocNo & "','" & vGenerateNumber & "',getdate(),'010','" & vUserID & "',1,0,'ทดแทนการเปลี่ยนคลังของเอกสาร')"
  gConnection.Execute vQuery
End If

vDocNo1 = UCase(Left(vDocNo, 3))
vDocGroup1 = vDocNo1
vRepType = "INV"
If vDocGroup1 = "ICV" Or vDocGroup1 = "ICN" Or vDocGroup1 = "IVD" Or vDocGroup1 = "IVN" Or vDocGroup1 = "IVM" Then
vRepID = 346
Else
vRepID = 342
End If
vCheck = 1
vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
                    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                        With CrystalReport310
                            .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
                            .ParameterFields(0) = "@DocNo;" & vDocNo & ";true"
                            .ParameterFields(1) = "@vCheck;" & vCheck & ";true"
                            .Destination = crptToWindow
                            .WindowState = crptMaximized
                            .Action = 1
                        End With
                    End If
                    vRecordset.Close
                    
vQuery = "select lastprintcount,whcode  from npmaster.dbo.np_paygoods where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vCount = Trim(vRecordset.Fields("lastprintcount").Value)
End If
vRecordset.Close

vCount = vCount + 1
vQuery = "Update npmaster.dbo.np_paygoods  set lastprintcount = " & vCount & " ,lastuserprint = '" & vUserID & "' where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "'"
gConnection.Execute vQuery
        
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If
End Sub

Public Sub PrintItem010_Zone015()
Dim vReportName As String
Dim vDocNo As String
Dim vRepID As Integer
Dim vRepType As String
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vCount As Integer
Dim vWHCode As String
Dim vWHCode1 As Integer
Dim vCheckNumber As String, vDocGroup1 As String
Dim vAutoNumber As String, vDocNo1, vHeader As String
Dim vGenerateNumber  As String
Dim vCheck As Integer

On Error GoTo ErrDescription

vDocNo = UCase(Trim(TXT3101.Text))
vWHCode = Trim("010C")
vQuery = "select * from npmaster.dbo.NP_PayGoods where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vCheckNumber = Trim(vRecordset.Fields("invoiceno").Value)
End If
vRecordset.Close

vQuery = "Update npmaster.dbo.NP_PayGoods set PrintedTimeStamp = getdate() where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' "
gConnection.Execute vQuery

If vCheckNumber = "" Then
   vQuery = "select header,autonumber  from npmaster.dbo.NP_Generate_DocNo where headertype = 8 "
  If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
      vAutoNumber = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
      vHeader = Trim(vRecordset.Fields("header").Value)
  End If
  vRecordset.Close
  vGenerateNumber = vHeader & "-" & vAutoNumber

  vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1 where headertype = 8 "
  gConnection.Execute vQuery
  
  vQuery = "insert npmaster.dbo.np_paygoods (invoiceno,paynumber,paydatetime,whcode,UserPrint,LastPrintCount,Checked,description) " _
  & " values ('" & vDocNo & "','" & vGenerateNumber & "',getdate(),'010C','" & vUserID & "',1,0,'ทดแทนการเปลี่ยนคลังของเอกสาร')"
  gConnection.Execute vQuery
End If

vDocNo1 = UCase(Left(vDocNo, 3))
vDocGroup1 = vDocNo1
vRepType = "INV"
If vDocGroup1 = "ICV" Or vDocGroup1 = "ICN" Or vDocGroup1 = "IVD" Or vDocGroup1 = "IVN" Or vDocGroup1 = "IVM" Then
vRepID = 348
Else
vRepID = 344
End If
vCheck = 1

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
                    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                        With CrystalReport310
                            .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
                            .ParameterFields(0) = "@DocNo;" & vDocNo & ";true"
                            .ParameterFields(1) = "@vCheck;" & vCheck & ";true"
                            .Destination = crptToWindow
                            .WindowState = crptMaximized
                            .Action = 1
                        End With
                    End If
                    vRecordset.Close
                    
vQuery = "select lastprintcount,whcode  from npmaster.dbo.np_paygoods where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vCount = Trim(vRecordset.Fields("lastprintcount").Value)
End If
vRecordset.Close

vCount = vCount + 1
vQuery = "Update npmaster.dbo.np_paygoods  set lastprintcount = " & vCount & " ,lastuserprint = '" & vUserID & "' where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "'"
gConnection.Execute vQuery
        
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If

End Sub


Public Sub PrintItem011()
        Dim vReportName As String
        Dim vDocNo As String
        Dim vRepID As Integer
        Dim vRepType As String
        Dim vQuery As String
        Dim vRecordset As New ADODB.Recordset
        Dim vCount As Integer
        Dim vWHCode As String
        Dim vWHCode1 As Integer
        Dim vCheckNumber As String, vDocNo1 As String
        Dim vAutoNumber As String, vDocGroup1 As String
        Dim vGenerateNumber, vHeader As String
        
        On Error GoTo ErrDescription
        
        vDocNo = Trim(TXT3101.Text)
        vWHCode = Trim(CMB3101.Text)
        vQuery = "select * from npmaster.dbo.NP_PayGoods where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vCheckNumber = Trim(vRecordset.Fields("invoiceno").Value)
        End If
        vRecordset.Close
        
        vQuery = "Update npmaster.dbo.NP_PayGoods set PrintedTimeStamp = getdate() where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' "
        gConnection.Execute vQuery
        
        If vCheckNumber = "" Then
                 vQuery = "select header,autonumber  from npmaster.dbo.NP_Generate_DocNo where headertype = 8 "
                If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                    vAutoNumber = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
                    vHeader = Trim(vRecordset.Fields("header").Value)
                End If
                vRecordset.Close
                vGenerateNumber = vHeader & "-" & vAutoNumber
        
                vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1 where headertype = 8 "
                gConnection.Execute vQuery
                
                vQuery = "insert npmaster.dbo.np_paygoods (invoiceno,paynumber,paydatetime,whcode,UserPrint,LastPrintCount,Checked,description) " _
                & " values ('" & vDocNo & "','" & vGenerateNumber & "',getdate(),'011','" & vUserID & "',1,0,'ทดแทนการเปลี่ยนคลังของเอกสาร')"
                gConnection.Execute vQuery
        End If
         vDocNo1 = UCase(Left(vDocNo, 3))
         vDocGroup1 = vDocNo1
        
        If vDocGroup1 = "ICV" Or vDocGroup1 = "ICN" Or vDocGroup1 = "IVD" Or vDocGroup1 = "IVN" Then
        vRepID = 128
        vRepType = "INV"
        Else
        vRepID = 94
        vRepType = "INV"
        End If

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
 'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
                            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                                With CrystalReport310
                                    .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
                                    .ParameterFields(0) = "@DocNo;" & vDocNo & ";true"
                                    .Destination = crptToWindow
                                    .WindowState = crptMaximized
                                    .Action = 1
                                End With
                            End If
                            vRecordset.Close
                            
        vQuery = "select lastprintcount,whcode  from npmaster.dbo.np_paygoods where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vCount = Trim(vRecordset.Fields("lastprintcount").Value)
        End If
        vRecordset.Close
        
        vCount = vCount + 1
        vQuery = "Update npmaster.dbo.np_paygoods  set lastprintcount = " & vCount & " ,lastuserprint = '" & vUserID & "' where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "'"
        gConnection.Execute vQuery
                
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If
End Sub

Public Sub PrintItem012()
        Dim vReportName As String
        Dim vDocNo As String
        Dim vRepID As Integer
        Dim vRepType As String
        Dim vQuery As String
        Dim vRecordset As New ADODB.Recordset
        Dim vCount As Integer
        Dim vWHCode As String
        Dim vWHCode1 As Integer
        Dim vCheckNumber As String, vDocNo1 As String
        Dim vAutoNumber As String, vDocGroup1 As String
        Dim vGenerateNumber, vHeader As String
        Dim vCheck As Integer
        
        On Error GoTo ErrDescription
        
        vDocNo = Trim(TXT3101.Text)
        vWHCode = Trim(CMB3101.Text)
        vQuery = "select * from npmaster.dbo.NP_PayGoods where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vCheckNumber = Trim(vRecordset.Fields("invoiceno").Value)
        End If
        vRecordset.Close
        
        vQuery = "Update npmaster.dbo.NP_PayGoods set PrintedTimeStamp = getdate() where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' "
        gConnection.Execute vQuery
        
        If vCheckNumber = "" Then
                 vQuery = "select header,autonumber  from npmaster.dbo.NP_Generate_DocNo where headertype = 8 "
                If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                    vAutoNumber = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
                    vHeader = Trim(vRecordset.Fields("header").Value)
                End If
                vRecordset.Close
                vGenerateNumber = vHeader & "-" & vAutoNumber
        
                vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1 where headertype = 8 "
                gConnection.Execute vQuery
                
                vQuery = "insert npmaster.dbo.np_paygoods (invoiceno,paynumber,paydatetime,whcode,UserPrint,LastPrintCount,Checked,description) " _
                & " values ('" & vDocNo & "','" & vGenerateNumber & "',getdate(),'012','" & vUserID & "',1,0,'ทดแทนการเปลี่ยนคลังของเอกสาร')"
                gConnection.Execute vQuery
        End If
        
        vDocNo1 = UCase(Left(vDocNo, 3))
        vDocGroup1 = vDocNo1
        If vDocGroup1 = "ICV" Or vDocGroup1 = "ICN" Or vDocGroup1 = "IVD" Or vDocGroup1 = "IVN" Then
        vRepID = 130
        vRepType = "INV"
        Else
        vRepID = 95
        vRepType = "INV"
        End If
vCheck = 1

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
 'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
                            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                                With CrystalReport310
                                    .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
                                    .ParameterFields(0) = "@DocNo;" & vDocNo & ";true"
                                    .ParameterFields(1) = "@vCheck;" & vCheck & ";true"
                                    .Destination = crptToWindow
                                    .WindowState = crptMaximized
                                    .Action = 1
                                End With
                            End If
                            vRecordset.Close
                            
        vQuery = "select lastprintcount,whcode  from npmaster.dbo.np_paygoods where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vCount = Trim(vRecordset.Fields("lastprintcount").Value)
        End If
        vRecordset.Close
        
        vCount = vCount + 1
        vQuery = "Update npmaster.dbo.np_paygoods  set lastprintcount = " & vCount & " ,lastuserprint = '" & vUserID & "' where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "'"
        gConnection.Execute vQuery
                
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If
End Sub

Public Sub PrintItem015()
   Dim vReportName As String
        Dim vDocNo As String
        Dim vRepID As Integer
        Dim vRepType As String
        Dim vQuery As String
        Dim vRecordset As New ADODB.Recordset
        Dim vCount As Integer
        Dim vWHCode As String
        Dim vWHCode1 As Integer
        Dim vCheckNumber As String, vDocNo1 As String
        Dim vAutoNumber As String, vDocGroup1 As String
        Dim vGenerateNumber, vHeader As String
        Dim vCheck As Integer
        
        On Error GoTo ErrDescription
        
        vDocNo = Trim(TXT3101.Text)
        vWHCode = Trim(CMB3101.Text)
        vQuery = "select * from npmaster.dbo.NP_PayGoods where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vCheckNumber = Trim(vRecordset.Fields("invoiceno").Value)
        End If
        vRecordset.Close
        
        vQuery = "Update npmaster.dbo.NP_PayGoods set PrintedTimeStamp = getdate() where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' "
        gConnection.Execute vQuery
        
        If vCheckNumber = "" Then
                 vQuery = "select header,autonumber  from npmaster.dbo.NP_Generate_DocNo where headertype = 8 "
                If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                    vAutoNumber = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
                    vHeader = Trim(vRecordset.Fields("header").Value)
                End If
                vRecordset.Close
                vGenerateNumber = vHeader & "-" & vAutoNumber
        
                vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1 where headertype = 8 "
                gConnection.Execute vQuery
                
                vQuery = "insert npmaster.dbo.np_paygoods (invoiceno,paynumber,paydatetime,whcode,UserPrint,LastPrintCount,Checked,description) " _
                & " values ('" & vDocNo & "','" & vGenerateNumber & "',getdate(),'015','" & vUserID & "',1,0,'ทดแทนการเปลี่ยนคลังของเอกสาร')"
                gConnection.Execute vQuery
        End If
        
        vDocNo1 = UCase(Left(vDocNo, 3))
        vDocGroup1 = vDocNo1

        If vDocGroup1 = "ICV" Or vDocGroup1 = "ICN" Or vDocGroup1 = "IVD" Or vDocGroup1 = "IVN" Then
        vRepID = 132
        vRepType = "INV"
        Else
        vRepID = 96
        vRepType = "INV"
        End If
vCheck = 1

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
 'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
                            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                                With CrystalReport310
                                    .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
                                    .ParameterFields(0) = "@DocNo;" & vDocNo & ";true"
                                    .ParameterFields(1) = "@vCheck;" & vCheck & ";true"
                                    .Destination = crptToWindow
                                    .WindowState = crptMaximized
                                    .Action = 1
                                End With
                            End If
                            vRecordset.Close
                            
        vQuery = "select lastprintcount,whcode  from npmaster.dbo.np_paygoods where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vCount = Trim(vRecordset.Fields("lastprintcount").Value)
        End If
        vRecordset.Close
        
        vCount = vCount + 1
        vQuery = "Update npmaster.dbo.np_paygoods  set lastprintcount = " & vCount & " ,lastuserprint = '" & vUserID & "' ,PrintedtimeStamp = getdate() where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "'"
        gConnection.Execute vQuery
                
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If
End Sub

Public Sub PrintItem097()
   Dim vReportName As String
        Dim vDocNo As String
        Dim vRepID As Integer
        Dim vRepType As String
        Dim vQuery As String
        Dim vRecordset As New ADODB.Recordset
        Dim vCount As Integer
        Dim vWHCode As String
        Dim vWHCode1 As Integer
        Dim vCheckNumber As String, vDocNo1 As String
        Dim vAutoNumber As String, vDocGroup1 As String
        Dim vGenerateNumber, vHeader As String
        Dim vCheck As Integer
        
        On Error GoTo ErrDescription
        
        vDocNo = Trim(TXT3101.Text)
        vWHCode = Trim(CMB3101.Text)
        vQuery = "select * from npmaster.dbo.NP_PayGoods where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vCheckNumber = Trim(vRecordset.Fields("invoiceno").Value)
        End If
        vRecordset.Close
        
        vQuery = "Update npmaster.dbo.NP_PayGoods set PrintedTimeStamp = getdate() where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' "
        gConnection.Execute vQuery
        
        If vCheckNumber = "" Then
                 vQuery = "select header,autonumber  from npmaster.dbo.NP_Generate_DocNo where headertype = 8 "
                If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                    vAutoNumber = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
                    vHeader = Trim(vRecordset.Fields("header").Value)
                End If
                vRecordset.Close
                vGenerateNumber = vHeader & "-" & vAutoNumber
        
                vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1 where headertype = 8 "
                gConnection.Execute vQuery
                
                vQuery = "insert npmaster.dbo.np_paygoods (invoiceno,paynumber,paydatetime,whcode,UserPrint,LastPrintCount,Checked,description) " _
                & " values ('" & vDocNo & "','" & vGenerateNumber & "',getdate(),'097','" & vUserID & "',1,0,'ทดแทนการเปลี่ยนคลังของเอกสาร')"
                gConnection.Execute vQuery
        End If
        
        vDocNo1 = UCase(Left(vDocNo, 3))
        vDocGroup1 = vDocNo1
        If vDocGroup1 = "ICV" Or vDocGroup1 = "ICN" Or vDocGroup1 = "IVD" Or vDocGroup1 = "IVN" Then
        vRepID = 134
        vRepType = "INV"
        Else
        vRepID = 98
        vRepType = "INV"
        End If
vCheck = 1

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
 'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
                            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                                With CrystalReport310
                                    .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
                                    .ParameterFields(0) = "@DocNo;" & vDocNo & ";true"
                                    .ParameterFields(1) = "@vCheck;" & vCheck & ";true"
                                    .Destination = crptToWindow
                                    .WindowState = crptMaximized
                                    .Action = 1
                                End With
                            End If
                            vRecordset.Close
                            
        vQuery = "select lastprintcount,whcode  from npmaster.dbo.np_paygoods where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vCount = Trim(vRecordset.Fields("lastprintcount").Value)
        End If
        vRecordset.Close
        
        vCount = vCount + 1
        vQuery = "Update npmaster.dbo.np_paygoods  set lastprintcount = " & vCount & " ,lastuserprint = '" & vUserID & "' where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "'"
        gConnection.Execute vQuery
                
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If
End Sub


Public Sub PrintItem014()
        Dim vReportName As String
        Dim vDocNo As String
        Dim vRepID As Integer
        Dim vRepType As String
        Dim vQuery As String
        Dim vRecordset As New ADODB.Recordset
        Dim vCount As Integer
        Dim vWHCode As String
        Dim vWHCode1 As Integer
        Dim vCheckNumber As String, vDocGroup1 As String
        Dim vAutoNumber As String, vDocNo1 As String
        Dim vGenerateNumber, vHeader As String
        Dim vCheck As Integer
        
        On Error GoTo ErrDescription
        
        vDocNo = Trim(TXT3101.Text)
        vWHCode = Trim(CMB3101.Text)
        vQuery = "select * from npmaster.dbo.NP_PayGoods where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vCheckNumber = Trim(vRecordset.Fields("invoiceno").Value)
        End If
        vRecordset.Close
        
        vQuery = "Update npmaster.dbo.NP_PayGoods set PrintedTimeStamp = getdate() where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' "
        gConnection.Execute vQuery
        
        If vCheckNumber = "" Then
                 vQuery = "select header,autonumber  from npmaster.dbo.NP_Generate_DocNo where headertype = 8 "
                If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                    vAutoNumber = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
                    vHeader = Trim(vRecordset.Fields("header").Value)
                End If
                vRecordset.Close
                vGenerateNumber = vHeader & "-" & vAutoNumber
        
                vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1 where headertype = 8 "
                gConnection.Execute vQuery
                
                vQuery = "insert npmaster.dbo.np_paygoods (invoiceno,paynumber,paydatetime,whcode,UserPrint,LastPrintCount,Checked,description) " _
                & " values ('" & vDocNo & "','" & vGenerateNumber & "',getdate(),'014','" & vUserID & "',1,0,'ทดแทนการเปลี่ยนคลังของเอกสาร')"
                gConnection.Execute vQuery
        End If
        
        vDocNo1 = UCase(Left(vDocNo, 3))
        vDocGroup1 = vDocNo1
        If vDocGroup1 = "ICV" Or vDocGroup1 = "ICN" Or vDocGroup1 = "IVD" Or vDocGroup1 = "IVN" Then
        vRepID = 172
        vRepType = "INV"
        Else
        vRepID = 171
        vRepType = "INV"
        End If
vCheck = 1

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
 'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
                            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                                With CrystalReport310
                                    .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
                                    .ParameterFields(0) = "@DocNo;" & vDocNo & ";true"
                                    .ParameterFields(1) = "@vCheck;" & vCheck & ";true"
                                    .Destination = crptToWindow
                                    .WindowState = crptMaximized
                                    .Action = 1
                                End With
                            End If
                            vRecordset.Close
                            
        vQuery = "select lastprintcount,whcode  from npmaster.dbo.np_paygoods where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vCount = Trim(vRecordset.Fields("lastprintcount").Value)
        End If
        vRecordset.Close
        
        vCount = vCount + 1
        vQuery = "Update npmaster.dbo.np_paygoods  set lastprintcount = " & vCount & " ,lastuserprint = '" & vUserID & "' where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "'"
        gConnection.Execute vQuery
                
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If
End Sub


Public Sub PrintItem020()
        Dim vReportName As String
        Dim vDocNo As String
        Dim vRepID As Integer
        Dim vRepType As String
        Dim vQuery As String
        Dim vRecordset As New ADODB.Recordset
        Dim vCount As Integer
        Dim vWHCode As String
        Dim vWHCode1 As Integer
        Dim vCheckNumber As String, vDocGroup1 As String
        Dim vAutoNumber As String, vDocNo1 As String
        Dim vGenerateNumber, vHeader As String
        Dim vCheck As Integer
        
        On Error GoTo ErrDescription
        
        vDocNo = Trim(TXT3101.Text)
        vWHCode = Trim(CMB3101.Text)
        vQuery = "select * from npmaster.dbo.NP_PayGoods where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vCheckNumber = Trim(vRecordset.Fields("invoiceno").Value)
        End If
        vRecordset.Close
        
        vQuery = "Update npmaster.dbo.NP_PayGoods set PrintedTimeStamp = getdate() where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' "
        gConnection.Execute vQuery
        
        If vCheckNumber = "" Then
                 vQuery = "select header,autonumber  from npmaster.dbo.NP_Generate_DocNo where headertype = 8 "
                If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                    vAutoNumber = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
                    vHeader = Trim(vRecordset.Fields("header").Value)
                End If
                vRecordset.Close
                vGenerateNumber = vHeader & "-" & vAutoNumber
        
                vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1 where headertype = 8 "
                gConnection.Execute vQuery
                
                vQuery = "insert npmaster.dbo.np_paygoods (invoiceno,paynumber,paydatetime,whcode,UserPrint,LastPrintCount,Checked,description) " _
                & " values ('" & vDocNo & "','" & vGenerateNumber & "',getdate(),'020','" & vUserID & "',1,0,'ทดแทนการเปลี่ยนคลังของเอกสาร')"
                gConnection.Execute vQuery
        End If
        
        vDocNo1 = UCase(Left(vDocNo, 3))
        vDocGroup1 = vDocNo1
        If vDocGroup1 = "ICV" Or vDocGroup1 = "ICN" Or vDocGroup1 = "IVD" Or vDocGroup1 = "IVN" Then
        vRepID = 196
        vRepType = "INV"
        Else
        vRepID = 195
        vRepType = "INV"
        End If
vCheck = 1

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
 'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
                            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                                With CrystalReport310
                                    .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
                                    .ParameterFields(0) = "@DocNo;" & vDocNo & ";true"
                                    .ParameterFields(1) = "@vCheck;" & vCheck & ";true"
                                    .Destination = crptToWindow
                                    .WindowState = crptMaximized
                                    .Action = 1
                                End With
                            End If
                            vRecordset.Close
                            
        vQuery = "select lastprintcount,whcode  from npmaster.dbo.np_paygoods where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vCount = Trim(vRecordset.Fields("lastprintcount").Value)
        End If
        vRecordset.Close
        
        vCount = vCount + 1
        vQuery = "Update npmaster.dbo.np_paygoods  set lastprintcount = " & vCount & " ,lastuserprint = '" & vUserID & "' where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "'"
        gConnection.Execute vQuery
                
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If
End Sub

Public Sub PrintItem016()
        Dim vReportName As String
        Dim vDocNo As String
        Dim vRepID As Integer
        Dim vRepType As String
        Dim vQuery As String
        Dim vRecordset As New ADODB.Recordset
        Dim vCount As Integer
        Dim vWHCode As String
        Dim vWHCode1 As Integer
        Dim vCheckNumber As String, vDocGroup1 As String
        Dim vAutoNumber As String, vDocNo1 As String
        Dim vGenerateNumber, vHeader As String
        Dim vCheck As Integer
        
        On Error GoTo ErrDescription
        
        vDocNo = Trim(TXT3101.Text)
        vWHCode = Trim(CMB3101.Text)
        vQuery = "select * from npmaster.dbo.NP_PayGoods where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vCheckNumber = Trim(vRecordset.Fields("invoiceno").Value)
        End If
        vRecordset.Close
        
        vQuery = "Update npmaster.dbo.NP_PayGoods set PrintedTimeStamp = getdate() where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' "
        gConnection.Execute vQuery
        
        If vCheckNumber = "" Then
                 vQuery = "select header,autonumber  from npmaster.dbo.NP_Generate_DocNo where headertype = 8 "
                If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                    vAutoNumber = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
                    vHeader = Trim(vRecordset.Fields("header").Value)
                End If
                vRecordset.Close
                vGenerateNumber = vHeader & "-" & vAutoNumber
        
                vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1 where headertype = 8 "
                gConnection.Execute vQuery
                
                vQuery = "insert npmaster.dbo.np_paygoods (invoiceno,paynumber,paydatetime,whcode,UserPrint,LastPrintCount,Checked,description) " _
                & " values ('" & vDocNo & "','" & vGenerateNumber & "',getdate(),'016','" & vUserID & "',1,0,'ทดแทนการเปลี่ยนคลังของเอกสาร')"
                gConnection.Execute vQuery
        End If
        
        vDocNo1 = UCase(Left(vDocNo, 3))
        vDocGroup1 = vDocNo1
        If vDocGroup1 = "ICV" Or vDocGroup1 = "ICN" Or vDocGroup1 = "IVD" Or vDocGroup1 = "IVN" Then
        vRepID = 219
        vRepType = "INV"
        Else
        vRepID = 218
        vRepType = "INV"
        End If
vCheck = 1

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
 'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
                            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                                With CrystalReport310
                                    .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
                                    .ParameterFields(0) = "@DocNo;" & vDocNo & ";true"
                                    .ParameterFields(1) = "@vCheck;" & vCheck & ";true"
                                    .Destination = crptToWindow
                                    .WindowState = crptMaximized
                                    .Action = 1
                                End With
                            End If
                            vRecordset.Close
                            
        vQuery = "select lastprintcount,whcode  from npmaster.dbo.np_paygoods where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vCount = Trim(vRecordset.Fields("lastprintcount").Value)
        End If
        vRecordset.Close
        
        vCount = vCount + 1
        vQuery = "Update npmaster.dbo.np_paygoods  set lastprintcount = " & vCount & " ,lastuserprint = '" & vUserID & "' where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "'"
        gConnection.Execute vQuery
                
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If

End Sub


Public Sub PrintItem070()
Dim vReportName As String
Dim vDocNo As String
Dim vRepID As Integer
Dim vRepType As String
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vCount As Integer
Dim vWHCode As String
Dim vWHCode1 As Integer
Dim vCheckNumber As String, vDocGroup1 As String
Dim vAutoNumber As String, vDocNo1 As String
Dim vGenerateNumber, vHeader As String
Dim vCheck As Integer

        On Error GoTo ErrDescription
        
        vDocNo = Trim(TXT3101.Text)
        vWHCode = Trim(CMB3101.Text)
        vQuery = "select * from npmaster.dbo.NP_PayGoods where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vCheckNumber = Trim(vRecordset.Fields("invoiceno").Value)
        End If
        vRecordset.Close
        
        vQuery = "Update npmaster.dbo.NP_PayGoods set PrintedTimeStamp = getdate() where invoiceno = '" & vDocNo & "'  and whcode = '" & vWHCode & "' "
        gConnection.Execute vQuery
        
        If vCheckNumber = "" Then
                 vQuery = "select header,autonumber  from npmaster.dbo.NP_Generate_DocNo where headertype = 8 "
                If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                    vAutoNumber = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
                    vHeader = Trim(vRecordset.Fields("header").Value)
                End If
                vRecordset.Close
                vGenerateNumber = vHeader & "-" & vAutoNumber
        
                vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1 where headertype = 8 "
                gConnection.Execute vQuery
                
                vQuery = "insert npmaster.dbo.np_paygoods (invoiceno,paynumber,paydatetime,whcode,UserPrint,LastPrintCount,Checked,description) " _
                & " values ('" & vDocNo & "','" & vGenerateNumber & "',getdate(),'070','" & vUserID & "',1,0,'ทดแทนการเปลี่ยนคลังของเอกสาร')"
                gConnection.Execute vQuery
        End If
        
        vDocNo1 = UCase(Left(vDocNo, 3))
        vDocGroup1 = vDocNo1
        If vDocGroup1 = "ICV" Or vDocGroup1 = "ICN" Or vDocGroup1 = "IVD" Or vDocGroup1 = "IVN" Then
        vRepID = 299
        vRepType = "INV"
        Else
        vRepID = 297
        vRepType = "INV"
        End If
vCheck = 1

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
 'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
                            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                                With CrystalReport310
                                    .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
                                    .ParameterFields(0) = "@DocNo;" & vDocNo & ";true"
                                    .ParameterFields(1) = "@vCheck;" & vCheck & ";true"
                                    .Destination = crptToWindow
                                    .WindowState = crptMaximized
                                    .Action = 1
                                End With
                            End If
                            vRecordset.Close
                            
        vQuery = "select lastprintcount,whcode  from npmaster.dbo.np_paygoods where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vCount = Trim(vRecordset.Fields("lastprintcount").Value)
        End If
        vRecordset.Close
        
        vCount = vCount + 1
        vQuery = "Update npmaster.dbo.np_paygoods  set lastprintcount = " & vCount & " ,lastuserprint = '" & vUserID & "' where invoiceno = '" & vDocNo & "' and whcode = '" & vWHCode & "'"
        gConnection.Execute vQuery
                
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If
End Sub

Public Sub PrintPayItem(vDocNo As String, vPickZoneGroup As String)
Dim vPrinterName As String
Dim printerObj As Printer
Dim vRecordset As New ADODB.Recordset
Dim vRecordset1 As New ADODB.Recordset
Dim vQuery As String
Dim vLineX As Integer
Dim vLineY As Integer
Dim vStartX As Integer
Dim vStartY As Integer
Dim i As Integer
Dim prnPrinter As Printer
Dim lngRetVal As Long
Dim Driver As String
Dim n As Integer
Dim vItemName As String
Dim vSoStatus As Integer
Dim vSelectPicked As Integer
Dim vGroupDocNo As String
Dim vPrinterID As Integer


vQuery = "exec dbo.USP_NP_SearchCheckPrinter " & vPrinterID & " "
If OpenDataBase(gConnection, vRecordset1, vQuery) <> 0 Then
vPrinterName = Trim(vRecordset1.Fields("printername").Value)
End If
vRecordset1.Close

For Each printerObj In Printers
If UCase(printerObj.DeviceName) = UCase(vPrinterName) Then
Set Printer = printerObj
Set printerObj = Nothing
Exit For
End If
Next


    vQuery = "exec dbo.USP_INV_PayItemSlip 1,'" & vDocNo & "','" & vPickZoneGroup & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 10
      Printer.CurrentX = 0
      Printer.CurrentY = 0
      Printer.Print Trim("_______________________________________________________________________________________")

      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 40
      Printer.FontBold = True
      Printer.CurrentX = 0
      Printer.CurrentY = 80
      Printer.Print "(ใบจ่ายสินค้า)"
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 10
      Printer.CurrentX = 0
      Printer.CurrentY = 0
      Printer.Print Trim("_______________________________________________________________________________________")

      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 50
      Printer.FontBold = True
      Printer.CurrentX = 1600
      Printer.CurrentY = 0
      Printer.Print Trim(vRecordset.Fields("docno").Value)
      
      Printer.Font.Name = "3 of 9 Barcode"
      Printer.Font.Size = 40
      Printer.FontBold = False
      Printer.CurrentX = 1200
      Printer.CurrentY = 1000
      Printer.Print "*" & Trim(vRecordset.Fields("docno").Value) & "*"
 
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 10
      Printer.CurrentX = 0
      Printer.CurrentY = 1400
      Printer.Print Trim("_______________________________________________________________________________________")
      
      

    Printer.Font.Name = "AngsanaUPC"
    Printer.Font.Size = 12
    Printer.CurrentX = 1500
    Printer.CurrentY = 1650
    Printer.Print Trim("ต้นฉบับ ใบจ่ายสินค้า")
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 11
      Printer.CurrentX = 0
      Printer.CurrentY = 1900
      Printer.Print Trim("วันที่คิว: ") & Trim(vRecordset.Fields("queuedate").Value)
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 10
      Printer.CurrentX = 2000
      Printer.CurrentY = 1900
      Printer.Print Trim("เลขที่เอกสาร : ") & Trim(vRecordset.Fields("docno").Value)
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 11
      Printer.CurrentX = 2200
      Printer.CurrentY = 2150
      Printer.Print Trim("วันที่เอกสาร : ") & Trim(vRecordset.Fields("docdate").Value)
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 11
      Printer.CurrentX = 0
      Printer.CurrentY = 2400
      Printer.Print Trim("รหัสลูกค้า : ") & Trim(vRecordset.Fields("arcode").Value)
      

      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 11
      Printer.CurrentX = 0
      Printer.CurrentY = 2650
      Printer.Print Trim("ชื่อลูกค้า : ") & Trim(vRecordset.Fields("name1").Value)
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 11
      Printer.CurrentX = 0
      Printer.CurrentY = 2900
      Printer.Print Trim("พนักงานขาย : ") & Trim(vRecordset.Fields("salecode").Value) & "/" & Trim(vRecordset.Fields("salename").Value)
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 11
      Printer.CurrentX = 0
      Printer.CurrentY = 3150
      
      If vRecordset.Fields("isconditionsend").Value = 0 Then
            Printer.Print Trim("วิธีการจัดส่ง : ") & Trim("รับเอง")
      Else
            Printer.Print Trim("วิธีการจัดส่ง : ") & Trim("ส่งให้")
      End If
                  
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 16
      Printer.FontBold = True
      Printer.CurrentX = 0
      Printer.CurrentY = 3400
      Printer.Print Trim("คลัง : ") & Trim(vRecordset.Fields("whcode").Value) & "/" & Trim(vRecordset.Fields("shelfcode").Value)

      If Trim(vRecordset.Fields("carlicense").Value) <> "" Then
        Printer.Font.Name = "AngsanaUPC"
        Printer.Font.Size = 16
        Printer.CurrentX = 1400
        Printer.CurrentY = 3400
        Printer.FontBold = True
        Printer.FontUnderline = True
        Printer.Print Trim("ทะเบียนรถขนส่ง : ") & Trim(vRecordset.Fields("carlicense").Value)
      End If
            
      If vSoStatus = 0 Then
         Printer.Font.Name = "AngsanaUPC"
         Printer.Font.Size = 14
         Printer.FontBold = True
         Printer.FontUnderline = False
         Printer.CurrentX = 0
         Printer.CurrentY = 3800
         Printer.Print Trim("เวลารับของ : ") & Trim(vRecordset.Fields("requesttime").Value)
      Else
         Printer.Font.Name = "AngsanaUPC"
         Printer.Font.Size = 14
         Printer.FontBold = True
         Printer.FontUnderline = False
         Printer.CurrentX = 0
         Printer.CurrentY = 3800
         Printer.Print Trim("วันที่ครบกำหนดรับของ : ") & Trim(vRecordset.Fields("duedate").Value)
      End If
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 14
      Printer.CurrentX = 0
      Printer.CurrentY = 4150
      Printer.Print Trim(vRecordset.Fields("customerzone").Value)
      
      vRecordset.MoveFirst
      vLineX = 50
      vLineY = 50
      
      Printer.Font.Name = "AngsanaUPC"
      Printer.Font.Size = 10
      Printer.CurrentX = 0
      Printer.CurrentY = Printer.CurrentY - 30
      Printer.FontBold = False
      Printer.FontUnderline = False
      Printer.Print Trim("-----------------------------------------------------------------------------------------------")
      n = 1
      While Not vRecordset.EOF
          Printer.Font.Size = 11
          
          Printer.CurrentX = 0
          Printer.CurrentY = Printer.CurrentY
          Printer.Print "ที่เก็บ :" & Trim(vRecordset.Fields("shelfcode1").Value) & "    " & Trim("OnHand") & "(" & Trim(vRecordset.Fields("shelfcode").Value) & ")" & ": " & Trim(vRecordset.Fields("qtylocation").Value) & "  " & Trim(vRecordset.Fields("stkunitcode").Value) & "     " & "รวมคลัง :" & Trim(vRecordset.Fields("StkWHCode").Value) & "  " & Trim(vRecordset.Fields("stkunitcode").Value)
          
          Printer.CurrentX = 0
          Printer.CurrentY = Printer.CurrentY
          Printer.Print n & ". " & "รหัสสินค้า :" & Trim(vRecordset.Fields("itemcode").Value) & "             " & " ขายชั้นเก็บ :" & Trim(vRecordset.Fields("shelfcode").Value)
          
          vItemName = Trim(vRecordset.Fields("itemname").Value) & Trim(vRecordset.Fields("descriptionline"))
          If Len(vItemName) <= 55 Then
             Printer.CurrentX = 0
             Printer.CurrentY = Printer.CurrentY
             Printer.Print "ชื่อสินค้า :" & vItemName
          Else
             Printer.CurrentX = 0
             Printer.CurrentY = Printer.CurrentY
             Printer.Print "ชื่อสินค้า :" & Left(vItemName, 55)
             
             Printer.CurrentX = 600
             Printer.CurrentY = Printer.CurrentY
             Printer.Print Right(vItemName, Len(vItemName) - 55)
          End If
          
          Printer.Font.Size = 13
          Printer.CurrentX = Printer.CurrentX + 15
          Printer.CurrentY = Printer.CurrentY + 100
          Printer.FontBold = True
          Printer.Print "ต้องการ :" & Format(Trim(vRecordset.Fields("qty").Value), "##,##0.00") & "    " & Trim(vRecordset.Fields("unitcode").Value) & "                         " & Trim("จัดได้ : ______________")
          
          Printer.CurrentX = 0
          Printer.CurrentY = Printer.CurrentY - 50
          Printer.FontBold = False
          Printer.Print Trim("-----------------------------------------------------------------------------------------------")
          
      vRecordset.MoveNext
      n = n + 1
      Wend
    End If
    vRecordset.Close
    Printer.Font.Size = 11
    Printer.CurrentX = 0
    Printer.CurrentY = Printer.CurrentY - 400
    Printer.Print Trim("_______________________________________________________________________________________________")
    
    Printer.CurrentX = 0
    Printer.CurrentY = Printer.CurrentY + vLineY
    Printer.Print "               ผู้จัดสินค้า                                             ผู้รับสินค้า"
    
    Printer.CurrentX = 0
    Printer.CurrentY = Printer.CurrentY + 150
    Printer.Print "         _____________                                    ______________"
          
    Printer.CurrentX = 0
    Printer.CurrentY = Printer.CurrentY
    Printer.Print Trim("______________________________________________________________________________________________")
    
    Printer.Font.Size = 10
    Printer.CurrentX = 0
    Printer.CurrentY = Printer.CurrentY
    Printer.Print Trim("วันที่พิมพ์ :") & Now & "          " & vPrinterName
    Printer.EndDoc
End Sub
