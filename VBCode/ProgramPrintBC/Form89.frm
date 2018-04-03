VERSION 5.00
Begin VB.Form Form89 
   Caption         =   "คำนวณจำนวนสินค้าในใบขอโอน"
   ClientHeight    =   8385
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form89.frx":0000
   ScaleHeight     =   8385
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CMDCal 
      Caption         =   "กดปุ่มคำนวณ"
      Height          =   615
      Left            =   4425
      TabIndex        =   1
      Top             =   2250
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox TXTCalTransfer 
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
      Height          =   465
      Left            =   2475
      TabIndex        =   0
      Top             =   1500
      Width           =   3540
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "เลขที่ใบขอโอนสินค้า"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   465
      Left            =   450
      TabIndex        =   3
      Top             =   1500
      Width           =   1965
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "คำนวณจำนวนสินค้าคงเหลือในใบขอโอนสินค้า"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   14.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   540
      Left            =   2475
      TabIndex        =   2
      Top             =   225
      Width           =   7590
   End
End
Attribute VB_Name = "Form89"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMDCal_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocNo As String
Dim vCheckDoc As String
Dim vItemCode As String

'On Error GoTo ErrDescription

vDocNo = Trim(TXTCalTransfer.Text)

vQuery = "select * from bcnp.dbo.vw_tf_bcstktransfer2 where docno = '" & vDocNo & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vRecordset.MoveFirst
        Do Until vRecordset.EOF
                If IsNull(Trim(vRecordset.Fields("mydescription").Value)) Then
                        vItemCode = Trim(vRecordset.Fields("itemcode").Value)
                        vQuery = "update bcnp.dbo.vw_tf_bcstktransfer2 " _
                            & " set mydescription =  CONVERT(char(10), qty) where docno = '" & vDocNo & "' and itemcode = '" & vItemCode & "' "
                        gConnection.Execute vQuery
                End If
        vRecordset.MoveNext
        Loop
End If
vRecordset.Close

vQuery = "select * from bcnp.dbo.vw_tf_bcstktransfer2 where docno = '" & vDocNo & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vRecordset.MoveFirst
        Do Until vRecordset.EOF
                If Trim(vRecordset.Fields("qty").Value) <> 0 Then
                        vItemCode = Trim(vRecordset.Fields("itemcode").Value)
                        vQuery = "update bcnp.dbo.vw_tf_bcstktransfer2 " _
                                        & " set qty =  convert(numeric(10),mydescription)-qtytransfer  where docno = '" & vDocNo & "' and itemcode = '" & vItemCode & "' "
                        gConnection.Execute vQuery
                End If
        vRecordset.MoveNext
        Loop
End If
MsgBox "ได้ทำการคำนวณจำนวนสินค้าที่ขอโอนเรียบร้อยแล้ว"
CMDCal.Visible = False
vRecordset.Close
TXTCalTransfer.Text = ""
End Sub

Private Sub TXTCalTransfer_KeyPress(KeyAscii As Integer)
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocNo As String
Dim vCheckDoc As String
Dim vCountItem As Integer

'On Error GoTo ErrDescription

If KeyAscii = 13 Then
        vDocNo = Trim(TXTCalTransfer.Text)
        vQuery = "select * from bcnp.dbo.vw_tf_bcstktransfer2 where docno = '" & vDocNo & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vCheckDoc = Trim(vRecordset.Fields("docno").Value)
        Else
            MsgBox "ไม่มีเอกสารนี้ในระบบ"
            Exit Sub
        End If
        vRecordset.Close

        vQuery = "select * from bcnp.dbo.vw_tf_bcstktransfer2 where docno = '" & vDocNo & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vCountItem = vRecordset.RecordCount
            vRecordset.MoveFirst
            Do Until vRecordset.EOF
            If vRecordset.Fields("qty").Value <> vRecordset.Fields("qtytransfer").Value Then
            MsgBox "มีสินค้า ที่สามารถดึงไปทำใบโอนสินค้าได้"
            CMDCal.Visible = True
            Exit Sub
            End If
            vRecordset.MoveNext
            Loop
        End If
        vRecordset.Close
        
        
End If
End Sub
