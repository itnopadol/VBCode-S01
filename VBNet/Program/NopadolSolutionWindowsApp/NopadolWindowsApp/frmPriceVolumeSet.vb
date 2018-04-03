Imports System.Windows.Forms
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports System.Globalization
Imports Microsoft.VisualBasic
Public Class frmPriceVolumeSet
    Dim QryString As String
    Dim iQryString As String
    Dim da As SqlDataAdapter
    Dim vReadQuery As SqlDataReader
    Dim ds As DataSet
    Dim dt As DataTable
    Dim n As Integer
    Dim i As Integer
    Dim cmd As SqlCommand
    Dim DeptCode As String
    Dim BrandCode As String
    Dim TypeCode As String
    Dim CateCode As String
    '-----------Var for Insert Update Master
    Dim vDocNo As String
    Dim vDocdate As DateTime
    Dim vStartDate As DateTime
    Dim vEndDate As DateTime
    Dim LineNumber As String
    Dim ItemNo As String
    Dim ItemName1 As String
    Dim ItemUnit As String
    Dim vPriceLP As Double
    Dim vPrice1 As Double 'ราคาที่ 1
    Dim vPriceLevel As Double 'ระดับราคา
    Dim vPercentSMP1 As Double '%ระดับราคาที่1
    Dim vSmartPoint1 As Double 'smartpoint1
    Dim vVolume As Double 'Volume
    Dim vDCPrice1 As Double
    Dim vPrice As Double
    Dim vPercentSMP As Double
    Dim vSmartPoint As Double
    Dim vMKcost As Double
    Dim vGPMKcost As Double
    Dim vAvgCostlot As Double
    Dim vGPAvgCostlot As Double
    Dim vPSDocNo As String
    Dim pNewdoc As Integer ' สถานะของเลขที่เอกสาร

    '-----------Var for Insert Update Master
    '---------Search Var----------
    Dim FdocNo As String
    Dim vTypeSL As String
    Dim vDate As DateTime
    Dim chkKey As Integer
    Dim vNewdocNo As String
    Dim iNewPSVdoc As String
    '----

    'Dim frmPriceVolumeSet As frmPriceVolumeSet


    Private Sub btnProduct_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnProduct.Click
        If Me.LPcbx.Checked = True Then
            lpStatus = 1
        Else
            lpStatus = 0
        End If
        vDate = DateAdd(DateInterval.Day, 1, Now.Date).ToShortDateString
        Dim ivLM3 As String
        Dim ivLM2 As String
        Dim vDC2 As String
        Dim vDCrs1 As String
        Dim vDC3 As String
        'Dim vDCrs1 As String
        Dim vDCrs2 As String
        Dim vd1 As String
        Dim vd2 As String
        Dim i1 As Double
        Dim i2 As Double
        Dim i3 As Double
        Dim i4 As Double
        '---------------------
        Dim rs1 As String
        Dim rs2 As String
        Dim vrs1 As String
        '--------------------
        ivLM2 = Me.txtVM2.Text
        ivLM3 = Me.txtVM3.Text
        vDC2 = Me.txtDC2.Text
        vDC3 = Me.txtDC3.Text
        vDCrs1 = CHKdcPC(vDC2)
        vDCrs2 = CHKdcPC(vDC3)
        On Error GoTo chkChar
        i1 = CDbl(ivLM2)
        i2 = CDbl(ivLM3)
        If Me.DTPStartDate.Value < vDate Then
            MsgBox("วันที่ปรับราคาต้องไม่น้อยกว่าวันพรุ่งนี้.กรุณากำหนดใหม่", MsgBoxStyle.Critical, "Error")
            Me.DTPStartDate.Focus()
        Else
            
            If Me.txtVM2.Text <> "" And Me.txtVM3.Text <> "" And Me.txtDC2.Text <> "" And Me.txtDC3.Text <> "" And Me.smpLV1.Value > 0 Then
               
                'Me.P01.Visible = False
                'Me.P02.Visible = True
                'Call GetDepartment()
                'Call GetBrand()
                'Call GetItemType()
                'Me.cbxCategory.Enabled = False
                'Me.GroupBox1.Enabled = False
                'Me.smpLV1.Enabled = False
                'Me.GB01.Enabled = False
                'Me.btnGenerate.Enabled = False
                '--------------------------
                vd1 = chkBPD(vDC2, vDC3)
                'vd2 = chkBPD(vDC3)
                'If (vd1 = "0" And vd2 = "i") Or (vd1 = "i" And vd2 = "x") Then
                If (vd1 = "0") Then
                    MsgBox("คุณกำหนดหน่วยส่วนลดไม่เหมือนกัน กรุณากำหนดใหม่", MsgBoxStyle.Critical, "Warnings")
                Else
                    If ivLM2 <> "" And ivLM3 <> "" And (CDbl(ivLM2) > 1) And (CDbl(ivLM3) >= CDbl(ivLM2)) And (CDbl(vDCrs2) >= CDbl(vDCrs1)) Then
                        rs1 = CHKVMDC(ivLM2, ivLM3)
                        rs2 = CHKVMDC(vDCrs1, vDCrs2)
                        vrs1 = CHKVMDC1(ivLM2, ivLM3, vDCrs1, vDCrs2)
                        If rs1 = 0 Or rs2 = 0 Then
                            MsgBox("ค่าของVolume 2 หรือ ค่าส่วนลดของราคาที่ 2 ต้องน้อยกว่า ค่าของระดับราคาที่ 3 กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Warnings")
                        ElseIf vrs1 = 0 Then
                            MsgBox("ค่าของ Volume และส่วนลดไม่ถูกต้อง กรุณาตรวจสอบใหม่", MsgBoxStyle.Critical, "Warnings")
                        Else
                            pvLV2 = Me.txtLv2.Text
                            pvLV3 = Me.txtLv3.Text
                            pvVM2 = Me.txtVM2.Text
                            pvVM3 = Me.txtVM3.Text
                            pvDC2 = Me.txtDC2.Text
                            pvDC3 = Me.txtDC3.Text
                            pvSMP2 = Me.txtSMTP2.Text
                            pvSMP3 = Me.txtSMTP3.Text
                            Me.Enabled = False
                            dlgPSVdocSearch.Show()
                            dlgPSVdocSearch.txtPSDoc.Focus()
                            Me.Text = "[::]เตรียมสินค้าสำหรับกำหนดส่วนลด"
                        End If
                    Else

                        MsgBox("ค่า Volume2 อาจน้อยกว่าหรือเท่ากับ 1 หรือ ค่าของข้อมูลระดับราคาที่ 3 น้อยกว่าระดับราคาที่ 2 กรุณาตรวจสอบอีกครั้ง", MsgBoxStyle.Critical, "Warnings")
                        'Me.txtVM2.Focus()
                    End If

                    'pvLV2 = Me.txtLv2.Text
                    'pvLV3 = Me.txtLv3.Text
                    'pvVM2 = Me.txtVM2.Text
                    'pvVM3 = Me.txtVM3.Text
                    'pvDC2 = Me.txtDC2.Text
                    'pvDC3 = Me.txtDC3.Text
                    'pvSMP2 = Me.txtSMTP2.Text
                    'pvSMP3 = Me.txtSMTP3.Text
                    'Me.Enabled = False
                    'dlgPSVdocSearch.Show()
                    'dlgPSVdocSearch.txtPSDoc.Focus()
                    'Me.Text = "[::]เตรียมสินค้าสำหรับกำหนดส่วนลด"
                End If

            Else
                MsgBox("คุณใส่ข้อมูลไม่ครบ..กรุณาตรวจสอบอีกครั้ง.", MsgBoxStyle.Critical, "Error")
            End If

chkChar:
            If Err.Description <> "" Then
                MsgBox("คุณป้อนข้อมูลไม่ถูกต้อง กรุณาป้อนใหม่", MsgBoxStyle.Critical, "Error")
            End If

        End If
    End Sub
    Private Function chkBPD(ByVal VN As String, ByVal VN1 As String) As String
        Dim i, i1 As Integer
        Dim x, x1 As Integer
        i = Microsoft.VisualBasic.Len(VN)
        i1 = Microsoft.VisualBasic.Len(VN1)
        i = InStr(VN, "%")
        i1 = InStr(VN1, "%")
        x = Microsoft.VisualBasic.Len(VN)
        x1 = Microsoft.VisualBasic.Len(VN1)
        x = InStr(VN, "#$@()*+=-/\,a-b'")
        x1 = InStr(VN, "#$@()*+=-/\,a-b'")

        If i = 0 And i1 = 0 Then
            VN = "1"
        ElseIf i = 1 And i1 = 1 Then
            VN = "1"
        ElseIf i > 1 Or i1 > 1 Then
            VN = "1"
        ElseIf x > 0 Or x1 > 0 Then
            VN = "0"
        Else
            VN = "0"
        End If
        Return VN
    End Function
    Private Function CHKVMDC(ByVal vC1 As Double, ByVal vC2 As Double) As Double
        If vC1 > vC2 Then
            vC1 = 0
        End If
        Return vC1
    End Function
    Private Function CHKVMDC1(ByVal v1 As Double, ByVal v2 As Double, ByVal d1 As Double, ByVal d2 As Double) As Double
        If v1 < v2 And d1 = d2 Then
            v1 = 0
        ElseIf v1 < v2 And d1 > d2 Then
            v1 = 0
        ElseIf v1 > v2 And d1 < d2 Then
            v1 = 0
        ElseIf v1 = v2 And d1 > d2 Then
            v1 = 0
        End If
        Return v1
    End Function
    Private Sub frmPriceVolumeSet_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.DTPdocDate.Text = Now.Date
        Me.DTPStartDate.Text = DateAdd(DateInterval.Day, 1, Now.Date)
        Me.DTPEndDate.Text = DateAdd(DateInterval.Year, 5, Now.Date)
        Call InitializeDataBase()
        'DateTimePickerFormat.Short.ToString("dd/mm/yyyy")
        Me.BtnGenVLM.Visible = False
        Me.PBNew.Visible = False
        Me.PBConfirm.Visible = False
        Me.okFind.Visible = False
        Me.txtDocno.ReadOnly = True
        Me.txtDocno.ReadOnly = True
        Me.txtDocno.BackColor = Color.LightSkyBlue
        Me.txtDocno.ForeColor = Color.Blue
        Me.txtDocno.Font.Bold.ToString()
        Me.txtLv2.Text = "2"
        Me.txtLv3.Text = "3"
        Me.txtLv2.BackColor = Color.LightSkyBlue
        Me.txtLv3.BackColor = Color.LightSkyBlue
        Me.txtLv2.ReadOnly = True
        Me.txtLv3.ReadOnly = True
        Me.txtVM2.BackColor = Color.White
        Me.txtVM3.BackColor = Color.White
        Me.txtDC2.BackColor = Color.White
        Me.txtDC3.BackColor = Color.White
        Me.txtSMTP2.Text = "0.5"
        Me.txtSMTP3.Text = "0.25"
        Me.txtSMTP2.BackColor = Color.White
        Me.txtSMTP3.BackColor = Color.White
        Me.Text = "[::] กำหนดราคาตามจำนวน Price Volume Set V.1"
        Me.btnPrint.Enabled = False
        Me.btnSave.Enabled = False
        Me.btnProduct.Enabled = False
        Me.btnSaveAS.Enabled = False
        'Me.btnSelect.Enabled = False
        Me.txtVM2.Enabled = False
        Me.txtVM3.Enabled = False
        Me.txtDC2.Enabled = False
        Me.txtDC3.Enabled = False
        Me.btnNewDoc.Focus()
        pNewdoc = 1

        
        'List Combobox at P02       
    End Sub

    
    Private Sub GetData()
        Dim i As Integer
        Dim dt1 As New DataTable("vData")
        Dim dr As DataRow
        Dim itemcode As String = "xxxx" 'x ค่าสมมุติ
        'Dim status As String
        If itemcode <> "" Then
            dt1.Columns.Add("ลำดับ", GetType(Integer))
            dt1.Columns.Add("รหัสสินค้า", GetType(String))
            dt1.Columns.Add("ชื่อสินค้า", GetType(String))
            dt1.Columns.Add("หน่วยขาย", GetType(String))
            dt1.Columns.Add("ราคาที่1", GetType(String))
            dt1.Columns.Add("ระดับราคา", GetType(String))
            dt1.Columns.Add("Volume", GetType(String))
            dt1.Columns.Add("ส่วนลด%จากราคาที่1", GetType(String))
            dt1.Columns.Add("%SmartPoint", GetType(String))
            dt1.Columns.Add("ทุนตลาดSaleVat", GetType(String))
            dt1.Columns.Add("GP ทุนตลาด", GetType(String))
            dt1.Columns.Add("ทุนเฉลี่ย LotSaleVat", GetType(String))
            dt.Columns.Add("GP ทุนเฉลี่ย Lot", GetType(String))
            QryString = "" 'รอ Query
            da = New SqlDataAdapter(QryString, vConnectionString)
            ds = New DataSet
            da.Fill(ds, "itemList")
            dt = ds.Tables("itemList")
            If dt.Rows.Count > 0 Then
                For i = 0 To dt.Rows.Count - 1
                    dr = dt.NewRow
                    dr("ลำดับ") = dt.Rows(i).Item("ลำดับ")
                    dr("รหัสสินค้า") = dt.Rows(i).Item("รหัสสินค้า")
                    dr("ชื่อสินค้า") = dt.Rows(i).Item("ชื่อสินค้า")
                    dr("หน่วยขาย") = dt.Rows(i).Item("หน่วยขาย")
                    dr("ราคาที่1") = dt.Rows(i).Item("ราคาที่1")
                Next
                Me.gvDetail.DataSource = dt1
                'กำหนดค่าของ Grid
                Me.gvDetail.Columns(0).ReadOnly = True
                'xxxxx
            End If

        End If

    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        ' ส่งค่าเอกสารใหม่ Insert
        Dim i As Integer
        Dim vi As Integer
        Dim idocDate As String
        Dim istartDate As String
        Dim iendDate As String
        Dim ierr As Integer

        'Open Transection
        QryString = "begin tran"
        cmd = New SqlCommand
        With cmd
            .CommandType = CommandType.Text
            .CommandText = QryString
            .Connection = vConnection
            .ExecuteNonQuery()
        End With
        On Error GoTo ErrConfirm
        vDocNo = Me.txtDocno.Text
        vDocdate = Me.DTPdocDate.Value
        vStartDate = Me.DTPStartDate.Value
        vEndDate = Me.DTPEndDate.Value
        idocDate = vDocdate.Day & "/" & vDocdate.Month & "/" & vDocdate.Year  'Format(vDocdate, "dd/mm/yyyy")
        istartDate = vStartDate.Day & "/" & vStartDate.Month & "/" & vStartDate.Year 'Format(vStartDate, "dd/mm/yyyy")
        iendDate = vEndDate.Day & "/" & vEndDate.Month & "/" & vEndDate.Year 'Format(vEndDate, "dd/mm/yyyy")
        'QryString = "exec dbo.USP_PS_PriceVolumeSet '" & sv & "','" & vDocNo & "','" & idocDate & "','" & istartDate & "','" & iendDate & "'"
        'cmd = New SqlCommand
        'With cmd
        '    .CommandType = CommandType.Text
        '    .CommandText = QryString
        '    .Connection = vConnection
        '    .ExecuteNonQuery()

        'End With
        '-----------------------
        QryString = "exec dbo.USP_PS_PriceVolumeSet '" & sv & "','" & vDocNo & "','" & idocDate & "','" & istartDate & "','" & iendDate & "'"
        da = New SqlDataAdapter(QryString, vConnection)
        ds = New DataSet
        da.Fill(ds, "ChkInsert")
        dt = ds.Tables("ChkInsert")
        If dt.Rows.Count > 0 Then
            ierr = dt.Rows(0).Item("IsError")
            If ierr = 1 Then
                On Error GoTo ErrConfirmChkInsert
            Else
                '-----------------------
                'MsgBox(Err.Description, MsgBoxStyle.Critical, "Error")
                If Me.gvDetail.Rows.Count <> 0 Then
                    For i = 0 To Me.gvDetail.Rows.Count - 1
                        vDocNo = Me.txtDocno.Text
                        LineNumber = CInt(Me.gvDetail.Rows(i).Cells(0).Value)
                        vPSDocNo = Me.gvDetail.Rows(i).Cells(1).Value
                        ItemNo = Me.gvDetail.Rows(i).Cells(2).Value
                        ItemName1 = Me.gvDetail.Rows(i).Cells(3).Value
                        ItemUnit = Me.gvDetail.Rows(i).Cells(4).Value
                        'Format(Int(Me.TextCash01.Text), "##,##0.00")
                        vPriceLP = Me.gvDetail.Rows(i).Cells(5).Value
                        vPrice1 = Me.gvDetail.Rows(i).Cells(6).Value
                        vPercentSMP1 = Me.gvDetail.Rows(i).Cells(7).Value
                        vSmartPoint1 = Me.gvDetail.Rows(i).Cells(8).Value
                        vPriceLevel = Me.gvDetail.Rows(i).Cells(9).Value
                        vVolume = Me.gvDetail.Rows(i).Cells(10).Value
                        vDCPrice1 = Me.gvDetail.Rows(i).Cells(11).Value
                        vPrice = Me.gvDetail.Rows(i).Cells(12).Value
                        vPercentSMP = Me.gvDetail.Rows(i).Cells(13).Value
                        vSmartPoint = Me.gvDetail.Rows(i).Cells(14).Value
                        vMKcost = Me.gvDetail.Rows(i).Cells(15).Value
                        vGPMKcost = Me.gvDetail.Rows(i).Cells(16).Value
                        vAvgCostlot = Me.gvDetail.Rows(i).Cells(18).Value
                        vGPAvgCostlot = Me.gvDetail.Rows(i).Cells(19).Value

                        If vPrice1 = 0 Then
                            vi = MsgBox("คุณไม่ได้กำหนดราคา1 ของสินค้ารหัส :" & ItemNo & " : " & ItemName1 & " หรือมีรายการสินค้าซ้ำ", MsgBoxStyle.Critical, "Error")
                            If vi = 1 Then
                                GoTo ErrPricr1
                            End If
                        End If
                        '1. @DocNo as varchar(25)'2. @Linenumber as smallint '3. ,@Itemcode as varchar(25)'4. ,@Itemname as varchar(200)'5. ,@Unitcode as varchar(15)'6. ,@Price1 as money '8. ,@Smartpoint1Percent as money '9. ,@Smartpoint1 as money'10. ,@PriceLevel as smallint'11. ,@Volume as money'12. ,@Price1Discount as money'13. ,@PriceSet as money'14. ,@SmartpointPercent money'15. ,@Smartpoint as money'16. ,@MarketCost as money '17. ,@marketcostGP as money '18. ,@LotAverageCost as money '19. ,@LotAverageCostGP as money
                        'exec dbo.USP_PS_PriceVolumeSetSub               'PSV5202-00006',1,'2104220','บานซิ้งค์คู่ Freedom สีขาวไอวอรี่ ตราเพชร','บาน',1,090.00,1.00,16.00,2.00,11,8.00,1,002.00,0.00,5.00,995.00,7.00,930.00,72.00

                        QryString = "exec dbo.USP_PS_PriceVolumeSetSub '" & vDocNo & "'," & LineNumber & ",'" & ItemNo & "','" & ItemName1 & "','" & ItemUnit & "'," & vPrice1 & "," & vPercentSMP1 & "," & vSmartPoint1 & "," & vPriceLevel & "," & vVolume & "," & vDCPrice1 & "," & vPrice & "," & vPercentSMP & "," & vSmartPoint & "," & vMKcost & "," & vGPMKcost & "," & vAvgCostlot & "," & vGPAvgCostlot & ",'" & vPSDocNo & "'," & vPriceLP & ""

                        'QryString += ",'" & vPrice1 & "','" & vPercentSMP1 & "','" & vSmartPoint1 & "','" & vPriceLevel & "','" & vVolume & "','" & vDCPrice1 & "'"
                        'QryString += ",'" & vPrice & "','" & vPercentSMP & "','" & vSmartPoint & "','" & vMKcost & "','" & vGPMKcost & "','" & vAvgCostlot & "','" & vGPAvgCostlot & "'"
                        cmd = New SqlCommand
                        With cmd
                            .CommandType = CommandType.Text
                            .CommandText = QryString
                            .Connection = vConnection
                            .ExecuteNonQuery()
                        End With

                    Next i

                    '-----------------------------------------------------------------
                    QryString = "exec dbo.USP_PS_PriceVolumeSetConfirm '" & vDocNo & "','" & 1 & "'"
                    With cmd
                        .CommandType = CommandType.Text
                        .CommandText = QryString
                        .Connection = vConnection
                        .ExecuteNonQuery()
                    End With
                    QryString = "commit tran"
                    With cmd
                        .CommandType = CommandType.Text
                        .CommandText = QryString
                        .Connection = vConnection
                        .ExecuteNonQuery()
                    End With
                    MsgBox("บันทึกข้อมูลเอกสารเลขที่ " & vDocNo & "เรียบร้อยแล้ว", MsgBoxStyle.Critical, "Information")
                    Dim xmsbx As String
                    xmsbx = MsgBox("คุณต้องการเคลียร์ฟอร์มหรือไม่", MsgBoxStyle.YesNo, "เลือกเคลียร์ข้อมูล")
                    If xmsbx = 6 Then
                        Call clearfrm()
                    End If
                    Me.btnSaveAS.Enabled = False
                End If
            End If
        End If
                Me.btnPrint.Enabled = True
                '-----------------------------------------------------------------
ErrPricr1:
                If vi = 1 Then
                    QryString = " rollback tran"
                    With cmd
                        .CommandType = CommandType.Text
                        .CommandText = QryString
                        .Connection = vConnection
                        .ExecuteNonQuery()
                    End With
                    MsgBox("ไม่สามารถบันทึกข้อมูลเอกสารเลขที่  :" & vDocNo & "", MsgBoxStyle.Critical, "Error")
                    'Me.txtDocno.fo()
                    'Me.gvDetail.CurrentRow.Cells(5).Selected

        End If

ErrConfirm:
                If Err.Description <> "" Then
                    QryString = " rollback tran"
                    With cmd
                        .CommandType = CommandType.Text
                        .CommandText = QryString
                        .Connection = vConnection
                        .ExecuteNonQuery()
                    End With
                    MsgBox("ไม่สามารถบันทึกข้อมูลเอกสารเลขที่  :" & vDocNo & "หรือมีรายการสินค้ามากกว่า 2 รายการ", MsgBoxStyle.Critical, "Error")
        End If
ErrConfirmChkInsert:
        If ierr = 1 Then
            QryString = " rollback tran"
            With cmd
                .CommandType = CommandType.Text
                .CommandText = QryString
                .Connection = vConnection
                .ExecuteNonQuery()
            End With
            MsgBox("ไม่สามารถบันทึกข้อมูลเอกสารเลขที่  :" & vDocNo & "เนื่องจากมีเลขที่เอกสารซ้ำ กรุณาสร้างเลขที่เอกสารใหม่", MsgBoxStyle.Critical, "Error")
            pNewdoc = 0
        End If
    End Sub

    Private Sub btnApprove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MsgBox("คุณยังไม่ได้เลือกรายการอนุมัติ", MsgBoxStyle.Critical, "Error")
        MsgBox("คุณได้ตรวจสอบข้อมูลเรียบร้อยแล้วและต้องการอนุมัติรายการนี้.", MsgBoxStyle.OkCancel, "Confirm Information")
    End Sub

    Private Sub btnNewDoc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNewDoc.Click
        If pNewdoc = 1 Then


            Me.PBNew.Visible = True
            Me.PBConfirm.Visible = False
            Me.btnPrint.Enabled = False
            Me.btnSave.Enabled = False
            Me.btnSearch.Enabled = True
            Me.txtLv2.Text = "2"
            Me.txtLv3.Text = "3"
            Me.txtSMTP2.Text = "0.5"
            Me.txtSMTP3.Text = "0.25"
            Me.txtVM2.Text = ""
            Me.txtVM3.Text = ""
            Me.txtDC2.Text = ""
            Me.txtDC3.Text = ""
            Me.txtVM2.Enabled = True
            Me.txtVM3.Enabled = True
            Me.txtDC2.Enabled = True
            Me.txtDC3.Enabled = True
            Me.txtVM2.ReadOnly = False
            Me.txtVM3.ReadOnly = False
            Me.txtDC2.ReadOnly = False
            Me.txtDC3.ReadOnly = False
            Me.btnProduct.Enabled = True
            'Me.txtVM2.BackColor = Color.LightSkyBlue
            ''Me.txtVM2.ForeColor = Color.White
            'Me.txtVM3.BackColor = Color.LightSkyBlue
            '' Me.txtVM3.ForeColor = Color.LightSkyBlue 
            'Me.txtDC2.BackColor = Color.LightSkyBlue
            '' Me.txtDC2.ForeColor = Color.White
            'Me.txtDC3.BackColor = Color.LightSkyBlue
            ''Me.txtDC3.ForeColor = Color.White
            Me.gvDetail.DataSource = Nothing
            '-----------------------
            Call GenNewdoc()
            Me.txtVM2.Focus()
        Else
            Call GenNewdoc()
        End If

    End Sub
    Private Sub GenNewdoc()
        Dim StrDate As String
        StrDate = Now.Day & "/" & Now.Month & "/" & Now.Year
        QryString = "exec USP_PS_PriceVolumeSetNewdoc " ''" & StrDate & "'
        da = New SqlDataAdapter(QryString, vConnectionString)
        ds = New DataSet
        da.Fill(ds, "docno")
        dt = ds.Tables("docno")
        If dt.Rows.Count <> 0 Then
            vNewdocNo = dt.Rows(0).Item("NewDocNo")
            Me.txtDocno.Text = vNewdocNo
        End If
    End Sub
    Private Sub clearfrm()
        Me.PBNew.Visible = True
        Me.PBConfirm.Visible = False
        Me.btnPrint.Enabled = False
        Me.btnSave.Enabled = False
        Me.btnSearch.Enabled = True
        Me.txtDocno.Text = ""
        Me.txtSMTP2.Text = "0.5"
        Me.txtSMTP3.Text = "0.25"
        Me.txtLv2.Text = "2"
        Me.txtLv3.Text = "3"
        Me.txtVM2.Text = ""
        Me.txtVM3.Text = ""
        Me.txtDC2.Text = ""
        Me.txtDC3.Text = ""
        Me.txtLv2.BackColor = Color.LightSkyBlue
        Me.txtLv3.BackColor = Color.LightSkyBlue
        Me.txtVM2.Enabled = True
        Me.txtVM3.Enabled = True
        Me.txtDC2.Enabled = True
        Me.txtDC3.Enabled = True
        Me.txtVM2.ReadOnly = False
        Me.txtVM3.ReadOnly = False
        Me.txtDC2.ReadOnly = False
        Me.txtDC3.ReadOnly = False
        Me.GroupBox1.Enabled = True
        Me.BtnGenVLM.Visible = False
        Me.btnProduct.Visible = True
        Me.btnProduct.Enabled = True
        Me.okFind.Visible = False
        Me.btnNewDoc.Visible = True
        Me.PBNew.Visible = False
        Me.PBConfirm.Visible = False
        Me.gvDetail.DataSource = Nothing
        Me.txtDocno.Focus()
    End Sub
    Private Sub GetDepartment()
        Dim daDP As New SqlDataAdapter("exec USP_PS_DepartmentList", vConnectionString)
        Dim dsDP As New DataSet
        Dim dtDP As DataTable
        daDP.Fill(dsDP, "Dpm")
        dtDP = dsDP.Tables("Dpm")
        Me.cbxDepartment.DataSource = dsDP.Tables("Dpm")
        Me.cbxDepartment.DisplayMember = ("Department")
        Me.cbxDepartment.ValueMember = ("Departmentcode")
        ' DeptCode = Me.cbxDepartment.ValueMember.ToString()

    End Sub
    Private Sub GetBrand()
        Dim daBND As New SqlDataAdapter("exec USP_PS_BrandList", vConnectionString)
        Dim dsBND As New DataSet
        Dim dtBND As New DataTable
        daBND.Fill(dsBND, "bnd")
        dtBND = dsBND.Tables("bnd")
        Me.cbxBrand.DataSource = dsBND.Tables("bnd")
        Me.cbxBrand.DisplayMember = ("Brand")
        Me.cbxBrand.ValueMember = ("BrandCode")
        'BrandCode = dtBND.Rows(0).Item("BrandCode")
    End Sub
    Private Sub GetItemType()
        Dim daITMTY As New SqlDataAdapter("exec USP_PS_ItemTypeList", vConnectionString)
        Dim dsITMTY As New DataSet
        Dim dtITMTY As New DataTable
        daITMTY.Fill(dsITMTY, "itmtype")
        dtITMTY = dsITMTY.Tables("itmtype")
        Me.cbxProductType.DataSource = dsITMTY.Tables("itmtype")
        Me.cbxProductType.DisplayMember = ("Typename")
        Me.cbxProductType.ValueMember = ("TypeCode")
        'TypeCode = dtITMTY.Rows(0).Item("TypeCode")
    End Sub
    Private Sub GetCategories()
        Dim daCATE As New SqlDataAdapter("exec USP_PS_CategoryList '" & DeptCode & "'", vConnectionString)
        Dim dsCATE As New DataSet
        Dim dtCATE As New DataTable
        daCATE.Fill(dsCATE, "cate")
        dtCATE = dsCATE.Tables("cate")
        Me.cbxCategory.DataSource = dsCATE.Tables("cate")
        Me.cbxCategory.DisplayMember = ("category")
        Me.cbxCategory.ValueMember = ("categorycode")
        'CateCode = dtCATE.Rows(0).Item("categorycode")
    End Sub

    Private Sub cbxDepartment_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbxDepartment.SelectedIndexChanged
        Me.cbxCategory.Enabled = True
        If Me.cbxDepartment.ValueMember <> "" Then
            DeptCode = Me.cbxDepartment.SelectedValue.ToString()
            Call GetCategories()
        End If
    End Sub

    Private Sub btnGenerate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGenerate.Click
        Dim iListView As ListViewItem
        Dim pr1 As String
        Dim mkcs As String
        Dim avmkcs As String
        Me.pgbItem.Minimum = 0
        'If Me.cbxCategory.ValueMember <> "" Then
        If Me.cbxCategory.SelectedIndex > 0 Then
            CateCode = Me.cbxCategory.SelectedValue.ToString()
        Else
            CateCode = ""
        End If

        BrandCode = Me.cbxBrand.SelectedValue.ToString()
        TypeCode = Me.cbxProductType.SelectedValue.ToString()

        iQryString = "exec USP_PS_PriceVolumeSetPrepare '" & DeptCode & "','" & CateCode & "','" & BrandCode & "','" & TypeCode & "','" & "'"
        da = New SqlDataAdapter(iQryString, vConnectionString)
        ds = New DataSet
        da.Fill(ds, "ivw")
        dt = ds.Tables("ivw")
        Me.pgbItem.Maximum = dt.Rows.Count - 1
        If dt.Rows.Count > 0 Then
            Me.LvProduct.Items.Clear()
            For n = 0 To dt.Rows.Count - 1
                iListView = Me.LvProduct.Items.Add(dt.Rows(n).Item("PSDocNo"))
                iListView.SubItems.Add(0).Text = dt.Rows(n).Item("itemcode")
                iListView.SubItems.Add(0).Text = dt.Rows(n).Item("itemname")
                iListView.SubItems.Add(0).Text = dt.Rows(n).Item("SaleUnitCode")
                pr1 = dt.Rows(n).Item("Price1").ToString()
                mkcs = dt.Rows(n).Item("MarketCostSaleVat").ToString()
                avmkcs = dt.Rows(n).Item("AveragecostlotSaleVat").ToString()
                If pr1 <> "" Then
                    iListView.SubItems.Add(0).Text = dt.Rows(n).Item("Price1").ToString()
                Else
                    iListView.SubItems.Add(0).Text = 0
                End If
                If mkcs <> "" Then
                    iListView.SubItems.Add(0).Text = dt.Rows(n).Item("MarketCostSaleVat").ToString()
                Else
                    iListView.SubItems.Add(0).Text = 0
                End If
                If avmkcs <> "" Then
                    iListView.SubItems.Add(0).Text = dt.Rows(n).Item("AveragecostlotSaleVat").ToString()
                Else
                    iListView.SubItems.Add(0).Text = 0
                End If
                Me.pgbItem.Value = n
            Next
            'End If
        End If

    End Sub

    'Private Sub vPGB()
    '    Dim px As Integer
    '    Me.pgbItem.Minimum = 0
    '    Me.pgbItem.Maximum = Me.LvProduct.Items.Count - 1
    '    If Me.LvProduct.Items.Count > 0 Then
    '        For px = 0 To Me.LvProduct.Items.Count - 1
    '            Me.pgbItem.Value = px
    '        Next
    '    End If
    'End Sub
    Private Sub btnSelect_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSelect.Click
        'for GridView
        Dim dtbl As New DataTable("Ptbl")
        Dim dr As DataRow
        Dim x As Integer
        Dim nml As Integer = 0
        'vLV2
        Dim nPrice2 As Double
        Dim mkvalue2 As Double
        Dim avMKlot2 As Double
        'vLV3
        Dim nPrice3 As Double
        Dim mkvalue3 As Double
        Dim avMKlot3 As Double
        'Dim gpAVmklot As Double
        Dim vPriceLV1 As Double
        Dim vSLPlv1 As Double
        dtbl.Columns.Add("No.", GetType(String))
        dtbl.Columns.Add("เลขที่เอกสารโครงสร้างราคา", GetType(String))
        dtbl.Columns.Add("รหัสสินค้า", GetType(String))
        dtbl.Columns.Add("ชื่อสินค้า", GetType(String))
        dtbl.Columns.Add("หน่วยขาย", GetType(String))
        dtbl.Columns.Add("ราคาที่1", GetType(String))
        dtbl.Columns.Add("%smartpoint ราคาที่1", GetType(String))
        dtbl.Columns.Add("SmartPoint1", GetType(String))
        dtbl.Columns.Add("ระดับราคา", GetType(String))
        dtbl.Columns.Add("Volume", GetType(String))
        dtbl.Columns.Add("%ส่วนลดจากราคาที่ 1", GetType(String))
        dtbl.Columns.Add("ราคา", GetType(String))
        dtbl.Columns.Add("%SmartPoint", GetType(String))
        dtbl.Columns.Add("SmartPoint", GetType(String))
        dtbl.Columns.Add("ทุนตลาดSaleVat", GetType(String))
        dtbl.Columns.Add("GP ทุนตลาด", GetType(String))
        dtbl.Columns.Add("ทุนเฉลี่ย LotSaleVat", GetType(String))
        dtbl.Columns.Add("GP ทุนเฉลี่ย Lot", GetType(String))

        For x = 0 To Me.LvProduct.Items.Count - 1
            If Me.LvProduct.Items(x).Checked = True Then
                '----------------------Price Level 2
                dr = dtbl.NewRow
                nml = nml + 1
                dr("No.") = nml
                dr("เลขที่เอกสารโครงสร้างราคา") = Me.LvProduct.Items(x).SubItems(0)
                dr("รหัสสินค้า") = Me.LvProduct.Items(x).SubItems(1).Text
                dr("ชื่อสินค้า") = Me.LvProduct.Items(x).SubItems(2).Text
                dr("หน่วยขาย") = Me.LvProduct.Items(x).SubItems(3).Text
                vPriceLV1 = CDbl(Me.LvProduct.Items(x).SubItems(4).Text)
                dr("ราคาที่1") = Format(Int(vPriceLV1), "##,##0.00")
                vSLPlv1 = Me.smpLV1.Value
                dr("%SmartPoint ราคาที่1") = Format(Int(vSLPlv1), "##,##0.00")
                dr("SmartPoint1") = Format(Int((vPriceLV1 * vSLPlv1) / 100), "##,##0.00")
                dr("ระดับราคา") = Format(Int(Me.txtLv2.Text), "##,##0.00")
                dr("Volume") = Format(Int(Me.txtVM2.Text), "##,##0.00")
                dr("%ส่วนลดจากราคาที่ 1") = Format(Int(Me.txtDC2.Text), "##,##0.00")
                nPrice2 = CDbl((Me.LvProduct.Items(x).SubItems(4).Text - ((Me.LvProduct.Items(x).SubItems(4).Text * Me.txtDC2.Text) / 100)))
                dr("ราคา") = Format(Int(nPrice2), "##,##0.00")
                dr("%SmartPoint") = Format(Int(Me.txtSMTP2.Text), "##,##0.00")
                dr("SmartPoint") = Format(Int(CDbl(((Me.LvProduct.Items(x).SubItems(4).Text * Me.txtSMTP2.Text) / 100))), "##,##0.0000")
                mkvalue2 = CDbl(Me.LvProduct.Items(x).SubItems(4).Text)
                dr("ทุนตลาดSaleVat") = Format(Int(mkvalue2), "##,##0.00")
                dr("GP ทุนตลาด") = Format(Int((nPrice2 - mkvalue2)), "##,##0.00")
                avMKlot2 = CDbl(Me.LvProduct.Items(x).SubItems(6).Text)
                dr("ทุนเฉลี่ย LotSaleVat") = Format(Int(avMKlot2), "##,##0.00")
                ' gpAVmklot = (Me.gvDetail.Columns(11).ToString - Me.gvDetail.Columns(16).ToString)
                dr("GP ทุนเฉลี่ย Lot") = Format(Int(nPrice2 - avMKlot2), "##,##0.00")
                dtbl.Rows.Add(dr)
                '----------------------Price Level 3
                dr = dtbl.NewRow
                nml = nml + 1
                dr("No.") = nml
                dr("เลขที่เอกสารโครงสร้างราคา") = Me.LvProduct.Items(x).SubItems(0)
                dr("รหัสสินค้า") = Me.LvProduct.Items(x).SubItems(1).Text
                dr("ชื่อสินค้า") = Me.LvProduct.Items(x).SubItems(2).Text
                dr("หน่วยขาย") = Me.LvProduct.Items(x).SubItems(3).Text
                vPriceLV1 = CDbl(Me.LvProduct.Items(x).SubItems(4).Text)
                dr("ราคาที่1") = Format(Int(vPriceLV1), "##,##0.00")
                vSLPlv1 = Me.smpLV1.Value
                dr("%SmartPoint ราคาที่1") = Format(Int(vSLPlv1), "##,##0.00")
                dr("SmartPoint1") = Format(Int((vPriceLV1 * vSLPlv1) / 100), "##,##0.00")
                dr("ระดับราคา") = Format(Int(Me.txtLv3.Text), "##,##0.00")
                dr("Volume") = Format(Int(Me.txtVM3.Text), "##,##0.00")
                dr("%ส่วนลดจากราคาที่ 1") = Format(Int(Me.txtDC3.Text), "##,##0.00")
                nPrice3 = Format(Int(CDbl((Me.LvProduct.Items(x).SubItems(4).Text - ((Me.LvProduct.Items(x).SubItems(4).Text * Me.txtDC3.Text) / 100)))), "##,##0.00")
                dr("ราคา") = Format(Int(nPrice3), "##,##0.00")
                dr("%SmartPoint") = Format(Int(Me.txtSMTP3.Text), "##,##0.00")
                dr("SmartPoint") = Format(Int(CDbl(((Me.LvProduct.Items(x).SubItems(4).Text * Me.txtSMTP3.Text) / 100))), "##,##0.0000")
                mkvalue3 = CDbl(Me.LvProduct.Items(x).SubItems(5).Text)
                dr("ทุนตลาดSaleVat") = Format(Int(mkvalue3), "##,##0.00")
                dr("GP ทุนตลาด") = Format(Int((nPrice3 - mkvalue3)), "##,##0.00")
                avMKlot3 = CDbl(Me.LvProduct.Items(x).SubItems(6).Text)
                dr("ทุนเฉลี่ย LotSaleVat") = Format(Int(avMKlot3), "##,##0.00")
                dr("GP ทุนเฉลี่ย Lot") = Format(Int(nPrice3 - avMKlot3), "##,##0.00")
                dtbl.Rows.Add(dr)
                Me.gvDetail.BackgroundColor = Color.Cyan
            End If
        Next
        Me.gvDetail.DataSource = dtbl
        Call gvdetailFormat()
        Me.P01.Visible = True
        Call gvReadonly()
        sv = "1"
        Me.Text = "[::] กำหนดราคาตามจำนวน"
        Me.btnPrint.Enabled = False
        Me.btnSave.Enabled = True
        Me.btnSearch.Enabled = False
    End Sub
    Public Sub gvdetailFormat()
        Me.gvDetail.Columns(0).Width = 30
        Me.gvDetail.Columns(1).Width = 75
        Me.gvDetail.Columns(2).Width = 75
        Me.gvDetail.Columns(3).Width = 200
        Me.gvDetail.Columns(4).Width = 50
        Me.gvDetail.Columns(5).Width = 70
        Me.gvDetail.Columns(6).Width = 70
        Me.gvDetail.Columns(7).Width = 50
        Me.gvDetail.Columns(8).Width = 25
        Me.gvDetail.Columns(9).Width = 40
        Me.gvDetail.Columns(10).Width = 30
        Me.gvDetail.Columns(11).Width = 80
        Me.gvDetail.Columns(12).Width = 40
        Me.gvDetail.Columns(13).Width = 70
        Me.gvDetail.Columns(14).Width = 70
        Me.gvDetail.Columns(15).Width = 70
        Me.gvDetail.Columns(16).Width = 70
        Me.gvDetail.Columns(17).Width = 70
        Me.gvDetail.Columns(18).Width = 70
        Me.gvDetail.Columns(19).Width = 70
        Me.gvDetail.Columns(20).Width = 70
        Me.gvDetail.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.gvDetail.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.gvDetail.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.gvDetail.Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.gvDetail.Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.gvDetail.Columns(9).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.gvDetail.Columns(10).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.gvDetail.Columns(11).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.gvDetail.Columns(12).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.gvDetail.Columns(13).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.gvDetail.Columns(14).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.gvDetail.Columns(15).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.gvDetail.Columns(16).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.gvDetail.Columns(17).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.gvDetail.Columns(18).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.gvDetail.Columns(19).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.gvDetail.Columns(20).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight


    End Sub
    Private Sub gvReadonly()
        Me.gvDetail.Columns(0).ReadOnly = True
        Me.gvDetail.Columns(1).ReadOnly = True
        Me.gvDetail.Columns(2).ReadOnly = True
        Me.gvDetail.Columns(3).ReadOnly = True
        'Me.gvDetail.Columns(4).ToolTipText = "คลิกเพื่อแก้ไขราคารายการนี้"
        'Me.gvDetail.Columns(4).DefaultCellStyle.ForeColor = Color.Blue
        Me.gvDetail.Columns(4).ReadOnly = True
        Me.gvDetail.Columns(5).ReadOnly = True
        Me.gvDetail.Columns(6).ReadOnly = True
        Me.gvDetail.Columns(7).DefaultCellStyle.ForeColor = Color.Blue
        Me.gvDetail.Columns(7).ReadOnly = False
        Me.gvDetail.Columns(7).ToolTipText = "คลิกเพื่อแก้ไข Volume รายการนี้"
        Me.gvDetail.Columns(8).ReadOnly = True
        Me.gvDetail.Columns(9).ReadOnly = True
        Me.gvDetail.Columns(10).ToolTipText = "คลิกเพื่อแก้ไขส่วนลดรายการนี้"
        Me.gvDetail.Columns(10).DefaultCellStyle.ForeColor = Color.Blue
        Me.gvDetail.Columns(11).ReadOnly = False
        Me.gvDetail.Columns(11).ToolTipText = "คลิกที่ช่องที่ต้องการแก้ไขราคา"
        Me.gvDetail.Columns(11).DefaultCellStyle.ForeColor = Color.Blue
        ' Me.gvDetail.Columns(12).ToolTipText = "คลิกเพื่อแก้ไข SmartPoint รายการนี้"
        'Me.gvDetail.Columns(12).DefaultCellStyle.ForeColor = Color.Blue
        Me.gvDetail.Columns(12).ReadOnly = False
        Me.gvDetail.Columns(12).DefaultCellStyle.ForeColor = Color.Blue
        Me.gvDetail.Columns(13).ReadOnly = False
        Me.gvDetail.Columns(13).DefaultCellStyle.ForeColor = Color.Blue
        Me.gvDetail.Columns(14).ReadOnly = True
        Me.gvDetail.Columns(15).ReadOnly = True
        Me.gvDetail.Columns(16).ReadOnly = True
        Me.gvDetail.Columns(17).ReadOnly = True
        Me.gvDetail.Columns(18).ReadOnly = True
        Me.gvDetail.Columns(19).ReadOnly = True
        Me.gvDetail.Columns(20).ReadOnly = True

    End Sub

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        publicFdocno = Me.txtDocno.Text
        frmPrintVolumeSet.Show()
    End Sub

    Private Sub btnSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearch.Click
        sv = 0
        Call clearfrm()
        dlgVolumeSearch.Show()
        Me.LPcbx.Enabled = False
        dlgVolumeSearch.TxtFind.Focus()
        'Me.btnNewDoc.Visible = False
        'Me.okFind.Visible = True
        'Me.okFind.Enabled = True
        'Me.txtDocno.ReadOnly = False
        ''Me.txtDocno.Focus()
        'Me.btnSaveAS.Enabled = False


     
    End Sub

    Private Sub okFind_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles okFind.Click
        publicFdocno = Me.txtDocno.Text
        'Me.Enabled = False
        dlgVolumeSearch.Show()
        
    End Sub

    Private Sub ckeckedAll_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ckeckedAll.CheckedChanged
        Dim xi As Integer
        If Me.LvProduct.Items.Count > 0 Then
            If Me.ckeckedAll.Checked = True Then
                For xi = 0 To Me.LvProduct.Items.Count - 1
                    Me.LvProduct.Items(xi).Checked = True
                Next
            Else
                For xi = 0 To Me.LvProduct.Items.Count - 1
                    Me.LvProduct.Items(xi).Checked = False
                Next
            End If

        End If
    End Sub
   
    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Call clearfrm()
        Me.DTPdocDate.Enabled = True
        Me.DTPStartDate.Enabled = True
        Me.DTPEndDate.Enabled = True
        Me.DTPStartDate.Text = DateAdd(DateInterval.Day, 1, Now.Date)
        Me.DTPEndDate.Text = DateAdd(DateInterval.Year, 5, Now.Date)
        Me.smpLV1.Enabled = True
        Me.LPcbx.Enabled = True
        'Me.LPcbx.Checked = True

    End Sub

    Private Sub btnexitP2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnexitP2.Click
        Me.P02.Visible = False
        Me.P01.Visible = True
    End Sub

    Private Sub txtDocno_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtDocno.KeyDown
        If e.KeyCode = Keys.Enter Then
            publicFdocno = Me.txtDocno.Text
            Me.Enabled = False
            dlgVolumeSearch.Show()
        End If
    End Sub

   
    Private Sub RdoPSDoc_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RdoPSDoc.CheckedChanged
        Me.GB01.Enabled = False
        Me.btnGenerate.Enabled = False
        dlgPSVdocSearch.Show()
    End Sub

    Private Sub RdoManual_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RdoManual.CheckedChanged
        Me.GB01.Enabled = True
        Me.GB01.Visible = True
        Me.btnGenerate.Enabled = True
    End Sub

    Private Function FNnewvDCM(ByVal vi As Double) As Double
        Dim a As Integer
        Dim i As Integer
        Dim vL As String
        Dim vc As Double
        Dim vx As String
        'Format(vx, "##.00")
        vx = CStr(vi)
        a = Microsoft.VisualBasic.Len(vx)
        i = InStr(vx, ".")
        If i = 0 Then
            i = 1
            vL = vx
        Else
            i = i
            vL = Microsoft.VisualBasic.Left(vx, i - 1)
        End If
        'vR = Microsoft.VisualBasic.Right(vx, a - i)
        vc = (vx - vL)
        If (vc >= 0.5 And vc < 1.0) Then
            vc = 1
            vx = (vL + vc) & "." & "00"

        ElseIf (vc > 0.0 And vc < 0.5) Then
            'vc = 0.5
            vx = vL & ".5"
        End If
        vi = CDbl(vx)
        Return vi
    End Function
    Private Function CHKVMDCK(ByVal v1 As Double, ByVal v2 As Double) As Double
        If v1 < v2 Then
            v1 = 0
        End If
        Return v1
    End Function

    Private Sub gvDetail_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles gvDetail.CellEndEdit
        Dim vNewPriceKey As Double
        'New var----------------------------
        Dim icl5 As Double = 0
        Dim icl6 As Double
        Dim icl7 As Double
        Dim icl8 As Double
        Dim icl10 As Double
        Dim icl11 As Double
        Dim icl12 As Double
        Dim icl13 As Double
        Dim icl14 As Double
        Dim icl16 As Double
        Dim icl19 As Double
        Dim vi1 As Integer = 0
        '------------------------
        Dim vivmx As Double
        Dim vivmn As Double
        Dim vivm3 As Double
        Dim vidc2 As Double
        Dim vidc3 As Double
        Dim vichklv As Integer
        Dim vPriceLP As Double

        '------------------------
        ' เก็บค่าจาก dg
        icl6 = Me.gvDetail.Item(6, gvDetail.CurrentRow.Index).Value 'คอลัม ราคาที่ 1
        icl7 = Me.gvDetail.Item(7, gvDetail.CurrentRow.Index).Value ' %SmartPoint1
        icl8 = Me.gvDetail.Item(8, gvDetail.CurrentRow.Index).Value ' SmartPoint 1
        icl10 = Me.gvDetail.Item(10, gvDetail.CurrentRow.Index).Value ' Volume
        icl12 = Me.gvDetail.Item(12, gvDetail.CurrentRow.Index).Value ' ราคา
        icl13 = Me.gvDetail.Item(13, gvDetail.CurrentRow.Index).Value ' %SmartPoint
        icl14 = Me.gvDetail.Item(14, gvDetail.CurrentRow.Index).Value ' SmartPoint
        icl16 = Me.gvDetail.Item(16, gvDetail.CurrentRow.Index).Value ' GP ทุนตลาด
        icl19 = Me.gvDetail.Item(19, gvDetail.CurrentRow.Index).Value ' GP ทุนเฉลี่ยตาม lot

        If Me.gvDetail.Rows.Count > 0 Then
            ' Key แก้ไข %SmartPoint ราคาที่ 1
            If e.ColumnIndex = 7 Then
                vPriceLP = Me.gvDetail.Item(5, gvDetail.CurrentRow.Index).Value
                Dim irsSMP1 As Double
                If Me.LPcbx.Checked = True Then
                    '------- คำนวณจาก ราคา LP ----------
                    irsSMP1 = ((vPrice * icl7) / 100)
                    Me.gvDetail.Item(8, gvDetail.CurrentRow.Index).Value = Format(irsSMP1, "##,##0.00")
                Else
                    '------- คำนวณ จากราคา ที่ 1
                    irsSMP1 = ((icl6 * icl7) / 100)
                    Me.gvDetail.Item(8, gvDetail.CurrentRow.Index).Value = Format(irsSMP1, "##,##0.00")
                End If
            End If

                'Key แก้ไข Volume  --------------------------------
            If e.ColumnIndex = 10 Then
                vichklv = Me.gvDetail.Item(9, gvDetail.CurrentRow.Index).Value
                vivmx = Me.gvDetail.Item(10, gvDetail.CurrentRow.Index).Value
                If Me.gvDetail.CurrentRow.Index > 0 Then
                    vivmn = Me.gvDetail.Item(10, gvDetail.CurrentRow.Index - 1).Value
                End If
                vivm3 = Me.gvDetail.Item(10, gvDetail.CurrentRow.Index + 1).Value
                If vichklv = 2 And (vivmx < 1 Or vivmx > vivm3) Then
                    MsgBox("ค่า Volume ต้องมากกว่า 1 หรือ ต้องน้อยกว่า Volume ระดับ 3", MsgBoxStyle.Critical, "Warnings")
                ElseIf vichklv = 3 And vivmx < vivmn Then
                    MsgBox("ค่าของ Volume ระดับราคาที่ 3 ต้องมากกว่า Volume ระดับราคาที่ 2", MsgBoxStyle.Critical, "Warnings")
                End If
            End If


            'Key แก้ไขส่วนลด --------------------------
            If e.ColumnIndex = 11 Then
                Dim rsx1 As Integer
                Dim rsx2 As Integer
                icl11 = Me.gvDetail.Item(11, gvDetail.CurrentRow.Index).Value ' ส่วนลดจากราคาที่1
                '----- ตรวจสอบค่าการป้อนส่วนลด -----------
                vichklv = Me.gvDetail.Item(9, gvDetail.CurrentRow.Index).Value
                vidc3 = Me.gvDetail.Item(11, gvDetail.CurrentRow.Index + 1).Value
                If Me.gvDetail.CurrentRow.Index > 0 Then
                    vidc2 = Me.gvDetail.Item(11, gvDetail.CurrentRow.Index - 1).Value
                End If
                rsx1 = CHKVMDC(icl11, vidc3)
                rsx2 = CHKVMDCK(icl11, vidc2)
                If vichklv = 2 And (icl11 <= 0 Or rsx1 = 0) Then
                    MsgBox("ค่าส่วนลดระดับที่ 2 ต้องมากกว่า 0 และ ต้องน้อยกว่าส่วนลดระดับที่ 3", MsgBoxStyle.Critical, "Warnings")
                    'If Me.gvDetail.CurrentRow.Index > 1 And icl11 > 0 Then
                    '    
                ElseIf vichklv = 3 And (icl11 <= 0 Or rsx2 = 0) Then
                    MsgBox("ค่าส่วนลดระดับที่ 3 ต้องมากกว่าระดับที่ 2 และต้องมากกว่า 0", MsgBoxStyle.Critical, "Warnings")
                    'ราคา
                Else
                    Dim irsPrice01 As Double = 0

                    If Me.LPcbx.Checked = True Then
                        'คำนวณจากราคา LP ------------------
                        vPriceLP = Me.gvDetail.Item(5, gvDetail.CurrentRow.Index).Value
                        irsPrice01 = (vPriceLP - ((vPriceLP * icl11) / 100))
                        vNewPriceKey = FNnewvDCM(irsPrice01)
                        Me.gvDetail.Item(12, gvDetail.CurrentRow.Index).Value = Format(vNewPriceKey, "##,##0.00")
                        ' SmartPoint
                        'Dim irsSMP As Double
                        'irsSMP = ((vNewPriceKey * icl13) / 100)
                        'Me.gvDetail.Item(14, gvDetail.CurrentRow.Index).Value = Format(irsSMP, "##,##0.0000")
                        ''GP ทุนตลาด
                        'Dim irsGPMK As Double
                        'irsGPMK = ((Me.gvDetail.Item(12, gvDetail.CurrentRow.Index).Value - Me.gvDetail.Item(15, gvDetail.CurrentRow.Index).Value) - Me.gvDetail.Item(14, gvDetail.CurrentRow.Index).Value)
                        'Me.gvDetail.Item(16, gvDetail.CurrentRow.Index).Value = Format(irsGPMK, "##,##0.00")
                        ''% GP ทุนตลาด
                        'Dim irsPCgpMK As Double
                        'irsPCgpMK = ((Me.gvDetail.Item(16, gvDetail.CurrentRow.Index).Value * 100) / Me.gvDetail.Item(15, gvDetail.CurrentRow.Index).Value)
                        'Me.gvDetail.Item(17, gvDetail.CurrentRow.Index).Value = Format(irsPCgpMK, "##,##0.00")
                        ''Gp ทุนเฉลี่ย Lot
                        'Dim irsGPAVlot
                        'irsGPAVlot = ((Me.gvDetail.Item(12, gvDetail.CurrentRow.Index).Value - Me.gvDetail.Item(18, gvDetail.CurrentRow.Index).Value) - Me.gvDetail.Item(14, gvDetail.CurrentRow.Index).Value)
                        'Me.gvDetail.Item(17, gvDetail.CurrentRow.Index).Value = Format(irsGPAVlot, "##,##0.00")
                        ''% GP ทุนเฉลี่ย
                        'Dim irsPCgpAV As Double
                        'gpx1 = Me.gvDetail.Item(19, gvDetail.CurrentRow.Index).Value
                        'gpx2 = Me.gvDetail.Item(18, gvDetail.CurrentRow.Index).Value
                        'irsPCgpAV = ((gpx1 * 100) / gpx2)
                        'Me.gvDetail.Item(20, gvDetail.CurrentRow.Index).Value = Format(irsPCgpAV, "##,##0.00")
                        '
                        'irsPCgpAV = ((Me.gvDetail.Item(19, gvDetail.CurrentRow.Index).Value * 100) / Me.gvDetail.Item(18, gvDetail.CurrentRow.Index).Value)
                        'Me.gvDetail.Item(20, gvDetail.CurrentRow.Index).Value = Format(irsPCgpAV, "##,##0.00")
                    Else
                        'คำนวณจากราคาที่ 1 ------------------------------
                        irsPrice01 = (icl6 - ((icl6 * icl11) / 100))
                        vNewPriceKey = FNnewvDCM(irsPrice01)
                        Me.gvDetail.Item(12, gvDetail.CurrentRow.Index).Value = Format(vNewPriceKey, "##,##0.00")
                        ' SmartPoint
                        'Dim irsSMP As Double
                        'irsSMP = ((vNewPriceKey * icl13) / 100)
                        'Me.gvDetail.Item(14, gvDetail.CurrentRow.Index).Value = Format(irsSMP, "##,##0.0000")
                        ''GP ทุนตลาด
                        'Dim irsGPMK As Double
                        'irsGPMK = ((Me.gvDetail.Item(12, gvDetail.CurrentRow.Index).Value - Me.gvDetail.Item(15, gvDetail.CurrentRow.Index).Value) - Me.gvDetail.Item(14, gvDetail.CurrentRow.Index).Value)
                        'Me.gvDetail.Item(16, gvDetail.CurrentRow.Index).Value = Format(irsGPMK, "##,##0.00")
                        ''% GP ทุนตลาด
                        'Dim irsPCgpMK As Double
                        'irsPCgpMK = ((Me.gvDetail.Item(16, gvDetail.CurrentRow.Index).Value * 100) / Me.gvDetail.Item(15, gvDetail.CurrentRow.Index).Value)
                        'Me.gvDetail.Item(17, gvDetail.CurrentRow.Index).Value = Format(irsPCgpMK, "##,##0.00")
                        ''Gp ทุนเฉลี่ย Lot
                        'Dim irsGPAVlot
                        'irsGPAVlot = ((Me.gvDetail.Item(12, gvDetail.CurrentRow.Index).Value - Me.gvDetail.Item(18, gvDetail.CurrentRow.Index).Value) - Me.gvDetail.Item(14, gvDetail.CurrentRow.Index).Value)
                        'Me.gvDetail.Item(17, gvDetail.CurrentRow.Index).Value = Format(irsGPAVlot, "##,##0.00")
                        ''% GP ทุนเฉลี่ย
                        'Dim irsPCgpAV As Double
                        'gpx1 = Me.gvDetail.Item(19, gvDetail.CurrentRow.Index).Value
                        'gpx2 = Me.gvDetail.Item(18, gvDetail.CurrentRow.Index).Value
                        'irsPCgpAV = ((gpx1 * 100) / gpx2)
                        'Me.gvDetail.Item(20, gvDetail.CurrentRow.Index).Value = Format(irsPCgpAV, "##,##0.00")
                        ' End If

                        '   End If
                    End If
                    Dim irsSMP As Double
                    irsSMP = ((vNewPriceKey * icl13) / 100)
                    Me.gvDetail.Item(14, gvDetail.CurrentRow.Index).Value = Format(irsSMP, "##,##0.0000")
                    'GP ทุนตลาด
                    Dim irsGPMK As Double
                    irsGPMK = ((Me.gvDetail.Item(12, gvDetail.CurrentRow.Index).Value - Me.gvDetail.Item(15, gvDetail.CurrentRow.Index).Value) - Me.gvDetail.Item(14, gvDetail.CurrentRow.Index).Value)
                    Me.gvDetail.Item(16, gvDetail.CurrentRow.Index).Value = Format(irsGPMK, "##,##0.00")
                    '% GP ทุนตลาด 17
                    Dim irsPCgpMK As Double
                    irsPCgpMK = ((Me.gvDetail.Item(16, gvDetail.CurrentRow.Index).Value * 100) / Me.gvDetail.Item(15, gvDetail.CurrentRow.Index).Value)
                    Me.gvDetail.Item(17, gvDetail.CurrentRow.Index).Value = Format(irsPCgpMK, "##,##0.00")
                    'Gp ทุนเฉลี่ย Lot 19
                    Dim irsGPAVlot
                    irsGPAVlot = ((Me.gvDetail.Item(12, gvDetail.CurrentRow.Index).Value - Me.gvDetail.Item(18, gvDetail.CurrentRow.Index).Value) - Me.gvDetail.Item(14, gvDetail.CurrentRow.Index).Value)
                    Me.gvDetail.Item(19, gvDetail.CurrentRow.Index).Value = Format(irsGPAVlot, "##,##0.00")
                    '% GP ทุนเฉลี่ย 20
                    Dim irsPCgpAV As Double
                    Dim gpx1 As Double
                    Dim gpx2 As Double
                    gpx1 = Me.gvDetail.Item(19, gvDetail.CurrentRow.Index).Value
                    gpx2 = Me.gvDetail.Item(18, gvDetail.CurrentRow.Index).Value
                    irsPCgpAV = ((gpx1 * 100) / gpx2)
                    Me.gvDetail.Item(20, gvDetail.CurrentRow.Index).Value = Format(irsPCgpAV, "##,##0.00")
                    ' vi1 = 1               
                    'End If
                End If
            End If
            'End If
            '    End If


            ' Key แก้ไขราคา---------------------------------------
            If e.ColumnIndex = 12 Then
                Dim irsDCT As Double = 0
                vPriceLP = Me.gvDetail.Item(5, gvDetail.CurrentRow.Index).Value
                If Me.LPcbx.Checked = True Then
                    icl12 = Me.gvDetail.Item(12, gvDetail.CurrentRow.Index).Value ' ราคา
                    irsDCT = (100 - ((icl12 * 100) / vPriceLP))
                    Me.gvDetail.Item(11, gvDetail.CurrentRow.Index).Value = Format(irsDCT, "##,##0.00")
                    ' SmartPoint
                    Dim irsSMP As Double
                    irsSMP = ((vNewPriceKey * icl13) / 100)
                    Me.gvDetail.Item(14, gvDetail.CurrentRow.Index).Value = Format(irsSMP, "##,##0.0000")
                    'GP ทุนตลาด
                    Dim irsGPMK As Double
                    irsGPMK = ((Me.gvDetail.Item(12, gvDetail.CurrentRow.Index).Value - Me.gvDetail.Item(15, gvDetail.CurrentRow.Index).Value) - Me.gvDetail.Item(14, gvDetail.CurrentRow.Index).Value)
                    Me.gvDetail.Item(16, gvDetail.CurrentRow.Index).Value = Format(irsGPMK, "##,##0.00")
                    '% GP ทุนตลาด
                    Dim irsPCgpMK As Double
                    irsPCgpMK = ((Me.gvDetail.Item(16, gvDetail.CurrentRow.Index).Value * 100) / Me.gvDetail.Item(15, gvDetail.CurrentRow.Index).Value)
                    Me.gvDetail.Item(17, gvDetail.CurrentRow.Index).Value = Format(irsPCgpMK, "##,##0.00")
                    'Gp ทุนเฉลี่ย Lot
                    Dim irsGPAVlot
                    irsGPAVlot = ((Me.gvDetail.Item(12, gvDetail.CurrentRow.Index).Value - Me.gvDetail.Item(18, gvDetail.CurrentRow.Index).Value) - Me.gvDetail.Item(14, gvDetail.CurrentRow.Index).Value)
                    Me.gvDetail.Item(19, gvDetail.CurrentRow.Index).Value = Format(irsGPAVlot, "##,##0.00")
                    '% GP ทุนเฉลี่ย
                    Dim irsPCgpAV As Double
                    irsPCgpAV = ((Me.gvDetail.Item(19, gvDetail.CurrentRow.Index).Value * 100) / Me.gvDetail.Item(18, gvDetail.CurrentRow.Index).Value)
                    Me.gvDetail.Item(20, gvDetail.CurrentRow.Index).Value = Format(irsPCgpAV, "##,##0.00")
                Else
                    icl12 = Me.gvDetail.Item(12, gvDetail.CurrentRow.Index).Value ' ราคา
                    irsDCT = (100 - ((icl12 * 100) / icl6))
                    Me.gvDetail.Item(11, gvDetail.CurrentRow.Index).Value = Format(irsDCT, "##,##0.00")
                    ' SmartPoint
                    Dim irsSMP As Double
                    irsSMP = ((vNewPriceKey * icl13) / 100)
                    Me.gvDetail.Item(14, gvDetail.CurrentRow.Index).Value = Format(irsSMP, "##,##0.0000")
                    'GP ทุนตลาด
                    Dim irsGPMK As Double
                    irsGPMK = ((Me.gvDetail.Item(12, gvDetail.CurrentRow.Index).Value - Me.gvDetail.Item(15, gvDetail.CurrentRow.Index).Value) - Me.gvDetail.Item(14, gvDetail.CurrentRow.Index).Value)
                    Me.gvDetail.Item(16, gvDetail.CurrentRow.Index).Value = Format(irsGPMK, "##,##0.00")
                    '% GP ทุนตลาด
                    Dim irsPCgpMK As Double
                    irsPCgpMK = ((Me.gvDetail.Item(16, gvDetail.CurrentRow.Index).Value * 100) / Me.gvDetail.Item(15, gvDetail.CurrentRow.Index).Value)
                    Me.gvDetail.Item(17, gvDetail.CurrentRow.Index).Value = Format(irsPCgpMK, "##,##0.00")
                    'Gp ทุนเฉลี่ย Lot
                    Dim irsGPAVlot
                    irsGPAVlot = ((Me.gvDetail.Item(12, gvDetail.CurrentRow.Index).Value - Me.gvDetail.Item(18, gvDetail.CurrentRow.Index).Value) - Me.gvDetail.Item(14, gvDetail.CurrentRow.Index).Value)
                    Me.gvDetail.Item(19, gvDetail.CurrentRow.Index).Value = Format(irsGPAVlot, "##,##0.00")
                    '% GP ทุนเฉลี่ย
                    Dim irsPCgpAV As Double
                    irsPCgpAV = ((Me.gvDetail.Item(19, gvDetail.CurrentRow.Index).Value * 100) / Me.gvDetail.Item(18, gvDetail.CurrentRow.Index).Value)
                    Me.gvDetail.Item(20, gvDetail.CurrentRow.Index).Value = Format(irsPCgpAV, "##,##0.00")

                End If
            End If


            ' Key แก้ไข %SmartPoint--------------------------------
            If e.ColumnIndex = 13 Then
                Dim irsSMPV As Double
                irsSMPV = ((Me.gvDetail.Item(12, gvDetail.CurrentRow.Index).Value * Me.gvDetail.Item(13, gvDetail.CurrentRow.Index).Value) / 100)
                Me.gvDetail.Item(14, gvDetail.CurrentRow.Index).Value = Format(irsSMPV, "##,##0.0000")
                'GP ทุนตลาด
                Dim irsGPMK As Double
                irsGPMK = ((Me.gvDetail.Item(12, gvDetail.CurrentRow.Index).Value - Me.gvDetail.Item(15, gvDetail.CurrentRow.Index).Value) - Me.gvDetail.Item(14, gvDetail.CurrentRow.Index).Value)
                Me.gvDetail.Item(16, gvDetail.CurrentRow.Index).Value = Format(irsGPMK, "##,##0.00")
                '% GP ทุนตลาด
                Dim irsPCgpMK As Double
                irsPCgpMK = ((Me.gvDetail.Item(16, gvDetail.CurrentRow.Index).Value * 100) / Me.gvDetail.Item(15, gvDetail.CurrentRow.Index).Value)
                Me.gvDetail.Item(17, gvDetail.CurrentRow.Index).Value = Format(irsPCgpMK, "##,##0.00")
                'Gp ทุนเฉลี่ย Lot
                Dim irsGPAVlot
                irsGPAVlot = ((Me.gvDetail.Item(12, gvDetail.CurrentRow.Index).Value - Me.gvDetail.Item(18, gvDetail.CurrentRow.Index).Value) - Me.gvDetail.Item(14, gvDetail.CurrentRow.Index).Value)
                Me.gvDetail.Item(19, gvDetail.CurrentRow.Index).Value = Format(irsGPAVlot, "##,##0.00")
                '% GP ทุนเฉลี่ย
                Dim irsPCgpAV As Double
                irsPCgpAV = ((Me.gvDetail.Item(19, gvDetail.CurrentRow.Index).Value * 100) / Me.gvDetail.Item(18, gvDetail.CurrentRow.Index).Value)
                Me.gvDetail.Item(20, gvDetail.CurrentRow.Index).Value = Format(irsPCgpAV, "##,##0.00")
            End If
        End If

    End Sub

    Private Sub gvDetail_CellEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles gvDetail.CellEnter
        'Dim vNewPriceKey As Double
        ''New var----------------------------
        'Dim icl5 As Double = 0
        'Dim icl6 As Double
        'Dim icl7 As Double
        'Dim icl9 As Double
        'Dim icl10 As Double
        'Dim icl11 As Double
        'Dim icl12 As Double
        'Dim icl13 As Double
        'Dim icl14 As Double
        'Dim icl16 As Double
        'Dim vi1 As Integer = 0
        ''--------------------------------------
        ''ราคาที่1
        'icl5 = Me.gvDetail.Item(5, gvDetail.CurrentRow.Index).Value 'คอลัม ราคาที่ 1
        'icl6 = Me.gvDetail.Item(6, gvDetail.CurrentRow.Index).Value ' %SmartPoint1
        'icl7 = Me.gvDetail.Item(7, gvDetail.CurrentRow.Index).Value ' SmartPoint 1
        'icl9 = Me.gvDetail.Item(9, gvDetail.CurrentRow.Index).Value ' Volume
        '' icl10 = Me.gvDetail.Item(10, gvDetail.CurrentRow.Index).Value ' ส่วนลดจากราคาที่1
        'icl11 = Me.gvDetail.Item(11, gvDetail.CurrentRow.Index).Value ' ราคา
        'icl12 = Me.gvDetail.Item(12, gvDetail.CurrentRow.Index).Value ' %SmartPoint
        'icl13 = Me.gvDetail.Item(13, gvDetail.CurrentRow.Index).Value ' SmartPoint
        'icl14 = Me.gvDetail.Item(14, gvDetail.CurrentRow.Index).Value ' GP ทุนตลาด
        'icl16 = Me.gvDetail.Item(16, gvDetail.CurrentRow.Index).Value ' GP ทุนเฉลี่ยตาม lot

        'If Me.gvDetail.Rows.Count > 0 Then
        '    ' Key แก้ไข %SmartPoint ราคาที่ 1
        '    If e.ColumnIndex = 6 Then
        '        Dim irsSMP1 As Double
        '        irsSMP1 = ((icl5 * icl6) / 100)
        '        Me.gvDetail.Item(7, gvDetail.CurrentRow.Index).Value = Format(irsSMP1, "##,##0.00")
        '    End If

        '    'Key แก้ไขส่วนลดจากราคาที่ 1--------------------------
        '    If e.ColumnIndex = 10 Then
        '        icl10 = Me.gvDetail.Item(10, gvDetail.CurrentRow.Index).Value ' ส่วนลดจากราคาที่1
        '        'ราคา
        '        Dim irsPrice01 As Double = 0
        '        irsPrice01 = (icl5 - ((icl5 * icl10) / 100))
        '        vNewPriceKey = FNnewvDCM(irsPrice01)
        '        Me.gvDetail.Item(11, gvDetail.CurrentRow.Index).Value = Format(vNewPriceKey, "##,##0.00")
        '        ' SmartPoint
        '        Dim irsSMP As Double
        '        irsSMP = ((vNewPriceKey * icl12) / 100)
        '        Me.gvDetail.Item(13, gvDetail.CurrentRow.Index).Value = Format(irsSMP, "##,##0.0000")
        '        'GP ทุนตลาด
        '        Dim irsGPMK As Double
        '        irsGPMK = (icl11 - icl14)
        '        Me.gvDetail.Item(15, gvDetail.CurrentRow.Index).Value = Format(irsGPMK, "##,##0.00")
        '        'Gp ทุนเฉลี่ย Lot
        '        Dim irsGPAVlot
        '        irsGPAVlot = (icl11 - icl16)
        '        Me.gvDetail.Item(17, gvDetail.CurrentRow.Index).Value = Format(irsGPAVlot, "##,##0.00")
        '        ' vi1 = 1

        '        ' Key แก้ไขราคา---------------------------------------
        '    ElseIf e.ColumnIndex = 11 Then
        '        Dim irsDCT As Double = 0
        '        icl11 = Me.gvDetail.Item(11, gvDetail.CurrentRow.Index).Value ' ราคา
        '        irsDCT = (100 - ((icl11 * 100) / icl5))
        '        Me.gvDetail.Item(10, gvDetail.CurrentRow.Index).Value = Format(irsDCT, "##,##0.00")
        '        'GP ทุนตลาด
        '        Dim irsGPMK As Double
        '        irsGPMK = (icl11 - icl14)
        '        Me.gvDetail.Item(15, gvDetail.CurrentRow.Index).Value = Format(irsGPMK, "##,##0.00")
        '        'Gp ทุนเฉลี่ย Lot
        '        Dim irsGPAVlot
        '        irsGPAVlot = (icl11 - icl16)
        '        Me.gvDetail.Item(17, gvDetail.CurrentRow.Index).Value = Format(irsGPAVlot, "##,##0.00")
        '    End If


        '    ' Key แก้ไข %SmartPoint--------------------------------
        '    If e.ColumnIndex = 12 Then
        '        Dim irsSMPV As Double
        '        irsSMPV = ((icl11 * icl12) / 100)
        '        Me.gvDetail.Item(13, gvDetail.CurrentRow.Index).Value = Format(irsSMPV, "##,##0.0000")
        '    End If

        'End If
    End Sub



    Private Sub gvDetail_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles gvDetail.CellValueChanged
        'Dim vPriceKey As Double
        'Dim vNewPriceKey As Double
        ''New var----------------------------
        'Dim icl5 As Double = 0
        'Dim icl6 As Double
        'Dim icl7 As Double
        'Dim icl9 As Double
        'Dim icl10 As Double
        'Dim icl11 As Double
        'Dim icl12 As Double
        'Dim icl13 As Double
        'Dim icl14 As Double
        'Dim icl16 As Double
        'Dim vi1 As Integer = 0


        ''--------------------------------------
        ''ราคาที่1
        'icl5 = Me.gvDetail.Item(5, gvDetail.CurrentRow.Index).Value 'คอลัม ราคาที่ 1
        'icl6 = Me.gvDetail.Item(6, gvDetail.CurrentRow.Index).Value ' %SmartPoint1
        'icl7 = Me.gvDetail.Item(7, gvDetail.CurrentRow.Index).Value ' SmartPoint 1
        'icl9 = Me.gvDetail.Item(9, gvDetail.CurrentRow.Index).Value ' Volume
        '' icl10 = Me.gvDetail.Item(10, gvDetail.CurrentRow.Index).Value ' ส่วนลดจากราคาที่1
        'icl11 = Me.gvDetail.Item(11, gvDetail.CurrentRow.Index).Value ' ราคา
        'icl12 = Me.gvDetail.Item(12, gvDetail.CurrentRow.Index).Value ' %SmartPoint
        'icl13 = Me.gvDetail.Item(13, gvDetail.CurrentRow.Index).Value ' SmartPoint
        'icl14 = Me.gvDetail.Item(14, gvDetail.CurrentRow.Index).Value ' GP ทุนตลาด
        'icl16 = Me.gvDetail.Item(16, gvDetail.CurrentRow.Index).Value ' GP ทุนเฉลี่ยตาม lot

        'If Me.gvDetail.Rows.Count > 0 Then
        '    ' Key แก้ไข %SmartPoint ราคาที่ 1
        '    If e.ColumnIndex = 6 Then
        '        Dim irsSMP1 As Double
        '        irsSMP1 = ((icl5 * icl6) / 100)
        '        Me.gvDetail.Item(7, gvDetail.CurrentRow.Index).Value = Format(irsSMP1, "##,##0.00")
        '    End If

        '    'Key แก้ไขส่วนลดจากราคาที่ 1--------------------------
        '    If e.ColumnIndex = 10 Then
        '        icl10 = Me.gvDetail.Item(10, gvDetail.CurrentRow.Index).Value ' ส่วนลดจากราคาที่1
        '        'ราคา
        '        Dim irsPrice01 As Double = 0
        '        irsPrice01 = (icl5 - ((icl5 * icl10) / 100))
        '        vNewPriceKey = FNnewvDCM(irsPrice01)
        '        Me.gvDetail.Item(11, gvDetail.CurrentRow.Index).Value = Format(vNewPriceKey, "##,##0.00")
        '        ' SmartPoint
        '        Dim irsSMP As Double
        '        irsSMP = ((vNewPriceKey * icl12) / 100)
        '        Me.gvDetail.Item(13, gvDetail.CurrentRow.Index).Value = Format(irsSMP, "##,##0.0000")
        '        'GP ทุนตลาด
        '        Dim irsGPMK As Double
        '        irsGPMK = (icl11 - icl14)
        '        Me.gvDetail.Item(15, gvDetail.CurrentRow.Index).Value = Format(irsGPMK, "##,##0.00")
        '        'Gp ทุนเฉลี่ย Lot
        '        Dim irsGPAVlot
        '        irsGPAVlot = (icl11 - icl16)
        '        Me.gvDetail.Item(17, gvDetail.CurrentRow.Index).Value = Format(irsGPAVlot, "##,##0.00")
        '        ' vi1 = 1

        '        ' Key แก้ไขราคา---------------------------------------
        '    ElseIf e.ColumnIndex = 11 Then
        '        Dim irsDCT As Double = 0
        '        icl11 = Me.gvDetail.Item(11, gvDetail.CurrentRow.Index).Value ' ราคา
        '        irsDCT = (100 - ((icl11 * 100) / icl5))
        '        Me.gvDetail.Item(10, gvDetail.CurrentRow.Index).Value = Format(irsDCT, "##,##0.00")
        '        'GP ทุนตลาด
        '        Dim irsGPMK As Double
        '        irsGPMK = (icl11 - icl14)
        '        Me.gvDetail.Item(15, gvDetail.CurrentRow.Index).Value = Format(irsGPMK, "##,##0.00")
        '        'Gp ทุนเฉลี่ย Lot
        '        Dim irsGPAVlot
        '        irsGPAVlot = (icl11 - icl16)
        '        Me.gvDetail.Item(17, gvDetail.CurrentRow.Index).Value = Format(irsGPAVlot, "##,##0.00")
        '    End If


        '    ' Key แก้ไข %SmartPoint--------------------------------
        '    If e.ColumnIndex = 12 Then
        '        Dim irsSMPV As Double
        '        irsSMPV = ((icl11 * icl12) / 100)
        '        Me.gvDetail.Item(13, gvDetail.CurrentRow.Index).Value = Format(irsSMPV, "##,##0.0000")
        '    End If

        'End If
        '---------------------------------------------------------------------------------------------------------------
        'If Me.gvDetail.Rows.Count > 0 Then
        '    If e.ColumnIndex = 5 Then
        '        'me.gvDetail .Item (5,gvdetail.CurrentRow .Index )
        '        'ราคา
        '        vPriceKey = Format(gvDetail.Item(5, gvDetail.CurrentRow.Index).Value - (gvDetail.Item(5, gvDetail.CurrentRow.Index).Value * gvDetail.Item(10, gvDetail.CurrentRow.Index).Value) / 100, "##,##0.00")
        '        vNewPriceKey = newvDCM(vPriceKey)
        '        Me.gvDetail.Item(11, gvDetail.CurrentRow.Index).Value = Format(vNewPriceKey, "##,##0.00")
        '        ' SmartPoint
        '        Me.gvDetail.Item(13, gvDetail.CurrentRow.Index).Value = Format((gvDetail.Item(5, gvDetail.CurrentRow.Index).Value * gvDetail.Item(12, gvDetail.CurrentRow.Index).Value) / 100, "##,##0.0000")
        '        ' SmartPoint ราคา 1
        '        Me.gvDetail.Item(7, gvDetail.CurrentRow.Index).Value = Format((gvDetail.Item(5, gvDetail.CurrentRow.Index).Value * gvDetail.Item(6, gvDetail.CurrentRow.Index).Value) / 100, "##,##0.0000")
        '        'GP ทุนตลาด
        '        Me.gvDetail.Item(15, gvDetail.CurrentRow.Index).Value = Format((gvDetail.Item(11, gvDetail.CurrentRow.Index).Value - gvDetail.Item(14, gvDetail.CurrentRow.Index).Value), "##,##0.00")
        '        'Gp ทุนเฉลี่ย Lot
        '        Me.gvDetail.Item(17, gvDetail.CurrentRow.Index).Value = Format((gvDetail.Item(11, gvDetail.CurrentRow.Index).Value - gvDetail.Item(16, gvDetail.CurrentRow.Index).Value), "##,##0.00")
        '    End If
        '    If e.ColumnIndex = 6 Then
        '        Me.gvDetail.Item(7, gvDetail.CurrentRow.Index).Value = Format((gvDetail.Item(5, gvDetail.CurrentRow.Index).Value * gvDetail.Item(6, gvDetail.CurrentRow.Index).Value) / 100, "##,##0.00")
        '    End If
        '    If e.ColumnIndex = 10 Then
        '        'ราคา
        '        vPriceKey = Format(gvDetail.Item(5, gvDetail.CurrentRow.Index).Value - (gvDetail.Item(5, gvDetail.CurrentRow.Index).Value * gvDetail.Item(10, gvDetail.CurrentRow.Index).Value) / 100, "##,##0.00")
        '        vNewPriceKey = newvDCM(vPriceKey)
        '        Me.gvDetail.Item(11, gvDetail.CurrentRow.Index).Value = Format(vNewPriceKey, "##,##0.00")
        '        ' SmartPoint
        '        Me.gvDetail.Item(13, gvDetail.CurrentRow.Index).Value = Format((gvDetail.Item(5, gvDetail.CurrentRow.Index).Value * gvDetail.Item(12, gvDetail.CurrentRow.Index).Value) / 100, "##,##0.0000")
        '        'GP ทุนตลาด
        '        Me.gvDetail.Item(15, gvDetail.CurrentRow.Index).Value = Format((gvDetail.Item(11, gvDetail.CurrentRow.Index).Value - gvDetail.Item(14, gvDetail.CurrentRow.Index).Value), "##,##0.00")
        '        'Gp ทุนเฉลี่ย Lot
        '        Me.gvDetail.Item(17, gvDetail.CurrentRow.Index).Value = Format((gvDetail.Item(11, gvDetail.CurrentRow.Index).Value - gvDetail.Item(16, gvDetail.CurrentRow.Index).Value), "##,##0.00")

        '    End If
        '    If e.ColumnIndex = 12 Then
        '        Me.gvDetail.Item(13, gvDetail.CurrentRow.Index).Value = Format((gvDetail.Item(6, gvDetail.CurrentRow.Index).Value * gvDetail.Item(12, gvDetail.CurrentRow.Index).Value) / 100, "##,##0.0000")
        '    End If
        'End If
        '-------------------------------------------------------------------------------6-03-2009
        'If Me.gvDetail.Rows.Count > 0 Then
        '    If e.ColumnIndex = 5 Then
        '        Me.gvDetail.Item(11, gvDetail.CurrentRow.Index).Value = (gvDetail.CurrentRow.Cells(5).Value * gvDetail.CurrentRow.Cells(10).Value) / 100
        '    End If
        'End If
        '        DataGridView1.CellValueChanged()

        '        If DataGridView1.Rows.Count > 0 Then

        '            If e.ColumnIndex = 5 Then

        '                Dim ReturnQty As Integer

        '                ReturnQty = DataGridView1.Item(5, DataGridView1.CurrentRow.Index).Value

        '                If ReturnQty = 0 Then

        '                    DataGridView1.Item(6, DataGridView1.CurrentRow.Index).Value = 0 'หน่วยละ

        '                    Zero789()

        '                Else

        '                    If ReturnQty <= DataGridView1.Item(3, DataGridView1.CurrentRow.Index).Value Then

        '                        DataGridView1.Item(6, DataGridView1.CurrentRow.Index).Value = DataGridView1.Item(4, DataGridView1.CurrentRow.Index).Value 'หน่วยละ

        '                        DataGridView1.Item(7, DataGridView1.CurrentRow.Index).Value = (DataGridView1.Item(4, 
        'DataGridView1.CurrentRow.Index).Value * DataGridView1.Item(10, DataGridView1.CurrentRow.Index).Value) / 100 'ส่วนลด GP
        '                        DataGridView1.Item(8, DataGridView1.CurrentRow.Index).Value = DataGridView1.Item(4, DataGridView1.CurrentRow.Index).Value - DataGridView1.Item(7, DataGridView1.CurrentRow.Index).Value 'ราคาสุทธิ

        '                        DataGridView1.Item(9, DataGridView1.CurrentRow.Index).Value = ReturnQty * DataGridView1.Item(8, DataGridView1.CurrentRow.Index).Value 'จำนวนเงิน

        '                    Else 'ถ้าคืนสินค้ามากกว่าที่มีอยู่

        '                        Msg("คืนสินค้ามากกว่าจำนวนที่มีอยู่ไม่ได้", MsgType.Invalid)

        '                        DataGridView1.Item(5, DataGridView1.CurrentRow.Index).Value = 0

        '                        DataGridView1.Item(6, DataGridView1.CurrentRow.Index).Value = 0 'หน่วยละ

        '                        Zero789()

        '                    End If

        '                End If

        '                Calculate()
        '            End If

    End Sub


    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    Private Sub txtVM2_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtVM2.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.txtVM3.Focus()
        End If
        'Dim ivm2 As String

        'ivm2 = Me.txtVM2.Text

        ''If ivm2 <> "" Then

        'If e.KeyCode = Keys.Enter Then

        '    If ivm2 <> "" And CDbl(ivm2) > 1 Then

        '        Me.txtVM3.Focus()

        '    Else

        '        MsgBox("ค่าของ Volume 2 ต้องมากกว่า 1", MsgBoxStyle.Critical, "Warnings")

        '        '  Me.txtVM2.Focus()

        '    End If

        'End If
        'End If
    End Sub

    Private Sub txtVM2_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtVM2.LostFocus
        'Me.txtVM3.Focus()
        'Dim ivm2i As String

        'ivm2i = Me.txtVM2.Text

        'If ivm2i <> "" And (CDbl(ivm2i) > 1) Then

        '    Me.txtVM3.Focus()

        'Else
        '    MsgBox("ค่าของ Volume 2 ต้องมากกว่า 1", MsgBoxStyle.Critical, "Warnings")
        '    'Me.txtVM2.Focus()

        'End If

    End Sub

    Private Sub txtVM3_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtVM3.LostFocus
        'Me.txtDC2.Focus()
        'Dim ivLM2 As String
        'Dim ivLM3 As String

        'ivLM2 = Me.txtVM2.Text
        'ivLM3 = Me.txtVM3.Text

        'If ivLM2 <> "" And ivLM3 <> "" Then

        '    If (CDbl(ivLM3) >= CDbl(ivLM2)) Then
        '        Me.txtDC2.Focus()
        '    Else
        '        MsgBox("ค่าของ Volume3 ต้องมากกว่าหรือเท่ากับ Volume2", MsgBoxStyle.Critical, "Error")
        '        '  Me.txtVM3.Focus()
        '    End If

        'End If
    End Sub

    Private Sub txtVM3_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtVM3.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.txtDC2.Focus()
        End If
        'Dim ivLM2 As String
        'Dim ivLM3 As String
        'ivLM2 = Me.txtVM2.Text
        'ivLM3 = Me.txtVM3.Text
        'If e.KeyCode = Keys.Enter Then
        '    If Me.txtVM2.Text <> "" And Me.txtVM3.Text <> "" And CDbl(ivLM2) >= CDbl(ivLM2) Then
        '        Me.txtDC2.Focus()
        '    Else
        '        MsgBox("ค่าของ Volume3 ต้องมากกว่าหรือเท่ากับ Volume2", MsgBoxStyle.Critical, "Error")
        '        Me.txtVM3.Focus()
        '    End If
        'End If

    End Sub

    Private Sub txtDC2_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtDC2.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.txtDC3.Focus()
        End If
        'Dim vDC2 As String
        'Dim vDCrs1 As String
        'vDC2 = Me.txtDC2.Text
        'If e.KeyCode = Keys.Enter Then
        '    vDCrs1 = CHKdcPC(vDC2)
        '    If vDC2 <> "" And CDbl(vDCrs1) > 0 Then
        '        Me.txtDC3.Focus()
        '    Else
        '        MsgBox("ส่วนลดต้องมีค่ามากกว่า 0 ", MsgBoxStyle.Critical, "Error")
        '        Me.txtDC2.Focus()
        '    End If
        'End If

    End Sub

    Private Sub txtDC3_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtDC3.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.btnProduct.Focus()
        End If
        'Dim vDC2 As String
        'Dim vDC3 As String
        'Dim vDCrs1 As String
        'Dim vDCrs2 As String
        'vDC2 = Me.txtDC2.Text
        'vDC3 = Me.txtDC3.Text

        'If e.KeyCode = Keys.Enter Then
        '    vDCrs1 = CHKdcPC(vDC2)
        '    vDCrs2 = CHKdcPC(vDC3)
        '    If vDC2 <> "" And vDC3 <> "" And (CDbl(vDCrs2) >= CDbl(vDCrs1)) Then
        '        Me.btnProduct.Focus()
        '    Else
        '        MsgBox("ส่วนลดของ Volume 3 ต้องมากกว่าหรือเท่ากับ Volume 2 ", MsgBoxStyle.Critical, "Error")
        '        Me.txtDC3.Focus()
        '    End If
        'End If

    End Sub
    Private Sub txtDC2_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDC2.LostFocus
        'Me.txtDC3.Focus()
        'Dim vDC2 As String
        'Dim vDCrs1 As String
        'vDC2 = Me.txtDC2.Text
        'If vDC2 <> "" Then
        '    vDCrs1 = CHKdcPC(vDC2)
        '    If CDbl(vDCrs1) > 0 Then
        '        Me.txtDC3.Focus()
        '    Else
        '        MsgBox("ส่วนลดต้องมีค่ามากกว่า 0 ", MsgBoxStyle.Critical, "Warnings")
        '        Me.txtDC2.Focus()
        '    End If
        'End If

    End Sub

    Private Sub txtDC3_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDC3.LostFocus
        'Me.btnProduct.Focus()
        'Dim vDC2 As String
        'Dim vDC3 As String
        'Dim vDCrs1 As String
        'Dim vDCrs2 As String
        'vDC2 = Me.txtDC2.Text
        'vDC3 = Me.txtDC3.Text
        'If vDC2 <> "" And vDC3 <> "" Then
        '    vDCrs1 = CHKdcPC(vDC2)
        '    vDCrs2 = CHKdcPC(vDC3)
        '    If (CDbl(vDCrs2) >= CDbl(vDCrs1)) Then
        '        Me.btnProduct.Focus()
        '    Else
        '        MsgBox("ส่วนลดของ Volume 3 ต้องมากกว่าหรือเท่ากับ Volume 2 ", MsgBoxStyle.Critical, "Error")
        '        Me.txtDC3.Focus()
        '    End If
        'End If
    End Sub

    Private Sub btnSaveAS_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSaveAS.Click
        Call GenNewdoc()
        iNewPSVdoc = Me.gvDetail.Item(1, gvDetail.Rows(1).Index).Value
        '  sv = "0"
        MsgBox("เลขที่เอกสารใหม่ของคุณคือ" & vNewdocNo & "", MsgBoxStyle.Information, "Information")
        Call gvReadonly()
        Me.gvDetail.AllowUserToDeleteRows = True
        Me.btnSave.Enabled = True
        Me.PBConfirm.Visible = False
        Me.PBNew.Visible = True
        sv = "1"
        Me.gvDetail.Enabled = False
        Me.txtVM2.Enabled = True
        Me.txtVM3.Enabled = True
        Me.txtDC2.Enabled = True
        Me.txtDC3.Enabled = True
        Me.txtSMTP2.Enabled = True
        Me.txtSMTP3.Enabled = True
        Me.BtnGenVLM.Visible = True
        Me.btnProduct.Visible = False
        Me.GroupBox1.Enabled = True
        Me.DTPdocDate.Enabled = True
        Me.DTPStartDate.Enabled = True
        Me.DTPEndDate.Enabled = True
        Me.DTPdocDate.Text = Now.Date
        Me.DTPStartDate.Text = DateAdd(DateInterval.Day, 1, Now.Date)
        Me.DTPEndDate.Text = DateAdd(DateInterval.Year, 5, Now.Date)
        Me.smpLV1.Enabled = True
        Me.smpLV1.Value = 1.0
        ' iNewPSVdoc = Me.gvDetail.Item(1, gvDetail.Rows(1).Index).Value


        ' Me.gvDetail.Columns(11).ReadOnly = False
        'Me.gvDetail.EditMode = DataGridViewEditMode.EditOnF2
    End Sub
    Private Function CHKdcPC(ByVal vD As String) As String
        Dim vx1 As String
        ' Dim vx2 As String
        'Dim vx3 As Double
        Dim i As Integer
        i = Microsoft.VisualBasic.Len(vD)
        'vx1 = vD
        i = InStr(vD, "%")
        If i = 0 Then
            vD = vD
        Else
            vx1 = Microsoft.VisualBasic.Left(vD, i - 1)
            ' CDbl(vx1)
            ' vx2 = (vx1 / 100)
            vD = vx1
        End If
        Return vD
    End Function
    Private Sub gvDetail_CurrentCellChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles gvDetail.CurrentCellChanged
        ''Dim vPriceKey As Double
        'Dim vNewPriceKey As Double
        ''New var----------------------------
        'Dim icl5 As Double = 0
        'Dim icl6 As Double
        'Dim icl7 As Double
        'Dim icl9 As Double
        'Dim icl10 As Double
        'Dim icl11 As Double
        'Dim icl12 As Double
        'Dim icl13 As Double
        'Dim icl14 As Double
        'Dim icl16 As Double
        'Dim vi1 As Integer
        'Dim vi2 As Integer

        ''--------------------------------------
        ''ราคาที่1
        'icl5 = Me.gvDetail.Item(5, gvDetail.CurrentRow.Index).Value 'คอลัม ราคาที่ 1
        'icl6 = Me.gvDetail.Item(6, gvDetail.CurrentRow.Index).Value ' %SmartPoint1
        'icl7 = Me.gvDetail.Item(7, gvDetail.CurrentRow.Index).Value ' SmartPoint 1
        'icl9 = Me.gvDetail.Item(9, gvDetail.CurrentRow.Index).Value ' Volume
        'icl10 = Me.gvDetail.Item(10, gvDetail.CurrentRow.Index).Value ' ส่วนลดจากราคาที่1
        'icl11 = Me.gvDetail.Item(11, gvDetail.CurrentRow.Index).Value ' ราคา
        'icl12 = Me.gvDetail.Item(12, gvDetail.CurrentRow.Index).Value ' %SmartPoint
        'icl13 = Me.gvDetail.Item(13, gvDetail.CurrentRow.Index).Value ' SmartPoint
        'icl14 = Me.gvDetail.Item(14, gvDetail.CurrentRow.Index).Value ' GP ทุนตลาด
        'icl16 = Me.gvDetail.Item(16, gvDetail.CurrentRow.Index).Value ' GP ทุนเฉลี่ยตาม lot
        'vi1 = Me.gvDetail.Item(10, gvDetail.CurrentRow.Index).ColumnIndex
        'vi2 = Me.gvDetail.Item(11, gvDetail.CurrentRow.Index).ColumnIndex

        'If Me.gvDetail.Rows.Count > 0 Then
        '    ' Key แก้ไข %SmartPoint ราคาที่ 1
        '    If e.ColumnIndex = 6 Then
        '        Dim irsSMP1 As Double
        '        irsSMP1 = ((icl5 * icl6) / 100)
        '        Me.gvDetail.Item(7, gvDetail.CurrentRow.Index).Value = Format(irsSMP1, "##,##0.00")
        '    End If

        '    'Key แก้ไขส่วนลดจากราคาที่ 1--------------------------
        '    If e.ColumnIndex = 10 Then
        '        icl10 = Me.gvDetail.Item(10, gvDetail.CurrentRow.Index).Value ' ส่วนลดจากราคาที่1
        '        'ราคา
        '        Dim irsPrice01 As Double = 0
        '        irsPrice01 = (icl5 - ((icl5 * icl10) / 100))
        '        vNewPriceKey = FNnewvDCM(irsPrice01)
        '        Me.gvDetail.Item(11, gvDetail.CurrentRow.Index).Value = Format(vNewPriceKey, "##,##0.00")
        '        ' SmartPoint
        '        Dim irsSMP As Double
        '        irsSMP = ((vNewPriceKey * icl12) / 100)
        '        Me.gvDetail.Item(13, gvDetail.CurrentRow.Index).Value = Format(irsSMP, "##,##0.0000")
        '        'GP ทุนตลาด
        '        Dim irsGPMK As Double
        '        irsGPMK = (icl11 - icl14)
        '        Me.gvDetail.Item(15, gvDetail.CurrentRow.Index).Value = Format(irsGPMK, "##,##0.00")
        '        'Gp ทุนเฉลี่ย Lot
        '        Dim irsGPAVlot
        '        irsGPAVlot = (icl11 - icl16)
        '        Me.gvDetail.Item(17, gvDetail.CurrentRow.Index).Value = Format(irsGPAVlot, "##,##0.00")

        '    End If



        '    ' Key แก้ไขราคา---------------------------------------
        '    If e.ColumnIndex = 11 Then
        '        Dim irsDCT As Double = 0
        '        icl11 = Me.gvDetail.Item(11, gvDetail.CurrentRow.Index).Value ' ราคา
        '        irsDCT = (100 - ((icl11 * 100) / icl5))
        '        Me.gvDetail.Item(10, gvDetail.CurrentRow.Index).Value = Format(irsDCT, "##,##0.00")
        '        'GP ทุนตลาด
        '        Dim irsGPMK As Double
        '        irsGPMK = (icl11 - icl14)
        '        Me.gvDetail.Item(15, gvDetail.CurrentRow.Index).Value = Format(irsGPMK, "##,##0.00")
        '        'Gp ทุนเฉลี่ย Lot
        '        Dim irsGPAVlot
        '        irsGPAVlot = (icl11 - icl16)
        '        Me.gvDetail.Item(17, gvDetail.CurrentRow.Index).Value = Format(irsGPAVlot, "##,##0.00")
        '    End If


        '    ' Key แก้ไข %SmartPoint--------------------------------
        '    If e.ColumnIndex = 12 Then
        '        Dim irsSMPV As Double
        '        irsSMPV = ((icl11 * icl12) / 100)
        '        Me.gvDetail.Item(13, gvDetail.CurrentRow.Index).Value = Format(irsSMPV, "##,##0.0000")
        '    End If

        'End If
    End Sub

    Private Sub DTPStartDate_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles DTPStartDate.KeyDown
        Dim vDate As Date
        If e.KeyCode = Keys.Enter Then
            vDate = DateAdd(DateInterval.Day, 1, Now.Date)
            If Me.DTPStartDate.Value < vDate Then
                MsgBox("วันที่ปรับราคาต้องไม่น้อยกว่าวันพรุ่งนี้", MsgBoxStyle.Critical, "Error")
            End If
        End If
    End Sub

    Private Sub DTPStartDate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DTPStartDate.ValueChanged
        'Dim vDate As Date
        'vDate = DateAdd(DateInterval.Day, 1, Now.Date)
        'If Me.DTPStartDate.Value < vDate Then
        '    MsgBox("วันที่ปรับราคาต้องไม่น้อยกว่าวันพรุ่งนี้", MsgBoxStyle.Critical, "Error")
        'End If
    End Sub

   
    Private Sub BtnGenVLM_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnGenVLM.Click
        Me.gvDetail.DataSource = Nothing
        vDate = DateAdd(DateInterval.Day, 1, Now.Date).ToShortDateString
        Dim ivLM3 As String
        Dim ivLM2 As String
        Dim vDC2 As String
        Dim vDCrs1 As String
        Dim vDC3 As String
        'Dim vDCrs1 As String
        Dim vDCrs2 As String
        Dim vd1 As String
        ' Dim vd2 As String
        Dim i1 As Double
        Dim i2 As Double
        'Dim i3 As Double
        'Dim i4 As Double
        '---------------------
        Dim rs1 As String
        Dim rs2 As String
        Dim vrs1 As String
        '--------------------
        Dim lpvDC1 As String
        Dim lpvDC2 As String
        Dim lpST As Integer
        '
        ivLM2 = Me.txtVM2.Text
        ivLM3 = Me.txtVM3.Text
        vDC2 = Me.txtDC2.Text
        vDC3 = Me.txtDC3.Text
        vDCrs1 = CHKdcPC(vDC2) ' เปลี่ยนเปอร์เซนต์ให้เป็นตัวเลข
        vDCrs2 = CHKdcPC(vDC3)
        'On Error GoTo chkChar1
        i1 = CDbl(ivLM2)
        i2 = CDbl(ivLM3)
        If Me.LPcbx.Checked = True Then
            lpST = 1
        End If
        If Me.DTPStartDate.Value < vDate Then
            MsgBox("วันที่ปรับราคาต้องไม่น้อยกว่าวันพรุ่งนี้.กรุณากำหนดใหม่", MsgBoxStyle.Critical, "Error")
            Me.DTPStartDate.Focus()
        Else

            If Me.txtVM2.Text <> "" And Me.txtVM3.Text <> "" And Me.txtDC2.Text <> "" And Me.txtDC3.Text <> "" And Me.smpLV1.Value > 0 Then
                vd1 = chkBPD(vDC2, vDC3)
                If (vd1 = "0") Then
                    MsgBox("คุณกำหนดหน่วยส่วนลดไม่เหมือนกัน กรุณากำหนดใหม่", MsgBoxStyle.Critical, "Warnings")
                Else
                    If ivLM2 <> "" And ivLM3 <> "" And (CDbl(ivLM2) > 1) And (CDbl(ivLM3) >= CDbl(ivLM2)) And (CDbl(vDCrs2) >= CDbl(vDCrs1)) Then
                        rs1 = CHKVMDC(ivLM2, ivLM3) ' ตรวจสอบการป้อน volume
                        rs2 = CHKVMDC(vDCrs1, vDCrs2) ' ตรวจสอบการป้อนส่วนลด
                        vrs1 = CHKVMDC1(ivLM2, ivLM3, vDCrs1, vDCrs2) 'ตรวจสอบการป้อนส่วนลด และ Volume
                        If rs1 = 0 Or rs2 = 0 Then
                            MsgBox("ค่าของVolume 2 หรือ ค่าส่วนลดของราคาที่ 2 ต้องน้อยกว่า ค่าของระดับราคาที่ 3 กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Warnings")
                        ElseIf vrs1 = 0 Then
                            MsgBox("ค่าของ Volume และส่วนลดไม่ถูกต้อง กรุณาตรวจสอบใหม่", MsgBoxStyle.Critical, "Warnings")
                        Else
                            pvLV2 = Me.txtLv2.Text
                            pvLV3 = Me.txtLv3.Text
                            pvVM2 = Me.txtVM2.Text
                            pvVM3 = Me.txtVM3.Text
                            pvDC2 = Me.txtDC2.Text
                            pvDC3 = Me.txtDC3.Text
                            pvSMP2 = Me.txtSMTP2.Text
                            pvSMP3 = Me.txtSMTP3.Text
                            '--------------------------------------
                            Dim dtbl As New DataTable("Ptbl")
                            Dim dr As DataRow
                            Dim n As Integer
                            Dim nml As Integer = 0
                            Dim i As Integer
                            Dim vnPrice2 As Double
                            Dim vnPrice3 As Double
                            'Dim vPSdocno As String
                            ''vLV2
                            'Dim nPrice2 As Double
                            ''vLV3
                            'Dim nPrice3 As Double
                            ''Dim gpAVmklot As Double
                            'Dim vPriceLV2 As Double
                            'Dim vPriceLV1 As Double
                            'Dim vSLPlv1 As Double
                            'Dim vSMP001 As Double
                            'Dim vSMP002 As Double
                            ''-----------%GP
                            'Dim GPmk1 As Double
                            'Dim vPCGPMK1 As Double
                            'Dim GPav1 As Double
                            'Dim vPCGPAV1 As Double
                            'Dim GPmk2 As Double
                            'Dim vPCGPMK2 As Double
                            'Dim vPCGPAV2 As Double
                            ''-----------
                            'Dim vPsmtpoint As Double
                            'Dim vPsmtpoint2 As Double
                            ''-----------
                            'Dim vFMcode As String
                            'Dim lpvDC1 As String
                            'Dim lpvDC2 As String
                            '-----------
                            dtbl.Columns.Add("No.", GetType(String))
                            dtbl.Columns.Add("เลขที่เอกสารโครงสร้างราคา", GetType(String))
                            dtbl.Columns.Add("รหัสสินค้า", GetType(String))
                            dtbl.Columns.Add("ชื่อสินค้า", GetType(String))
                            dtbl.Columns.Add("หน่วยขาย", GetType(String))
                            dtbl.Columns.Add("ราคา LP", GetType(String))
                            dtbl.Columns.Add("ราคาที่1", GetType(String))
                            dtbl.Columns.Add("%smartpoint ราคาที่1", GetType(String))
                            dtbl.Columns.Add("SmartPoint1", GetType(String))
                            dtbl.Columns.Add("ระดับราคา", GetType(String))
                            dtbl.Columns.Add("Volume", GetType(String))
                            dtbl.Columns.Add("%ส่วนลดจากราคาที่ 1", GetType(String))
                            dtbl.Columns.Add("ราคา", GetType(String))
                            dtbl.Columns.Add("%SmartPoint", GetType(String))
                            dtbl.Columns.Add("SmartPoint", GetType(String))
                            dtbl.Columns.Add("ทุนตลาดSaleVat", GetType(String))
                            dtbl.Columns.Add("GP ทุนตลาดSaleVat", GetType(String))
                            dtbl.Columns.Add("%GP ทุนตลาด", GetType(String))
                            dtbl.Columns.Add("ทุนเฉลี่ย LotSaleVat", GetType(String))
                            dtbl.Columns.Add("GP ทุนเฉลี่ย LotSaleVat", GetType(String))
                            dtbl.Columns.Add("%GP ทุนเฉลี่ย", GetType(String))
                            lpvDC1 = pvDC2
                            lpvDC2 = pvDC3

                            QryString = "exec dbo.USP_PS_PriceVolumeSetPrepare '" & "','" & "','" & "','" & "','" & iNewPSVdoc & "'"
                            da = New SqlDataAdapter(QryString, vConnectionString)
                            ds = New DataSet
                            da.Fill(ds, "psdoc")
                            dt = ds.Tables("psdoc")
                            If dt.Rows.Count > 0 Then
                                'do for count dt
                                '----------------------Price Level 2
                                'pvLV2 = Me.txtLV2.Text
                                'pvLV3 = Me.txtLv3.Text
                                'pvVM2 = Me.txtVM2.Text
                                'pvVM3 = Me.txtVM3.Text
                                'pvDC2 = Me.txtDC2.Text
                                'pvDC3 = Me.txtDC3.Text
                                'pvSMP2 = Me.txtSMTP2.Text
                                'pvSMP3 = Me.txtSMTP3.Text
                                For i = 0 To dt.Rows.Count - 1
                                    dr = dtbl.NewRow
                                    nml = nml + 1
                                    dr("No.") = nml
                                    dr("เลขที่เอกสารโครงสร้างราคา") = dt.Rows(i).Item("Docno")
                                    dr("รหัสสินค้า") = dt.Rows(i).Item("itemcode")
                                    dr("ชื่อสินค้า") = dt.Rows(i).Item("itemname")
                                    dr("หน่วยขาย") = dt.Rows(i).Item("SaleUnitcode")
                                    'ราคาจากโรงงาน
                                    Dim vLPPrice As Double
                                    vLPPrice = CDbl(dt.Rows(i).Item("Price_LP"))
                                    dr("ราคา LP") = Format(vLPPrice, "##,##0.00")
                                    'Price1
                                    Dim vPriceLV1 As Double
                                    vPriceLV1 = CDbl(dt.Rows(i).Item("Price1"))

                                    dr("ราคาที่1") = Format(vPriceLV1, "##,##0.00")
                                    'เช็คสินค้าวัสดุก่อสร้าง
                                    Dim vFMcode As String
                                    Dim vSLPlv1 As Double
                                    Dim vPsmtpoint As Double
                                    Dim vPsmtpoint2 As Double
                                    vFMcode = dt.Rows(i).Item("Familycode")
                                    If vFMcode = 10000 Then
                                        vSLPlv1 = 0.0
                                        vPsmtpoint = 0.0
                                        vPsmtpoint2 = 0.0
                                    Else
                                        vPsmtpoint = CDbl(pvSMP2)
                                        vPsmtpoint2 = CDbl(pvSMP3)
                                        vSLPlv1 = Me.smpLV1.Value
                                    End If
                                    Dim xSmartPoint As Double
                                    '------------------------------------------------
                                    xSmartPoint = ((vPriceLV1 * vSLPlv1) / 100)
                                    dr("%SmartPoint ราคาที่1") = Format(vSLPlv1, "##,##0.00")
                                    dr("SmartPoint1") = Format(xSmartPoint, "##,##0.0000") '---------smartpoint 1
                                    dr("ระดับราคา") = Format(Int(pvLV2), "##,##0") '--------------Level2
                                    dr("Volume") = Format(Int(pvVM2), "##,##0.00") '----------Volume 2
                                    Dim ivPDC2 As String
                                    ivPDC2 = CHKdiscount(lpvDC1) 'ค่าเปอร์เซ็นต์
                                    Dim va As Double
                                    Dim vPCDC01 As Double
                                    Dim nPrice2 As Double
                                    If ivPDC2 = lpvDC1 Then
                                        va = ((ivPDC2 * 100) / vPriceLV1)
                                        vPCDC01 = Format(va, "##,##0.00") '---------%ส่วนลด2
                                        dr("%ส่วนลดจากราคาที่ 1") = vPCDC01
                                        If lpST = 1 Then
                                            nPrice2 = (vLPPrice - ivPDC2)
                                        Else
                                            nPrice2 = (vPriceLV1 - ivPDC2)
                                        End If
                                        Format(nPrice2, "##,##0.00")
                                        vnPrice2 = chkDCM(nPrice2)
                                        'xx = Format(CStr(nPrice2)
                                        'ราคา1*ส่วนลด%ราคา1/100
                                        dr("ราคา") = Format(vnPrice2, "##,##0.00")
                                    Else
                                        vPCDC01 = Format(CDbl(ivPDC2), "##,##0.00") '---------%ส่วนลด2
                                        dr("%ส่วนลดจากราคาที่ 1") = vPCDC01
                                        If lpST = 1 Then
                                            nPrice2 = (vLPPrice - ((vLPPrice * vPCDC01) / 100))
                                        Else
                                            nPrice2 = (vPriceLV1 - ((vPriceLV1 * vPCDC01) / 100))
                                        End If
                                        Format(nPrice2, "##,##0.00")
                                        vnPrice2 = chkDCM(nPrice2)
                                        'xx = Format(CStr(nPrice2)
                                        'ราคา1*ส่วนลด%ราคา1/100
                                        dr("ราคา") = Format(vnPrice2, "##,##0.00")
                                    End If
                                    dr("%SmartPoint") = Format(vPsmtpoint, "##,##0.00")
                                    'ราคา1*%smartpoint/100
                                    Dim vSMP001 As Double
                                    Dim GPmk1 As Double
                                    Dim vPCGPMK1 As Double
                                    Dim GPav1 As Double
                                    Dim vPCGPAV1 As Double
                                    vSMP001 = ((vnPrice2 * vPsmtpoint) / 100)
                                    dr("SmartPoint") = Format(vSMP001, "##,##0.0000") '---------------smartpoint 2
                                    Dim vMKcostSV1 As Double
                                    vMKcostSV1 = dt.Rows(i).Item("marketcostSaleVat")
                                    dr("ทุนตลาดSaleVat") = Format(vMKcostSV1, "##,##0.00")
                                    GPmk1 = ((vnPrice2 - vMKcostSV1) - vSMP001)
                                    dr("GP ทุนตลาดSaleVat") = Format(GPmk1, "##,##0.00")
                                    vPCGPMK1 = ((GPmk1 * 100) / vMKcostSV1)
                                    dr("%GP ทุนตลาด") = Format(vPCGPMK1, "##,##0.00")
                                    Dim vAVcostLotSV1 As Double
                                    vAVcostLotSV1 = dt.Rows(i).Item("AveragecostLotSaleVat")
                                    dr("ทุนเฉลี่ย LotSaleVat") = Format(vAVcostLotSV1, "##,##0.00")
                                    ' gpAVmklot = (Me.gvDetail.Columns(11).ToString - Me.gvDetail.Columns(16).ToString)
                                    GPav1 = ((vnPrice2 - vAVcostLotSV1) - vSMP001)
                                    dr("GP ทุนเฉลี่ย LotSaleVat") = Format(GPav1, "##,##0.00")
                                    vPCGPAV1 = ((GPav1 * 100) / vAVcostLotSV1)
                                    dr("%GP ทุนเฉลี่ย") = Format(vPCGPAV1, "##,##0.00")
                                    dtbl.Rows.Add(dr)

                                    '----------------------Price Level 3------------------------------------------------------------------------------------
                                    dr = dtbl.NewRow
                                    nml = nml + 1
                                    dr("No.") = nml
                                    dr("เลขที่เอกสารโครงสร้างราคา") = dt.Rows(i).Item("Docno")
                                    dr("รหัสสินค้า") = dt.Rows(i).Item("itemcode")
                                    dr("ชื่อสินค้า") = dt.Rows(i).Item("itemname")
                                    dr("หน่วยขาย") = dt.Rows(i).Item("SaleUnitcode")
                                    'ราคาจากโรงงาน
                                    Dim vLPPrice2 As Double
                                    vLPPrice2 = CDbl(dt.Rows(i).Item("Price_LP"))
                                    dr("ราคา LP") = Format(vLPPrice2, "##,##0.00")
                                    'Price1
                                    Dim vPriceLV2 As Double
                                    vPriceLV2 = CDbl(dt.Rows(i).Item("Price1"))
                                    dr("ราคาที่1") = Format(vPriceLV2, "##,##0.00")
                                    vPriceLV1 = CDbl(dt.Rows(i).Item("Price1"))
                                    dr("ราคาที่1") = Format(vPriceLV1, "##,##0.00")
                                    'vSLPlv1 = frmPriceVolumeSet.smpLV1.Value
                                    '------------------------------------------------
                                    xSmartPoint = ((vPriceLV1 * vSLPlv1) / 100)
                                    dr("%SmartPoint ราคาที่1") = Format(vSLPlv1, "##,##0.00")
                                    dr("SmartPoint1") = Format(xSmartPoint, "##,##0.0000") '---------smartpoint 1
                                    dr("ระดับราคา") = Format(Int(pvLV3), "##,##0") '--------Level 3
                                    dr("Volume") = Format(Int(pvVM3), "##,##0.00") '--------Volume3
                                    Dim vPCDC02 As Double
                                    Dim ivPDC3 As String
                                    ivPDC3 = CHKdiscount(lpvDC2)
                                    Dim vx As Double
                                    Dim nPrice3 As Double
                                    If ivPDC3 = lpvDC2 Then
                                        vx = ((ivPDC3 * 100) / vPriceLV2)
                                        vPCDC02 = Format(vx, "##,##0.00") '--------%ส่วนลด3
                                        dr("%ส่วนลดจากราคาที่ 1") = vPCDC02
                                        If lpST = 1 Then
                                            nPrice3 = vLPPrice2 - ivPDC3
                                        Else
                                            nPrice3 = vPriceLV2 - ivPDC3
                                        End If

                                        'ราคา1*ส่วนลด%ราคา1/100
                                        vnPrice3 = chkDCM(nPrice3)
                                        dr("ราคา") = Format(vnPrice3, "##,##0.00")
                                    Else
                                        vPCDC02 = Format(CDbl(ivPDC3), "##,##0.00") '--------%ส่วนลด3
                                        dr("%ส่วนลดจากราคาที่ 1") = vPCDC02
                                        If lpST = 1 Then
                                            nPrice3 = vLPPrice2 - ((vLPPrice2 * vPCDC02) / 100)
                                        Else
                                            nPrice3 = vPriceLV2 - ((vPriceLV2 * vPCDC02) / 100)
                                        End If

                                        'ราคา1*ส่วนลด%ราคา1/100
                                        vnPrice3 = chkDCM(nPrice3)
                                        dr("ราคา") = Format(vnPrice3, "##,##0.00")
                                    End If
                                    dr("%SmartPoint") = vPsmtpoint2
                                    'ราคา1*%smartpoint/100
                                    Dim vSMP002 As Double
                                    Dim GPmk2 As Double
                                    Dim vPCGPMK2 As Double
                                    Dim vPCGPAV2 As Double
                                    vSMP002 = ((vnPrice3 * vPsmtpoint2) / 100)
                                    dr("SmartPoint") = Format(vSMP002, "##,##0.0000") '---------smartpoint 3
                                    Dim vMKcostSV2 As Double
                                    vMKcostSV2 = dt.Rows(i).Item("marketcostSaleVat")
                                    dr("ทุนตลาดSaleVat") = Format(vMKcostSV2, "##,##0.00")
                                    GPmk2 = ((vnPrice3 - vMKcostSV2) - vSMP002)
                                    dr("GP ทุนตลาดSaleVat") = Format(GPmk2, "##,##0.00")
                                    vPCGPMK2 = ((GPmk2 * 100) / vMKcostSV2)
                                    dr("%GP ทุนตลาด") = Format(vPCGPMK2, "##,##0.00")
                                    Dim vAVcostLotSV2 As Double
                                    vAVcostLotSV2 = dt.Rows(i).Item("AveragecostLotSaleVat")
                                    dr("ทุนเฉลี่ย LotSaleVat") = Format(vAVcostLotSV2, "##,##0.00")
                                    GPav1 = ((vnPrice3 - vAVcostLotSV2) - vSMP002)
                                    dr("GP ทุนเฉลี่ย LotSaleVat") = Format(GPav1, "##,##0.00")
                                    vPCGPAV2 = ((GPav1 * 100) / vAVcostLotSV2)
                                    dr("%GP ทุนเฉลี่ย") = Format(vPCGPAV2, "##,##0.00")
                                    dtbl.Rows.Add(dr)
                                    Me.gvDetail.BackgroundColor = Color.Cyan
                                Next
                            End If


                            Me.gvDetail.DataSource = dtbl
                            For n = 0 To Me.gvDetail.Rows.Count - 1
                                If n Mod 2 = 0 Then
                                    Me.gvDetail.Rows(n).DefaultCellStyle.BackColor = Color.SkyBlue
                                End If
                            Next
                            Call gvReadonly()
                            sv = "1"
                            Me.gvDetail.Enabled = True
                            Me.GroupBox1.Enabled = False
                            Me.gvDetail.AllowUserToDeleteRows = True
                            Me.LPcbx.Enabled = False

                            'If Me.gvDetail.Rows.Count > 0 Then
                            '    Me.gvDetail.Enabled = True
                            '    For i = 0 To Me.gvDetail.Rows.Count - 1

                            '        Dim NsmpLV1 As Double
                            '        NsmpLV1 = Me.smpLV1.Value
                            '        Me.gvDetail.Item(6, gvDetail.CurrentRow.Index).Value = NsmpLV1
                            '        'smartpoint1
                            '        Dim NPrice1 As Double
                            '        NPrice1 = Me.gvDetail.Item(5, gvDetail.CurrentRow.Index).Value
                            '        Dim Nsmp As Double
                            '        Nsmp = ((NPrice1 * NsmpLV1) / 100)
                            '        Me.gvDetail.Item(7, gvDetail.CurrentRow.Index).Value = Nsmp

                            '        '  For n = 0 To Me.gvDetail.Rows.Count - 1

                            '        If i Mod 2 = 0 Then

                            '            'Me.gvDetail.Item(9, gvDetail.CurrentRow.Index).Value = Me.txtVM3.Text
                            '            'Me.gvDetail.Item(10, gvDetail.CurrentRow.Index).Value = vDCrs2
                            '            'Me.gvDetail.Item(12, gvDetail.CurrentRow.Index).Value = Me.txtSMTP3.Text
                            '            ''ราคา
                            '            'nPrice = ((NPrice1 * vDCrs1) / 100)
                            '            'nwPrice = FNnewvDCM(nPrice)
                            '            'Me.gvDetail.Item(11, gvDetail.CurrentRow.Index).Value = nwPrice
                            '            Me.gvDetail.Item(9, gvDetail.CurrentRow.Index).Value = Me.txtVM3.Text
                            '            Me.gvDetail.Item(10, gvDetail.CurrentRow.Index).Value = vDCrs2
                            '            Me.gvDetail.Item(12, gvDetail.CurrentRow.Index).Value = Me.txtSMTP3.Text
                            '            'SmartPoint 1
                            '            Dim NPrice1v As Double
                            '            NPrice1v = Me.gvDetail.Item(5, gvDetail.CurrentRow.Index).Value
                            '            Dim nwPcSMP1v As Double
                            '            Dim nwSMP1v As Double
                            '            nwSMP1v = ((NPrice1v * NsmpLV1) / 100)
                            '            Me.gvDetail.Item(7, gvDetail.CurrentRow.Index).Value = nwSMP1v

                            '            Dim nwPcSMPv As Double
                            '            Dim nwSMPv As Double

                            '            'ราคา
                            '            Dim nPricev As Double
                            '            Dim nwPricev As Double
                            '            nPricev = NPrice1v - ((NPrice1v * vDCrs2) / 100)
                            '            nwPricev = FNnewvDCM(nPricev)
                            '            Me.gvDetail.Item(11, gvDetail.CurrentRow.Index).Value = Format(nwPricev, "##,##0.00")

                            '            'Smart Point
                            '            nwPcSMPv = CDbl(Me.txtSMTP2.Text)
                            '            nwSMPv = ((nwPricev * nwPcSMPv) / 100)
                            '            'GP Market
                            '            Dim nwMKv As Double
                            '            Dim nwGPMKv As Double
                            '            nwMKv = Me.gvDetail.Item(14, gvDetail.CurrentRow.Index).Value
                            '            nwGPMKv = (nwPricev - nwMKv - nwSMPv)
                            '            Me.gvDetail.Item(15, gvDetail.CurrentRow.Index).Value = nwGPMKv
                            '            '%GP Market
                            '            Dim nwPCgpMKv As Double
                            '            nwPCgpMKv = ((nwGPMKv * 100) / nwMKv)
                            '            Me.gvDetail.Item(16, gvDetail.CurrentRow.Index).Value = nwPCgpMKv

                            '            'GP Average--------------
                            '            Dim nwAVv As Double
                            '            Dim nwGPavv As Double
                            '            nwAVv = Me.gvDetail.Item(17, gvDetail.CurrentRow.Index).Value
                            '            nwGPavv = (nwPricev - nwAVv - nwSMPv)
                            '            Me.gvDetail.Item(18, gvDetail.CurrentRow.Index).Value = nwGPavv
                            '            '%GP Average
                            '            Dim nwPCgpAVv As Double
                            '            nwPCgpAVv = ((nwGPavv * 100) / nwAVv)
                            '            Me.gvDetail.Item(19, gvDetail.CurrentRow.Index).Value = nwPCgpAVv

                            '        Else
                            '            Dim nPrice As Double
                            '            Dim nwPrice As Double
                            '            Me.gvDetail.Item(9, gvDetail.CurrentRow.Index).Value = Me.txtVM2.Text
                            '            Me.gvDetail.Item(10, gvDetail.CurrentRow.Index).Value = vDCrs1
                            '            Me.gvDetail.Item(12, gvDetail.CurrentRow.Index).Value = Me.txtSMTP2.Text
                            '            'SmartPoint 1
                            '            Dim nwPcSMP1 As Double
                            '            Dim nwSMP1 As Double
                            '            nwSMP1 = ((NPrice1 * NsmpLV1) / 100)
                            '            Me.gvDetail.Item(7, gvDetail.CurrentRow.Index).Value = nwSMP1

                            '            Dim nwPcSMP As Double
                            '            Dim nwSMP As Double

                            '            'ราคา
                            '            nPrice = NPrice1 - ((NPrice1 * vDCrs1) / 100)
                            '            nwPrice = FNnewvDCM(nPrice)
                            '            Me.gvDetail.Item(11, gvDetail.CurrentRow.Index).Value = nwPrice

                            '            'Smart Point
                            '            nwPcSMP = CDbl(Me.txtSMTP2.Text)
                            '            nwSMP = ((nwPrice * nwPcSMP) / 100)
                            '            'GP Market
                            '            Dim nwMK As Double
                            '            Dim nwGPMK As Double
                            '            nwMK = Me.gvDetail.Item(14, gvDetail.CurrentRow.Index).Value
                            '            nwGPMK = (nwPrice - nwMK - nwSMP)
                            '            Me.gvDetail.Item(15, gvDetail.CurrentRow.Index).Value = nwGPMK
                            '            '%GP Market
                            '            Dim nwPCgpMK As Double
                            '            nwPCgpMK = ((nwGPMK * 100) / nwMK)
                            '            Me.gvDetail.Item(16, gvDetail.CurrentRow.Index).Value = nwPCgpMK

                            '            'GP Average--------------
                            '            Dim nwAV As Double
                            '            Dim nwGPav As Double
                            '            nwAV = Me.gvDetail.Item(17, gvDetail.CurrentRow.Index).Value
                            '            nwGPav = (nwPrice - nwAV - nwSMP)
                            '            Me.gvDetail.Item(18, gvDetail.CurrentRow.Index).Value = nwGPav
                            '            '%GP Average
                            '            Dim nwPCgpAV As Double
                            '            nwPCgpAV = ((nwGPav * 100) / nwAV)
                            '            Me.gvDetail.Item(19, gvDetail.CurrentRow.Index).Value = nwPCgpAV

                            '            '----------

                            '        End If

                            '        ' Next n
                            '    Next i
                            'End If
                        End If
                    Else
                        MsgBox("ค่า Volume2 อาจน้อยกว่าหรือเท่ากับ 1 หรือ ค่าของข้อมูลระดับราคาที่ 3 น้อยกว่าระดับราคาที่ 2 กรุณาตรวจสอบอีกครั้ง", MsgBoxStyle.Critical, "Warnings")

                    End If

                End If

            Else
                MsgBox("คุณใส่ข้อมูลไม่ครบ..กรุณาตรวจสอบอีกครั้ง.", MsgBoxStyle.Critical, "Error")
            End If


chkChar1:
            If Err.Description <> "" Then
                MsgBox("คุณป้อนข้อมูลไม่ถูกต้อง กรุณาป้อนใหม่", MsgBoxStyle.Critical, "Error")
            End If
        End If
        '-------------        
    End Sub
    Private Function CHKdiscount(ByVal vD As String) As String
        Dim vx1 As String
        ' Dim vx2 As String
        'Dim vx3 As Double
        Dim i As Integer
        i = Microsoft.VisualBasic.Len(vD)
        'vx1 = vD
        i = InStr(vD, "%")

        If i = 0 Then
            vD = vD
        Else
            vx1 = Microsoft.VisualBasic.Left(vD, i - 1)
            ' CDbl(vx1)
            vD = vx1
        End If
        Return vD
    End Function
    Private Function chkDCM(ByVal vx As String) As String
        Dim a As Integer
        Dim i As Integer
        Dim vL As String
        Dim vc As Double
        'Format(vx, "##.00")
        a = Microsoft.VisualBasic.Len(vx)
        i = InStr(vx, ".")
        If i = 0 Then
            i = 1
            vL = vx
        Else
            i = i
            vL = Microsoft.VisualBasic.Left(vx, i - 1)
        End If
        'vR = Microsoft.VisualBasic.Right(vx, a - i)

        vc = (vx - vL)

        If (vc >= 0.5 And vc < 1.0) Then
            vc = 1
            vx = (vL + vc) & "." & "00"

        ElseIf (vc > 0.0 And vc < 0.5) Then
            '  vc = 0.5
            vx = vL & ".5"
        End If
        Return vx
    End Function

End Class