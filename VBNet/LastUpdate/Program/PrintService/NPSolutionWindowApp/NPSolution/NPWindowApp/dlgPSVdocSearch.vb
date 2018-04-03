Imports System.Windows.Forms
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports System.Globalization
Imports Microsoft.VisualBasic


Public Class dlgPSVdocSearch
    Dim QryString As String
    Dim vSearch As String
    Dim da As SqlDataAdapter
    Dim ds As DataSet
    Dim ipsLSV As New ListViewItem
    Dim i As Integer
    Dim txtLV2 As String
    '-------------------------
    Dim vPSdocno As String
    'vLV2
    Dim nPrice2 As Double
    'vLV3
    Dim nPrice3 As Double
    'Dim gpAVmklot As Double
    Dim vPriceLV2 As Double
    Dim vPriceLV1 As Double
    Dim vSLPlv1 As Double
    Dim vSMP001 As Double
    Dim vSMP002 As Double
    '-----------%GP
    Dim GPmk1 As Double
    Dim vPCGPMK1 As Double
    Dim GPav1 As Double
    Dim vPCGPAV1 As Double
    Dim GPmk2 As Double
    Dim vPCGPMK2 As Double
    Dim vPCGPAV2 As Double
    '-----------
    Dim vPsmtpoint As Double
    Dim vPsmtpoint2 As Double
    '-----------
    Dim vFMcode As String
    Dim lpvDC1 As String
    Dim lpvDC2 As String
    Dim vPCDC01 As Double
    Dim vLPs As Integer
    '------------------------


    Private Sub btnOKps_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOKps.Click
        Call GetSearchLV()
    End Sub
    Private Sub GetSearchLV()
        Dim iLSV As New ListViewItem
        Dim ivUser As String
        Me.LVps.Items.Clear()
        vSearch = Me.txtPSDoc.Text
        If vUserID = "nueng" Or vUserID = "panuvich" Or vUserID = "komkrithc" Then
            ivUser = ""
        Else
            ivUser = vUserID
        End If
        QryString = "exec dbo.USP_PS_PriceVolumeSetFromStructure '" & vSearch & "','" & ivUser & "'"
        da = New SqlDataAdapter(QryString, vConnectionString)
        ds = New DataSet
        da.Fill(ds, "TBsch")
        dt = ds.Tables("TBsch")
        If dt.Rows.Count <> 0 Then
            For i = 0 To dt.Rows.Count - 1
                iLSV = Me.LVps.Items.Add(dt.Rows(i).Item("psDocno"))
                iLSV.SubItems.Add(0).Text = dt.Rows(i).Item("Docdate")
                If dt.Rows(i).Item("Ownercode") Is DBNull.Value = False Then
                    iLSV.SubItems.Add(0).Text = CInt(dt.Rows(i).Item("Ownercode"))
                Else
                    iLSV.SubItems.Add(0).Text = "ไม่มีรหัสผู้กำหนด"
                End If
                If dt.Rows(i).Item("userid") Is DBNull.Value = False Then
                    iLSV.SubItems.Add(0).Text = dt.Rows(i).Item("userid")
                Else
                    iLSV.SubItems.Add(0).Text = "ไม่มี User Id"
                End If
                If dt.Rows(i).Item("Owner") Is DBNull.Value = False Then
                    iLSV.SubItems.Add(0).Text = dt.Rows(i).Item("Owner")
                Else
                    iLSV.SubItems.Add(0).Text = "ไม่มีชื่อผู้กำหนด"
                End If
            Next
        End If

    End Sub

    Private Sub dlgPSVdocSearch_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        frmPriceVolumeSet.Enabled = True
    End Sub

    Private Sub dlgPSVdocSearch_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call InitializeDataBase()
        Me.WindowState = FormWindowState.Normal
        Me.StartPosition = FormStartPosition.CenterScreen
        txtLV2 = frmPriceVolumeSet.txtDC2.Text
        publicFdocno = frmPriceVolumeSet.txtDocno.Text
        Me.txtPSDoc.Focus()
        vLPs = lpStatus

        'Call getHeadDoc()
    End Sub
    Private Sub gvformat()
        frmPriceVolumeSet.gvDetail.Columns(0).Width = 30
        frmPriceVolumeSet.gvDetail.Columns(1).Width = 75
        frmPriceVolumeSet.gvDetail.Columns(2).Width = 75
        frmPriceVolumeSet.gvDetail.Columns(3).Width = 200
        frmPriceVolumeSet.gvDetail.Columns(4).Width = 50
        frmPriceVolumeSet.gvDetail.Columns(5).Width = 70
        frmPriceVolumeSet.gvDetail.Columns(6).Width = 70
        frmPriceVolumeSet.gvDetail.Columns(7).Width = 50
        frmPriceVolumeSet.gvDetail.Columns(8).Width = 25
        frmPriceVolumeSet.gvDetail.Columns(9).Width = 40
        frmPriceVolumeSet.gvDetail.Columns(10).Width = 30
        frmPriceVolumeSet.gvDetail.Columns(11).Width = 80
        frmPriceVolumeSet.gvDetail.Columns(12).Width = 40
        frmPriceVolumeSet.gvDetail.Columns(13).Width = 70
        frmPriceVolumeSet.gvDetail.Columns(14).Width = 70
        frmPriceVolumeSet.gvDetail.Columns(15).Width = 70
        frmPriceVolumeSet.gvDetail.Columns(16).Width = 70
        frmPriceVolumeSet.gvDetail.Columns(17).Width = 70
        frmPriceVolumeSet.gvDetail.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        frmPriceVolumeSet.gvDetail.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        frmPriceVolumeSet.gvDetail.Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        frmPriceVolumeSet.gvDetail.Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        frmPriceVolumeSet.gvDetail.Columns(9).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        frmPriceVolumeSet.gvDetail.Columns(10).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        frmPriceVolumeSet.gvDetail.Columns(11).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        frmPriceVolumeSet.gvDetail.Columns(12).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        frmPriceVolumeSet.gvDetail.Columns(13).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        frmPriceVolumeSet.gvDetail.Columns(14).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        frmPriceVolumeSet.gvDetail.Columns(15).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        frmPriceVolumeSet.gvDetail.Columns(16).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        frmPriceVolumeSet.gvDetail.Columns(17).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        frmPriceVolumeSet.gvDetail.Columns(18).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        frmPriceVolumeSet.gvDetail.Columns(19).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        frmPriceVolumeSet.gvDetail.Columns(20).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
    End Sub
    Private Sub gvReadonly()
        frmPriceVolumeSet.gvDetail.Columns(0).ReadOnly = True
        frmPriceVolumeSet.gvDetail.Columns(1).ReadOnly = True
        frmPriceVolumeSet.gvDetail.Columns(2).ReadOnly = True
        frmPriceVolumeSet.gvDetail.Columns(3).ReadOnly = True
        frmPriceVolumeSet.gvDetail.Columns(4).ReadOnly = True
        frmPriceVolumeSet.gvDetail.Columns(5).ReadOnly = True
        frmPriceVolumeSet.gvDetail.Columns(6).ReadOnly = True
        frmPriceVolumeSet.gvDetail.Columns(7).ReadOnly = True
        frmPriceVolumeSet.gvDetail.Columns(8).ReadOnly = True
        frmPriceVolumeSet.gvDetail.Columns(9).ReadOnly = True
        frmPriceVolumeSet.gvDetail.Columns(10).ReadOnly = True
        frmPriceVolumeSet.gvDetail.Columns(11).ReadOnly = True
        frmPriceVolumeSet.gvDetail.Columns(12).ReadOnly = True
        frmPriceVolumeSet.gvDetail.Columns(13).ReadOnly = True
        frmPriceVolumeSet.gvDetail.Columns(14).ReadOnly = True
        frmPriceVolumeSet.gvDetail.Columns(15).ReadOnly = True
        frmPriceVolumeSet.gvDetail.Columns(16).ReadOnly = True
        frmPriceVolumeSet.gvDetail.Columns(17).ReadOnly = True

    End Sub
    Private Sub gvReadonly1()
        frmPriceVolumeSet.gvDetail.Columns(0).ReadOnly = True
        frmPriceVolumeSet.gvDetail.Columns(1).ReadOnly = True
        frmPriceVolumeSet.gvDetail.Columns(2).ReadOnly = True
        frmPriceVolumeSet.gvDetail.Columns(3).ReadOnly = True
        frmPriceVolumeSet.gvDetail.Columns(4).ReadOnly = True
        frmPriceVolumeSet.gvDetail.Columns(5).ReadOnly = True
        frmPriceVolumeSet.gvDetail.Columns(6).ReadOnly = True
        frmPriceVolumeSet.gvDetail.Columns(7).ToolTipText = "คลิกแก้ไข %SmartPoint1"
        frmPriceVolumeSet.gvDetail.Columns(7).DefaultCellStyle.ForeColor = Color.Blue
        frmPriceVolumeSet.gvDetail.Columns(8).ReadOnly = True
        frmPriceVolumeSet.gvDetail.Columns(9).ReadOnly = True
        frmPriceVolumeSet.gvDetail.Columns(10).ToolTipText = "คลิกเพื่อแก้ไข Volume รายการนี้"
        frmPriceVolumeSet.gvDetail.Columns(10).DefaultCellStyle.ForeColor = Color.Blue
        frmPriceVolumeSet.gvDetail.Columns(11).ToolTipText = "คลิกเพื่อแก้ไขส่วนลดรายการนี้"
        frmPriceVolumeSet.gvDetail.Columns(11).DefaultCellStyle.ForeColor = Color.Blue
        frmPriceVolumeSet.gvDetail.Columns(12).DefaultCellStyle.ForeColor = Color.Blue
        'frmPriceVolumeSet.gvDetail.Columns(11).ReadOnly = True
        frmPriceVolumeSet.gvDetail.Columns(13).ToolTipText = "คลิกเพื่อแก้ไข SmartPoint รายการนี้"
        frmPriceVolumeSet.gvDetail.Columns(13).DefaultCellStyle.ForeColor = Color.Blue
        frmPriceVolumeSet.gvDetail.Columns(14).ReadOnly = True
        frmPriceVolumeSet.gvDetail.Columns(15).ReadOnly = True
        frmPriceVolumeSet.gvDetail.Columns(16).ReadOnly = True
        frmPriceVolumeSet.gvDetail.Columns(17).ReadOnly = True
        frmPriceVolumeSet.gvDetail.Columns(18).ReadOnly = True
        frmPriceVolumeSet.gvDetail.Columns(19).ReadOnly = True
        frmPriceVolumeSet.gvDetail.Columns(20).ReadOnly = True
        'frmPriceVolumeSet.gvDetail.Columns(0).DefaultCellStyle.BackColor = Color.LightYellow
        'frmPriceVolumeSet.gvDetail.Columns(1).DefaultCellStyle.BackColor = Color.LightYellow
        'frmPriceVolumeSet.gvDetail.Columns(2).DefaultCellStyle.BackColor = Color.LightYellow
        'frmPriceVolumeSet.gvDetail.Columns(3).DefaultCellStyle.BackColor = Color.LightYellow
        'frmPriceVolumeSet.gvDetail.Columns(4).DefaultCellStyle.BackColor = Color.LightYellow
        ' frmPriceVolumeSet.gvDetail.Columns(5).DefaultCellStyle.BackColor = Color.LightSalmon
        'frmPriceVolumeSet.gvDetail.Columns(7).DefaultCellStyle.BackColor = Color.LightYellow
        'frmPriceVolumeSet.gvDetail.Columns(8).DefaultCellStyle.BackColor = Color.LightYellow
        'frmPriceVolumeSet.gvDetail.Columns(11).DefaultCellStyle.BackColor = Color.LightSalmon
        'frmPriceVolumeSet.gvDetail.Columns(13).DefaultCellStyle.BackColor = Color.LightYellow
        'frmPriceVolumeSet.gvDetail.Columns(14).DefaultCellStyle.BackColor = Color.LightYellow
        'frmPriceVolumeSet.gvDetail.Columns(15).DefaultCellStyle.BackColor = Color.LightYellow
        'frmPriceVolumeSet.gvDetail.Columns(16).DefaultCellStyle.BackColor = Color.LightYellow
        'frmPriceVolumeSet.gvDetail.Columns(17).DefaultCellStyle.BackColor = Color.LightYellow


    End Sub

    Private Sub btnImport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnImport.Click
        Dim dtbl As New DataTable("Ptbl")
        Dim dr As DataRow
        Dim x, n As Integer
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
        For x = 0 To Me.LVps.Items.Count - 1
            If Me.LVps.Items(x).Checked = True Then
                vPSdocno = Me.LVps.Items(x).SubItems(0).Text
                QryString = "exec dbo.USP_PS_PriceVolumeSetPrepare '" & "','" & "','" & "','" & "','" & vPSdocno & "'"
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
                        Dim vLPPrice As Double
                        vLPPrice = CDbl(dt.Rows(i).Item("Price_LP"))
                        dr("ราคา LP") = Format(vLPPrice, "##,##0.00")
                        'Price1
                        vPriceLV1 = CDbl(dt.Rows(i).Item("Price1"))
                        dr("ราคาที่1") = Format(vPriceLV1, "##,##0.00")
                        'เช็คสินค้าวัสดุก่อสร้าง
                        vFMcode = dt.Rows(i).Item("Familycode")
                        If vFMcode = 10000 Then
                            vSLPlv1 = 0.0
                            vPsmtpoint = 0.0
                            vPsmtpoint2 = 0.0
                        Else
                            vPsmtpoint = CDbl(pvSMP2)
                            vPsmtpoint2 = CDbl(pvSMP3)
                            vSLPlv1 = frmPriceVolumeSet.smpLV1.Value
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
                        ' Dim vPCDC01 As Double

                        If ivPDC2 = lpvDC1 Then
                            va = ((ivPDC2 * 100) / vPriceLV1)
                            vPCDC01 = Format(va, "##,##0.00") '---------%ส่วนลด2
                            dr("%ส่วนลดจากราคาที่ 1") = vPCDC01
                            If vLPs = 1 Then
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
                            '---------เลือกรายการลดจากราคา LP หรือ จากราคาที่ 1
                            If vLPs = 1 Then
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
                        dr("%SmartPoint") = Format(vPsmtpoint, "##,##0.0000")
                            'ราคา1*%smartpoint/100
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
                        'ราคาโรงงาน
                        Dim vLPPrice2 As Double
                        vLPPrice2 = CDbl(dt.Rows(i).Item("Price_LP"))
                        dr("ราคา LP") = Format(vLPPrice2, "##,##0.00")
                            'Price1
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
                        If ivPDC3 = lpvDC2 Then
                            vx = ((ivPDC3 * 100) / vPriceLV2)
                            vPCDC02 = Format(vx, "##,##0.00") '--------%ส่วนลด3
                            dr("%ส่วนลดจากราคาที่ 1") = vPCDC02
                            If vLPs = 1 Then
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
                            If vLPs = 1 Then
                                nPrice3 = vLPPrice2 - ((vLPPrice2 * vPCDC02) / 100)
                            Else
                                nPrice3 = vPriceLV2 - ((vPriceLV2 * vPCDC02) / 100)
                            End If

                            'ราคา1*ส่วนลด%ราคา1/100
                            vnPrice3 = chkDCM(nPrice3)
                            dr("ราคา") = Format(vnPrice3, "##,##0.00")
                        End If
                        dr("%SmartPoint") = Format(vPsmtpoint2, "##,##0.0000")
                        'ราคา1*%smartpoint/100
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
                        frmPriceVolumeSet.gvDetail.BackgroundColor = Color.Cyan
                    Next
                End If
            End If
        Next
        frmPriceVolumeSet.gvDetail.DataSource = dtbl
        For n = 0 To frmPriceVolumeSet.gvDetail.Rows.Count - 1
            If n Mod 2 = 0 Then
                frmPriceVolumeSet.gvDetail.Rows(n).DefaultCellStyle.BackColor = Color.SkyBlue
            End If
        Next
        Call gvformat()
        frmPriceVolumeSet.Enabled = True
        Call gvReadonly1()
        sv = "1"
        frmPriceVolumeSet.Text = "[::] กำหนดราคาตามจำนวน"
        frmPriceVolumeSet.btnProduct.Enabled = False
        frmPriceVolumeSet.btnPrint.Enabled = False
        frmPriceVolumeSet.btnSave.Enabled = True
        frmPriceVolumeSet.GroupBox1.Enabled = False
        frmPriceVolumeSet.Enabled = True
        frmPriceVolumeSet.Enabled = False
        frmPriceVolumeSet.LPcbx.Enabled = False
        frmPriceVolumeSet.gvDetail.AllowUserToDeleteRows = True
        Me.Close()
    End Sub
    Private Function CHKdPrice(ByVal vD As String) As String
        Dim vx1 As String
        'Dim vx2 As String
        Dim vx3 As Double
        Dim i As Integer
        i = Microsoft.VisualBasic.Len(vD)
        'vx1 = vD
        i = InStr(vD, "%")

        If i = 0 Then
            vD = (vPriceLV1 - ((vPriceLV1 * vPCDC01) / 100))
        Else
            vx1 = Microsoft.VisualBasic.Left(vD, i - 1)
            ' CDbl(vx1)
            vx3 = (vPriceLV1 - ((vPriceLV1 * vx1) / 100))
            vD = vx3
        End If
        Return vD
    End Function
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
    Public Function CalcAmount(ByVal vPriceSetAmount As Double, ByVal vPercent As Double) As Double
        CalcAmount = (vPriceSetAmount * vPercent) / 100
    End Function

    Private Sub txtPSDoc_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtPSDoc.KeyDown
        If e.KeyCode = Keys.Enter Then
            vSearch = Me.txtPSDoc.Text
            Call GetSearchLV()
        End If
    End Sub


    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        frmPriceVolumeSet.Enabled = True
        Me.Close()
    End Sub

    Private Sub cbxAll_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbxAll.CheckedChanged
        Dim xi As Integer
        If Me.LVps.Items.Count > 0 Then
            If Me.cbxAll.Checked = True Then
                For xi = 0 To Me.LVps.Items.Count - 1
                    Me.LVps.Items(xi).Checked = True
                Next
            Else
                For xi = 0 To Me.LVps.Items.Count - 1
                    Me.LVps.Items(xi).Checked = False
                Next
            End If

        End If
    End Sub

   
End Class
