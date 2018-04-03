Imports System.Windows.Forms
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports System.Globalization
Imports Microsoft.VisualBasic


Public Class dlgVolumeSearch
    Dim QryString As String
    Dim FNdocno As String
    Dim da As SqlDataAdapter
    Dim ds As DataSet
    Dim iLv As New ListViewItem
    Dim i, n As Integer
    Dim xsc As String

    Private Sub dlgVolumeSearch_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        frmPriceVolumeSet.Enabled = True
    End Sub

    Private Sub dlgVolumeSearch_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call InitializeDataBase()
        'Me.TxtFind.Focus()
        'publicFdocno = frmPriceVolumeSet.txtDocno.Text
        'Call getHeadDoc()

    End Sub
    Private Sub getHeadDoc()
        Dim cfStatus As String
        Me.LVsearch.Items.Clear()
        QryString = "exec dbo.USP_PS_PriceVolumeSetSearch '" & publicFdocno & "'"
        da = New SqlDataAdapter(QryString, vConnectionString)
        ds = New DataSet
        da.Fill(ds, "TBsch")
        dt = ds.Tables("TBsch")
        If dt.Rows.Count <> 0 Then
            For i = 0 To dt.Rows.Count - 1
                'ListView = Me.LvProduct.Items.Add(dt.Rows(n).Item("itemcode"))
                'iListView.SubItems.Add(0).Text = dt.Rows(n).Item("itemname")
                iLv = Me.LVsearch.Items.Add(dt.Rows(i).Item("DocNo"))
                iLv.SubItems.Add(0).Text = dt.Rows(i).Item("DocDate")
                iLv.SubItems.Add(0).Text = dt.Rows(i).Item("PsDocno")
                iLv.SubItems.Add(0).Text = dt.Rows(i).Item("begindate")
                iLv.SubItems.Add(0).Text = dt.Rows(i).Item("enddate")
                cfStatus = dt.Rows(i).Item("isconfirm")
                If cfStatus = 0 Then
                    iLv.SubItems.Add(0).Text = "No"
                    iLv.SubItems.Add(0).ForeColor = Color.Red
                Else
                    iLv.SubItems.Add(0).Text = "Yes"
                    iLv.SubItems.Add(0).ForeColor = Color.Green
                End If
            Next
        End If

    End Sub

    Private Sub LVsearch_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles LVsearch.MouseDoubleClick
        Dim tbFN As New DataTable("Ftbl")
        Dim drw As DataRow
        Dim FdocNo As String
        Dim vPriceLevel As Integer
        Dim vLNnum As Integer
        Dim MKsv1 As Double
        Dim GPmk1 As Double
        Dim vPCGPMK1 As Double
        Dim AVsv1 As Double
        Dim GPav1 As Double
        Dim vPCGPAV1 As Double
        'Dim vLPprice As Double

        ' If Me.LVsearch.Items(0).Selected = True Or Me.LVsearch.Items(1).Selected = True Or Me.LVsearch.Items(2).Selected = True Or Me.LVsearch.Items(3).Selected = True Then
        If Me.LVsearch.SelectedItems.Count <> 0 Then
            frmPriceVolumeSet.GroupBox1.Enabled = False
            frmPriceVolumeSet.DTPdocDate.Enabled = False
            frmPriceVolumeSet.smpLV1.Enabled = False
            frmPriceVolumeSet.DTPStartDate.Enabled = False
            frmPriceVolumeSet.DTPEndDate.Enabled = False
            frmPriceVolumeSet.btnProduct.Enabled = False
            FdocNo = Me.LVsearch.SelectedItems(0).Text.ToString
            QryString = "exec dbo.USP_PS_PriceVolumeSetSearchsub '" & FdocNo & "'"
            da = New SqlDataAdapter(QryString, vConnectionString)
            ds = New DataSet
            da.Fill(ds, "Fdocpv")
            dt = ds.Tables("Fdocpv")
            If dt.Rows.Count > 0 Then
                frmPriceVolumeSet.txtDocno.Text = dt.Rows(0).Item("Docno")
                frmPriceVolumeSet.DTPdocDate.Value = dt.Rows(0).Item("DocDate")
                frmPriceVolumeSet.DTPStartDate.Value = dt.Rows(0).Item("begindate")
                frmPriceVolumeSet.DTPEndDate.Value = dt.Rows(0).Item("enddate")
                frmPriceVolumeSet.smpLV1.Value = CInt(dt.Rows(0).Item("SmartPoint1percent"))
                frmPriceVolumeSet.txtVM2.Text = Format(Int(dt.Rows(0).Item("Volume")), "##,##0.00")
                frmPriceVolumeSet.txtDC2.Text = Format(Int(dt.Rows(0).Item("Price1Discount")), "##,##0.00") & "%"
                PublicDocStatus = dt.Rows(0).Item("IsConfirm")
                'header gridview
                tbFN.Columns.Add("No.", GetType(String))
                tbFN.Columns.Add("เลขที่เอกสารโครงสร้างราคา", GetType(String))
                tbFN.Columns.Add("รหัสสินค้า", GetType(String))
                tbFN.Columns.Add("ชื่อสินค้า", GetType(String))
                tbFN.Columns.Add("หน่วยขาย", GetType(String))
                tbFN.Columns.Add("ราคา LP", GetType(String))
                tbFN.Columns.Add("ราคาที่1", GetType(String))
                tbFN.Columns.Add("%smartpoint ราคาที่1", GetType(String))
                tbFN.Columns.Add("SmartPoint1", GetType(String))
                tbFN.Columns.Add("ระดับราคา", GetType(String))
                tbFN.Columns.Add("Volume", GetType(String))
                tbFN.Columns.Add("%ส่วนลดจากราคาที่ 1", GetType(String))
                tbFN.Columns.Add("ราคา", GetType(String))
                tbFN.Columns.Add("%SmartPoint", GetType(String))
                tbFN.Columns.Add("SmartPoint", GetType(String))
                tbFN.Columns.Add("ทุนตลาดSaleVat", GetType(String))
                tbFN.Columns.Add("GP ทุนตลาด", GetType(String))
                tbFN.Columns.Add("%GP ทุนตลาด", GetType(String))
                tbFN.Columns.Add("ทุนเฉลี่ย LotSaleVat", GetType(String))
                tbFN.Columns.Add("GP ทุนเฉลี่ย Lot", GetType(String))
                tbFN.Columns.Add("%GP ทุนเฉลี่ย", GetType(String))
                For i = 0 To dt.Rows.Count - 1
                    drw = tbFN.NewRow
                    vLNnum = dt.Rows(i).Item("Linenumber")
                    drw("No.") = Format(vLNnum, "00#")
                    drw("เลขที่เอกสารโครงสร้างราคา") = dt.Rows(i).Item("PSDocNo")
                    drw("รหัสสินค้า") = dt.Rows(i).Item("Itemcode")
                    drw("ชื่อสินค้า") = dt.Rows(i).Item("Itemname")
                    drw("หน่วยขาย") = dt.Rows(i).Item("Unitcode")
                    'vLPprice = dt.Rows(i).Item("Price_LP")
                    If dt.Rows(i).Item("Price_LP") IsNot DBNull.Value Then
                        drw("ราคา LP") = Format(Int(dt.Rows(i).Item("Price_LP")), "##,##0.00")
                    Else
                        drw("ราคา LP") = 0.0
                    End If
                    drw("ราคาที่1") = Format(Int(dt.Rows(i).Item("Price1")), "##,##0.00")
                    drw("%smartpoint ราคาที่1") = dt.Rows(i).Item("SmartPoint1percent")
                    drw("SmartPoint1") = dt.Rows(i).Item("SmartPoint1")
                    vPriceLevel = dt.Rows(i).Item("Pricelevel")
                    drw("ระดับราคา") = Format(Int(vPriceLevel), "##,##0") '-----ระดับราคา
                    drw("Volume") = Format(Int(dt.Rows(i).Item("Volume")), "##,##0.00")
                    drw("%ส่วนลดจากราคาที่ 1") = Format(dt.Rows(i).Item("Price1Discount"), "##0.00")
                    drw("ราคา") = Format(Int(dt.Rows(i).Item("Priceset")), "##,##0.00")
                    drw("%SmartPoint") = dt.Rows(i).Item("SmartpointPercent")
                    drw("SmartPoint") = dt.Rows(i).Item("SmartPoint")
                    MKsv1 = dt.Rows(i).Item("Marketcost") 'ทุนตลาด
                    drw("ทุนตลาดSaleVat") = Format(Int(MKsv1), "##,##0.00")
                    GPmk1 = dt.Rows(i).Item("MarketcostGP") 'GP ทุนตลาด
                    drw("GP ทุนตลาด") = Format(Int(GPmk1), "##,##0.00")
                    vPCGPMK1 = ((GPmk1 * 100) / MKsv1) '% GP ทุนตลาด
                    drw("%GP ทุนตลาด") = Format(Int(vPCGPMK1), "##,##0.00")
                    AVsv1 = dt.Rows(i).Item("LotAverageCost")
                    drw("ทุนเฉลี่ย LotSaleVat") = Format(Int(AVsv1), "##,##0.00")
                    GPav1 = dt.Rows(i).Item("LotAverageCostGP")
                    drw("GP ทุนเฉลี่ย Lot") = Format(Int(GPav1), "##,##0.00")
                    vPCGPAV1 = ((GPav1 * 100) / AVsv1)
                    drw("%GP ทุนเฉลี่ย") = Format(Int(vPCGPAV1), "##,##0.00")
                    tbFN.Rows.Add(drw)
                    frmPriceVolumeSet.gvDetail.DataSource = tbFN
                    'frmPriceVolumeSet.gvDetail.SortOrde
                    For n = 0 To frmPriceVolumeSet.gvDetail.Rows.Count - 1
                        If n Mod 2 = 0 Then
                            frmPriceVolumeSet.gvDetail.Rows(n).DefaultCellStyle.BackColor = Color.SkyBlue
                        End If
                    Next
                    If frmPriceVolumeSet.gvDetail.Rows.Count > 1 Then
                        'Dim vx As Integer
                        ' vx = frmPriceVolumeSet.gvDetail.Rows(1).Cells(8).Value
                        frmPriceVolumeSet.txtVM3.Text = frmPriceVolumeSet.gvDetail.Rows(1).Cells(10).Value
                        frmPriceVolumeSet.txtDC3.Text = frmPriceVolumeSet.gvDetail.Rows(1).Cells(11).Value & "%"
                    Else
                        frmPriceVolumeSet.txtVM3.Text = "0.00"
                        frmPriceVolumeSet.txtDC3.Text = "0.00"
                    End If
                Next
            End If
            If PublicDocStatus = 0 Then
                frmPriceVolumeSet.PBNew.Visible = True
                frmPriceVolumeSet.PBConfirm.Visible = False
                frmPriceVolumeSet.btnSave.Enabled = True
                frmPriceVolumeSet.btnProduct.Enabled = False
                frmPriceVolumeSet.btnSaveAS.Enabled = False
                Call gvReadonly1()
                frmPriceVolumeSet.gvDetail.AllowUserToDeleteRows = True
                frmPriceVolumeSet.LPcbx.Enabled = True
                sv = 0

            ElseIf PublicDocStatus = 1 Then
                frmPriceVolumeSet.PBNew.Visible = False
                frmPriceVolumeSet.PBConfirm.Visible = True
                frmPriceVolumeSet.btnSave.Enabled = False
                frmPriceVolumeSet.btnProduct.Enabled = False
                frmPriceVolumeSet.btnSaveAS.Enabled = True
                frmPriceVolumeSet.btnSaveAS.Enabled = True
                frmPriceVolumeSet.LPcbx.Enabled = False
                Call gvReadonly()
                sv = 0
            Else
                frmPriceVolumeSet.PBNew.Visible = True
                frmPriceVolumeSet.PBConfirm.Visible = False
                frmPriceVolumeSet.LPcbx.Enabled = False
            End If
            Call gvformat()
            frmPriceVolumeSet.btnPrint.Enabled = True
        Else
            MsgBox("คุณไม่ได้เลือกข้อมูล", MsgBoxStyle.Critical, "Error")
        End If
        frmPriceVolumeSet.Enabled = True
        frmPriceVolumeSet.okFind.Enabled = False
        frmPriceVolumeSet.GroupBox1.Enabled = False
        Me.Close()
        'xsc = Me.Close()
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
        frmPriceVolumeSet.gvDetail.Columns(18).Width = 70
        frmPriceVolumeSet.gvDetail.Columns(19).Width = 70
        frmPriceVolumeSet.gvDetail.Columns(20).Width = 70
        frmPriceVolumeSet.gvDetail.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
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
        frmPriceVolumeSet.gvDetail.Columns(18).ReadOnly = True
        frmPriceVolumeSet.gvDetail.Columns(19).ReadOnly = True
        frmPriceVolumeSet.gvDetail.Columns(20).ReadOnly = True

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
        frmPriceVolumeSet.gvDetail.Columns(7).ReadOnly = False
        frmPriceVolumeSet.gvDetail.Columns(8).ReadOnly = True
        frmPriceVolumeSet.gvDetail.Columns(9).ReadOnly = True
        frmPriceVolumeSet.gvDetail.Columns(10).ToolTipText = "คลิกเพื่อแก้ไข Volume รายการนี้"
        frmPriceVolumeSet.gvDetail.Columns(10).DefaultCellStyle.ForeColor = Color.Blue
        frmPriceVolumeSet.gvDetail.Columns(11).ToolTipText = "คลิกเพื่อแก้ไขส่วนลดรายการนี้"
        frmPriceVolumeSet.gvDetail.Columns(11).DefaultCellStyle.ForeColor = Color.Blue
        frmPriceVolumeSet.gvDetail.Columns(12).ReadOnly = False
        frmPriceVolumeSet.gvDetail.Columns(12).DefaultCellStyle.ForeColor = Color.Blue
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
        'frmPriceVolumeSet.gvDetail.Columns(5).DefaultCellStyle.BackColor = Color.LightSalmon
        'frmPriceVolumeSet.gvDetail.Columns(7).DefaultCellStyle.BackColor = Color.LightYellow
        'frmPriceVolumeSet.gvDetail.Columns(8).DefaultCellStyle.BackColor = Color.LightYellow
        'frmPriceVolumeSet.gvDetail.Columns(11).DefaultCellStyle.BackColor = Color.LightSalmon
        'frmPriceVolumeSet.gvDetail.Columns(13).DefaultCellStyle.BackColor = Color.LightYellow
        'frmPriceVolumeSet.gvDetail.Columns(14).DefaultCellStyle.BackColor = Color.LightYellow
        'frmPriceVolumeSet.gvDetail.Columns(15).DefaultCellStyle.BackColor = Color.LightYellow
        'frmPriceVolumeSet.gvDetail.Columns(16).DefaultCellStyle.BackColor = Color.LightYellow
        'frmPriceVolumeSet.gvDetail.Columns(17).DefaultCellStyle.BackColor = Color.LightYellow

    End Sub

    Private Sub TxtFind_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtFind.KeyDown
        If e.KeyCode = Keys.Enter Then
            publicFdocno = Me.TxtFind.Text
            Call getHeadDoc()
        End If
    End Sub

    Private Sub btnSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearch.Click
        publicFdocno = Me.TxtFind.Text
        Call getHeadDoc()
    End Sub
End Class
