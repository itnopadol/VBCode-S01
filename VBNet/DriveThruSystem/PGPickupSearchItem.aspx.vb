Imports System.Data
Imports System.Data.SqlClient
Imports System.Net
Imports System.Web
Imports System.Web.SessionState

Partial Class PGPickupSearchItem
    Inherits System.Web.UI.Page

    Dim vConnectionString As String
    Dim vConnection As SqlConnection
    Dim vCommand As SqlCommand
    Dim da As SqlDataAdapter
    Dim ds As DataSet
    Dim dt As New DataTable
    Dim dv As DataView
    Dim vQuery As String
    Dim vdtSession As New DataTable

    Dim vRecBarCode As String
    Dim vRecLicense As String
    Dim vRecSaleCode As String
    Dim vRecNetAmount As Double
    Dim vRecPointID As Integer


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
  
        If Request.QueryString.Count > 0 Then
            Dim vScript As String = "<SCRIPT language='javascript'>form1.TBQTY.focus();</SCRIPT>"
            Page.RegisterStartupScript("focus", vScript)

            vRecBarCode = Request.QueryString("BarCodeID").ToString()
            vRecLicense = Request.QueryString("License").ToString()
            vRecSaleCode = Request.QueryString("SaleCode").ToString()
            vRecNetAmount = Request.QueryString("NetAmount").ToString()
            vRecPointID = Request.QueryString("PointID").ToString()

            Me.LBLBar.Text = vRecBarCode
            Me.LBLLicense.Text = vRecLicense
            Me.LBLSaleCode.Text = vRecSaleCode
            Me.LBLPointID.Text = vRecPointID
            Me.LBLNetAmount.Text = vRecNetAmount

            vQuery = "exec bcnp.dbo.usp_MB_SearchBarcode '" & vRecBarCode & "' "
            Dim vService As New WebReference.WebServiceCalc
            Dim ds As DataSet = vService.vGetQueryAnlyzer(vQuery)
            If ds.Tables(0).Rows.Count > 0 Then
                Me.LBLItem.Text = ds.Tables(0).Rows(0)("itemcode").ToString
                Me.LBLItemName.Text = ds.Tables(0).Rows(0)("itemname").ToString
                Me.LBLRemain.Text = Format(Int(ds.Tables(0).Rows(0)("stock").ToString), "##,##0")
                Me.LBLUnit.Text = ds.Tables(0).Rows(0)("unitcode").ToString
                Me.LBLPrice.Text = Format(Int(ds.Tables(0).Rows(0)("price").ToString), "##,##0")
            End If
        End If
    End Sub

    Protected Sub TBQTY_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TBQTY.TextChanged
        Dim vQTY As Double
        Dim vItemCode As String
        Dim vItemName As String
        Dim vUnitCode As String
        Dim vOnHand As Double
        Dim vLicense As String
        Dim vSaleCode As String
        Dim vPointID As String
        Dim vPrice As Double
        Dim vNetAmount As Double
        Dim vItemAmount As Double
        Dim vCalcNetAmount As Double

        Dim i As Integer
        Dim dr As DataRow
        Dim vCount As Integer

        On Error Resume Next

        If Me.TBQTY.Text <> "" Then

            vQTY = Me.TBQTY.Text
            vItemCode = Me.LBLItem.Text
            vItemName = Me.LBLItemName.Text
            vUnitCode = Me.LBLUnit.Text
            vOnHand = Me.LBLRemain.Text
            vLicense = Me.LBLLicense.Text
            vSaleCode = Me.LBLSaleCode.Text
            vPointID = Me.LBLPointID.Text
            vPrice = Me.LBLPrice.Text
            vNetAmount = Me.LBLNetAmount.Text

            vItemAmount = vPrice * vQTY
            vCalcNetAmount = vNetAmount + vItemAmount

            Response.Redirect("PGPickup.aspx?License=" & vLicense & "&SaleCode=" & vSaleCode & "&PointID=" & vPointID & "&QTY=" & vQTY & "&BarCodeID=" & vItemCode & "&ItemName=" & vItemName & "&UnitCode=" & vUnitCode & "&OnHand=" & vOnHand & "&Price=" & vPrice & "&NetAmount=" & vCalcNetAmount)

        End If
    End Sub

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BTNPrevoius.Click
        Dim vQTY As Double
        Dim vItemCode As String
        Dim vItemName As String
        Dim vUnitCode As String
        Dim vOnHand As Double
        Dim vLicense As String
        Dim vSaleCode As String
        Dim vPointID As String
        Dim vPrice As Double
        Dim vNetAmount As Double
        Dim vItemAmount As Double
        Dim vCalcNetAmount As Double

        Dim i As Integer
        Dim dr As DataRow
        Dim vCount As Integer

        On Error Resume Next

        vQTY = 0
        vItemCode = ""
        vItemName = ""
        vUnitCode = ""
        vOnHand = 0
        vLicense = Me.LBLLicense.Text
        vSaleCode = Me.LBLSaleCode.Text
        vPointID = Me.LBLPointID.Text
        vPrice = 0
        vNetAmount = Me.LBLNetAmount.Text

        vCalcNetAmount = vNetAmount

        Response.Redirect("PGPickup.aspx?License=" & vLicense & "&SaleCode=" & vSaleCode & "&PointID=" & vPointID & "&QTY=" & vQTY & "&BarCodeID=" & vItemCode & "&ItemName=" & vItemName & "&UnitCode=" & vUnitCode & "&OnHand=" & vOnHand & "&Price=" & vPrice & "&NetAmount=" & vCalcNetAmount)


    End Sub
End Class
