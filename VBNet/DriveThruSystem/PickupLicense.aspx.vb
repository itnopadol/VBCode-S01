Imports System.data
Imports System.Data.SqlClient
Imports System.IO
Imports vb6 = Microsoft.VisualBasic
Imports System.Globalization
Imports System.Windows.Forms
Imports System.Web.SessionState
Imports System.Net
Imports System.Web
Partial Class PickupApp
    Inherits System.Web.UI.Page

    Dim vRecUserID As String
    Dim vRecSaleCode As String
    Dim vRecPointID As String

    Dim vLicense As String
    Dim vQTY As Double
    Dim vItemCode As String
    Dim vItemName As String
    Dim vUnitCode As String
    Dim vOnHand As Double
    Dim vSaleCode As String
    Dim vPointID As String
    Dim vPrice As Double
    Dim vNetAmount As Double

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Request.QueryString.Count > 0 Then
            Dim vScript As String = "<SCRIPT language='javascript'>form1.TBLicense.focus();</SCRIPT>"
            Page.RegisterStartupScript("focus", vScript)

            vRecUserID = Request.QueryString("UserID").ToString()
            vRecSaleCode = Request.QueryString("SaleCode").ToString()
            vRecpointID = Request.QueryString("PointID").ToString()

            Me.LBLUserID.Text = vRecUserID
            Me.LBLPointID.Text = vRecpointID
            Me.LBLSaleCode.Text = vRecSaleCode
            Me.Label2.Text = "ยินดีต้อนรับ คุณ " & "" & vRecSaleCode
        End If
    End Sub

    Protected Sub TBLicense_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TBLicense.TextChanged
        If Me.TBLicense.Text <> "" Then
            vLicense = Me.TBLicense.Text
            vSaleCode = Me.LBLSaleCode.Text
            vPointID = Me.LBLPointID.Text

            Response.Redirect("PGPickup.aspx?License=" & vLicense & "&SaleCode=" & vSaleCode & "&PointID=" & vPointID & "&QTY=" & vQTY & "&BarCodeID=" & vItemCode & "&ItemName=" & vItemName & "&UnitCode=" & vUnitCode & "&OnHand=" & vOnHand & "&Price=" & vPrice & "&NetAmount=" & vNetAmount)
        End If
    End Sub

    Protected Sub LinkButton1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles LinkButton1.Click
        If Me.TBLicense.Text <> "" Then
            vLicense = Me.TBLicense.Text
            vSaleCode = Me.LBLSaleCode.Text
            vPointID = Me.LBLPointID.Text

            Response.Redirect("PGPickup.aspx?License=" & vLicense & "&SaleCode=" & vSaleCode & "&PointID=" & vPointID & "&QTY=" & vQTY & "&BarCodeID=" & vItemCode & "&ItemName=" & vItemName & "&UnitCode=" & vUnitCode & "&OnHand=" & vOnHand & "&Price=" & vPrice & "&NetAmount=" & vNetAmount)
        End If
    End Sub
End Class
