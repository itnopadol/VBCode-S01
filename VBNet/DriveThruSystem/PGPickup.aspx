<%@ Page Language="VB" AutoEventWireup="false" CodeFile="PGPickup.aspx.vb" Inherits="PGPickup" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Untitled Page</title>
    <link href="/aspnet_client/System_Web/2_0_50727/CrystalReportWebFormViewer3/css/default.css"
        rel="stylesheet" type="text/css" />
    <link href="/aspnet_client/System_Web/2_0_50727/CrystalReportWebFormViewer3/css/default.css"
        rel="stylesheet" type="text/css" />
    <link href="/aspnet_client/System_Web/2_0_50727/CrystalReportWebFormViewer3/css/default.css"
        rel="stylesheet" type="text/css" />
    <link href="/aspnet_client/System_Web/2_0_50727/CrystalReportWebFormViewer3/css/default.css"
        rel="stylesheet" type="text/css" />
    <link href="/aspnet_client/System_Web/2_0_50727/CrystalReportWebFormViewer3/css/default.css"
        rel="stylesheet" type="text/css" />
    <link href="/aspnet_client/System_Web/2_0_50727/CrystalReportWebFormViewer3/css/default.css"
        rel="stylesheet" type="text/css" />
    <link href="/aspnet_client/System_Web/2_0_50727/CrystalReportWebFormViewer3/css/default.css"
        rel="stylesheet" type="text/css" />
    <link href="/aspnet_client/System_Web/2_0_50727/CrystalReportWebFormViewer3/css/default.css"
        rel="stylesheet" type="text/css" />
    <link href="/aspnet_client/System_Web/2_0_50727/CrystalReportWebFormViewer3/css/default.css"
        rel="stylesheet" type="text/css" />
    <link href="/aspnet_client/System_Web/2_0_50727/CrystalReportWebFormViewer3/css/default.css"
        rel="stylesheet" type="text/css" />
    <link href="/aspnet_client/System_Web/2_0_50727/CrystalReportWebFormViewer3/css/default.css"
        rel="stylesheet" type="text/css" />
    <link href="/aspnet_client/System_Web/2_0_50727/CrystalReportWebFormViewer3/css/default.css"
        rel="stylesheet" type="text/css" />
</head>



<script language="javascript" type ="text/javascript">   

</script> 




<body topmargin ="0px" leftmargin = "0px">
    <form id="form1" runat="server">
    <div>
        <table style="width: 200px">
            <tr>
                <td style="width: 1px; text-align: right; border-right: #000099 1px solid; border-top: #000099 1px solid; margin: 1px; border-left: #000099 1px solid; border-bottom: #000099 1px solid; height: 7px; background-color: #000099;">
                    <asp:Label ID="Label2" runat="server" Font-Size="6.5pt" Text="ทะเบียนรถ :" Width="66px" Height="3px" ForeColor="White"></asp:Label></td>
                <td style="width: 22px; border-right: #000099 1px solid; border-top: #000099 1px solid; margin: 1px; border-left: #000099 1px solid; border-bottom: #000099 1px solid; height: 7px;">
                    <asp:Label ID="LBLLicense" runat="server" Font-Size="7pt" Width="141px" Height="1px"></asp:Label></td>
            </tr>
            <tr>
                <td style="width: 1px; text-align: right; border-right: #000099 1px solid; border-top: #000099 1px solid; margin: 1px; border-left: #000099 1px solid; border-bottom: #000099 1px solid; background-color: #000099;">
                    <asp:Label ID="Label3" runat="server" Font-Size="6.5pt" Text="ลูกค้า :" Width="64px" Height="1px" ForeColor="White"></asp:Label></td>
                <td style="width: 22px; border-right: #000099 1px solid; border-top: #000099 1px solid; margin: 1px; border-left: #000099 1px solid; border-bottom: #000099 1px solid;">
                    <asp:Label ID="LBLArCode" runat="server" Font-Size="7pt" Text="1/เงินสด" Width="140px"></asp:Label></td>
            </tr>
            <tr>
                <td style="width: 1px; text-align: right; border-right: #000099 1px solid; border-top: #000099 1px solid; margin: 1px; border-left: #000099 1px solid; border-bottom: #000099 1px solid; background-color: #000099;">
                    <asp:Label ID="Label4" runat="server" Font-Size="6.5pt" Text="พนง.ขาย :" Width="63px" Height="1px" ForeColor="White"></asp:Label></td>
                <td style="width: 22px; border-right: #000099 1px solid; border-top: #000099 1px solid; margin: 1px; border-left: #000099 1px solid; border-bottom: #000099 1px solid;">
                    <asp:Label ID="LBLSaleCode" runat="server" Font-Size="7pt" Width="140px"></asp:Label></td>
            </tr>
            <tr>
                <td style="width: 1px; text-align: right; border-right: #000099 1px solid; border-top: #000099 1px solid; margin: 1px; border-left: #000099 1px solid; border-bottom: #000099 1px solid; height: 2px; background-color: #000099;">
                    <asp:Label ID="Label1" runat="server" Font-Size="6.5pt" Text="บาร์โค้ด :" Width="65px" Height="1px" ForeColor="White"></asp:Label></td>
                <td style="width: 22px; height: 2px;">
                    <asp:TextBox ID="TBBarCode" runat="server" Font-Size="6.5pt" Width="49px" BorderWidth="1px"></asp:TextBox>&nbsp;
                </td>
            </tr>
            <tr>
                <td style="border-right: #000099 1px solid; border-top: #000099 1px solid; margin: 1px;
                    border-left: #000099 1px solid; border-bottom: #000099 1px solid;
                    background-color: #ff9933; text-align: left; vertical-align: top; height: 2px;" colspan="2">
                    <asp:Label ID="LBLMessage" runat="server" Font-Size="7pt" Style="vertical-align: top;
                        text-align: center" Width="214px" Height="17px"></asp:Label></td>
            </tr>
            <tr>
                <td colspan="2" style="border-right: #000099 1px solid; border-top: #000099 1px solid; margin: 1px; border-left: #000099 1px solid; border-bottom: #000099 1px solid">
                    <asp:GridView ID="GridView2" runat="server" Width="215px" Font-Size="6.5pt">
                        <Columns>
                            <asp:CommandField HeaderText="Del" ShowDeleteButton="True" DeleteText="Del" />
                        </Columns>
                        <HeaderStyle BackColor="Navy" ForeColor="White" />
                        <AlternatingRowStyle BackColor="#FF8000" />
                    </asp:GridView><asp:GridView ID="GridView1" runat="server" Width="215px" Font-Size="6.5pt" Visible="False">
                    </asp:GridView>
                    <asp:Label ID="LBLDocNo" runat="server" Font-Size="7pt" Height="5px" Width="1px"></asp:Label>
                    <asp:Label ID="LBLUserID" runat="server" Font-Size="6.5pt" Width="65px" Visible="False"></asp:Label>
                    <asp:Label ID="LBLPointID" runat="server" Font-Size="6.5pt" Width="65px" Visible="False"></asp:Label></td>
            </tr>
            <tr>
                <td style="border-right: #000099 1px solid; border-top: #000099 1px solid; margin: 1px;
                    border-left: #000099 1px solid; width: 1px; border-bottom: #000099 1px solid; background-color: #ff9933">
                    <asp:Label ID="Label5" runat="server" Font-Size="6.5pt" Height="1px" Style="text-align: right"
                        Text="มูลค่าสินค้า :" Width="65px"></asp:Label></td>
                <td style="border-right: #000099 1px solid; border-top: #000099 1px solid; margin: 1px;
                    border-left: #000099 1px solid; width: 22px; border-bottom: #000099 1px solid; background-color: #ff9933; text-align: right">
                    <asp:Label ID="LBLNetAmount" runat="server" Font-Size="7pt" Width="142px" style="text-align: right"></asp:Label></td>
            </tr>
            <tr>
                <td style="width: 1px; vertical-align: top; text-align: left;">
                    <asp:Button ID="BTNPrevoius" runat="server" Font-Bold="True" Font-Size="8pt" Text="l<<" /></td>
                <td style="width: 22px; text-align: right; vertical-align: top;">
                    <asp:Button ID="BTNSave" runat="server" Font-Bold="True" Font-Size="8pt" Text="บันทึก" />
                </td>
            </tr>
            <tr>
                <td style="width: 1px">
                </td>
                <td style="width: 22px; text-align: right">
                    <asp:LinkButton ID="LinkButton2" runat="server" Visible="False">LinkButton</asp:LinkButton></td>
            </tr>
        </table>
    
    </div>
    </form>
</body>
</html>
