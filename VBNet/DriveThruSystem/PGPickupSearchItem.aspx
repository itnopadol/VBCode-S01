<%@ Page Language="VB" AutoEventWireup="false" CodeFile="PGPickupSearchItem.aspx.vb" Inherits="PGPickupSearchItem" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Untitled Page</title>
</head>
<body topmargin="0px" leftmargin="0px">
    <form id="form1" runat="server">
    <div>
        <table style="width: 230px; border-right: #000099 1px solid; border-top: #000099 1px solid; margin: 1px; border-left: #000099 1px solid; border-bottom: #000099 1px solid;">
            <tr>
                <td colspan="4" style="border-right: #000099 1px solid; border-top: #000099 1px solid;
                    margin: 1px; border-left: #000099 1px solid; border-bottom: #000099 1px solid;
                    height: 3px; background-color: #ff6600">
                    <asp:Label ID="Label8" runat="server" Font-Size="8pt" ForeColor="White" Style="vertical-align: middle;
                        text-align: center" Text="กรอกจำนวนสินค้าที่ต้องการขาย" Width="258px"></asp:Label></td>
            </tr>
            <tr>
                <td style="width: 40px; border-right: #000099 1px solid; border-top: #000099 1px solid; margin: 1px; border-left: #000099 1px solid; border-bottom: #000099 1px solid; background-color: #000099;">
                    <asp:Label ID="Label5" runat="server" Font-Bold="False" Font-Size="8pt" Style="text-align: right"
                        Text="บาร์โค้ด :" Width="60px" ForeColor="White"></asp:Label></td>
                <td colspan="3" style="border-right: #000099 1px solid; border-top: #000099 1px solid; margin: 1px; border-left: #000099 1px solid; border-bottom: #000099 1px solid; background-color: white">
                    <asp:Label ID="LBLBar" runat="server" Font-Bold="False" Font-Size="8pt" Style="text-align: left"
                        Width="190px"></asp:Label></td>
            </tr>
            <tr>
                <td style="width: 40px; border-right: #000099 1px solid; border-top: #000099 1px solid; margin: 1px; border-left: #000099 1px solid; border-bottom: #000099 1px solid; background-color: #000099;">
                    <asp:Label ID="Label1" runat="server" Font-Bold="False" Font-Size="8pt" Style="text-align: right"
                        Text="รหัสสินค้า :" Width="60px" ForeColor="White"></asp:Label></td>
                <td colspan="3" style="border-right: #000099 1px solid; border-top: #000099 1px solid; margin: 1px; border-left: #000099 1px solid; border-bottom: #000099 1px solid; background-color: white">
                    <asp:Label ID="LBLItem" runat="server" Font-Bold="False" Font-Size="8pt" Style="text-align: left"
                        Width="190px"></asp:Label></td>
            </tr>
            <tr>
                <td style="width: 40px; border-right: #000099 1px solid; border-top: #000099 1px solid; margin: 1px; border-left: #000099 1px solid; border-bottom: #000099 1px solid; background-color: #000099;">
                    <asp:Label ID="Label2" runat="server" Font-Bold="False" Font-Size="8pt" Style="text-align: right"
                        Text="ชื่อสินค้า :" Width="60px" ForeColor="White"></asp:Label></td>
                <td colspan="3" style="border-right: #000099 1px solid; border-top: #000099 1px solid; margin: 1px; border-left: #000099 1px solid; border-bottom: #000099 1px solid; background-color: white">
                    <asp:Label ID="LBLItemName" runat="server" Font-Bold="False" Font-Size="8pt" Style="text-align: left"
                        Width="189px"></asp:Label></td>
            </tr>
            <tr>
                <td style="width: 40px; border-right: #000099 1px solid; border-top: #000099 1px solid; margin: 1px; border-left: #000099 1px solid; border-bottom: #000099 1px solid; background-color: #000099;">
                    <asp:Label ID="Label3" runat="server" Font-Bold="False" Font-Size="8pt" Style="text-align: right"
                        Text="ราคา :" Width="60px" ForeColor="White"></asp:Label></td>
                <td style="width: 103px; border-right: #000099 1px solid; border-top: #000099 1px solid; margin: 1px; border-left: #000099 1px solid; border-bottom: #000099 1px solid; background-color: white;">
                    <asp:Label ID="LBLPrice" runat="server" Font-Bold="False" Font-Size="8pt" Style="text-align: left"
                        Width="66px"></asp:Label></td>
                <td style="width: 36px; border-right: #000099 1px solid; border-top: #000099 1px solid; margin: 1px; border-left: #000099 1px solid; border-bottom: #000099 1px solid; background-color: #000099;">
                    <asp:Label ID="Label4" runat="server" Font-Bold="False" Font-Size="8pt" Style="text-align: right"
                        Text="หน่วย :" Width="54px" ForeColor="White"></asp:Label></td>
                <td style="border-right: #000099 1px solid; border-top: #000099 1px solid; margin: 1px; border-left: #000099 1px solid; width: 63px; border-bottom: #000099 1px solid; background-color: white;">
                    <asp:Label ID="LBLUnit" runat="server" Font-Bold="False" Font-Size="8pt" Style="text-align: left"
                        Width="57px"></asp:Label></td>
            </tr>
            <tr>
                <td style="width: 40px; border-right: #000099 1px solid; border-top: #000099 1px solid; margin: 1px; border-left: #000099 1px solid; border-bottom: #000099 1px solid; background-color: #000099;">
                    <asp:Label ID="Label7" runat="server" Font-Bold="False" Font-Size="8pt" Style="text-align: right"
                        Text="คงเหลือ :" Width="60px" ForeColor="White"></asp:Label></td>
                <td style="width: 103px; border-right: #000099 1px solid; border-top: #000099 1px solid; margin: 1px; border-left: #000099 1px solid; border-bottom: #000099 1px solid; background-color: white;">
                    <asp:Label ID="LBLRemain" runat="server" Font-Bold="False" Font-Size="8pt" Style="text-align: right"
                        Width="22px"></asp:Label></td>
                <td style="width: 36px">
                </td>
                <td style="width: 63px">
                    <asp:Label ID="LBLLicense" runat="server" Font-Size="7pt" Width="56px" Visible="False"></asp:Label></td>
            </tr>
            <tr>
                <td style="width: 40px; border-right: #000099 1px solid; border-top: #000099 1px solid; margin: 1px; border-left: #000099 1px solid; border-bottom: #000099 1px solid; background-color: #000099;">
                    <asp:Label ID="Label6" runat="server" Font-Bold="False" Font-Size="8pt" Style="text-align: right"
                        Text="ต้องการ :" Width="60px" ForeColor="White"></asp:Label></td>
                <td style="width: 103px; border-top-width: 1px; border-left-width: 1px; border-left-color: #000099; border-bottom-width: 1px; border-bottom-color: #000099; border-top-color: #000099; border-right-width: 1px; border-right-color: #000099;">
                    <asp:TextBox ID="TBQTY" runat="server" Font-Size="7pt" Width="60px" Height="12px" BackColor="#FFFFC0" BorderStyle="Solid" BorderWidth="1px" ForeColor="Black"></asp:TextBox></td>
                <td style="width: 36px;">
                    <asp:Label ID="LBLPointID" runat="server" Font-Size="7pt" Width="56px" Visible="False"></asp:Label></td>
                <td style="width: 63px;">
                    <asp:Label ID="LBLSaleCode" runat="server" Font-Size="7pt" Width="56px" Visible="False"></asp:Label></td>
            </tr>
            <tr>
                <td colspan="4" style="border-right: #000099 1px solid; border-top: #000099 1px solid; margin: 1px; border-left: #000099 1px solid; border-bottom: #000099 1px solid">
                    <asp:GridView ID="GridView1" runat="server" Font-Size="7pt" Width="259px">
                    </asp:GridView>
                </td>
            </tr>
            <tr>
                <td colspan="4" style="border-right: #000099 1px solid; border-top: #000099 1px solid;
                    margin: 1px; border-left: #000099 1px solid; border-bottom: #000099 1px solid; vertical-align: top; background-color: #ff6600; text-align: left;">
                    <asp:Button ID="BTNPrevoius" runat="server" Font-Bold="True" Font-Size="8pt" Text="l<<" />
                    <asp:Label ID="LBLNetAmount" runat="server" Font-Size="7pt" Width="102px" Font-Bold="True" Visible="False"></asp:Label>
                </td>
            </tr>
        </table>
    
    </div>
    </form>
</body>
</html>
