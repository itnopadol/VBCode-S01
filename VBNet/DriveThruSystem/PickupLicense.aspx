<%@ Page Language="VB" AutoEventWireup="false" CodeFile="PickupLicense.aspx.vb" Inherits="PickupApp" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Untitled Page</title>
</head>
<body topmargin="0px" leftmargin="0px">
    <form id="form1" runat="server">
    <div>
        <table style="width: 230px">
            <tr>
                <td colspan="4" style="height: 15px; background-color: #ff6600">
                </td>
            </tr>
            <tr>
                <td colspan="4">
                    <asp:Label ID="Label3" runat="server" Font-Size="Large" Style="text-align: center"
                        Text="ºÃÔÉÑ· ¹¾´Å¾Ò¹Ôª ¨Ó¡Ñ´" Width="223px"></asp:Label></td>
            </tr>
            <tr>
                <td colspan="4" style="height: 11px; background-color: #000099">
                </td>
            </tr>
            <tr>
                <td colspan="4" style="height: 11px; background-color: #ff6600">
                    <asp:Label ID="Label2" runat="server" Font-Bold="True" Font-Size="7pt" Width="222px"></asp:Label></td>
            </tr>
            <tr>
                <td style="width: 5px; text-align: right;">
                    <asp:Label ID="Label1" runat="server" Font-Size="7pt" Text="·ÐàºÕÂ¹Ã¶ :" Width="85px"></asp:Label></td>
                <td style="width: 37px">
                    <asp:TextBox ID="TBLicense" runat="server" Font-Size="7pt" Width="53px"></asp:TextBox></td>
                <td style="width: 6px">
                </td>
                <td style="width: 6px">
                </td>
            </tr>
            <tr>
                <td style="width: 5px">
                    <asp:Label ID="LBLUserID" runat="server" Font-Size="7pt" Width="65px" Visible="False"></asp:Label></td>
                <td style="width: 37px">
                    <asp:Label ID="LBLPointID" runat="server" Font-Size="7pt" Width="65px" Visible="False"></asp:Label></td>
                <td style="width: 6px">
                </td>
                <td style="width: 6px">
                </td>
            </tr>
            <tr>
                <td style="width: 5px">
                    <asp:Label ID="LBLSaleCode" runat="server" Font-Size="7pt" Width="65px" Visible="False"></asp:Label></td>
                <td style="width: 37px">
                    <asp:LinkButton ID="LinkButton1" runat="server" Font-Bold="True" Font-Size="8pt">Next</asp:LinkButton></td>
                <td style="width: 6px">
                </td>
                <td style="width: 6px">
                </td>
            </tr>
            <tr>
                <td colspan="4" style="background-color: #000099; height: 15px;">
                    </td>
            </tr>
            <tr>
                <td colspan="4" style="height: 15px; background-color: #000099">
                </td>
            </tr>
            <tr>
                <td colspan="4" style="height: 15px; background-color: #000099">
                </td>
            </tr>
            <tr>
                <td colspan="4" style="height: 15px; background-color: #000099">
                </td>
            </tr>
        </table>
    
    </div>
    </form>
</body>
</html>
