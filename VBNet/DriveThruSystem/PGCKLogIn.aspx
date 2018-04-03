<%@ Page Language="VB" AutoEventWireup="false" CodeFile="PGCKLogIn.aspx.vb" Inherits="PGCKLogIn" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Untitled Page</title>
</head>
<body topmargin ="0px" leftmargin = "0px">
    <form id="form1" runat="server">
    <div>
        <table style="width: 166px; height: 71px">
            <tr>
                <td colspan="3" style="height: 15px; background-color: #000099">
                </td>
            </tr>
            <tr>
                <td colspan="3">
                    <asp:Label ID="Label3" runat="server" Font-Size="Large" Style="text-align: center"
                        Text="บริษัท นพดลพานิช จำกัด" Width="225px"></asp:Label></td>
            </tr>
            <tr>
                <td colspan="3" style="height: 15px; background-color: #000099">
                </td>
            </tr>
            <tr>
                <td style="width: 95px">
                </td>
                <td style="width: 117px">
                </td>
                <td style="width: 53px">
                </td>
            </tr>
            <tr>
                <td colspan="3">
                    <asp:Label ID="Label1" runat="server" Font-Size="7pt" Text="Label" Width="225px"></asp:Label></td>
            </tr>
            <tr>
                <td style="width: 95px; height: 21px">
                    <asp:Label ID="Label2" runat="server" Font-Bold="False" Font-Size="8pt" Height="18px"
                        Style="text-align: right" Text="รหัสผู้ใช้งาน :" Width="90px"></asp:Label></td>
                <td colspan="2" style="height: 21px">
                    <asp:TextBox ID="TBUserID" runat="server" Font-Bold="True" Font-Size="6pt" onkeypress="onEnter0();"
                        Width="121px"></asp:TextBox></td>
            </tr>
            <tr>
                <td style="width: 95px; height: 15px">
                    <asp:Label ID="Label4" runat="server" Font-Bold="False" Font-Size="8pt" Height="18px"
                        Style="text-align: right" Text="รหัสผ่าน :" Width="90px"></asp:Label></td>
                <td colspan="2" style="height: 15px">
                    <asp:TextBox ID="TBPassword" runat="server" Font-Bold="True" Font-Size="6pt" onkeypress="onEnter0();"
                        TextMode="Password" Width="121px"></asp:TextBox></td>
            </tr>
            <tr>
                <td style="width: 95px; height: 15px">
                </td>
                <td colspan="2" style="height: 15px">
                    <asp:Button ID="Button1" runat="server" Font-Bold="True" Font-Size="7pt" Text="LogIn" /></td>
            </tr>
            <tr>
                <td colspan="3" style="height: 15px; background-color: #ff6600">
                </td>
            </tr>
            <tr>
                <td colspan="3" style="height: 15px; background-color: #ff6600">
                </td>
            </tr>
            <tr>
                <td colspan="3" style="height: 15px; background-color: #ff6600">
                </td>
            </tr>
            <tr>
                <td colspan="3" style="height: 15px; background-color: #ff6600">
                </td>
            </tr>
            <tr>
                <td style="width: 95px; height: 21px">
                </td>
                <td style="width: 117px; height: 21px">
                </td>
                <td style="width: 53px; height: 21px">
                </td>
            </tr>
        </table>
    
    </div>
    </form>
</body>
</html>
