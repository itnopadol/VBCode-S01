<%@ Page Language="VB" AutoEventWireup="false" CodeFile="PGCheckOut.aspx.vb" Inherits="PGCheckOut" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Untitled Page</title>
</head>
<body topmargin ="0px" leftmargin = "0px">
    <form id="form1" runat="server">
    <div>
        <table style="width: 224px">
            <tr>
                <td style="width: 60px">
                </td>
                <td style="width: 4px">
                </td>
                <td style="width: 72px">
                </td>
            </tr>
            <tr>
                <td style="width: 60px">
                    <asp:Label ID="Label2" runat="server" Font-Bold="False" Font-Size="8pt" Height="18px"
                        Style="text-align: right" Text="ทะเบียนรถ :" Width="72px"></asp:Label></td>
                <td style="width: 4px">
                </td>
                <td style="width: 72px">
                    <asp:TextBox ID="TBUserID" runat="server" Font-Bold="True" Font-Size="6pt" onkeypress="onEnter0();"
                        Width="121px"></asp:TextBox></td>
            </tr>
            <tr>
                <td style="width: 60px">
                </td>
                <td style="width: 4px">
                </td>
                <td style="width: 72px">
                </td>
            </tr>
            <tr>
                <td colspan="3">
                    <asp:GridView ID="GridView1" runat="server" BackColor="White" Font-Size="7pt" Width="217px">
                        <PagerTemplate>
                            <asp:CheckBox ID="CBSelect" runat="server" />
                        </PagerTemplate>
                        <HeaderStyle BackColor="Navy" ForeColor="White" />
                        <AlternatingRowStyle BackColor="#FF8000" />
                    </asp:GridView>
                </td>
            </tr>
            <tr>
                <td style="width: 60px">
                </td>
                <td style="width: 4px">
                </td>
                <td style="width: 72px">
                </td>
            </tr>
            <tr>
                <td style="width: 60px">
                </td>
                <td style="width: 4px">
                </td>
                <td style="width: 72px">
                </td>
            </tr>
            <tr>
                <td style="width: 60px">
                </td>
                <td style="width: 4px">
                </td>
                <td style="width: 72px">
                </td>
            </tr>
            <tr>
                <td style="width: 60px">
                </td>
                <td style="width: 4px">
                </td>
                <td style="width: 72px">
                </td>
            </tr>
            <tr>
                <td style="width: 60px">
                </td>
                <td style="width: 4px">
                </td>
                <td style="width: 72px">
                </td>
            </tr>
        </table>
    
    </div>
    </form>
</body>
</html>
