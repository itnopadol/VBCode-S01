<%@ Page Language="VB" AutoEventWireup="false" CodeFile="LogInForm.aspx.vb" Inherits="_Default" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Untitled Page</title>
</head>
<body  topmargin = "0px" leftmargin = "0px">
    <form id="form1" runat="server">
    <div>
        <table style="width: 1255px; height: 570px;">
            <tr>
                <td style="width: 93px; height: 203px; vertical-align: top; text-align: left;">
                    <table style="width: 124px">
                        <tr>
                            <td>
                                <img src="Picture/Installer.jpg" style="width: 441px; height: 566px" /></td>
                        </tr>
                    </table>
                </td>
                <td style="width: 559px; height: 203px; vertical-align: top; text-align: left;">
                    <table style="width: 335px">
                        <tr>
                            <td style="width: 226px">
                            </td>
                            <td style="width: 3px">
                            </td>
                        </tr>
                        <tr>
                            <td style="vertical-align: top; width: 226px; background-color: #000000; text-align: center">
                                <asp:Label ID="Label1" runat="server" Font-Bold="True" Font-Size="12pt" Font-Underline="True"
                                    ForeColor="White" Text="กรอก ข้อมูลเข้าใช้งานโปรแกรม โอนข้อมูล" Width="309px"></asp:Label></td>
                            <td style="width: 3px">
                            </td>
                        </tr>
                        <tr>
                            <td style="width: 226px; height: 32px; background-color: #ff6600">
                            </td>
                            <td style="width: 3px; height: 32px">
                            </td>
                        </tr>
                        <tr>
                            <td style="vertical-align: top; width: 226px; height: 190px; background-color: #000099;
                                text-align: left">
                                <table style="width: 319px; height: 147px">
                                    <tr>
                                        <td style="vertical-align: top; width: 127px; text-align: right">
                                            <asp:Label ID="Label2" runat="server" Font-Bold="True" Font-Size="10pt" ForeColor="White"
                                                Text="ชื่อ เซอร์เวอร์ :" Width="114px"></asp:Label></td>
                                        <td style="width: 16px">
                                        </td>
                                        <td style="width: 136px">
                                            <asp:TextBox ID="TBServer" runat="server" ReadOnly="True" Width="131px" BackColor="#FFFF80">NEBULA</asp:TextBox></td>
                                        <td>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td style="vertical-align: top; width: 127px; text-align: right">
                                            <asp:Label ID="Label3" runat="server" Font-Bold="True" Font-Size="10pt" ForeColor="White"
                                                Text="ชื่อ ฐานข้อมูล :" Width="111px"></asp:Label></td>
                                        <td style="width: 16px">
                                        </td>
                                        <td style="width: 136px">
                                            <asp:TextBox ID="TBDatabase" runat="server" ReadOnly="True" Width="130px" BackColor="#FFFF80">BCNP</asp:TextBox></td>
                                        <td>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td style="vertical-align: top; width: 127px; text-align: right">
                                            <asp:Label ID="Label4" runat="server" Font-Bold="True" Font-Size="10pt" ForeColor="White"
                                                Text="ชื่อ ผู้ใช้งาน :" Width="117px"></asp:Label></td>
                                        <td style="width: 16px">
                                        </td>
                                        <td style="width: 136px">
                                            <asp:TextBox ID="TBUserID" runat="server" Width="130px"></asp:TextBox></td>
                                        <td>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td style="vertical-align: top; width: 127px; text-align: right">
                                            <asp:Label ID="Label5" runat="server" Font-Bold="True" Font-Size="10pt" ForeColor="White"
                                                Text="รหัสผ่าน :" Width="111px"></asp:Label></td>
                                        <td style="width: 16px">
                                        </td>
                                        <td style="width: 136px">
                                            <asp:TextBox ID="TBPassword" runat="server" Width="130px" TextMode="Password"></asp:TextBox></td>
                                        <td>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td style="width: 127px">
                                        </td>
                                        <td style="width: 16px">
                                        </td>
                                        <td style="vertical-align: middle; width: 136px; text-align: right">
                                            <asp:Button ID="BTNLogIn" runat="server" Font-Bold="True" Font-Size="10pt" Text="ตกลง" /></td>
                                        <td style="vertical-align: middle; text-align: right">
                                        </td>
                                    </tr>
                                </table>
                            </td>
                            <td style="width: 3px; height: 190px">
                            </td>
                        </tr>
                        <tr>
                            <td style="width: 226px; height: 21px">
                            </td>
                            <td style="width: 3px; height: 21px">
                            </td>
                        </tr>
                        <tr>
                            <td style="width: 226px; height: 288px; background-color: #ff6600">
                            </td>
                            <td style="width: 3px; height: 288px">
                            </td>
                        </tr>
                    </table>
                </td>
                <td style="vertical-align: bottom; width: 205px; height: 203px; text-align: right">
                </td>
            </tr>
        </table>
        <br />
    
    </div>
    </form>
</body>
</html>
