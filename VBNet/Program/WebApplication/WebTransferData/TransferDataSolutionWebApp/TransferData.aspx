<%@ Page Language="VB" AutoEventWireup="false" CodeFile="TransferData.aspx.vb" Inherits="TransferData" %>

<%@ Register Assembly="System.Web.Extensions, Version=1.0.61025.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35"
    Namespace="System.Web.UI" TagPrefix="aspp" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Untitled Page</title>
</head>
<body topmargin = "0px" leftmargin = "0px" style="background-color: #000000">
    <form id="form1" runat="server">
    <div>
        <table style="width: 1243px; height: 320px">
            <tr>
                <td style="vertical-align: top; width: 297px; text-align: right">
                    <asp:TextBox ID="TextBox13" runat="server" BackColor="#FFFF80" Font-Bold="True" ReadOnly="True"
                        Visible="False" Width="120px"></asp:TextBox></td>
                <td style="vertical-align: top; width: 282px; text-align: left">
                </td>
                <td style="width: 440px">
                </td>
            </tr>
            <tr>
                <td style="vertical-align: top; width: 297px; text-align: left">
                </td>
                <td style="vertical-align: top; width: 282px; text-align: left">
                </td>
                <td style="width: 440px">
                </td>
            </tr>
            <tr>
                <td style="vertical-align: top; width: 297px; text-align: left">
                    <asp:Label ID="Label2" runat="server" Font-Bold="True" Style="vertical-align: top;
                        text-align: right; text-decoration: underline;" Text="������ ����ͧ�鹷ҧ" Width="175px" ForeColor="White"></asp:Label></td>
                <td style="vertical-align: top; width: 282px; text-align: left">
                    <asp:Label ID="Label4" runat="server" Font-Bold="True" Text="���͡�������ͧ����͹������"
                        Width="316px" ForeColor="White" style="text-decoration: underline"></asp:Label></td>
                <td style="width: 440px">
                    <asp:Label ID="Label1" runat="server" Font-Bold="True" ForeColor="White" Text="�š�� �͹������"
                        Width="316px" style="text-decoration: underline"></asp:Label></td>
            </tr>
            <tr>
                <td style="vertical-align: top; width: 297px; height: 24px; text-align: left">
                </td>
                <td style="vertical-align: top; width: 282px; height: 24px; text-align: left">
                </td>
                <td style="height: 24px; width: 440px;">
                </td>
            </tr>
            <tr>
                <td style="vertical-align: top; width: 297px; text-align: left; height: 148px; background-color: #000099;">
                    <table style="vertical-align: top; width: 313px; text-align: right">
                        <tr>
                            <td style="vertical-align: top; width: 170px; text-align: right">
                                <asp:Label ID="Label5" runat="server" Text="���� ���������鹷ҧ :" Width="155px" ForeColor="White"></asp:Label></td>
                            <td>
                            </td>
                            <td style="vertical-align: top; text-align: left">
                                <asp:TextBox ID="TextBox1" runat="server" Width="120px" BackColor="#FFFF80" Font-Bold="True" ReadOnly="True"></asp:TextBox></td>
                        </tr>
                        <tr>
                            <td style="width: 170px">
                                <asp:Label ID="Label6" runat="server" Text="���� �ҹ�����ŵ鹷ҧ :" Width="154px" ForeColor="White"></asp:Label></td>
                            <td>
                            </td>
                            <td style="vertical-align: top; text-align: left">
                                <asp:TextBox ID="TextBox2" runat="server" Width="120px" BackColor="#FFFF80" Font-Bold="True" ReadOnly="True"></asp:TextBox></td>
                        </tr>
                        <tr>
                            <td style="width: 170px">
                                <asp:Label ID="Label7" runat="server" Text="���� �����ҹ�鹷ҧ :" Width="154px" ForeColor="White"></asp:Label></td>
                            <td>
                            </td>
                            <td style="vertical-align: top; text-align: left">
                                <asp:TextBox ID="TextBox3" runat="server" Width="120px" BackColor="#FFFF80" Font-Bold="True" ReadOnly="True"></asp:TextBox></td>
                        </tr>
                        <tr>
                            <td style="vertical-align: top; width: 170px; height: 21px; text-align: right">
                                <asp:Label ID="Label8" runat="server" Text="���ʼ�ҹ�鹷ҧ :" Width="154px" ForeColor="White"></asp:Label></td>
                            <td style="height: 21px">
                            </td>
                            <td style="height: 21px; vertical-align: top; text-align: left;">
                                <asp:TextBox ID="TextBox4" runat="server" Width="120px" BackColor="#FFFF80" Font-Bold="True" ForeColor="#FFFF80" ReadOnly="True"></asp:TextBox></td>
                        </tr>
                    </table>
                    <br />
                </td>
                <td style="vertical-align: top; width: 282px; text-align: left; height: 148px; background-color: #ff9900;">
                    <table style="width: 505px; vertical-align: top; text-align: right;">
                        <tr>
                            <td style="width: 99px; vertical-align: middle; height: 17px; text-align: right;">
                                <asp:Label ID="Label17" runat="server" Font-Bold="True" Font-Size="10pt" Font-Underline="True"
                                    Style="vertical-align: top; text-align: right" Text="������ ��鹷ҧ������ :"
                                    Width="145px"></asp:Label></td>
                            <td style="vertical-align: middle; width: 253px; text-align: right; height: 17px;"><asp:DropDownList ID="DropDownList2" runat="server" Width="250px" AutoPostBack="True">
                                <asp:ListItem Value="0">�͹��������Ң�</asp:ListItem>
                                <asp:ListItem Value="1">�͹�����Ũҡ�Ң��� �ӹѡ�ҹ�˭�</asp:ListItem>
                            </asp:DropDownList></td>
                            <td style="height: 17px">
                            </td>
                        </tr>
                        <tr>
                            <td style="width: 99px; height: 8px;">
                                <asp:Label ID="Label13" runat="server" Style="vertical-align: middle; text-align: right"
                                    Text="������ ����͹������ :" Width="145px" Font-Bold="True" Font-Size="10pt" Font-Underline="True"></asp:Label></td>
                            <td style="width: 253px; height: 8px;">
                                <asp:DropDownList ID="DropDownList1" runat="server" Width="250px" AutoPostBack="True">
                                    <asp:ListItem Value="0">�͡��� ���觢��/�ͧ�Թ���</asp:ListItem>
                                    <asp:ListItem Value="1">�͡��� ���觫����Թ���</asp:ListItem>
                                    <asp:ListItem Value="2">�͡��� ��͹�Թ���</asp:ListItem>
                                    <asp:ListItem Value="3">����¹�Թ���</asp:ListItem>
                                    <asp:ListItem Value="4">�͡��� ��Ţ��</asp:ListItem>
                                </asp:DropDownList></td>
                            <td style="height: 8px">
                            </td>
                        </tr>
                        <tr>
                            <td style="width: 99px">
                            </td>
                            <td style="width: 253px"></td>
                            <td style="vertical-align: top; text-align: left">
                                &nbsp;</td>
                        </tr>
                        <tr>
                            <td style="width: 99px">
                            </td>
                            <td style="vertical-align: top; text-align: right; width: 253px;">
                                <asp:Button ID="Button2" runat="server" Font-Bold="True" Font-Size="8pt" Text="���͡������" /></td>
                            <td>
                            </td>
                        </tr>
                    </table>
                </td>
                <td style="width: 440px; background-color: white; vertical-align: top; text-align: left;" rowspan="8">
                    <asp:DataGrid ID="DataGrid1" runat="server" BackColor="White" Font-Size="10pt" Width="414px">
                        <AlternatingItemStyle BackColor="#FF8000" BorderColor="White" ForeColor="Black" />
                        <HeaderStyle BackColor="#0000C0" BorderColor="White" ForeColor="White" />
                    </asp:DataGrid></td>
            </tr>
            <tr>
                <td style="vertical-align: top; width: 297px; text-align: left">
                </td>
                <td style="vertical-align: top; width: 282px; text-align: left">
                </td>
            </tr>
            <tr>
                <td style="vertical-align: top; width: 297px; text-align: left">
                </td>
                <td style="vertical-align: top; width: 282px; text-align: left">
                </td>
            </tr>
            <tr>
                <td style="vertical-align: top; width: 297px; text-align: left">
                    <asp:Label ID="Label3" runat="server" Font-Bold="True" Style="vertical-align: top;
                        text-align: right; text-decoration: underline;" Text="������ ����ͧ���·ҧ" Width="175px" ForeColor="White"></asp:Label></td>
                <td style="vertical-align: top; width: 282px; text-align: left">
                    <asp:Label ID="Label14" runat="server" Font-Bold="True" ForeColor="White" Text="��͡ �����š���͹"
                        Width="316px" style="text-decoration: underline"></asp:Label></td>
            </tr>
            <tr>
                <td style="vertical-align: top; width: 297px; height: 16px; text-align: left">
                </td>
                <td style="vertical-align: top; width: 282px; height: 16px; text-align: left">
                </td>
            </tr>
            <tr>
                <td style="vertical-align: top; width: 297px; height: 28px; text-align: left; background-color: #000099;">
                    <table style="width: 314px">
                        <tr>
                            <td style="vertical-align: top; width: 170px; text-align: right">
                                <asp:Label ID="Label9" runat="server" Text="���� �����������·ҧ :" Width="170px" ForeColor="White"></asp:Label></td>
                            <td style="width: 5px">
                            </td>
                            <td style="width: 158px">
                                <asp:TextBox ID="TextBox5" runat="server" Width="120px" BackColor="#FFFF80" Font-Bold="True" ReadOnly="True">S02DB</asp:TextBox></td>
                        </tr>
                        <tr>
                            <td style="vertical-align: top; width: 170px; text-align: right">
                                <asp:Label ID="Label10" runat="server" Text="���� �ҹ�����Ż��·ҧ :" Width="168px" ForeColor="White"></asp:Label></td>
                            <td style="width: 5px">
                            </td>
                            <td style="width: 158px">
                                <asp:TextBox ID="TextBox6" runat="server" Width="120px" BackColor="#FFFF80" Font-Bold="True" ReadOnly="True">BCNP</asp:TextBox></td>
                        </tr>
                        <tr>
                            <td style="vertical-align: top; width: 170px; text-align: right">
                                <asp:Label ID="Label11" runat="server" Text="���� �����ҹ���·ҧ :" Width="168px" ForeColor="White"></asp:Label></td>
                            <td style="width: 5px">
                            </td>
                            <td style="width: 158px">
                                <asp:TextBox ID="TextBox7" runat="server" Width="120px" BackColor="#FFFF80" Font-Bold="True" ReadOnly="True">sa</asp:TextBox></td>
                        </tr>
                        <tr>
                            <td style="vertical-align: top; width: 170px; text-align: right">
                                <asp:Label ID="Label12" runat="server" Text="���ʼ�ҹ���·ҧ :" ForeColor="White"></asp:Label></td>
                            <td style="width: 5px">
                            </td>
                            <td style="width: 158px">
                                <asp:TextBox ID="TextBox8" runat="server" Width="120px" BackColor="#FFFF80" Font-Bold="True" ForeColor="#FFFF80" ReadOnly="True">[ibdkifu</asp:TextBox></td>
                        </tr>
                    </table>
                    <br />
                </td>
                <td style="vertical-align: top; width: 282px; height: 28px; text-align: left; background-color: #ff9900;">
                    <table style="width: 506px">
                        <tr>
                            <td style="vertical-align: top; width: 47px; text-align: left">
                                <asp:Label ID="Label15" runat="server" Font-Bold="True" Font-Size="10pt" ForeColor="Black"
                                    Text="�Ţ����͡��� " Font-Underline="True" Visible="False" style="vertical-align: middle; text-align: left" Width="85px"></asp:Label>
                                <asp:TextBox ID="TextBox9" runat="server" Enabled="False" Visible="False" Width="120px"></asp:TextBox>
                            </td>
                            <td style="vertical-align: bottom; width: 29px; text-align: left">
                                <asp:TextBox ID="TextBox11" runat="server" Enabled="False" Visible="False" Width="360px"></asp:TextBox></td>
                        </tr>
                        <tr>
                            <td style="width: 47px; vertical-align: top; text-align: left;">
                                <asp:Label ID="Label16" runat="server" Font-Bold="True" Font-Size="10pt" Font-Underline="True"
                                    ForeColor="Black" Text="�����Թ��� " Visible="False" style="vertical-align: middle; text-align: left" Width="85px"></asp:Label>
                                <asp:TextBox ID="TextBox10" runat="server" Width="120px" Enabled="False" Visible="False"></asp:TextBox></td>
                            <td style="vertical-align: bottom; width: 29px; text-align: left">
                                <asp:TextBox ID="TextBox12" runat="server" Enabled="False" Visible="False" Width="360px"></asp:TextBox></td>
                        </tr>
                        <tr>
                            <td style="width: 47px"><asp:Button ID="Button3" runat="server" Font-Bold="True" Font-Size="8pt" Text="��Ǩ������" Visible="False" Width="127px" /></td>
                            <td style="width: 29px">
                                <asp:Button ID="Button4" runat="server" Font-Bold="True" Font-Size="8pt" Text="�����������" Visible="False" Width="85px" /></td>
                        </tr>
                        <tr>
                            <td style="width: 47px">
                            </td>
                            <td style="width: 29px">
                            </td>
                        </tr>
                        <tr>
                            <td style="width: 47px">
                                <asp:Button ID="Button1" runat="server" Font-Bold="True" Font-Size="8pt" Text="�͹������" Visible="False" Width="127px" /></td>
                            <td style="width: 29px">
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td style="vertical-align: top; width: 297px; height: 29px; background-color: #ffffff;
                    text-align: left">
                </td>
                <td style="vertical-align: top; width: 282px; height: 29px; background-color: #ffffff;
                    text-align: left">
                </td>
            </tr>
            <tr>
                <td style="vertical-align: top; width: 297px; height: 140px; background-color: white;
                    text-align: left">
                    <table style="width: 314px">
                        <tr>
                            <td style="width: 180px">
                            </td>
                            <td>
                            </td>
                        </tr>
                        <tr>
                            <td style="width: 180px">
                            </td>
                            <td>
                            </td>
                        </tr>
                        <tr>
                            <td style="width: 180px">
                            </td>
                            <td>
                            </td>
                        </tr>
                    </table>
                </td>
                <td style="vertical-align: top; width: 282px; height: 140px; background-color: white;
                    text-align: left">
                </td>
            </tr>
        </table>
    
    </div>
    </form>
</body>
</html>
