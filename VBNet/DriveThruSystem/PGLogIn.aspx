<%@ Page Language="VB" AutoEventWireup="false" CodeFile="PGLogIn.aspx.vb" Inherits="PGLogIn" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Untitled Page</title>
    
    <script type="text/javascript">
/*  
function onEnter0(){
document.getElementById("TextBox1").onkeypress = function() {
if(window.event.keyCode == 13) {
document.getElementById("TextBox2").focus();
return false;
}
return true;
};
document.getElementById("TextBox2").onkeypress = function() {
if(window.event.keyCode == 13) {
document.getElementById("TextBox3").focus();
return false;
}
return true;
};
document.getElementById("TextBox3").onkeypress = function() {
if(window.event.keyCode == 13) {
document.getElementById("TextBox4").focus();
return false;
}
return true;
};
document.getElementById("TextBox4").onkeypress = function() {
if(window.event.keyCode == 13) {
document.getElementById("").focus();
return false;
}
return true;
};
document.getElementById("TBUserID").onkeypress = function() {
if(window.event.keyCode == 13) {
document.getElementById("TextBox1").focus();
return false;
}
return true;
};
}
 */
    </script>

</head>
<body topmargin ="0px" leftmargin = "0px">
    <form id="form1" runat="server">
    <div>
        <table style="width: 159px">
            <tr>
                <td colspan="2" style="background-color: #ff6600; height: 15px;">
                </td>
            </tr>
            <tr>
                <td colspan="2">
                    <asp:Label ID="Label3" runat="server" Font-Size="Large" Text="บริษัท นพดลพานิช จำกัด"
                        Width="223px" style="text-align: center"></asp:Label></td>
            </tr>
            <tr>
                <td colspan="2" style="background-color: #000099; height: 15px;">
                </td>
            </tr>
            <tr>
                <td style="width: 5px; vertical-align: top; text-align: right;">
                    <asp:Label ID="Label2" runat="server" Font-Bold="False" Font-Size="8pt" Style="text-align: right"
                        Text="จุดเข้าใช้งาน :" Width="90px" Height="19px"></asp:Label></td>
                <td style="width: 22px; vertical-align: top; text-align: left;">
                    <asp:DropDownList ID="DDLPoint" runat="server" Font-Bold="True" Font-Size="8pt" Width="56px">
                    </asp:DropDownList></td>
            </tr>
            <tr>
                <td colspan="2">
                </td>
            </tr>
            <tr>
                <td colspan="2" style="background-color: #000099; height: 15px;">
                </td>
            </tr>
            <tr>
                <td style="width: 5px; vertical-align: top; text-align: right;">
                    <asp:Label ID="Label1" runat="server" Font-Bold="False" Font-Size="8pt" Style="text-align: right"
                        Text="รหัสพนักงาน :" Width="90px" Height="18px"></asp:Label></td>
                <td style="width: 22px; vertical-align: top; text-align: left;">
                    <asp:TextBox ID="TBUserID" runat="server" Font-Bold="True" Font-Size="6pt" Width="121px" onkeypress="onEnter0();"></asp:TextBox></td>
            </tr>
            <tr>
                <td colspan="2" style="background-color: #ff6600; height: 15px;">
                    <asp:Label ID="LBLMessage" runat="server" Font-Size="7pt" Width="222px"></asp:Label></td>
            </tr>
            <tr>
                <td style="width: 5px">
                </td>
                <td style="width: 22px; vertical-align: top; text-align: right;">
                    <asp:Button ID="Button1" runat="server" Font-Bold="True" Font-Size="7pt" Text="LogIn" /></td>
            </tr>
            <tr>
                <td style="height: 15px; background-color: #000099;" colspan="2">
                    </td>
            </tr>
            <tr>
                <td colspan="2" style="height: 15px; background-color: #000099">
                    </td>
            </tr>
            <tr>
                <td colspan="2" style="height: 15px; background-color: #000099">
                    </td>
            </tr>
            <tr>
                <td colspan="2" style="height: 15px; background-color: #000099">
                    </td>
            </tr>
        </table>
    
    </div>
    </form>
</body>
</html>
