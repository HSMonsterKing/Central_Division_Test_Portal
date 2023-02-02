<%@ Page Title="電子公布欄" Language="VB" MasterPageFile="./MasterPage.master" AutoEventWireup="false" CodeFile="電子公布欄.aspx.vb" Inherits="電子公布欄" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" Runat="Server">
    <style type="text/css">
    .auto-style1 {
        width: 100%;
    }
</style>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
    <table class="auto-style1">
    <tr>
        <td style="font-family: 標楷體; text-align: left; width: 100%">
            <asp:Button ID="Button24" runat="server" Text="Button" />
            <asp:Label ID="Label2" runat="server" Text="Label"></asp:Label>
            <asp:Label ID="Label3" runat="server" Text="Label"></asp:Label>
        </td>
    </tr>
    <tr>
        <td style="font-family: 標楷體; text-align: left; width: 100%">
            <asp:GridView ID="GridView1" runat="server">
            </asp:GridView>
        </td>
    </tr>
    <tr>
        <td style="font-family: 標楷體; text-align: left; width: 100%">&nbsp;</td>
    </tr>
</table>
</asp:Content>

