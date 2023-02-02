<%@ Page Language="VB" MasterPageFile="./MasterPage.master" MaintainScrollPositionOnPostback="true" AutoEventWireup="false" CodeFile="照片.aspx.vb" Inherits="照片"%>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server" >
    <asp:ScriptManager ID="ScriptManager1" runat="server" EnablePageMethods = "true" EnableScriptGlobalization="True" EnableScriptLocalization="True"/>
    <script src='js/照片.js'></script>
    <link rel="stylesheet" runat="server" media="screen" href="css\照片.css"/>
    <style>
    </style>
    <asp:Panel ID="Panel1" runat="server" CssClass="Panel1" >
        <asp:Button ID="返回" runat="server" Text="返回" CssClass="GreenButton" OnClick="返回_OnClick" Visible="False"/>
        <asp:Button ID="原始照片" runat="server" Text="原始照片" CssClass="GreenButton" OnClick="原始照片_OnClick" Visible="True"/><BR>
        <asp:Image ID="Image1" runat="server" Text="" CssClass="Image1" Visible="True"/>
        <asp:Image ID="O_Image" runat="server" Text="" CssClass="Image1" Visible="False"/>
        <asp:Label ID="Label1" runat="server" Text="" CssClass="GreenLabel"/>
        <asp:Label ID="Label2" runat="server" Text="" CssClass="RedLabel"/>
    </asp:Panel>
</asp:Content>
