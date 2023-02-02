<%@ Page Language="VB" MasterPageFile="./MasterPage.master" MaintainScrollPositionOnPostback="true" AutoEventWireup="false" CodeFile="測試.aspx.vb" Inherits="測試"%>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server" >
    <asp:ScriptManager ID="ScriptManager1" runat="server" EnablePageMethods = "true" EnableScriptGlobalization="True" EnableScriptLocalization="True"/>
    <script src='js/測試.js'></script>
    <link rel="stylesheet" runat="server" media="screen" href="css\測試.css"/>
    <style>
    </style>
    <div><h1><a id="Title" href="測試.aspx">測試PDF<a></h1></div>
    <asp:Panel ID="Panel1" runat="server" CssClass="Panel1" >
        <asp:FileUpload ID="FileUpload1" runat="server" AllowMultiple="True" CssClass="UploadButton"/>
        <asp:Button ID="測試" runat="server" Text="轉換" CssClass="GreenButton" OnClick="Test" Visible="False"/>
        <asp:Label ID="Label1" runat="server" Text="" CssClass="GreenLabel"/>
        <asp:Label ID="Label2" runat="server" Text="" CssClass="RedLabel"/>
    </asp:Panel>
    <asp:Panel ID="Panel3" runat="server" CssClass="Panel3" >
    </asp:Panel>
</asp:Content>
