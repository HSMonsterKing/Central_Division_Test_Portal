<%@ Page Language="VB" MasterPageFile="~/MasterPage.master" MaintainScrollPositionOnPostback="true" AutoEventWireup="false" CodeFile="中分入口網.aspx.vb" Inherits="中分入口網"%>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server" >
    <asp:ScriptManager ID="ScriptManager1" runat="server" EnablePageMethods = "true" EnableScriptGlobalization="True" EnableScriptLocalization="True"/>
    <script src='js/中分入口網.js'></script>
    <link rel="stylesheet" runat="server" media="screen" href="css\中分入口網.css"/>
    <style>
    </style>
    <div><h1><a id="Title" href="中分入口網.aspx">中分秘書室入口網<a></h1></div>
    <asp:Panel ID="Panel1" runat="server" CssClass="Panel1" >
        <asp:Button ID="零用金" runat="server" Text="零用金" OnClick="零用金_Click" CssClass="GreenButton"/>
        <asp:Button ID="水電設備管理" runat="server" Text="水電設備管理" OnClick="水電設備管理_Click" CssClass="GreenButton"/>
        <asp:Button ID="出納對帳系統" runat="server" Text="出納對帳系統" OnClick="出納對帳系統_Click" Visible="false"  CssClass="GreenButton"/>
        <!-- 附件:<asp:FileUpload ID="FileUpload1" runat="server" AllowMultiple="True" CssClass="UploadButton"/> -->
        <asp:Button ID="大宗郵件" runat="server" Text="大宗郵件" OnClick="大宗郵件_Click" Visible="false" CssClass="GreenButton"/>
        <asp:Button ID="稽催寄送" runat="server" Text="稽催寄送" OnClick="稽催寄送_Click" Visible="false" CssClass="GreenButton"/>
        <asp:Label ID="Label1" runat="server" Text="" CssClass="GreenLabel"/>
        <asp:Label ID="Label2" runat="server" Text="" CssClass="RedLabel"/>
    </asp:Panel>
</asp:Content>


