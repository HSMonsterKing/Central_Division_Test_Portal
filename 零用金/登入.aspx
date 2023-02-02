<%@ Page Language="VB" MasterPageFile="./MasterPage.master" MaintainScrollPositionOnPostback="true" AutoEventWireup="false" CodeFile="登入.aspx.vb" Inherits="登入" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
    <asp:ScriptManager ID="ScriptManager1" runat="server" EnablePageMethods = "true" EnableScriptGlobalization="True" EnableScriptLocalization="True"/>
    <script src='js/登入.js'></script>
    <link rel="stylesheet" runat="server" media="screen" href="css\登入.css"/>
<form onsubmit="return false;">
  <div class="padlock">
    <div class="padlock__hook">
      <div class="padlock__hook-body"></div>
      <div class="padlock__hook-body"></div>
    </div>
    <div class="padlock__body">
      <div class="padlock__face">
        <div class="padlock__eye padlock__eye--left"></div>
        <div class="padlock__eye padlock__eye--right"></div>
        <div class="padlock__mouth padlock__mouth--one"></div>
        <div class="padlock__mouth padlock__mouth--two"></div>
        <div class="padlock__mouth padlock__mouth--three"></div>
      </div>
    </div>
  </div>
  <div class="app">
    <h1>You logged in! 🎉</h1>
    <button class="logout-button" type="reset">Logout</button>
  </div><span class="logout-message">You have logged out.</span>
</form>
    <div><h1><a id="Title" href="登入.aspx">登入</a></h1></div>
    <asp:Panel ID="Panel1" runat="server" DefaultButton="登入" CssClass="Panel1">
        帳號:<asp:TextBox ID="帳號" runat="server" Maxlength=20 CssClass="Input1"/><br>
        密碼:<asp:TextBox ID="密碼" runat="server" Maxlength=20 CssClass="Input1" TextMode="password" /><br>
        <asp:Button ID="登入" runat="server" Text="登入" OnClick="登入_Click" CssClass="login-button"/>
        <asp:Button ID="登出" runat="server" Text="登出" OnClick="登出_Click" CssClass="logout-button" Visible="false"/>
        <asp:Button ID="密碼轉換" runat="server" Text="密碼轉換" OnClick="密碼轉換_Click" CssClass="GreenButton" Visible="false"/>
        <asp:Label ID="Label1" runat="server" Text="" CssClass=""/>
        <asp:Label ID="Label2" runat="server" Text="" CssClass=""/>
    </asp:Panel>
    <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString='<%$ ConnectionStrings:ApplicationServices%>'
        SelectCommand="" 
        Insertcommand="" 
        UpdateCommand=""
        DeleteCommand="">
        <SelectParameters>
        </SelectParameters>
        <InsertParameters>
        </InsertParameters>
        <UpdateParameters>
        </UpdateParameters>
        <DeleteParameters>
        </DeleteParameters>
    </asp:SqlDataSource>
</asp:Content>
