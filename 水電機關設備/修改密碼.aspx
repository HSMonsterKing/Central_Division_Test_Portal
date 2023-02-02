<%@ Page Language="VB" MasterPageFile="./MasterPage.master" MaintainScrollPositionOnPostback="true" AutoEventWireup="false" CodeFile="修改密碼.aspx.vb" Inherits="修改密碼" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
    <asp:ScriptManager ID="ScriptManager1" runat="server" EnablePageMethods = "true" EnableScriptGlobalization="True" EnableScriptLocalization="True"/>
    <script src='js/修改密碼.js'></script>
    <link rel="stylesheet" runat="server" media="screen" href="css\修改密碼.css"/>
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
    <div><h1><a id="Title" href="修改密碼.aspx">修改密碼</a></h1></div>
    <asp:Panel ID="Panel1" runat="server" DefaultButton="修改密碼" CssClass="Panel1">
        <asp:Label ID="RedLabel" runat="server" Text="" CssClass="RedLabel1"/><br>
        如要離開請回上一頁<br>
        新密碼（大寫，小寫，數字/特殊字符和最少12個字符<br>
        帳號:<asp:TextBox ID="帳號" runat="server" Maxlength=20 CssClass="Input1"/><br>
        密碼:<asp:TextBox ID="密碼" runat="server" Maxlength=20 CssClass="Input1" TextMode="password" title="格式錯誤" pattern="(?=^.{12,}$)((?=.*\d)(?=.*\W+))(?![.\n])(?=.*[A-Z])(?=.*[a-z]).*$" /><br>
        確認密碼:<asp:TextBox ID="確認密碼" runat="server" Maxlength=20 CssClass="Input1" TextMode="password" /><br>
        <asp:Button ID="修改密碼" runat="server" Text="修改密碼" OnClick="修改密碼_Click" CssClass="login-button"/><br>
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
