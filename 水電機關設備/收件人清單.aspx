﻿<%@ Page Language="VB" MasterPageFile="./MasterPage.master" MaintainScrollPositionOnPostback="true" AutoEventWireup="false" CodeFile="收件人清單.aspx.vb" Inherits="收件人清單"%>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server" >
    <asp:ScriptManager ID="ScriptManager1" runat="server" EnablePageMethods = "true" EnableScriptGlobalization="True" EnableScriptLocalization="True"/>
    <script src='js/收件人清單.js'></script>
    <link rel="stylesheet" runat="server" media="screen" href="css\收件人清單.css"/>
    <style>
    </style>
    <div><h1><a id="Title" href="收件人清單.aspx">收件人清單<a></h1></div>
    <asp:Panel ID="Panel1" runat="server" CssClass="Panel1">
        <asp:Button ID="新增" runat="server" Text="新增一列" OnClick="Insert" CssClass="GreenButton"/>
        <asp:Button ID="存檔" runat="server" Text="存檔" OnClick="Update" CssClass="GreenButton"/>
        <asp:Button ID="測試" runat="server" Text="測試" OnClick="test" CssClass="GreenButton" Visible="false"/>
        <asp:Label ID="Label1" runat="server" Text="" CssClass="GreenLabel"/>
        <asp:Label ID="Label2" runat="server" Text="" CssClass="RedLabel"/>
    </asp:Panel>
    <asp:Panel ID="Panel3" runat="server" CssClass="Panel3" DefaultButton="存檔">
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="false" CellPadding="1" HorizontalAlign="Center" DataSourceID="SqlDataSource1" DataKeyNames="id" GridLines="None" ShowHeaderWhenEmpty="false" AllowPaging="True" PageSize="15" AllowSorting="True" PagerSettings-PageButtonCount=25 EnableModelValidation="True" AutoGenerateEditButton="false">
            <Columns>
                <asp:TemplateField HeaderText="id" Visible="false">
                    <ItemTemplate>
                        <asp:TextBox ID="id" runat="server" Text='<%# Bind("id") %>' Maxlength=0 Enabled="False" CssClass="TextBox id"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="名字">
                    <ItemTemplate>
                    <asp:TextBox ID="名字" runat="server" Text='<%# Bind("名字") %>' Maxlength=0 Enabled="True" CssClass="TextBox 名字"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="">
                    <ItemTemplate>
                        <asp:Button ID="刪除" runat="server" Text="刪除" CommandName="刪除" OnClientClick="return confirm('確定刪除?')" Enabled="True" CssClass="RedButton"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Left" CssClass="Item"/>
                </asp:TemplateField>
            </Columns>
            <HeaderStyle BackColor="Green" Font-Bold="True" ForeColor="White" CssClass="Header"/>
            <RowStyle BackColor="White" CssClass="Row"/>
            <AlternatingRowStyle/>
            <SelectedRowStyle/>
            <EditRowStyle/>
            <PagerStyle BackColor="Green" HorizontalAlign="Center" CssClass="Pager"/>
            <FooterStyle/>
            <PagerSettings  Mode="NumericFirstLast" FirstPageText="<<" PreviousPageText="<" NextPageText=">" LastPageText=">>" />
        </asp:GridView>
    </asp:Panel>
    <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString='<%$ ConnectionStrings:ApplicationServices%>'
        SelectCommand="
            SELECT * FROM 收件人清單
            ORDER BY id"
        Insertcommand="INSERT INTO 收件人清單 (名字) VALUES ('')"
        UpdateCommand=""
        DeleteCommand="">
        <SelectParameters>
        </SelectParameters>
        <InsertParameters>
        </InsertParameters>
        <UpdateParameters>
        </UpdateParameters>
        <DeleteParameters>
            <asp:Parameter Name="id"/>
        </DeleteParameters>
    </asp:SqlDataSource>
</asp:Content>
