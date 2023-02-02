<%@ Page Language="VB" MasterPageFile="./MasterPage.master" MaintainScrollPositionOnPostback="true" AutoEventWireup="false" CodeFile="分配明細表.aspx.vb" Inherits="分配明細表" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
    <asp:ScriptManager ID="ScriptManager1" runat="server" EnablePageMethods = "true" EnableScriptGlobalization="True" EnableScriptLocalization="True"/>
    <script src='js/分配明細表.js'></script>
    <link rel="stylesheet" runat="server" media="screen" href="css\分配明細表.css"/>
    <div><h1><a id="Title" href="分配明細表.aspx">分配明細表<a></h1></div>
    <asp:Panel ID="Panel1" runat="server" DefaultButton="搜尋" CssClass="Panel1">
        <asp:TextBox ID="年" runat="server" Maxlength=3 CssClass="Input2"/>年
        <asp:TextBox ID="月" runat="server" Maxlength=2 CssClass="Input2"/>月
        <asp:TextBox ID="日" runat="server" Maxlength=2 CssClass="Input2"/>日
        <asp:Button ID="搜尋" runat="server" Text="搜尋" CssClass="GreenButton"/>
        <asp:Button ID="新增" runat="server" Text="新增一列" OnClick="Insert" CssClass="GreenButton"/>
        <asp:Button ID="存檔" runat="server" Text="存檔" OnClick="Update" CssClass="GreenButton"/>
        <asp:Button ID="刪除" runat="server" Text="刪除末列" OnClick="Delete" OnClientClick="return confirm('確定刪除?')" CssClass="RedButton"/>
        <asp:Label ID="Label1" runat="server" Text="" CssClass="GreenLabel"/>
        <asp:Label ID="Label2" runat="server" Text="" CssClass="RedLabel"/>
    </asp:Panel>
    <asp:Panel ID="Panel3" runat="server" DefaultButton="存檔" CssClass="Panel3">
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="false" CellPadding="1" HorizontalAlign="Center" DataSourceID="SqlDataSource1" DataKeyNames="id" GridLines="None" ShowHeaderWhenEmpty="false" AllowPaging="True" PageSize="15" AllowSorting="True" PagerSettings-PageButtonCount=25 EnableModelValidation="True" AutoGenerateEditButton="false">
            <Columns>
                <asp:TemplateField HeaderText="id" Visible="false">
                    <ItemTemplate>
                        <asp:TextBox ID="id" runat="server" Text='<%# Bind("id") %>' Maxlength=0 Enabled="False" CssClass="TextBox id"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="零用金使用單位">
                    <ItemTemplate>
                        <asp:TextBox ID="零用金使用單位" runat="server" Text='<%# Bind("零用金使用單位") %>' Maxlength=0 Enabled="True" CssClass="TextBox 零用金使用單位"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="金額">
                    <ItemTemplate>
                        <asp:TextBox ID="金額" runat="server" Text='<%# Bind("金額", "{0:c0}") %>' Maxlength=0 Enabled="True" CssClass="TextBox 金額"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="累計">
                    <ItemTemplate>
                        <asp:Label ID="累計" runat="server" Text='<%# Eval("累計", "{0:c0}") %>' Maxlength=0 Enabled="True" CssClass="Label 累計"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
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
            SELECT * FROM 分配明細表
            order by id "
        Insertcommand="INSERT INTO 分配明細表 (金額) VALUES (0)"
        UpdateCommand=""
        DeleteCommand="DELETE FROM 收支備查簿 WHERE id=@id">
        <SelectParameters>
            <asp:ControlParameter ControlID="年" ConvertEmptyStringToNull="False" Name="年" PropertyName="Text" Type="String"/>
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
