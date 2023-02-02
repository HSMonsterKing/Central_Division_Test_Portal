<%@ Page Language="VB" MasterPageFile="./MasterPage.master" MaintainScrollPositionOnPostback="true" AutoEventWireup="False" CodeFile="對帳.aspx.vb" Inherits="對帳" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" runat="Server">
    <asp:ScriptManager ID="ScriptManager1" runat="server" EnablePageMethods = "true" EnableScriptGlobalization="True" EnableScriptLocalization="True"/>
    <link rel="stylesheet" runat="server" media="screen" href="css\對帳.css"/>
    <div>
        <h1>
            <a ID="Title" href="對帳.aspx">對帳<a>
            <asp:DropDownList ID="DropDownList1" runat="server" AutoPostBack="True" OnSelectedIndexChanged="DropDownList1_SelectedIndexChanged" CssClass="DropDownList"/>
        </h1>
    </div>
    <asp:Panel ID="Panel1" runat="server" DefaultButton="Button1" CssClass="Panel1">
        結帳日期<asp:DropDownList ID="結帳日期a" runat="server" DataSourceID="SqlDataSource2" DataTextField="結帳日期" DataValueField="結帳日期" CssClass="DropDownList"/>~<asp:DropDownList ID="結帳日期b" runat="server" DataSourceID="SqlDataSource2" DataTextField="結帳日期" DataValueField="結帳日期" CssClass="DropDownList"/>
        <asp:Button ID="Button1" runat="server" Text="對帳" OnClick="Download" OnClientClick="return confirm('確認無誤?');" CssClass="GreenButton"/>
        <asp:Label ID="Label1" runat="server" Text="" CssClass="GreenLabel"/>
        <asp:Label ID="Label2" runat="server" Text="" CssClass="RedLabel"/>
    </asp:Panel>
    <asp:Panel ID="Panel2" runat="server" DefaultButton="Button1" CssClass="Panel2">
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" CellPadding="1" HorizontalAlign="Center" DataSourceID="SqlDataSource1" DataKeyNames="id" GridLines="None" ShowHeaderWhenEmpty="False" AllowPaging="True" PageSize="20" AllowSorting="False" PagerSettings-PageButtonCount=25 EnableModelValidation="True" AutoGenerateEditButton="False">
            <Columns>
                <asp:TemplateField HeaderText="">
                    <ItemTemplate>
                        <asp:Label ID="Label1" runat="server" Text='<%# Container.DataItemIndex + 1 %>' CssClass="Label Label1"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="ID" Visible="False">
                    <ItemTemplate>
                        <asp:Label ID="Label2" runat="server" Text='<%# Bind("id") %>' CssClass="Label Label2"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="上傳時間">
                    <ItemTemplate>
                        <asp:TextBox ID="上傳時間" runat="server" Text='<%# Bind("上傳時間") %>' Maxlength=300 Enabled="False" CssClass="TextBox 上傳時間"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="檔名">
                    <ItemTemplate>
                        <asp:HyperLink ID="檔名" runat="server" NavigateUrl='<%# Bind("檔名") %>' Text='<%# Bind("檔名") %>' CssClass="檔名"></asp:HyperLink>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="備註">
                    <ItemTemplate>
                        <asp:TextBox ID="備註" runat="server" Text='<%# Bind("備註") %>' Maxlength=300 Enabled="False" CssClass="TextBox 備註"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="起">
                    <ItemTemplate>
                        <asp:TextBox ID="起" runat="server" Text='<%# Bind("起") %>' Maxlength=300 Enabled="False" CssClass="TextBox 起"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="迄">
                    <ItemTemplate>
                        <asp:TextBox ID="迄" runat="server" Text='<%# Bind("迄") %>' Maxlength=300 Enabled="False" CssClass="TextBox 迄"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="選取">
                    <ItemTemplate>
                        <asp:CheckBox ID="CheckBox1" runat="server" Checked='<%# If(Eval("選取").ToString() = "True", True, False) %>' AutoPostBack="True" OnCheckedChanged="CheckBox1_CheckedChanged"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="">
                    <ItemTemplate>
                        <asp:Button ID="刪除" runat="server" Text="刪除" CommandName="CustomDelete" OnClientClick="return confirm('確定刪除?')" CssClass="RedButton"/>
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
        SelectCommand=
            "SELECT 
            id, 
            帳戶, 
            FORMAT(上傳時間, 'yyyy-MM-dd tt hh:mm:ss') AS 上傳時間, 
            檔名, 
            備註, 
            FORMAT(起, 'yyyy-MM-dd') AS 起, 
            FORMAT(迄, 'yyyy-MM-dd') AS 迄, 
            選取 
            FROM 對帳 WHERE 帳戶 = @選項 
            ORDER BY 迄, 
            CASE WHEN 備註 = '帳戶明細' THEN 0 ELSE 1 END, 
            CASE WHEN 備註 = '專案代收查詢' THEN 0 ELSE 1 END"
        DeleteCommand="DELETE FROM 對帳 WHERE id = @id">
        <SelectParameters>
            <asp:ControlParameter ControlID="DropDownList1" ConvertEmptyStringToNull="False" Name="選項" PropertyName="Text" Type="String"/>
        </SelectParameters>
        <DeleteParameters>
            <asp:Parameter Name="id" ConvertEmptyStringToNull="False" Type="String"/>
        </DeleteParameters>
    </asp:SqlDataSource>
    <asp:SqlDataSource ID="SqlDataSource2" runat="server" ConnectionString='<%$ ConnectionStrings:ApplicationServices%>'
        SelectCommand="SELECT DISTINCT CONVERT(varchar(4), YEAR(結帳日期) - 1911) + FORMAT(結帳日期, '/MM/dd') AS 結帳日期 FROM 日報表 ORDER BY 結帳日期 DESC">
    </asp:SqlDataSource>
</asp:Content>


