<%@ Page Language="VB" MasterPageFile="./MasterPage.master" MaintainScrollPositionOnPostback="true" AutoEventWireup="false" CodeFile="科目清單.aspx.vb" Inherits="科目清單" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
    <asp:ScriptManager ID="ScriptManager1" runat="server" EnablePageMethods = "true" EnableScriptGlobalization="True" EnableScriptLocalization="True"/>
    <script src='js/科目清單.js'></script>
    <link rel="stylesheet" runat="server" media="screen" href="css\科目清單.css"/>
    <div><h1><a id="Title" href="科目清單.aspx">科目清單<a></h1></div>
    <asp:Panel ID="Panel1" runat="server" DefaultButton="搜尋" CssClass="Panel1">
        收尋科目<asp:TextBox ID="科目" runat="server" Maxlength=0 CssClass="Input1"/>
        <asp:Button ID="搜尋" runat="server" Text="搜尋" CssClass="GreenButton"/>
        <asp:Button ID="新增" runat="server" Text="新增一列" OnClick="Insert" CssClass="GreenButton"/>
        <asp:Button ID="存檔" runat="server" Text="存檔" OnClick="Update" CssClass="GreenButton"/>
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
                <asp:TemplateField HeaderText="科目">
                    <ItemTemplate>
                        <asp:TextBox ID="科目" runat="server" Text='<%# Bind("科目") %>' Maxlength=0 Enabled="True" CssClass="TextBox 科目"/>
                        <!-- 備註 -->
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="">
                    <ItemTemplate>
                        <asp:Button ID="刪除" runat="server" Text="刪除" CommandName="刪除" OnClientClick="return confirm('確定刪除?如此資科目前已使用其資料將一併刪除')" Enabled="True" CssClass="RedButton"/>
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
            SELECT * FROM 科目表 
            WHERE (TRIM(@科目) = '' OR 科目 LIKE '%' + @科目 + '%')
            ORDER BY CASE WHEN 科目 IS NULL THEN 1 ELSE 0 END, 科目" 
        Insertcommand="INSERT INTO 科目表 (科目) VALUES (Null)" 
        UpdateCommand="" 
        DeleteCommand="DELETE FROM 科目表 WHERE id=@id">
        <SelectParameters>
            <asp:ControlParameter ControlID="科目" ConvertEmptyStringToNull="False" Name="科目" PropertyName="Text" Type="String"/>
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
