<%@ Page Language="VB" MasterPageFile="./MasterPage.master" MaintainScrollPositionOnPostback="true" AutoEventWireup="false" CodeFile="取號.aspx.vb" Inherits="取號" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
    <asp:ScriptManager ID="ScriptManager1" runat="server" EnablePageMethods = "true" EnableScriptGlobalization="True" EnableScriptLocalization="True"/>
    <script src='js/取號.js'></script>
    <link rel="stylesheet" runat="server" media="screen" href="css\取號.css"/>
    <div><h1><a id="Title" href="取號.aspx">取號<a></h1></div>
    <asp:Panel ID="Panel1" runat="server" DefaultButton="搜尋" CssClass="Panel1">
        年<asp:TextBox ID="年" runat="server" Maxlength=3 CssClass="Input2"/>
        種類<asp:DropDownList ID="_種類" runat="server" AutoPostBack="True" CssClass="DropDownList">
            <asp:ListItem Text="A" Value="A"></asp:ListItem>
            <asp:ListItem Text="B" Value="B"></asp:ListItem>
            <asp:ListItem Text="XZ" Value="XZ"></asp:ListItem>
        </asp:DropDownList>
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
                <asp:TemplateField HeaderText="單位別">
                    <ItemTemplate>
                        <asp:DropDownList ID="單位別" runat="server" Text='<%# Bind("單位別") %>' AutoPostBack="True" CssClass="DropDownList">
                        <asp:ListItem Text="" Value=""></asp:ListItem>
                        <asp:ListItem Text="中區交通控制中心" Value="中區交通控制中心"></asp:ListItem>
                        <asp:ListItem Text="交通管理科" Value="交通管理科"></asp:ListItem>
                        <asp:ListItem Text="工務科" Value="工務科"></asp:ListItem>
                        <asp:ListItem Text="分局長室" Value="分局長室"></asp:ListItem>
                        <asp:ListItem Text="主計室" Value="主計室"></asp:ListItem>
                        <asp:ListItem Text="政風室" Value="政風室"></asp:ListItem>
                        <asp:ListItem Text="業務科" Value="業務科"></asp:ListItem>
                        <asp:ListItem Text="秘書室" Value="秘書室"></asp:ListItem>
                        <asp:ListItem Text="機料及保養場" Value="機料及保養場"></asp:ListItem>
                        <asp:ListItem Text="人事室" Value="人事室"></asp:ListItem>
                        <asp:ListItem Text="勞安科" Value="勞安科"></asp:ListItem>
                        <asp:ListItem Text="5工務段" Value="5工務段"></asp:ListItem>
                        <asp:ListItem Text="5服務區" Value="5服務區"></asp:ListItem>
                    </asp:DropDownList>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="承辦人">
                    <ItemTemplate>
                        <asp:DropDownList ID="承辦人" runat="server" Text='<%# Bind("承辦人") %>' AutoPostBack="True" CssClass="DropDownList">
                        <asp:ListItem Text="" Value=""></asp:ListItem>
                        <asp:ListItem Text="陳春綢" Value="陳春綢"></asp:ListItem>
                        <asp:ListItem Text="江嘉珊" Value="江嘉珊"></asp:ListItem>
                        <asp:ListItem Text="洪孟恬" Value="洪孟恬"></asp:ListItem>
                        <asp:ListItem Text="彭金杏" Value="彭金杏"></asp:ListItem>
                        <asp:ListItem Text="柯佳妮" Value="柯佳妮"></asp:ListItem>
                        <asp:ListItem Text="藍雅燕" Value="藍雅燕"></asp:ListItem>
                    </asp:DropDownList>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="種類">
                    <ItemTemplate>
                        <asp:TextBox ID="種類" runat="server" Text='<%# Bind("種類") %>' Maxlength=0 Enabled="True" CssClass="TextBox 種類"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="號數">
                    <ItemTemplate>
                        <asp:TextBox ID="號數" runat="server" Text='<%# Bind("號數", "{0:000}") %>' Maxlength=0 Enabled="True" CssClass="TextBox 號數"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="">
                    <ItemTemplate>
                        <asp:Button ID="取號" runat="server" Text="取號" CommandName="取號" Enabled='<%# If (Eval("號數").ToString = "", 1, 0) %>' CssClass="GreenButton"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Left" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="備註">
                    <ItemTemplate>
                        <asp:TextBox ID="備註" runat="server" Text='<%# Bind("備註") %>' Maxlength=0 Enabled="True" CssClass="TextBox 備註"/>
                        <ajaxToolkit:AutoCompleteExtender ID="備註自動" runat="server" TargetControlID="備註" 
                            CompletionSetCount="10000" CompletionInterval="50" EnableCaching="true" MinimumPrefixLength="1" 
                            ServiceMethod="GetMyList" 
                            CompletionListCssClass="CompletionList" 
                            CompletionListItemCssClass="CompletionListItem" 
                            CompletionListHighlightedItemCssClass="CompletionListHighlightedItem"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="">
                    <ItemTemplate>
                        <asp:Button ID="刪除" runat="server" Text="刪除" CommandName="Delete" OnClientClick="return confirm('確定刪除?')" Enabled="True" CssClass="RedButton"/>
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
            SELECT * FROM 收支備查簿 
            WHERE 年 = @年 
            AND 取號 = 1 
            AND _種類 = @_種類 
            order by 號數" 
        Insertcommand="INSERT INTO 收支備查簿 (年, 取號, _種類) VALUES (@年, 1, @_種類)" 
        UpdateCommand=""
        DeleteCommand="DELETE FROM 收支備查簿 WHERE id=@id">
        <SelectParameters>
            <asp:ControlParameter ControlID="年" ConvertEmptyStringToNull="False" Name="年" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="_種類" ConvertEmptyStringToNull="False" Name="_種類" PropertyName="Text" Type="String"/>
        </SelectParameters>
        <InsertParameters>
            <asp:ControlParameter ControlID="年" ConvertEmptyStringToNull="False" Name="年" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="_種類" ConvertEmptyStringToNull="False" Name="_種類" PropertyName="Text" Type="String"/>
        </InsertParameters>
        <UpdateParameters>
        </UpdateParameters>
        <DeleteParameters>
            <asp:Parameter Name="id"/>
        </DeleteParameters>
    </asp:SqlDataSource>
</asp:Content>
