<%@ Page Language="VB" MasterPageFile="./MasterPage.master" MaintainScrollPositionOnPostback="true" AutoEventWireup="false" CodeFile="濾心日誌.aspx.vb" Inherits="濾心日誌"%>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server" >
    <asp:ScriptManager ID="ScriptManager1" runat="server" EnablePageMethods = "true" EnableScriptGlobalization="True" EnableScriptLocalization="True"/>
    <script src='js/濾心日誌.js'></script>
    <link rel="stylesheet" runat="server" media="screen" href="css\濾心日誌.css"/>
    <style>
    </style>
    <div><h1><a id="Title" href="濾心日誌.aspx">濾心日誌<a></h1></div>
    <asp:Panel ID="Panel1" runat="server" DefaultButton="搜尋" CssClass="Panel1" >
        建置日期:<asp:DropDownList ID="年" runat="server" AutoPostBack="True" DataSourceID="SqlDataSource2" DataTextField="民國年" DataValueField="西元年" CssClass="DropDownList"/>年
    <asp:DropDownList ID="月" runat="server" AutoPostBack="True" CssClass="DropDownList">
        <asp:ListItem Text="" Value=""></asp:ListItem>
        <asp:ListItem Text="1" Value="1"></asp:ListItem>
        <asp:ListItem Text="2" Value="2"></asp:ListItem>
        <asp:ListItem Text="3" Value="3"></asp:ListItem>
        <asp:ListItem Text="4" Value="4"></asp:ListItem>
        <asp:ListItem Text="5" Value="5"></asp:ListItem>
        <asp:ListItem Text="6" Value="6"></asp:ListItem>
        <asp:ListItem Text="7" Value="7"></asp:ListItem>
        <asp:ListItem Text="8" Value="8"></asp:ListItem>
        <asp:ListItem Text="9" Value="9"></asp:ListItem>
        <asp:ListItem Text="10" Value="10"></asp:ListItem>
        <asp:ListItem Text="11" Value="11"></asp:ListItem>
        <asp:ListItem Text="12" Value="12"></asp:ListItem>
    </asp:DropDownList>月
        <asp:Button ID="搜尋" runat="server" Text="搜尋" CssClass="GreenButton" />
        <asp:Button ID="測試" runat="server" Text="轉換" CssClass="GreenButton" OnClick="Test" Visible="False"/>
        <asp:Label ID="Label1" runat="server" Text="" CssClass="GreenLabel"/>
        <asp:Label ID="Label2" runat="server" Text="" CssClass="RedLabel"/>
    </asp:Panel>
    <asp:Panel ID="Panel3" runat="server" CssClass="Panel3">
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="false" CellPadding="1" HorizontalAlign="Center" DataSourceID="SqlDataSource1" DataKeyNames="id" GridLines="None" ShowHeaderWhenEmpty="false" AllowPaging="True" PageSize="15" AllowSorting="True" PagerSettings-PageButtonCount=25 EnableModelValidation="True" AutoGenerateEditButton="false">
            <Columns>
                <asp:TemplateField HeaderText="id" Visible="false">
                    <ItemTemplate>
                        <asp:Label ID="id" runat="server" Text='<%# Eval("id") %>' Maxlength=0 Enabled="False" CssClass="Label id"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="濾心ID" Visible="false">
                    <ItemTemplate>
                        <asp:Label ID="濾心ID" runat="server" Text='<%# Eval("濾心ID") %>' Maxlength=0 Enabled="False" CssClass="Label 濾心ID"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="列">
                    <ItemTemplate>
                        <asp:Label ID="_列" runat="server" Text='' Maxlength=0 CssClass="Label _列"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="項目">
                    <ItemTemplate>
                        <asp:Label ID="項目" runat="server" Text='<%# Eval("項目") %>' Maxlength=0 Enabled="True" CssClass="Label 項目"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="品名型號">
                    <ItemTemplate>
                        <asp:Label ID="品名型號" runat="server" Text='<%# Eval("品名型號") %>' Maxlength=0 Enabled="True" TextMode="MultiLine" CssClass="Label 品名型號"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="更換日期">
                    <ItemTemplate>
                        <asp:Label ID="更換日期" runat="server" Text='<%# If(IsDate(Eval("更換日期")), (Year(Eval("更換日期"))-1911).ToString() & Eval("更換日期", "{0:/MM/dd}"), "") %>' Maxlength=0 Enabled="True" CssClass="Label 更換日期"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="更換地點">
                    <ItemTemplate>
                        <asp:TextBox ID="更換地點" runat="server" Text='<%# Bind("更換地點") %>' Maxlength=0 Enabled="True" TextMode="MultiLine" CssClass="TextBox 更換地點"/>
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
            SELECT 濾心更換_日誌.*,項目,品名型號,更換地點
	        FROM 濾心更換_日誌
            Inner Join 濾心清單表 On 濾心更換_日誌.濾心ID=濾心清單表.ID
            Where (濾心更換_日誌.更換日期 BETWEEN
                        TRY_PARSE(STR(TRIM(@年))+'/'+IIF(TRIM(@月)='','1',TRIM(@月))+'/01' AS date)
                    AND
                        TRY_PARSE(STR(TRIM(@年))+'/'+IIF(TRIM(@月)='','12',TRIM(@月))+'/'+STR(Day(EOMONTH((STR(TRIM(@年)))+'/'+IIF(TRIM(@月)='','12',TRIM(@月))+'/01'))) AS date)
                    )
            ORDER BY 濾心更換_日誌.更換日期"
        Insertcommand=""
        UpdateCommand=""
        DeleteCommand="">
        <SelectParameters>
            <asp:ControlParameter ControlID="年" ConvertEmptyStringToNull="False" Name="年" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="月" ConvertEmptyStringToNull="False" Name="月" PropertyName="Text" Type="String"/>
        </SelectParameters>
        <InsertParameters>
        </InsertParameters>
        <UpdateParameters>
        </UpdateParameters>
        <DeleteParameters>
        </DeleteParameters>
    </asp:SqlDataSource>
    <asp:SqlDataSource ID="SqlDataSource2" runat="server" ConnectionString='<%$ ConnectionStrings:ApplicationServices%>'
        SelectCommand="
            SELECT
                DISTINCT YEAR(更換日期) - 1911 AS 民國年,
                YEAR(更換日期) AS 西元年
            FROM 濾心更換_日誌
            ORDER BY 民國年"
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
