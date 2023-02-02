<%@ Page Language="VB" MasterPageFile="./MasterPage.master" MaintainScrollPositionOnPostback="true" AutoEventWireup="false" CodeFile="濾心清單.aspx.vb" Inherits="濾心清單"%>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server" >
    <asp:ScriptManager ID="ScriptManager1" runat="server" EnablePageMethods = "true" EnableScriptGlobalization="True" EnableScriptLocalization="True"/>
    <script src='js/濾心清單.js'></script>
    <link rel="stylesheet" runat="server" media="screen" href="css\濾心清單.css"/>
    <style>
    </style>
    <div><h1><a id="Title" href="濾心清單.aspx">濾心更換週期<a></h1></div>
    <asp:Panel ID="Panel1" runat="server" DefaultButton="搜尋" CssClass="Panel1" >
        <asp:Button ID="搜尋" runat="server" Text="搜尋" CssClass="GreenButton" Visible="False"/>
        <asp:Button ID="新增" runat="server" Text="新增一頁" OnClick="Insert" CssClass="GreenButton"/>
        <asp:Button ID="存檔" runat="server" Text="存檔" OnClick="Update" CssClass="GreenButton"/>
        <asp:Button ID="下載" runat="server" Text="下載報價單" OnClick="Download" CssClass="GreenButton"/>
        <asp:Button ID="下載2" runat="server" Text="下載更換清單" OnClick="Download2" CssClass="GreenButton"/>
        <asp:Button ID="刪除" runat="server" Text="刪除末頁" OnClick="Delete" OnClientClick="return confirm('確定刪除?')" CssClass="RedButton"/>
        狀態:<asp:DropDownList ID="狀態" runat="server" AutoPostBack="True" CssClass="DropDownList">
            <asp:ListItem Text="全部" Value=""></asp:ListItem>
            <asp:ListItem Text="下次需更換" Value="下次需更換"></asp:ListItem>
            <asp:ListItem Text="下次無需更換" Value="下次無需更換"></asp:ListItem>
        </asp:DropDownList>
        <%--需更換資料:<asp:CheckBox ID="需更換資料" runat="server" AutoPostBack="True" CssClass="input"/>--%>
        <asp:Button ID="測試" runat="server" Text="轉換" CssClass="GreenButton" OnClick="Test" Visible="False"/>
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
                <asp:TemplateField HeaderText="列">
                    <ItemTemplate>
                        <asp:Label ID="_列" runat="server" Text='<%# Eval("_列") %>' Maxlength=0 CssClass="Label _列"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="編號">
                    <ItemTemplate>
                        <asp:Label ID="編號" runat="server" Text='<%# Eval("編號") %>' Maxlength=0  CssClass="Label 編號"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="項目">
                    <ItemTemplate>
                        <asp:TextBox ID="項目" runat="server" Text='<%# Bind("項目") %>' Maxlength=0 Enabled="True" CssClass="TextBox 項目"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="品名型號">
                    <ItemTemplate>
                        <asp:TextBox ID="品名型號" runat="server" Text='<%# Bind("品名型號") %>' Maxlength=0 Enabled="True" TextMode="MultiLine" CssClass="TextBox 品名型號"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="更換週期">
                    <ItemTemplate>
                        <asp:TextBox ID="更換週期" runat="server" Text='<%# Bind("更換週期") %>' Maxlength=0 Enabled="True" placeholder="ex:1個月" CssClass="TextBox 更換週期"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="上次更換">
                    <ItemTemplate>
                        <asp:TextBox ID="上次更換" runat="server" Text='<%# If(IsDate(Eval("上次更換")), (Year(Eval("上次更換"))-1911).ToString() & Eval("上次更換", "{0:/M}"), "") %>' Maxlength=0 Enabled="True"  placeholder="ex:111/7" CssClass="TextBox 上次更換"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="下次更換">
                    <ItemTemplate>
                        <asp:Label ID="下次更換" runat="server" Text='<%# If(IsDate(Eval("上次更換")) AND NOT(Eval("更換週期")Is DBNull.Value) ,If(IsNumeric(Replace(Eval("更換週期"),"個月","")),(Year(CDate(DATEADD("m",CInt(Replace(Eval("更換週期"),"個月","")),CDate(Eval("上次更換")))))-1911).ToString() & "/" & (Month(CDate(DATEADD("m",CInt(Replace(Eval("更換週期"),"個月","")),CDate(Eval("上次更換")))))).ToString() , "") , "") %>' Maxlength=0 Enabled="True" CssClass="Label 下次更換"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="更換日期">
                    <ItemTemplate>
                        <asp:TextBox ID="更換日期" runat="server" Text='<%# If(IsDate(Eval("更換日期")), (Year(Eval("更換日期"))-1911).ToString() & Eval("更換日期", "{0:/MM/dd}"), "") %>' Maxlength=0 Enabled="True" CssClass="TextBox 更換日期"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="單價">
                    <ItemTemplate>
                        <asp:TextBox ID="單價" runat="server" Text='<%# Bind("單價", "{0:n0}") %>' Maxlength=0 Enabled="True" CssClass="TextBox 單價"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="數量">
                    <ItemTemplate>
                        <asp:TextBox ID="數量" runat="server" Text='<%# Bind("數量") %>' Maxlength=0 Enabled="True" CssClass="TextBox 數量"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="合計">
                    <ItemTemplate>
                        <asp:Label ID="合計" runat="server" Text='<%# If(NOT(Eval("單價")Is DBNull.Value) AND NOT(Eval("數量")Is DBNull.Value),(CInt(Eval("單價")*Eval("數量"))).ToString("n0"),"") %>' Maxlength=0 Enabled="True" CssClass="Label 合計"/>
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
            SELECT *
	        FROM 濾心清單表
             Where (@狀態='' OR 
				(@狀態='下次需更換' 
					AND (REPLACE(更換週期,'個月','')-DateDiff(m,上次更換,GETDATE())) in
					(select min(REPLACE(更換週期,'個月','')-DateDiff(m,上次更換,GETDATE())) FROM 濾心清單表)
             OR (@狀態='下次無需更換' AND NOT((REPLACE(更換週期,'個月','')-DateDiff(m,上次更換,GETDATE())) in
                (select min(REPLACE(更換週期,'個月','')-DateDiff(m,上次更換,GETDATE()))FROM 濾心清單表)))))
            ORDER BY _頁, _列"
        Insertcommand="INSERT INTO 設備 (數量) VALUES (0)"
        UpdateCommand=""
        DeleteCommand="">
        <SelectParameters>
            <asp:ControlParameter ControlID="狀態" ConvertEmptyStringToNull="False" Name="狀態" PropertyName="Text" Type="String"/>
        </SelectParameters>
        <InsertParameters>
        </InsertParameters>
        <UpdateParameters>
        </UpdateParameters>
        <DeleteParameters>
        </DeleteParameters>
    </asp:SqlDataSource>
</asp:Content>
