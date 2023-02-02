<%@ Page Language="VB" MasterPageFile="./MasterPage.master" MaintainScrollPositionOnPostback="true" AutoEventWireup="false" CodeFile="維護契約.aspx.vb" Inherits="維護契約"%>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
    <asp:ScriptManager ID="ScriptManager1" runat="server" EnablePageMethods = "true" EnableScriptGlobalization="True" EnableScriptLocalization="True"/>
    <script src='js/維護契約.js'></script>
    <link rel="stylesheet" runat="server" media="screen" href="css\維護契約.css"/>
    <style>
    </style>
    <div><h1><a id="Title" href="維護契約.aspx">維護契約<a></h1></div>
    <asp:TextBox ID="ID_維修契約" runat="server" Maxlength=3 CssClass="Input1" visible="False"/>
    <asp:Panel ID="Panel1" runat="server" DefaultButton="搜尋" CssClass="Panel1">
        廠商<asp:TextBox ID="廠商" runat="server" Maxlength=3 CssClass="Input1"/>
        <ajaxToolkit:AutoCompleteExtEnder ID="廠商自動" runat="server" TargetControlID="廠商"
                            CompletionSetCount="10000" CompletionInterval="50" EnableCaching="true" MinimumPrefixLength="1"
                            ServiceMethod="GetMyList"
                            CompletionListCssClass="CompletionList"
                            CompletionListItemCssClass="CompletionListItem"
                            CompletionListHighlightedItemCssClass="CompletionListHighlightedItem"/>
        <asp:Button ID="搜尋" runat="server" Text="搜尋" CssClass="GreenButton"/>
        <asp:Button ID="新增" runat="server" Text="新增一頁" OnClick="Insert" CssClass="GreenButton"/>
        <asp:Button ID="存檔" runat="server" Text="存檔" OnClick="Update" CssClass="GreenButton"/>
        <asp:Button ID="刪除" runat="server" Text="刪除末頁" OnClick="Delete" OnClientClick="return confirm('確定刪除?')" CssClass="RedButton"/>
        <asp:FileUpload ID="FileUpload1" runat="server" AllowMultiple="True" CssClass="UploadButton" />
        <asp:Button ID="測試" runat="server" Text="測試" OnClick="test" CssClass="GreenButton" Visible="false"/>
        <asp:Label ID="Label1" runat="server" Text="" CssClass="GreenLabel"/>
        <asp:Label ID="Label2" runat="server" Text="" CssClass="RedLabel"/>
    </asp:Panel>
    <asp:Panel ID="Panel3" runat="server" CssClass="Panel3" DefaultButton="存檔">
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="false" CellPadding="1" HorizontalAlign="Center" DataSourceID="SqlDataSource1" DataKeyNames="id" GridLines="None" ShowHeaderWhenEmpty="false" AllowPaging="True" PageSize="15" AllowSorting="True" PagerSettings-PageButtonCount=25 EnableModelValidation="True" AutoGenerateEditButton="false">
            <Columns>
            <asp:TemplateField HeaderText="" Visible="False">
                    <ItemTemplate>
                        <asp:TextBox ID="id" runat="server" Text='<%# Bind("id") %>' Maxlength=0 Enabled="False" CssClass="TextBox id"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="編號">
                    <ItemTemplate>
                        <asp:TextBox ID="_列" runat="server" Text='<%# Bind("_列") %>' Maxlength=0 Enabled="False" CssClass="TextBox _列"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="契約名稱">
                    <ItemTemplate>
                        <asp:TextBox ID="契約名稱" runat="server" Text='<%# Bind("契約名稱") %>' Maxlength=0 Enabled="True" CssClass="TextBox 契約名稱"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="ID_廠商" Visible="False">
                    <ItemTemplate>
                        <asp:TextBox ID="ID_廠商" runat="server" Text='<%# Bind("ID_廠商") %>' Maxlength=0 Enabled="True" CssClass="TextBox ID_廠商"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="維護廠商">
                    <ItemTemplate>
                        <asp:TextBox ID="維護廠商" runat="server" Text='<%# Bind("廠商") %>' Maxlength=0 Enabled="True" CssClass="TextBox 維護廠商"/>
                        <ajaxToolkit:AutoCompleteExtEnder ID="維護廠商自動" runat="server" TargetControlID="維護廠商"
                            CompletionSetCount="10000" CompletionInterval="50" EnableCaching="true" MinimumPrefixLength="1"
                            ServiceMethod="GetMyList"
                            CompletionListCssClass="CompletionList"
                            CompletionListItemCssClass="CompletionListItem"
                            CompletionListHighlightedItemCssClass="CompletionListHighlightedItem"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="廠商電話">
                    <ItemTemplate>
                        <asp:TextBox ID="廠商電話" runat="server" Text='<%# Bind("電話") %>' Maxlength=0 Enabled="True" CssClass="TextBox 廠商電話"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="維護內容">
                    <ItemTemplate>
                        <asp:TextBox ID="維護內容" runat="server" Text='<%# Bind("維護內容") %>' Maxlength=0 Enabled="True" CssClass="TextBox 維護內容"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="維護頻率">
                    <ItemTemplate>
                        <asp:DropDownList ID="維護頻率" runat="server" Text='<%# Bind("維護頻率") %>' AutoPostBack="True" CssClass="DropDownList">
                            <asp:ListItem Text="" Value=""></asp:ListItem>
                            <asp:ListItem Text="每月" Value="每月"></asp:ListItem>
                            <asp:ListItem Text="每季" Value="每季"></asp:ListItem>
                            <asp:ListItem Text="每半年" Value="每半年"></asp:ListItem>
                            <asp:ListItem Text="每年" Value="每年"></asp:ListItem>
                        </asp:DropDownList>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="上傳維護紀錄">
                    <ItemTemplate>
                        <asp:TextBox ID="ID_維護紀錄" runat="server" Text='<%# Bind("ID_維護紀錄") %>' Maxlength=0 Enabled="True" CssClass="TextBox" visible="False"/>
                        <asp:Button ID="上傳資料" runat="server" Text="上傳資料" CommandName="上傳資料" Enabled="True" CssClass="GreenButton"/>
                        <asp:Button ID="維護紀錄" runat="server" Text="維護紀錄" CommandName="維護紀錄" Enabled="True" CssClass="GreenButton"/>
                        </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="備註">
                    <ItemTemplate>
                        <asp:TextBox ID="備註" runat="server" Text='<%# Bind("備註") %>' Maxlength=0 Enabled="True" CssClass="TextBox 備註"/>
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
    <asp:Panel ID="Panel4" runat="server" CssClass="Panel4" DefaultButton="存檔" visible="False">
        <BR><asp:Button ID="返回" runat="server" Text="返回"  OnClick="返回_Click" CssClass="GreenButton"/>
        <BR>契約名稱:<asp:Label ID="Label4_1" runat="server" Text="" />
        <asp:GridView ID="GridView2" runat="server" AutoGenerateColumns="false" CellPadding="1" HorizontalAlign="Center" DataSourceID="SqlDataSource2" DataKeyNames="id" GridLines="None" ShowHeaderWhenEmpty="false" AllowPaging="True" PageSize="15" AllowSorting="True" PagerSettings-PageButtonCount=25 EnableModelValidation="True" AutoGenerateEditButton="false">
            <Columns>
            <asp:TemplateField HeaderText="" Visible="false">
                    <ItemTemplate>
                        <asp:TextBox ID="id" runat="server" Text='<%# Bind("id") %>' Maxlength=0 Enabled="False" CssClass="TextBox id"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="檔名">
                    <ItemTemplate>
                        <asp:HyperLink ID="檔名" runat="server" NavigateUrl='<%# If(Eval("維護契約_維修紀錄").ToString = "", "", "data/\維修契約維護紀錄/" & Eval("維護契約_維修紀錄")) %>' Text='<%# Eval("維護契約_維修紀錄") %>' CssClass="檔名"></asp:HyperLink>
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
            SELECT * FROM 維護契約表 left Join 廠商資料 on 維護契約表.ID_廠商 =廠商資料.ID
            WHERE (''=TRIM(@廠商) OR 廠商資料.廠商 LIKE N'%'+TRIM(@廠商)+'%')
            ORDER BY 維護契約表.ID"
        Insertcommand=""
        UpdateCommand=""
        DeleteCommand="">
        <SelectParameters>
            <asp:ControlParameter ControlID="廠商" ConvertEmptyStringToNull="False" Name="廠商" PropertyName="Text" Type="String"/>
        </SelectParameters>
        <InsertParameters>
        </InsertParameters>
        <UpdateParameters>
        </UpdateParameters>
        <DeleteParameters>
            <asp:Parameter Name="id"/>
        </DeleteParameters>
    </asp:SqlDataSource>
    <asp:SqlDataSource ID="SqlDataSource2" runat="server" ConnectionString='<%$ ConnectionStrings:ApplicationServices%>'
        SelectCommand="
            SELECT * FROM 維護契約表_維護紀錄
            WHERE TRIM(@ID_維修契約)<>'' AND 維護契約表_維護紀錄.ID_維修契約 = TRIM(@ID_維修契約)
            ORDER BY 維護契約表_維護紀錄.ID"
        Insertcommand=""
        UpdateCommand=""
        DeleteCommand="Delete From 維護契約表_維護紀錄">
        <SelectParameters>
            <asp:ControlParameter ControlID="ID_維修契約" ConvertEmptyStringToNull="False" Name="ID_維修契約" PropertyName="Text" Type="String"/>
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