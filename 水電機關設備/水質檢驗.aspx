<%@ Page Language="VB" MasterPageFile="./MasterPage.master" MaintainScrollPositionOnPostback="true" AutoEventWireup="false" CodeFile="水質檢驗.aspx.vb" Inherits="水質檢驗"%>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server" >
    <asp:ScriptManager ID="ScriptManager1" runat="server" EnablePageMethods = "true" EnableScriptGlobalization="True" EnableScriptLocalization="True"/>
    <script src='js/水質檢驗.js'></script>
    <link rel="stylesheet" runat="server" media="screen" href="css\水質檢驗.css"/>
    <style>
    </style>
    <div><h1><a id="Title" href="水質檢驗.aspx">水質檢驗週期<a></h1></div>
    <asp:Panel ID="Panel1" runat="server" DefaultButton="搜尋" CssClass="Panel1" >
        <asp:Button ID="搜尋" runat="server" Text="搜尋" CssClass="GreenButton" Visible="False"/>
        <asp:Button ID="新增" runat="server" Text="新增一頁" OnClick="Insert" CssClass="GreenButton"/>
        <asp:Button ID="存檔" runat="server" Text="存檔" OnClick="Update" CssClass="GreenButton"/>
        <asp:Button ID="下載" runat="server" Text="下載檢驗單" OnClick="Download" CssClass="GreenButton"/>
        <asp:Button ID="刪除" runat="server" Text="刪除末頁" OnClick="Delete" OnClientClick="return confirm('確定刪除?')" CssClass="RedButton"/>
        <%--狀態:<asp:DropDownList ID="狀態" runat="server" AutoPostBack="True" CssClass="DropDownList">
            <asp:ListItem Text="全部" Value=""></asp:ListItem>
            <asp:ListItem Text="需檢驗" Value="需檢驗"></asp:ListItem>
            <asp:ListItem Text="無需檢驗" Value="無需檢驗"></asp:ListItem>
        </asp:DropDownList>
        需檢驗資料:<asp:CheckBox ID="需檢驗資料" runat="server" AutoPostBack="True" CssClass="input"/>--%>
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
                <%--<asp:TemplateField HeaderText="編號">
                    <ItemTemplate>
                        <asp:Label ID="編號" runat="server" Text='<%# Eval("編號") %>' Maxlength=0  CssClass="Label 編號"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>--%>
                <asp:TemplateField HeaderText="檢驗週期">
                    <ItemTemplate>
                        <asp:TextBox ID="檢驗週期" runat="server" Text='<%# Bind("檢驗週期") %>' Maxlength=0 Enabled="True" CssClass="TextBox 檢驗週期"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <%--<asp:TemplateField HeaderText="上次檢驗">
                    <ItemTemplate>
                        <asp:TextBox ID="上次檢驗" runat="server" Text='<%# If(IsDate(Eval("上次檢驗")), (Year(Eval("上次檢驗"))-1911).ToString() & Eval("上次檢驗", "{0:/MM/dd}"), "") %>' Maxlength=0 Enabled="True" CssClass="TextBox 上次檢驗"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="下次檢驗">
                    <ItemTemplate>
                        <asp:Label ID="下次檢驗" runat="server" Text='<%# If(IsDate(Eval("上次檢驗")) AND NOT(Eval("檢驗週期")Is DBNull.Value),(Year(CDate(DATEADD("m",CInt(left(Eval("檢驗週期"),1)),CDate(Eval("上次檢驗")))))-1911).ToString() & (CDate(DATEADD("m",CInt(left(Eval("檢驗週期"),1)),CDate(Eval("上次檢驗"))))).ToString("/MM/dd") , "") %>' Maxlength=0 Enabled="True" CssClass="Label 下次檢驗"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>--%>
                <asp:TemplateField HeaderText="檢驗月份">
                    <ItemTemplate>
                        <asp:TextBox ID="檢驗日期" runat="server" Text='<%# If(IsDate(Eval("檢驗日期")), (Year(Eval("檢驗日期"))-1911).ToString() & Eval("檢驗日期", "{0:-M}"), "") %>' Maxlength=0 Enabled="True" placeholder="ex:111-1" CssClass="TextBox 檢驗日期"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="檢驗地點">
                    <ItemTemplate>
                        <asp:TextBox ID="檢驗地點" runat="server" Text='<%# Bind("檢驗地點") %>' Maxlength=0 Enabled="True" TextMode="MultiLine" CssClass="TextBox 檢驗地點"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="檢驗項目">
                    <ItemTemplate>
                        <asp:TextBox ID="檢驗項目" runat="server" Text='<%# Bind("檢驗項目") %>' Maxlength=0 Enabled="True" TextMode="MultiLine" CssClass="TextBox 檢驗項目"/>
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
            FROM 水質檢驗表
            --Where (@狀態='' OR (@狀態='需檢驗' AND (REPLACE(檢驗週期,'個月',''))<=DateDiff(m,上次檢驗,GETDATE()))OR (@狀態='無需檢驗' AND Not((REPLACE(檢驗週期,'個月',''))<=DateDiff(m,上次檢驗,GETDATE()))))
            ORDER BY _頁, _列"
        Insertcommand="INSERT INTO 設備 (數量) VALUES (0)"
        UpdateCommand=""
        DeleteCommand="">
        <SelectParameters>
            <%--<asp:ControlParameter ControlID="狀態" ConvertEmptyStringToNull="False" Name="狀態" PropertyName="Text" Type="String"/>--%>
        </SelectParameters>
        <InsertParameters>
        </InsertParameters>
        <UpdateParameters>
        </UpdateParameters>
        <DeleteParameters>
        </DeleteParameters>
    </asp:SqlDataSource>
</asp:Content>
