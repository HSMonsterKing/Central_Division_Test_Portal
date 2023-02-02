<%@ Page Language="VB" MasterPageFile="./MasterPage.master" MaintainScrollPositionOnPostback="true" AutoEventWireup="false" CodeFile="更新資料.aspx.vb" Inherits="更新資料"%>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server" >
    <asp:ScriptManager ID="ScriptManager1" runat="server" EnablePageMethods = "true" EnableScriptGlobalization="True" EnableScriptLocalization="True"/>
    <script src='js/更新資料.js'></script>
    <link rel="stylesheet" runat="server" media="screen" href="css\更新資料.css"/>
    <style>
    </style>
    <div><h1><a id="Title" href="更新資料.aspx">更新資料<a></h1></div>
    <asp:Panel ID="Panel1" runat="server" DefaultButton="搜尋" CssClass="Panel1">    
    更新版本:<asp:DropDownList ID="更新版本" runat="server" AutoPostBack="True" DataSourceID="SqlDataSource3" DataTextField="更新版本" DataValueField="更新版本" CssClass="DropDownList"/>
    更新日期:<asp:DropDownList ID="年" runat="server" AutoPostBack="True" DataSourceID="SqlDataSource2" DataTextField="民國年" DataValueField="西元年" CssClass="DropDownList"/>年
    <asp:DropDownList ID="月1" runat="server" AutoPostBack="True" CssClass="DropDownList">
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
    <asp:DropDownList ID="日1" runat="server" AutoPostBack="True" CssClass="DropDownList">
    </asp:DropDownList>日~
    <asp:DropDownList ID="月2" runat="server" AutoPostBack="True" CssClass="DropDownList">
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
    <asp:DropDownList ID="日2" runat="server" AutoPostBack="True" CssClass="DropDownList">
    </asp:DropDownList>日
    內容<asp:TextBox ID="內容" runat="server" Text='' Maxlength=0 Enabled="True" CssClass="TextBox Input1"/>
        <asp:Button ID="搜尋" runat="server" Text="搜尋" CssClass="GreenButton"/>
        <asp:Button ID="新增" runat="server" Text="新增一列" OnClick="Insert" Visible="False" CssClass="GreenButton"/>
        <asp:Button ID="存檔" runat="server" Text="存檔" OnClick="Update" Visible="False" CssClass="GreenButton"/>
        <asp:Button ID="測試" runat="server" Text="測試" OnClick="test" CssClass="GreenButton" Visible="False"/>
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
                <asp:TemplateField HeaderText="更新日期">
                    <ItemTemplate>
                        <asp:TextBox ID="更新日期" runat="server" Text='<%# If(IsDate(Eval("更新日期")), (Year(Eval("更新日期"))-1911).ToString() & Eval("更新日期", "{0:/MM/dd}"), "") %>' Maxlength=0 Enabled="False" CssClass="TextBox 更新日期"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="更新版本">
                    <ItemTemplate>
                        <asp:TextBox ID="更新版本" runat="server" Text='<%# Bind("更新版本") %>' Maxlength=0 Enabled="False" CssClass="TextBox 更新版本"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="內容">
                    <ItemTemplate>
                        <asp:TextBox ID="內容" runat="server" Text='<%# Bind("內容") %>' Maxlength=0 Enabled="False" TextMode="MultiLine" CssClass="TextBox 內容"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="" Visible="false">
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
            FROM 更新資料表
            WHERE (''=TRIM(@更新版本) OR 更新版本 LIKE N'%'+TRIM(@更新版本)+'%')
            AND (''=TRIM(@內容) OR 內容 LIKE N'%'+TRIM(@內容)+'%')
            AND (TRIM(@年)=''
                OR  (更新日期 BETWEEN
                        TRY_PARSE(STR(TRIM(@年))+'/'+IIF(TRIM(@月1)='','1',TRIM(@月1))+'/'+IIF(TRIM(@日1)='','1',TRIM(@日1)) AS date)
                    AND
                        TRY_PARSE(STR(TRIM(@年))+'/'+IIF(TRIM(@月2)='','12',TRIM(@月2))+'/'+IIF(TRIM(@日2)='',
                        STR(Day(EOMONTH((STR(TRIM(@年)))+'/'+IIF(TRIM(@月2)='','12',TRIM(@月2))+'/01')))
                        ,TRIM(@日2)) AS date)
                    )
                )
            Order By id Desc"
        Insertcommand=""
        UpdateCommand=""
        DeleteCommand="">
        <SelectParameters>
            <asp:ControlParameter ControlID="更新版本" ConvertEmptyStringToNull="False" Name="更新版本" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="內容" ConvertEmptyStringToNull="False" Name="內容" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="年" ConvertEmptyStringToNull="False" Name="年" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="月1" ConvertEmptyStringToNull="False" Name="月1" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="日1" ConvertEmptyStringToNull="False" Name="日1" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="月2" ConvertEmptyStringToNull="False" Name="月2" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="日2" ConvertEmptyStringToNull="False" Name="日2" PropertyName="Text" Type="String"/>
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
            SELECT
                DISTINCT YEAR(更新日期) - 1911 AS 民國年,
                YEAR(更新日期) AS 西元年
            FROM 更新資料表
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
    <asp:SqlDataSource ID="SqlDataSource3" runat="server" ConnectionString='<%$ ConnectionStrings:ApplicationServices%>'
        SelectCommand="SELECT 更新版本
            FROM 更新資料表"
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
