<%@ Page Language="VB" MasterPageFile="./MasterPage.master" MaintainScrollPositionOnPostback="true" AutoEventWireup="false" CodeFile="查詢作業.aspx.vb" Inherits="查詢作業"%>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server" >
    <asp:ScriptManager ID="ScriptManager1" runat="server" EnablePageMethods = "true" EnableScriptGlobalization="True" EnableScriptLocalization="True"/>
    <script src='js/查詢作業.js'></script>
    <link rel="stylesheet" runat="server" media="screen" href="css\查詢作業.css"/>
    <style>
    </style>
    <div><h1><a id="Title" href="查詢作業.aspx">查詢作業<a></h1></div>
    <asp:Panel ID="Panel1" runat="server" DefaultButton="搜尋" CssClass="Panel1" >
    品項:<asp:DropDownList ID="品項" runat="server" AutoPostBack="True" DataSourceID="SqlDataSource2" DataTextField="品項" DataValueField="id" CssClass="DropDownList"/>
    建置日期:<asp:DropDownList ID="年" runat="server" AutoPostBack="True" DataSourceID="SqlDataSource3" DataTextField="民國年" DataValueField="西元年" CssClass="DropDownList"/>年
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
    型號<asp:TextBox ID="型號" runat="server" Text='' Maxlength=0 Enabled="True" CssClass="TextBox Input2"/>
    存置地點<asp:TextBox ID="存置地點" runat="server" Text='' Maxlength=0 Enabled="True" CssClass="TextBox Input2"/>
    維護單位<asp:TextBox ID="維護單位" runat="server" Text='' Maxlength=0 Enabled="True" CssClass="TextBox Input2"/>
        <asp:Button ID="清空" runat="server" Text="清空"  OnClick="Clear_Click"  CssClass="GreenButton"/>
        <asp:Button ID="搜尋" runat="server" Text="搜尋" CssClass="GreenButton"/>
        <asp:Button ID="測試" runat="server" Text="測試" OnClick="test" CssClass="GreenButton" Visible="false"/>
        <asp:Label ID="Label1" runat="server" Text="" CssClass="GreenLabel"/>
        <asp:Label ID="Label2" runat="server" Text="" CssClass="RedLabel"/>
    </asp:Panel>
    <asp:Panel ID="Panel3" runat="server" CssClass="Panel3">
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="false" CellPadding="1" HorizontalAlign="Center" DataSourceID="SqlDataSource1" DataKeyNames="id" GridLines="None" ShowHeaderWhenEmpty="false" AllowPaging="True" PageSize="15" AllowSorting="True" PagerSettings-PageButtonCount=25 EnableModelValidation="True" AutoGenerateEditButton="false">
            <Columns>
                <asp:TemplateField HeaderText="id" Visible="false">
                    <ItemTemplate>
                        <asp:label ID="id" runat="server" Text='<%# Eval("id") %>' Maxlength=0 Enabled="False" CssClass="label id"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="" Visible="false">
                    <ItemTemplate>
                        <asp:label ID="_列" runat="server" Text='<%# Eval("_列") %>' Maxlength=0 Enabled="False" CssClass="label _列"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="設備編號">
                    <ItemTemplate>
                        <asp:TextBox ID="設備編號" runat="server" Text='<%# Bind("設備編號") %>' Maxlength=0 Enabled="False" CssClass="TextBox 設備編號"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="品項">
                    <ItemTemplate>
                        <asp:DropDownList ID="品項" runat="server" Text='<%# Bind("ID_品項") %>' DataSourceID="SqlDataSource2" DataTextField="品項" DataValueField="id" Enabled="False" CssClass="DropDownList"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="建置日期">
                    <ItemTemplate>
                        <asp:label ID="建置日期" runat="server" Text='<%# If(IsDate(Eval("建置日期")), (Year(Eval("建置日期"))-1911).ToString() & Eval("建置日期", "{0:/MM/dd}"), "") %>' Maxlength=0 Enabled="True" CssClass="label 建置日期"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="型號">
                    <ItemTemplate>
                        <asp:label ID="型號" runat="server" Text='<%# Eval("型號") %>' Maxlength=0 Enabled="True" CssClass="label 型號"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="存置地點">
                    <ItemTemplate>
                        <asp:label ID="存置地點" runat="server" Text='<%# Eval("存置地點") %>' Maxlength=0 Enabled="True" CssClass="label 存置地點"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="維護單位">
                    <ItemTemplate>
                        <asp:label ID="維護單位" runat="server" Text='<%# Eval("維護單位") %>' Maxlength=0 Enabled="True" CssClass="label 維護單位"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="照片">
                    <ItemTemplate>
                        <asp:ImageButton ID="照片縮圖" runat="server" ImageUrl='<%# If(Eval("照片縮圖").ToString = "", "", Eval("照片縮圖")) %>' CommandName="照片圖" AutoPostBack="False" OnClientClick="" Visible='<%# If(Eval("照片縮圖").ToString = "", "False", "True") %>' CssClass="label 照片"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="維護紀錄">
                    <ItemTemplate>
                        <asp:Button ID="維護紀錄" runat="server" Text="維護紀錄" CommandName="維護紀錄" Enabled="True" CssClass="GreenButton" />
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="備註">
                    <ItemTemplate>
                        <asp:label ID="備註" runat="server" Text='<%# Eval("備註") %>' Maxlength=0 Enabled="True" CssClass="label 備註"/>
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
            SELECT id,_列,_頁,設備編號,Id_品項,建置日期,型號,存置地點,維護單位,照片縮圖,維護紀錄,備註
            FROM 水電機關設備資料表
            WHERE ('19'=TRIM(@品項) OR (ID_品項 LIKE N'%'+TRIM(@品項)+'%'))
            AND (ID_品項<>'19')
            And (''=TRIM(@型號) OR 型號 LIKE N'%'+TRIM(@型號)+'%')
            AND (''=TRIM(@存置地點) OR 存置地點 LIKE N'%'+TRIM(@存置地點)+'%')
            AND (''=TRIM(@維護單位) OR 維護單位 LIKE N'%'+TRIM(@維護單位)+'%')
            AND (TRIM(@年)=''
                OR  (建置日期 BETWEEN
                        TRY_PARSE(STR(TRIM(@年))+'/'+IIF(TRIM(@月1)='','1',TRIM(@月1))+'/'+IIF(TRIM(@日1)='','1',TRIM(@日1)) AS date)
                    AND
                        TRY_PARSE(STR(TRIM(@年))+'/'+IIF(TRIM(@月2)='','12',TRIM(@月2))+'/'+IIF(TRIM(@日2)='',
                        STR(Day(EOMONTH((STR(TRIM(@年)))+'/'+IIF(TRIM(@月2)='','12',TRIM(@月2))+'/01')))
                        ,TRIM(@日2)) AS date)
                    )
                )"
        Insertcommand=""
        UpdateCommand=""
        DeleteCommand="">
        <SelectParameters>
            <asp:ControlParameter ControlID="存置地點" ConvertEmptyStringToNull="False" Name="存置地點" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="品項" ConvertEmptyStringToNull="False" Name="品項" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="維護單位" ConvertEmptyStringToNull="False" Name="維護單位" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="型號" ConvertEmptyStringToNull="False" Name="型號" PropertyName="Text" Type="String"/>
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
        </DeleteParameters>
    </asp:SqlDataSource>
    <asp:SqlDataSource ID="SqlDataSource2" runat="server" ConnectionString='<%$ ConnectionStrings:ApplicationServices%>'
        SelectCommand="
            SELECT * FROM 品項資料表
            ORDER BY 品項"
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
        SelectCommand="
            SELECT
                DISTINCT YEAR(建置日期) - 1911 AS 民國年,
                YEAR(建置日期) AS 西元年
            FROM 水電機關設備資料表
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
