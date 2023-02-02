<%@ Page Language="VB" MasterPageFile="./MasterPage.master" MaintainScrollPositionOnPostback="true" AutoEventWireup="false" CodeFile="付款查詢.aspx.vb" Inherits="付款查詢" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
    <asp:ScriptManager ID="ScriptManager1" runat="server" EnablePageMethods = "true" EnableScriptGlobalization="True" EnableScriptLocalization="True"/>
    <link rel="stylesheet" runat="server" media="screen" href="css\付款查詢.css"/>
    <div>
        <h1>
            <a ID="Title" href="付款查詢.aspx">付款查詢<a>
            <asp:DropDownList ID="選項" runat="server" AutoPostBack="True" OnSelectedIndexChanged="選項_SelectedIndexChanged" CssClass="DropDownList">
                <asp:ListItem Text="405" Value="405"></asp:ListItem>
                <asp:ListItem Text="409" Value="409"></asp:ListItem>
            </asp:DropDownList>
        </h1>
    </div>
    <asp:Panel ID="Panel1" runat="server" DefaultButton="搜尋" CssClass="Panel1">
        每頁筆數<asp:TextBox ID="每頁筆數" runat="server" Text="10" CssClass="Input 每頁筆數"/>
        年<asp:DropDownList ID="年A" runat="server" AutoPostBack="True" DataSourceID="SqlDataSource2" DataTextField="年" DataValueField="年" OnSelectedIndexChanged="年A_SelectedIndexChanged" CssClass="DropDownList"/>~<asp:DropDownList ID="年B" runat="server" AutoPostBack="True" DataSourceID="SqlDataSource2" DataTextField="年" DataValueField="年" OnSelectedIndexChanged="年B_SelectedIndexChanged" CssClass="DropDownList"/>
        傳票號碼<asp:TextBox ID="傳票號碼A" runat="server" CssClass="Input 傳票號碼"/>~<asp:TextBox ID="傳票號碼B" runat="server" CssClass="Input 傳票號碼"/>
        摘要<asp:TextBox ID="摘要" runat="server" CssClass="Input 摘要"/>
        付款日<asp:TextBox ID="付款日A" runat="server" CssClass="Input 付款日"/>~<asp:TextBox ID="付款日B" runat="server" CssClass="Input 付款日"/>
        金額<asp:TextBox ID="金額" runat="server" CssClass="Input 金額"/>
        支票編號<asp:TextBox ID="支票編號" runat="server" CssClass="Input 支票編號"/>
        <br>
        銀行名稱<asp:TextBox ID="銀行名稱" runat="server" CssClass="Input 銀行名稱"/>
        帳號<asp:TextBox ID="帳號" runat="server" CssClass="Input 帳號"/>
        名稱<asp:TextBox ID="名稱" runat="server" CssClass="Input 名稱"/>
        <ajaxToolkit:AutoCompleteExtender ID="AutoCompleteExtender1" runat="server" TargetControlID="名稱" 
            CompletionSetCount="10000" CompletionInterval="50" EnableCaching="true" MinimumPrefixLength="1" 
            ServiceMethod="GetMyList" 
            CompletionListCssClass="CompletionList" 
            CompletionListItemCssClass="CompletionListItem" 
            CompletionListHighlightedItemCssClass="CompletionListHighlightedItem"/>
        <asp:Button ID="搜尋" runat="server" Text="搜尋" CssClass="GreenButton"/>
        <asp:Button ID="下載" runat="server" Text="下載" OnClick="Download" CssClass="GreenButton"/>
        <div style="display:inline-block;width:275px;text-align:left;">
            <asp:GridView ID="GridView3" runat="server" HorizontalAlign="Center" DataSourceID="SqlDataSource3" GridLines="None" ShowHeader="false" CssClass="GridView3"/>
        </div>
    </asp:Panel>
    <asp:Panel ID="Panel2" runat="server" CssClass="Panel2">
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="false" CellPadding="1" HorizontalAlign="Center" DataSourceID="SqlDataSource1" DataKeyNames="id" GridLines="None" ShowHeaderWhenEmpty="false" AllowPaging="True" PageSize="20" AllowSorting="True" PagerSettings-PageButtonCount=25 EnableModelValidation="True" AutoGenerateEditButton="false">
            <Columns>
                <asp:TemplateField Visible="false">
                    <ItemTemplate>
                        <asp:Label ID="Label0" runat="server" Text='<%# Bind("id") %>'/>
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="">
                    <ItemTemplate>
                        <asp:Label ID="Label1" runat="server" Text='<%# Container.DataItemIndex + 1 %>' CssClass="Label Label1"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="傳票號碼">
                    <ItemTemplate>
                        <asp:TextBox ID="傳票號碼" runat="server" Text='<%# Bind("傳票號碼") %>' TextMode="MultiLine" CssClass="TextBox 傳票號碼t"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="摘要">
                    <ItemTemplate>
                        <asp:TextBox ID="摘要" runat="server" Text='<%# Bind("摘要") %>' TextMode="MultiLine" CssClass="TextBox 摘要t"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="付款日">
                    <ItemTemplate>
                        <asp:TextBox ID="付款日" runat="server" Text='<%# Bind("付款日") %>' TextMode="MultiLine" CssClass="TextBox 付款日t"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="金額">
                    <ItemTemplate>
                        <asp:TextBox ID="金額" runat="server" Text='<%# Eval("金額", "{0:n0}") %>' TextMode="MultiLine" CssClass="TextBox 金額t"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="支票編號">
                    <ItemTemplate>
                        <asp:TextBox ID="支票編號" runat="server" Text='<%# Bind("支票編號") %>' TextMode="MultiLine" CssClass="TextBox 支票編號t"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="名稱">
                    <ItemTemplate>
                        <asp:TextBox ID="名稱" runat="server" Text='<%# Bind("名稱") %>' TextMode="MultiLine" CssClass="TextBox 名稱t"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="銀行名稱">
                    <ItemTemplate>
                        <asp:TextBox ID="銀行名稱" runat="server" Text='<%# Bind("銀行名稱") %>' TextMode="MultiLine" CssClass="TextBox 銀行名稱t"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="帳號">
                    <ItemTemplate>
                        <asp:TextBox ID="帳號" runat="server" Text='<%# Bind("帳號") %>' TextMode="MultiLine" CssClass="TextBox 帳號t"/>
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
            <PagerSettings  Mode="NumericFirstLast" FirstPageText="<<" PreviousPageText="<" NextPageText=">" LastPageText=">>"/>
        </asp:GridView>
    </asp:Panel>
    <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString='<%$ ConnectionStrings:ApplicationServices%>'
        SelectCommand=" 
            IF @選項 = '405' 
                BEGIN 
                    SELECT 
                        傳票資料.id, 
                        傳票資料.傳票號碼, 
                        現金備查簿.會計科目及摘要 AS 摘要, 
                        傳票資料.預付日期 AS 付款日, 
                        傳票資料.支出金額 AS 金額, 
                        傳票資料.登錄序號 AS 支票編號, 
                        傳票資料.名稱, 
                        傳票資料.匯入銀行名稱 AS 銀行名稱, 
                        傳票資料.匯入帳號 AS 帳號 
                    FROM 傳票資料 INNER JOIN 現金備查簿 ON 傳票資料.年 = 現金備查簿.年 AND 傳票資料.傳票號碼 = 現金備查簿.傳票號碼 
                    WHERE (傳票資料.年 BETWEEN TRIM(@年A) AND TRIM(@年B)) 
                    AND (((''=TRIM(@傳票號碼A) OR 傳票資料.傳票號碼 LIKE N'%'+TRIM(@傳票號碼A)) AND (''=TRIM(@傳票號碼B) OR 傳票資料.傳票號碼 LIKE N'%'+TRIM(@傳票號碼B))) 
                        OR (LEN(TRIM(@傳票號碼A)) = 7 AND LEN(TRIM(@傳票號碼B)) = 7 AND 傳票資料.傳票號碼 BETWEEN TRIM(@傳票號碼A) AND TRIM(@傳票號碼B)) 
                        OR (LEN(TRIM(@傳票號碼A)) = 7 AND LEN(TRIM(@傳票號碼B)) BETWEEN 1 AND 6 AND 傳票資料.傳票號碼 BETWEEN TRIM(@傳票號碼A) AND LEFT(TRIM(@傳票號碼A), 1) + RIGHT('000000' + TRIM(@傳票號碼B), 6)) 
                        OR (LEN(TRIM(@傳票號碼A)) BETWEEN 1 AND 6 AND LEN(TRIM(@傳票號碼B)) = 7 AND 傳票資料.傳票號碼 BETWEEN LEFT(TRIM(@傳票號碼B), 1) + RIGHT('000000' + TRIM(@傳票號碼A), 6) AND TRIM(@傳票號碼B)) 
                        OR (LEN(TRIM(@傳票號碼A)) BETWEEN 1 AND 6 AND LEN(TRIM(@傳票號碼B)) BETWEEN 1 AND 6 AND RIGHT(傳票資料.傳票號碼, 6) BETWEEN RIGHT('000000' + TRIM(@傳票號碼A), 6) AND RIGHT('000000' + TRIM(@傳票號碼B), 6))) 
                    AND (''=TRIM(@摘要) OR 現金備查簿.會計科目及摘要 LIKE N'%' + TRIM(@摘要) + '%') 
                    AND (((''=TRIM(@付款日A) OR 傳票資料.預付日期 LIKE N'%'+TRIM(@付款日A)) AND (''=TRIM(@付款日B) OR 傳票資料.預付日期 LIKE N'%'+TRIM(@付款日B))) 
                        OR (LEN(TRIM(@付款日A)) BETWEEN 1 AND 7 AND LEN(TRIM(@付款日B)) BETWEEN 1 AND 7 AND CAST(RIGHT(傳票資料.預付日期, 4) AS int) BETWEEN CAST(RIGHT(TRIM(@付款日A), 4) AS int) AND CAST(RIGHT(TRIM(@付款日B), 4)AS int))) 
                    AND (''=TRIM(@金額) OR 傳票資料.支出金額 LIKE REPLACE(TRIM(@金額), ',', '')) 
                    AND (''=TRIM(@支票編號) OR 傳票資料.登錄序號 LIKE N'%' + TRIM(@支票編號) + '%') 
                    AND (''=TRIM(@名稱) OR 傳票資料.名稱 LIKE N'%' + TRIM(@名稱) + '%') 
                    AND (''=TRIM(@銀行名稱) OR 傳票資料.匯入銀行名稱 LIKE N'%' + TRIM(@銀行名稱) + '%') 
                    AND (''=TRIM(@帳號) OR 傳票資料.匯入帳號 LIKE N'%'+TRIM(@帳號)+'%') 
                    AND (現金備查簿.收入金額405 > 0 OR 現金備查簿.支出金額405 > 0) 
                    AND (傳票資料.支出金額 > 0 AND 傳票資料.支出金額 IS NOT NULL) 
                    ORDER BY 
                    CASE WHEN 傳票資料.傳票號碼 IS NULL THEN 1 ELSE 0 END, 傳票資料.傳票號碼, 
                    CASE WHEN 傳票資料.預付日期 IS NULL THEN 1 ELSE 0 END, 傳票資料.預付日期, 
                    CASE WHEN 傳票資料.名稱 IS NULL THEN 1 ELSE 0 END, 傳票資料.名稱 
                END 
            ELSE IF @選項 = '409' 
                BEGIN 
                    SELECT 
                        傳票資料.id, 
                        傳票資料.傳票號碼, 
                        現金備查簿.會計科目及摘要 AS 摘要, 
                        傳票資料.預付日期 AS 付款日, 
                        傳票資料.支出金額 AS 金額, 
                        傳票資料.登錄序號 AS 支票編號, 
                        傳票資料.名稱, 
                        傳票資料.匯入銀行名稱 AS 銀行名稱, 
                        傳票資料.匯入帳號 AS 帳號 
                    FROM 傳票資料 INNER JOIN 現金備查簿 ON 傳票資料.年 = 現金備查簿.年 AND 傳票資料.傳票號碼 = 現金備查簿.傳票號碼 
                    WHERE (傳票資料.年 BETWEEN TRIM(@年A) AND TRIM(@年B)) 
                    AND (((''=TRIM(@傳票號碼A) OR 傳票資料.傳票號碼 LIKE N'%'+TRIM(@傳票號碼A)) AND (''=TRIM(@傳票號碼B) OR 傳票資料.傳票號碼 LIKE N'%'+TRIM(@傳票號碼B))) 
                        OR (LEN(TRIM(@傳票號碼A)) = 7 AND LEN(TRIM(@傳票號碼B)) = 7 AND 傳票資料.傳票號碼 BETWEEN TRIM(@傳票號碼A) AND TRIM(@傳票號碼B)) 
                        OR (LEN(TRIM(@傳票號碼A)) = 7 AND LEN(TRIM(@傳票號碼B)) BETWEEN 1 AND 6 AND 傳票資料.傳票號碼 BETWEEN TRIM(@傳票號碼A) AND LEFT(TRIM(@傳票號碼A), 1) + RIGHT('000000' + TRIM(@傳票號碼B), 6)) 
                        OR (LEN(TRIM(@傳票號碼A)) BETWEEN 1 AND 6 AND LEN(TRIM(@傳票號碼B)) = 7 AND 傳票資料.傳票號碼 BETWEEN LEFT(TRIM(@傳票號碼B), 1) + RIGHT('000000' + TRIM(@傳票號碼A), 6) AND TRIM(@傳票號碼B)) 
                        OR (LEN(TRIM(@傳票號碼A)) BETWEEN 1 AND 6 AND LEN(TRIM(@傳票號碼B)) BETWEEN 1 AND 6 AND RIGHT(傳票資料.傳票號碼, 6) BETWEEN RIGHT('000000' + TRIM(@傳票號碼A), 6) AND RIGHT('000000' + TRIM(@傳票號碼B), 6))) 
                    AND (''=TRIM(@摘要) OR 現金備查簿.會計科目及摘要 LIKE N'%' + TRIM(@摘要) + '%') 
                    AND (((''=TRIM(@付款日A) OR 傳票資料.預付日期 LIKE N'%'+TRIM(@付款日A)) AND (''=TRIM(@付款日B) OR 傳票資料.預付日期 LIKE N'%'+TRIM(@付款日B))) 
                        OR (LEN(TRIM(@付款日A)) BETWEEN 1 AND 7 AND LEN(TRIM(@付款日B)) BETWEEN 1 AND 7 AND CAST(RIGHT(傳票資料.預付日期, 4) AS int) BETWEEN CAST(RIGHT(TRIM(@付款日A), 4) AS int) AND CAST(RIGHT(TRIM(@付款日B), 4)AS int))) 
                    AND (''=TRIM(@金額) OR 傳票資料.支出金額 LIKE REPLACE(TRIM(@金額), ',', '')) 
                    AND (''=TRIM(@支票編號) OR 傳票資料.登錄序號 LIKE N'%' + TRIM(@支票編號) + '%') 
                    AND (''=TRIM(@名稱) OR 傳票資料.名稱 LIKE N'%' + TRIM(@名稱) + '%') 
                    AND (''=TRIM(@銀行名稱) OR 傳票資料.匯入銀行名稱 LIKE N'%' + TRIM(@銀行名稱) + '%') 
                    AND (''=TRIM(@帳號) OR 傳票資料.匯入帳號 LIKE N'%'+TRIM(@帳號)+'%') 
                    AND (現金備查簿.支出金額409 > 0) 
                    ORDER BY 
                    CASE WHEN 傳票資料.傳票號碼 IS NULL THEN 1 ELSE 0 END, 傳票資料.傳票號碼, 
                    CASE WHEN 傳票資料.預付日期 IS NULL THEN 1 ELSE 0 END, 傳票資料.預付日期, 
                    CASE WHEN 傳票資料.名稱 IS NULL THEN 1 ELSE 0 END, 傳票資料.名稱 
                END">
        <SelectParameters>
            <asp:ControlParameter ControlID="選項" ConvertEmptyStringToNull="False" Name="選項" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="年A" ConvertEmptyStringToNull="False" Name="年A" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="年B" ConvertEmptyStringToNull="False" Name="年B" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="傳票號碼A" ConvertEmptyStringToNull="False" Name="傳票號碼A" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="傳票號碼B" ConvertEmptyStringToNull="False" Name="傳票號碼B" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="摘要" ConvertEmptyStringToNull="False" Name="摘要" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="付款日A" ConvertEmptyStringToNull="False" Name="付款日A" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="付款日B" ConvertEmptyStringToNull="False" Name="付款日B" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="金額" ConvertEmptyStringToNull="False" Name="金額" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="支票編號" ConvertEmptyStringToNull="False" Name="支票編號" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="名稱" ConvertEmptyStringToNull="False" Name="名稱" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="銀行名稱" ConvertEmptyStringToNull="False" Name="銀行名稱" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="帳號" ConvertEmptyStringToNull="False" Name="帳號" PropertyName="Text" Type="String"/>
        </SelectParameters>
    </asp:SqlDataSource>
    <asp:SqlDataSource ID="SqlDataSource2" runat="server" ConnectionString='<%$ ConnectionStrings:ApplicationServices%>'
        SelectCommand="SELECT DISTINCT 年 FROM 傳票資料 ORDER BY 年 DESC">
    </asp:SqlDataSource>
    <asp:SqlDataSource ID="SqlDataSource3" runat="server" ConnectionString='<%$ ConnectionStrings:ApplicationServices%>'
        SelectCommand=" 
            IF @選項 = '405' 
                BEGIN 
                    SELECT '總數量:' + STR(COUNT(*)) + '筆 總金額: ' + REPLACE(CONVERT(varchar(30), CONVERT(money, SUM(傳票資料.支出金額)), 1), '.00', '元') 
                    FROM 傳票資料 INNER JOIN 現金備查簿 ON 傳票資料.年 = 現金備查簿.年 AND 傳票資料.傳票號碼 = 現金備查簿.傳票號碼 
                    WHERE (傳票資料.年 BETWEEN TRIM(@年A) AND TRIM(@年B)) 
                    AND (((''=TRIM(@傳票號碼A) OR 傳票資料.傳票號碼 LIKE N'%'+TRIM(@傳票號碼A)) AND (''=TRIM(@傳票號碼B) OR 傳票資料.傳票號碼 LIKE N'%'+TRIM(@傳票號碼B))) 
                        OR (LEN(TRIM(@傳票號碼A)) = 7 AND LEN(TRIM(@傳票號碼B)) = 7 AND 傳票資料.傳票號碼 BETWEEN TRIM(@傳票號碼A) AND TRIM(@傳票號碼B)) 
                        OR (LEN(TRIM(@傳票號碼A)) = 7 AND LEN(TRIM(@傳票號碼B)) BETWEEN 1 AND 6 AND 傳票資料.傳票號碼 BETWEEN TRIM(@傳票號碼A) AND LEFT(TRIM(@傳票號碼A), 1) + RIGHT('000000' + TRIM(@傳票號碼B), 6)) 
                        OR (LEN(TRIM(@傳票號碼A)) BETWEEN 1 AND 6 AND LEN(TRIM(@傳票號碼B)) = 7 AND 傳票資料.傳票號碼 BETWEEN LEFT(TRIM(@傳票號碼B), 1) + RIGHT('000000' + TRIM(@傳票號碼A), 6) AND TRIM(@傳票號碼B)) 
                        OR (LEN(TRIM(@傳票號碼A)) BETWEEN 1 AND 6 AND LEN(TRIM(@傳票號碼B)) BETWEEN 1 AND 6 AND RIGHT(傳票資料.傳票號碼, 6) BETWEEN RIGHT('000000' + TRIM(@傳票號碼A), 6) AND RIGHT('000000' + TRIM(@傳票號碼B), 6))) 
                    AND (''=TRIM(@摘要) OR 現金備查簿.會計科目及摘要 LIKE N'%' + TRIM(@摘要) + '%') 
                    AND (((''=TRIM(@付款日A) OR 傳票資料.預付日期 LIKE N'%'+TRIM(@付款日A)) AND (''=TRIM(@付款日B) OR 傳票資料.預付日期 LIKE N'%'+TRIM(@付款日B))) 
                        OR (LEN(TRIM(@付款日A)) BETWEEN 1 AND 7 AND LEN(TRIM(@付款日B)) BETWEEN 1 AND 7 AND CAST(RIGHT(傳票資料.預付日期, 4) AS int) BETWEEN CAST(RIGHT(TRIM(@付款日A), 4) AS int) AND CAST(RIGHT(TRIM(@付款日B), 4)AS int))) 
                    AND (''=TRIM(@金額) OR 傳票資料.支出金額 LIKE REPLACE(TRIM(@金額), ',', '')) 
                    AND (''=TRIM(@支票編號) OR 傳票資料.登錄序號 LIKE N'%' + TRIM(@支票編號) + '%') 
                    AND (''=TRIM(@名稱) OR 傳票資料.名稱 LIKE N'%' + TRIM(@名稱) + '%') 
                    AND (''=TRIM(@銀行名稱) OR 傳票資料.匯入銀行名稱 LIKE N'%' + TRIM(@銀行名稱) + '%') 
                    AND (''=TRIM(@帳號) OR 傳票資料.匯入帳號 LIKE N'%'+TRIM(@帳號)+'%') 
                    AND (現金備查簿.收入金額405 > 0 OR 現金備查簿.支出金額405 > 0) 
                    AND (傳票資料.支出金額 > 0 AND 傳票資料.支出金額 IS NOT NULL) 
                END 
            ELSE IF @選項 = '409' 
                BEGIN 
                    SELECT '總數量:' + STR(COUNT(*)) + '筆 總金額: ' + REPLACE(CONVERT(varchar(30), CONVERT(money, SUM(傳票資料.支出金額)), 1), '.00', '元') 
                    FROM 傳票資料 INNER JOIN 現金備查簿 ON 傳票資料.年 = 現金備查簿.年 AND 傳票資料.傳票號碼 = 現金備查簿.傳票號碼 
                    WHERE (傳票資料.年 BETWEEN TRIM(@年A) AND TRIM(@年B)) 
                    AND (((''=TRIM(@傳票號碼A) OR 傳票資料.傳票號碼 LIKE N'%'+TRIM(@傳票號碼A)) AND (''=TRIM(@傳票號碼B) OR 傳票資料.傳票號碼 LIKE N'%'+TRIM(@傳票號碼B))) 
                        OR (LEN(TRIM(@傳票號碼A)) = 7 AND LEN(TRIM(@傳票號碼B)) = 7 AND 傳票資料.傳票號碼 BETWEEN TRIM(@傳票號碼A) AND TRIM(@傳票號碼B)) 
                        OR (LEN(TRIM(@傳票號碼A)) = 7 AND LEN(TRIM(@傳票號碼B)) BETWEEN 1 AND 6 AND 傳票資料.傳票號碼 BETWEEN TRIM(@傳票號碼A) AND LEFT(TRIM(@傳票號碼A), 1) + RIGHT('000000' + TRIM(@傳票號碼B), 6)) 
                        OR (LEN(TRIM(@傳票號碼A)) BETWEEN 1 AND 6 AND LEN(TRIM(@傳票號碼B)) = 7 AND 傳票資料.傳票號碼 BETWEEN LEFT(TRIM(@傳票號碼B), 1) + RIGHT('000000' + TRIM(@傳票號碼A), 6) AND TRIM(@傳票號碼B)) 
                        OR (LEN(TRIM(@傳票號碼A)) BETWEEN 1 AND 6 AND LEN(TRIM(@傳票號碼B)) BETWEEN 1 AND 6 AND RIGHT(傳票資料.傳票號碼, 6) BETWEEN RIGHT('000000' + TRIM(@傳票號碼A), 6) AND RIGHT('000000' + TRIM(@傳票號碼B), 6))) 
                    AND (''=TRIM(@摘要) OR 現金備查簿.會計科目及摘要 LIKE N'%' + TRIM(@摘要) + '%') 
                    AND (((''=TRIM(@付款日A) OR 傳票資料.預付日期 LIKE N'%'+TRIM(@付款日A)) AND (''=TRIM(@付款日B) OR 傳票資料.預付日期 LIKE N'%'+TRIM(@付款日B))) 
                        OR (LEN(TRIM(@付款日A)) BETWEEN 1 AND 7 AND LEN(TRIM(@付款日B)) BETWEEN 1 AND 7 AND CAST(RIGHT(傳票資料.預付日期, 4) AS int) BETWEEN CAST(RIGHT(TRIM(@付款日A), 4) AS int) AND CAST(RIGHT(TRIM(@付款日B), 4)AS int))) 
                    AND (''=TRIM(@金額) OR 傳票資料.支出金額 LIKE REPLACE(TRIM(@金額), ',', '')) 
                    AND (''=TRIM(@支票編號) OR 傳票資料.登錄序號 LIKE N'%' + TRIM(@支票編號) + '%') 
                    AND (''=TRIM(@名稱) OR 傳票資料.名稱 LIKE N'%' + TRIM(@名稱) + '%') 
                    AND (''=TRIM(@銀行名稱) OR 傳票資料.匯入銀行名稱 LIKE N'%' + TRIM(@銀行名稱) + '%') 
                    AND (''=TRIM(@帳號) OR 傳票資料.匯入帳號 LIKE N'%'+TRIM(@帳號)+'%') 
                    AND (現金備查簿.支出金額409 > 0) 
                END">
        <SelectParameters>
            <asp:ControlParameter ControlID="選項" ConvertEmptyStringToNull="False" Name="選項" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="年A" ConvertEmptyStringToNull="False" Name="年A" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="年B" ConvertEmptyStringToNull="False" Name="年B" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="傳票號碼A" ConvertEmptyStringToNull="False" Name="傳票號碼A" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="傳票號碼B" ConvertEmptyStringToNull="False" Name="傳票號碼B" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="摘要" ConvertEmptyStringToNull="False" Name="摘要" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="付款日A" ConvertEmptyStringToNull="False" Name="付款日A" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="付款日B" ConvertEmptyStringToNull="False" Name="付款日B" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="金額" ConvertEmptyStringToNull="False" Name="金額" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="支票編號" ConvertEmptyStringToNull="False" Name="支票編號" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="名稱" ConvertEmptyStringToNull="False" Name="名稱" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="銀行名稱" ConvertEmptyStringToNull="False" Name="銀行名稱" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="帳號" ConvertEmptyStringToNull="False" Name="帳號" PropertyName="Text" Type="String"/>
        </SelectParameters>
    </asp:SqlDataSource>
</asp:Content>
