<%@ Page Language="VB" MasterPageFile="./MasterPage.master" MaintainScrollPositionOnPostback="true" AutoEventWireup="false" CodeFile="核銷支出明細備查簿.aspx.vb" Inherits="核銷支出明細備查簿" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
    <asp:ScriptManager ID="ScriptManager1" runat="server" EnablePageMethods = "true" EnableScriptGlobalization="True" EnableScriptLocalization="True"/>
    <script src='js/核銷支出明細備查簿.js'></script>
    <link rel="stylesheet" runat="server" media="screen" href="css\核銷支出明細備查簿.css"/>
    <div><h1><a id="Title" href="核銷支出明細備查簿.aspx">核銷送交主計室明細表<a></h1></div>
    <asp:Panel ID="Panel1" runat="server" DefaultButton="搜尋" CssClass="Panel1">
        年<asp:TextBox ID="年" runat="server" Maxlength=3 CssClass="Input2"/>
        種類<asp:DropDownList ID="_種類" runat="server" AutoPostBack="True" CssClass="DropDownList">
            <asp:ListItem Text="A" Value="A"></asp:ListItem>
            <asp:ListItem Text="B" Value="B"></asp:ListItem>
            <asp:ListItem Text="XZ" Value="XZ"></asp:ListItem>
        </asp:DropDownList>
        等帳日期
        <asp:DropDownList ID="月1" runat="server" AutoPostBack="True" CssClass="DropDownList">
            <asp:ListItem Text="" Value=""></asp:ListItem>
            <asp:ListItem Text="1" Value="1" Selected="True"></asp:ListItem>
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
            <asp:ListItem Text="12" Value="12" Selected="True"></asp:ListItem>
        </asp:DropDownList>月
        <asp:DropDownList ID="日2" runat="server" AutoPostBack="True" CssClass="DropDownList">
        </asp:DropDownList>日
        號數<asp:TextBox ID="號數1" runat="server" OnTextChanged="號數1_TextChanged" Maxlength=3 CssClass="Input2"/>
        ~號數<asp:TextBox ID="號數2" runat="server" OnTextChanged="號數2_TextChanged" Maxlength=3 CssClass="Input2"/>
        <BR>
        商號<asp:TextBox ID="商號" runat="server" Text='' Maxlength=0 Enabled="True" CssClass="Input2"/>
        <ajaxToolkit:AutoCompleteExtender ID="商號自動" runat="server" TargetControlID="商號" 
                            CompletionSetCount="10000" CompletionInterval="50" EnableCaching="true" MinimumPrefixLength="1" 
                            ServiceMethod="GetMyList" 
                            CompletionListCssClass="CompletionList" 
                            CompletionListItemCssClass="CompletionListItem" 
                            CompletionListHighlightedItemCssClass="CompletionListHighlightedItem"/>
        摘要<asp:TextBox ID="摘要" runat="server" Text='' Maxlength=0 Enabled="True" CssClass="Input1"/>
        <ajaxToolkit:AutoCompleteExtender ID="摘要自動" runat="server" TargetControlID="摘要" 
                            CompletionSetCount="10000" CompletionInterval="50" EnableCaching="true" MinimumPrefixLength="1" 
                            ServiceMethod="GetMyList" 
                            CompletionListCssClass="CompletionList" 
                            CompletionListItemCssClass="CompletionListItem" 
                            CompletionListHighlightedItemCssClass="CompletionListHighlightedItem"/>
        狀態<asp:DropDownList ID="狀態" runat="server" AutoPostBack="True" CssClass="DropDownList">
            <asp:ListItem Text="已經主任審核" Value="False"></asp:ListItem>
            <asp:ListItem Text="已經主計室審核" Value="True"></asp:ListItem>
            <asp:ListItem Text="未經主任審核及所有資料" Value=""></asp:ListItem>
        </asp:DropDownList>
        <asp:Button ID="搜尋" runat="server" Text="搜尋" CssClass="GreenButton"/>
        <asp:Button ID="清空" runat="server" Text="清空" OnClick="Clear_Click" CssClass="GreenButton"/>
        <asp:Button ID="下載" runat="server" Text="下載"  OnClick="Download" CssClass="GreenButton"/>
        <asp:Button ID="測試" runat="server" Text="測試" OnClick="test" CssClass="GreenButton" Visible="false"/><BR>
        <asp:Label ID="Label1" runat="server" Text="" CssClass="GreenLabel"/>
        <asp:Label ID="Label2" runat="server" Text="" CssClass="RedLabel"/>
    </asp:Panel>
    <asp:Panel ID="Panel3" runat="server" DefaultButton="" CssClass="Panel3">
    全選:<asp:CheckBox ID="全選" runat="server" OnCheckedChanged="全選_CheckedChanged" AutoPostBack="True"/>
        <asp:Button ID="送交" runat="server" Text="送交" OnClick="send" CssClass="GreenButton"/>
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="false" CellPadding="1" HorizontalAlign="Center" DataSourceID="SqlDataSource1" DataKeyNames="id" GridLines="None" ShowHeaderWhenEmpty="false" AllowPaging="True" PageSize="20" AllowSorting="True" PagerSettings-PageButtonCount=25 EnableModelValidation="True" AutoGenerateEditButton="false">
            <Columns>
                <asp:TemplateField HeaderText="id" Visible="false">
                    <ItemTemplate>
                        <asp:Label ID="id" runat="server" Text='<%# Eval("id") %>' Maxlength=0 Enabled="False" CssClass="Label id"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="" Visible="false">
                    <ItemTemplate>
                        <asp:Label ID="_列" runat="server" Text='<%# Eval("_列") %>' Maxlength=0 Enabled="False" CssClass="Label _列"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="等帳日期">
                    <ItemTemplate>
                        <asp:Label ID="等帳日期" runat="server" Text='<%# Eval("年","{0:000}")+If(Eval("月").ToString = "","XX",Eval("月","{0:00}"))+If(Eval("日").ToString = "","XX",Eval("日","{0:00}")) %>' Maxlength=0 Enabled="True" CssClass="Label 等帳日期"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="號數">
                    <ItemTemplate>
                        <asp:Label ID="號數" runat="server" Text='<%# Eval("號數", "{0:000}") %>' Maxlength=0 Enabled="True" CssClass="Label 號數"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="廠商名稱">
                    <ItemTemplate>
                        <asp:Label ID="廠商名稱" runat="server" Text='<%# Eval("商號") %>' Maxlength=0 Enabled="True" CssClass="Label 廠商名稱"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="金額">
                    <ItemTemplate>
                        <asp:Label ID="金額" runat="server" Text='<%# Eval("支出2", "{0:c0}") %>' Maxlength=0 Enabled="True" CssClass="Label 金額"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="摘要">
                    <ItemTemplate>
                        <asp:Label ID="摘要" runat="server" Text='<%# Eval("摘要") %>' TextMode="MultiLine" Maxlength=0 Enabled="True" CssClass="Label 摘要"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="勾選送交主計室">
                    <ItemTemplate>
                    <asp:CheckBox ID="勾選" runat="server" AutoPostBack="True" OnCheckedChanged="勾選_CheckedChanged" Enabled='<%# If ((Eval("過審").ToString()="True" And Eval("鎖定").ToString()="True") ,If (Eval("回覆").ToString() = "False", True, False),False) %>'/>
                    <asp:TextBox ID="主計室日期" runat="server" Text='<%# If (IsDate(Eval("送交主計室日期")), Eval("送交主計室日期"), "") %>' Maxlength=0 Enabled="False" CssClass="TextBox 等帳日期"/><%--此段因資料庫中的資料為文字非日期，所以無法套用國民年，無法套用Label會有排版問題--%>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="狀態">
                    <ItemTemplate>
                    <asp:Label ID="審核狀態" runat="server"  Text='<%# 
                    If (Eval("鎖定").ToString = "True", 
                    If (Eval("過審").ToString = "True", 
                    If (Eval("回覆").ToString = "True", "主計室通過", 
                    If (Eval("送交主計室日期").ToString="", "通過", "送交主計室")),"已送審"), 
                    If (Eval("送出").ToString = "True", "駁回", "未送審"))%>' Maxlength=0 CssClass="Label 審核狀態"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="主計室回復情況">
                    <ItemTemplate>
                    <asp:RadioButtonList ID="回覆R" runat="server" AutoPostBack="True" Enabled="False" SelectedIndex='<%# If (Eval("回覆").ToString() = "True", 0, 1) %>' RepeatDirection="Horizontal">
                    <asp:ListItem>完成</asp:ListItem>
                    <asp:ListItem>未完成</asp:ListItem>
                    </asp:RadioButtonList>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="主計室簽核">
                    <ItemTemplate>
                    <asp:ImageButton ID="主計室簽核" runat="server" ImageUrl='<%# If (Eval("簽章").ToString = "", "", 
                    If (Eval("簽章").ToString = "2808", "./image/png/丁燕雪.png",
                    If (Eval("簽章").ToString = "2897", "./image/png/林容如.png",
                    If (Eval("簽章").ToString = "2808_1", "./image/png/主計室職章1.png", 
                    If (Eval("簽章").ToString = "2897_1", "./image/png/主計室職章2.png", ""
                    ))))) %>' AutoPostBack="False" OnClientClick="Alert('123456');return false;" visible='<%# If (Eval("簽章").ToString = "" Or Eval("回覆").ToString = "False", "False", "True") %>' CssClass="TextBox 主計室核章"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="勾選下載">
                    <ItemTemplate>
                    <asp:CheckBox ID="勾選下載" runat="server" AutoPostBack="True" OnCheckedChanged="勾選下載_CheckedChanged" />
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
            SELECT * ,(CASE WHEN (left(REPLACE(REPLACE(a.摘要,' ',''),CHAR(13)+CHAR(10),''), 10) Like substring(REPLACE(REPLACE(b.摘要,' ',''),CHAR(13)+CHAR(10),''), 3, 10)
				AND a.摘要 not Like '行政訴訟%' ) 
				OR (a.id='417' And b.id='499') 
				THEN (a.支出-b.收入) 
				ELSE a.支出 END) As 支出2 
            FROM 收支備查簿 As a left Join 收支備查簿 As b
            ON (a.年=b.年 And (a.號數=b.號數 or b.號數 Is NULL AND (left(REPLACE(REPLACE(a.摘要,' ',''),CHAR(13)+CHAR(10),''), 10) Like substring(REPLACE(REPLACE(b.摘要,' ',''),CHAR(13)+CHAR(10),''), 3, 10))))
			AND b.摘要 Like'_回%'
			AND b._種類 = @_種類 
			AND b.收入 > 0
			left join 日誌
            ON a.id=日誌.id
            AND 動作='主計室通過'
            Where a.年 = @年
            AND a.取號 = 0
            AND a._種類 = @_種類
            AND a.支出 > 0
            AND (CONVERT(date,STR(a.年+1911)+STR(a.月)+STR(a.日)) 
            BETWEEN CONVERT(date,STR(TRIM(@年)+1911)+STR(ISNULL(NULLIF(TRIM(@月1),''),'1'))+STR(ISNULL(NULLIF(TRIM(@日1),''),'1')))
            AND CONVERT(date,STR(@年+1911)+STR(ISNULL(NULLIF(TRIM(@月2),''),'12'))+STR(ISNULL(NULLIF(TRIM(@日2),''),STR(Day(EOMONTH(STR(TRIM(@年)+1911)+'/'+STR(ISNULL(NULLIF(TRIM(@月2),''),'12'))+'/01'))))))
			OR a.月 Is NULL OR a.日 Is NULL)
            AND ((''=TRIM(@號數1) OR ''=TRIM(@號數2))
            OR ( a.號數 BETWEEN 
            SUBSTRING(TRIM(@號數1), PATINDEX('%[^0]%', TRIM(@號數1)), 3) AND 
            SUBSTRING(TRIM(@號數2), PATINDEX('%[^0]%', TRIM(@號數2)), 3)))
            AND a.號數 Is Not Null
            AND ((a.摘要<>'本月小計' AND a.摘要<>'累計至本月') or a.摘要 is null)
            AND (''=TRIM(@商號) OR a.商號 LIKE N'%'+TRIM(@商號)+'%')
            AND (''=TRIM(@摘要) OR a.摘要 LIKE N'%'+TRIM(@摘要)+'%')
            AND (''=TRIM(@狀態) OR (a.過審 = 1 And a.鎖定 = 1 AND a.回覆 = TRIM(@狀態)))
            ORDER BY a.回覆,a.過審,a.鎖定,case when a.送交主計室日期 is null then 0 else 1 end ,a.送交主計室日期 desc,a._頁, a._列"
        Insertcommand="" 
        UpdateCommand=""
        DeleteCommand="">
        <SelectParameters>
            <asp:ControlParameter ControlID="年" ConvertEmptyStringToNull="False" Name="年" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="_種類" ConvertEmptyStringToNull="False" Name="_種類" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="月1" ConvertEmptyStringToNull="False" Name="月1" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="日1" ConvertEmptyStringToNull="False" Name="日1" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="月2" ConvertEmptyStringToNull="False" Name="月2" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="日2" ConvertEmptyStringToNull="False" Name="日2" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="號數1" ConvertEmptyStringToNull="False" Name="號數1" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="號數2" ConvertEmptyStringToNull="False" Name="號數2" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="商號" ConvertEmptyStringToNull="False" Name="商號" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="摘要" ConvertEmptyStringToNull="False" Name="摘要" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="狀態" ConvertEmptyStringToNull="False" Name="狀態" PropertyName="Text" Type="String"/>
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
