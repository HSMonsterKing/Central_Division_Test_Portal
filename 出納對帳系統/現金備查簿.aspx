<%@ Page Language="VB" MasterPageFile="./MasterPage.master" MaintainScrollPositionOnPostback="true" AutoEventWireup="false" CodeFile="現金備查簿.aspx.vb" Inherits="現金備查簿" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
    <asp:ScriptManager ID="ScriptManager1" runat="server" EnablePageMethods = "true" EnableScriptGlobalization="True" EnableScriptLocalization="True"/>
    <script src='js/現金備查簿.js'></script>
    <link rel="stylesheet" runat="server" media="screen" href="css\現金備查簿.css"/>
    <div><h1><a id="Title" href="現金備查簿.aspx">現金備查簿<a></h1></div>
    <asp:Panel ID="Panel1" runat="server" DefaultButton="Button1" CssClass="Panel1">
        每頁筆數<asp:TextBox ID="PageSize" runat="server" CssClass="Input" Text="10"/>
        年<asp:TextBox ID="TextBox1" runat="server" Maxlength=3 CssClass="Input"/>
        傳票號碼<asp:TextBox ID="傳票號碼a" runat="server" Maxlength=7 CssClass="Input"/>~<asp:TextBox ID="傳票號碼b" runat="server" Maxlength=7 CssClass="Input"/>
        <asp:Panel ID="Panel2" runat="server" DefaultButton="Button1" CssClass="Panel2">
            摘要<asp:TextBox ID="摘要" runat="server" CssClass="Input"/>
            付款日<asp:TextBox ID="TextBox3" runat="server" CssClass="Input"/>
            支票編號<asp:TextBox ID="TextBox4" runat="server" CssClass="Input"/>
            金額<asp:TextBox ID="金額" runat="server" CssClass="Input"/>
            廠商及備註<asp:TextBox ID="TextBox5" runat="server" CssClass="Input"/>
            <ajaxToolkit:AutoCompleteExtender ID="廠商自動" runat="server" TargetControlID="TextBox5" 
            CompletionSetCount="10000" CompletionInterval="50" EnableCaching="true" MinimumPrefixLength="1" 
            ServiceMethod="GetMyList" 
            CompletionListCssClass="CompletionList" 
            CompletionListItemCssClass="CompletionListItem" 
            CompletionListHighlightedItemCssClass="CompletionListHighlightedItem"/>
        </asp:Panel>
        <asp:Button ID="Button1" runat="server" Text="搜尋" CssClass="GreenButton"/>
        <asp:Button ID="Button2" runat="server" Text="新增" OnClick="Insert" CssClass="GreenButton"/>
        <asp:Button ID="Button3" runat="server" Text="存檔" OnClick="Update" CssClass="GreenButton"/>
        <asp:Button ID="Button4" runat="server" Text="下載" OnClick="Download" CssClass="GreenButton"/>
        <asp:Label ID="Label1" runat="server" Text="" CssClass="GreenLabel"/>
        <asp:Label ID="Label2" runat="server" Text="" CssClass="RedLabel"/>
    </asp:Panel>
    <asp:Panel ID="Panel3" runat="server" DefaultButton="Button3" CssClass="Panel3">
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="false" CellPadding="1" HorizontalAlign="Center" DataSourceID="SqlDataSource1" DataKeyNames="id" GridLines="None" ShowHeaderWhenEmpty="false" AllowPaging="True" PageSize="20" AllowSorting="True" PagerSettings-PageButtonCount=25 EnableModelValidation="True" AutoGenerateEditButton="false">
            <Columns>
                <asp:TemplateField Visible="false">
                    <ItemTemplate>
                        <asp:Label ID="Label1" runat="server" Text='<%# Bind("id") %>'/>
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="">
                    <ItemTemplate>
                        <asp:Label ID="Label2" runat="server" Text='<%# Container.DataItemIndex + 1 %>' CssClass="Label Label2"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="序號">
                    <ItemTemplate>
                        <asp:Label ID="Label3" runat="server" Text='<%# Bind("序號") %>' CssClass="Label Label3"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="年月日">
                    <ItemTemplate>
                        <asp:Label ID="Label4" runat="server" Text='<%# If(IsDate(Eval("結帳日期")), (Year(Eval("結帳日期"))-1911).ToString() & Eval("結帳日期", "{0:/MM/dd}"), "") %>' CssClass="Label Label4"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="種類">
                    <ItemTemplate>
                        <asp:TextBox ID="TextBox1" runat="server" Text='<%# Bind("種類") %>' Maxlength=1 CssClass="TextBox TextBox1"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="傳票號碼">
                    <ItemTemplate>
                        <asp:TextBox ID="TextBox2" runat="server" Text='<%# Bind("傳票號碼") %>' Maxlength=7 CssClass="TextBox TextBox2"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="摘　　　　　要">
                    <ItemTemplate>
                        <asp:TextBox ID="TextBox3" runat="server" Text='<%# Bind("會計科目及摘要") %>' Maxlength=500 TextMode="MultiLine" CssClass="TextBox TextBox3"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="付款日">
                    <ItemTemplate>
                        <asp:TextBox ID="TextBox4" runat="server" Text='<%# Bind("付款日") %>' Maxlength=300 TextMode="MultiLine" CssClass="TextBox TextBox4"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="支票編號">
                    <ItemTemplate>
                        <asp:TextBox ID="TextBox5" runat="server" Text='<%# Bind("支票編號") %>' Maxlength=300 TextMode="MultiLine" CssClass="TextBox TextBox5"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="<br>收入">
                    <ItemTemplate>
                        <asp:TextBox ID="TextBox6" runat="server" Text='<%# Eval("收入金額405", "{0:n0}") %>' Maxlength=11 CssClass="TextBox TextBox6"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="<span>土銀405<br>支出</span>">
                    <ItemTemplate>
                        <asp:TextBox ID="TextBox7" runat="server" Text='<%# Eval("支出金額405", "{0:n0}") %>' Maxlength=11 CssClass="TextBox TextBox7"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="<br>餘額">
                    <ItemTemplate>
                        <asp:TextBox ID="TextBox8" runat="server" Text='<%# Eval("餘額405", "{0:n0}") %>' Maxlength=11 CssClass="TextBox TextBox8"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField　HeaderText="">
                    <ItemTemplate>
                        <asp:Button ID="Button1" runat="server" Text="⇄" OnClientClick="switchtext(this);return false;" CausesValidation="false" CssClass="GreenButton"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="<br>收入">
                    <ItemTemplate>
                        <asp:TextBox ID="TextBox9" runat="server" Text='<%# Eval("收入金額409", "{0:n0}") %>' Maxlength=11 CssClass="TextBox TextBox9"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="中國信託409<br>支出">
                    <ItemTemplate>
                        <asp:TextBox ID="TextBox10" runat="server" Text='<%# Eval("支出金額409", "{0:n0}") %>' Maxlength=11 CssClass="TextBox TextBox10"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="<br>餘額">
                    <ItemTemplate>
                        <asp:TextBox ID="TextBox11" runat="server" Text='<%# Eval("餘額409", "{0:n0}") %>' Maxlength=11 CssClass="TextBox TextBox11"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="廠商及備註">
                    <ItemTemplate>
                        <asp:TextBox ID="TextBox12" runat="server" Text='<%# Bind("廠商及備註") %>' Maxlength=300 TextMode="MultiLine" CssClass="TextBox TextBox12"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="黏存單">
                    <ItemTemplate>
                        <asp:Button ID="Button2" runat="server" Text="下載" CommandName="Download" CssClass="GreenButton"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="">
                    <ItemTemplate>
                        <asp:Button ID="Button3" runat="server" Text="刪除" CommandName="Delete" OnClientClick="return confirm('確定刪除?')" CssClass="RedButton"/>
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
        SelectCommand="SELECT * 
        FROM 現金備查簿 
        WHERE (''=TRIM(@_年) OR 年 LIKE TRIM(@_年)) 
        AND (((''=TRIM(@傳票號碼A) OR 傳票號碼 LIKE N'%'+TRIM(@傳票號碼A)) AND (''=TRIM(@傳票號碼B) OR 傳票號碼 LIKE N'%'+TRIM(@傳票號碼B))) 
            OR (LEN(TRIM(@傳票號碼A)) = 7 AND LEN(TRIM(@傳票號碼B)) = 7 AND 傳票號碼 BETWEEN TRIM(@傳票號碼A) AND TRIM(@傳票號碼B)) 
            OR (LEN(TRIM(@傳票號碼A)) = 7 AND LEN(TRIM(@傳票號碼B)) BETWEEN 1 AND 6 AND 傳票號碼 BETWEEN TRIM(@傳票號碼A) AND LEFT(TRIM(@傳票號碼A), 1) + RIGHT('000000' + TRIM(@傳票號碼B), 6)) 
            OR (LEN(TRIM(@傳票號碼A)) BETWEEN 1 AND 6 AND LEN(TRIM(@傳票號碼B)) = 7 AND 傳票號碼 BETWEEN LEFT(TRIM(@傳票號碼B), 1) + RIGHT('000000' + TRIM(@傳票號碼A), 6) AND TRIM(@傳票號碼B)) 
            OR (LEN(TRIM(@傳票號碼A)) BETWEEN 1 AND 6 AND LEN(TRIM(@傳票號碼B)) BETWEEN 1 AND 6 AND RIGHT(傳票號碼, 6) BETWEEN RIGHT('000000' + TRIM(@傳票號碼A), 6) AND RIGHT('000000' + TRIM(@傳票號碼B), 6))) 
        AND (''=TRIM(@_會計科目及摘要) OR 會計科目及摘要 LIKE N'%'+TRIM(@_會計科目及摘要)+'%') 
        AND (''=TRIM(@_付款日) OR 付款日 LIKE N'%'+TRIM(@_付款日)+'%') 
        AND (''=TRIM(@_金額) 
            OR 收入金額405 LIKE Replace(TRIM(@_金額), ',', '') 
            OR 支出金額405 LIKE Replace(TRIM(@_金額), ',', '') 
            OR 餘額405 LIKE Replace(TRIM(@_金額), ',', '') 
            OR 收入金額409 LIKE Replace(TRIM(@_金額), ',', '') 
            OR 支出金額409 LIKE Replace(TRIM(@_金額), ',', '') 
            OR 餘額409 LIKE Replace(TRIM(@_金額), ',', '')) 
        AND (''=TRIM(@_支票編號) OR 支票編號 LIKE N'%'+TRIM(@_支票編號)+'%') 
        AND (''=TRIM(@_廠商及備註) OR 廠商及備註 LIKE N'%'+TRIM(@_廠商及備註)+'%') 
        ORDER BY 年, CASE WHEN 序號 IS NULL THEN 1 ELSE 0 END, 序號, CASE WHEN 傳票號碼 IS NULL THEN 1 ELSE 0 END, 傳票號碼" 
        Insertcommand="INSERT INTO 現金備查簿 (年) VALUES(TRIM(@_年))" 
        UpdateCommand="
            UPDATE 現金備查簿 SET 
            序號 = NULLIF(ISNULL(N''+@序號+'', 序號), ''), 
            結帳日期 = NULLIF(ISNULL(N''+@結帳日期+'', 結帳日期), ''), 
            種類 = NULLIF(ISNULL(N''+@種類+'', 種類), ''), 
            傳票號碼 = NULLIF(ISNULL(N''+@傳票號碼+'', 傳票號碼), ''), 
            會計科目及摘要 = NULLIF(ISNULL(N''+@會計科目及摘要+'', 會計科目及摘要), ''), 
            支票編號 = NULLIF(ISNULL(N''+@支票編號+'', 支票編號), ''), 
            收入金額405 = NULLIF(ISNULL(REPLACE(N''+@收入金額405+'', ',', ''), 收入金額405), ''), 
            支出金額405 = NULLIF(ISNULL(REPLACE(N''+@支出金額405+'', ',', ''), 支出金額405), ''), 
            餘額405 = NULLIF(ISNULL(REPLACE(N''+@餘額405+'', ',', ''), 餘額405), ''), 
            收入金額409 = NULLIF(ISNULL(REPLACE(N''+@收入金額409+'', ',', ''), 收入金額409), ''), 
            支出金額409 = NULLIF(ISNULL(REPLACE(N''+@支出金額409+'', ',', ''), 支出金額409), ''), 
            餘額409 = NULLIF(ISNULL(REPLACE(N''+@餘額409+'', ',', ''), 餘額409), ''), 
            廠商及備註 = NULLIF(ISNULL(N''+@廠商及備註+'', 廠商及備註), ''), 
            付款日 = NULLIF(ISNULL(N''+@付款日+'', 付款日), ''), 
            年 = NULLIF(ISNULL(N''+@年+'', 年), ''), 
            傳票送出納檔名 = NULLIF(ISNULL(N''+@傳票送出納檔名+'', 傳票送出納檔名), '') 
            WHERE id = @id" 
        DeleteCommand="DELETE FROM 現金備查簿 WHERE id=@id">
        <SelectParameters>
            <asp:ControlParameter ControlID="TextBox1" ConvertEmptyStringToNull="false" Name="_年" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="傳票號碼a" ConvertEmptyStringToNull="false" Name="傳票號碼a" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="傳票號碼b" ConvertEmptyStringToNull="false" Name="傳票號碼b" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="摘要" ConvertEmptyStringToNull="false" Name="_會計科目及摘要" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="TextBox3" ConvertEmptyStringToNull="false" Name="_付款日" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="金額" ConvertEmptyStringToNull="false" Name="_金額" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="TextBox4" ConvertEmptyStringToNull="false" Name="_支票編號" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="TextBox5" ConvertEmptyStringToNull="false" Name="_廠商及備註" PropertyName="Text" Type="String"/>
        </SelectParameters>
        <InsertParameters>
            <asp:ControlParameter ControlID="TextBox1" ConvertEmptyStringToNull="false" Name="_年" PropertyName="Text" Type="String"/>
        </InsertParameters>
        <UpdateParameters>
            <asp:Parameter Name="id" ConvertEmptyStringToNull="False" Type="String"/>
            <asp:Parameter Name="序號" ConvertEmptyStringToNull="False" Type="String"/>
            <asp:Parameter Name="結帳日期" ConvertEmptyStringToNull="False" Type="String"/>
            <asp:Parameter Name="種類" ConvertEmptyStringToNull="False" Type="String"/>
            <asp:Parameter Name="傳票號碼" ConvertEmptyStringToNull="False" Type="String"/>
            <asp:Parameter Name="會計科目及摘要" ConvertEmptyStringToNull="False" Type="String"/>
            <asp:Parameter Name="支票編號" ConvertEmptyStringToNull="False" Type="String"/>
            <asp:Parameter Name="收入金額405" ConvertEmptyStringToNull="False" Type="String"/>
            <asp:Parameter Name="支出金額405" ConvertEmptyStringToNull="False" Type="String"/>
            <asp:Parameter Name="餘額405" ConvertEmptyStringToNull="False" Type="String"/>
            <asp:Parameter Name="收入金額409" ConvertEmptyStringToNull="False" Type="String"/>
            <asp:Parameter Name="支出金額409" ConvertEmptyStringToNull="False" Type="String"/>
            <asp:Parameter Name="餘額409" ConvertEmptyStringToNull="False" Type="String"/>
            <asp:Parameter Name="廠商及備註" ConvertEmptyStringToNull="False" Type="String"/>
            <asp:Parameter Name="付款日" ConvertEmptyStringToNull="False" Type="String"/>
            <asp:Parameter Name="年" ConvertEmptyStringToNull="False" Type="String"/>
            <asp:Parameter Name="傳票送出納檔名" ConvertEmptyStringToNull="False" Type="String"/>
        </UpdateParameters>
        <DeleteParameters>
            <asp:Parameter Name="id"/>
        </DeleteParameters>
    </asp:SqlDataSource>
</asp:Content>


