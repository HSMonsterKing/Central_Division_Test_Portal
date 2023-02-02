<%@ Page Language="VB" MasterPageFile="./MasterPage.master" MaintainScrollPositionOnPostback="true" AutoEventWireup="false" CodeFile="保管品.aspx.vb" Inherits="保管品" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
    <link rel="stylesheet" runat="server" media="screen" href="css\保管品.css"/>
    <div>
        <h1>
            <a ID="Title" href="保管品.aspx">保管品<a>
            <asp:DropDownList ID="DropDownList1" runat="server" AutoPostBack="True" OnSelectedIndexChanged="DropDownList1_SelectedIndexChanged" CssClass="DropDownList DropDownList1">
                <asp:ListItem Text="保證書明細表" Value="保證書明細表"></asp:ListItem>
                <asp:ListItem Text="保證書紀錄簿" Value="保證書紀錄簿"></asp:ListItem>
                <asp:ListItem Text="定存單明細表" Value="定存單明細表"></asp:ListItem>
                <asp:ListItem Text="定存單紀錄簿" Value="定存單紀錄簿"></asp:ListItem>
            </asp:DropDownList>
        </h1>
    </div>
    <asp:Panel ID="Panel1" runat="server" DefaultButton="Button1" CssClass="Panel1">
        每頁筆數<asp:TextBox ID="Input1" runat="server" CssClass="Input Input1" Text="10"/>
        <asp:Panel ID="Panel1_1" runat="server" DefaultButton="Button1" CssClass="Panel1_1">
            序號<asp:TextBox ID="Input2" runat="server" CssClass="Input Input2"/>
        </asp:Panel>
        日期<asp:TextBox ID="Input3" runat="server" CssClass="Input Input3"/>
        <asp:Panel ID="Panel1_2" runat="server" DefaultButton="Button1" CssClass="Panel1_2">
            收/支<asp:DropDownList ID="DropDownList2" runat="server" AutoPostBack="True" CssClass="DropDownList DropDownList2">
                <asp:ListItem Text="" Value=""></asp:ListItem>
                <asp:ListItem Text="收" Value="0"></asp:ListItem>
                <asp:ListItem Text="支" Value="1"></asp:ListItem>
            </asp:DropDownList>
        </asp:Panel>
        金額<asp:TextBox ID="Input4" runat="server" CssClass="Input Input4"/>
        其他<asp:TextBox ID="Input5" runat="server" CssClass="Input Input5"/>
        <asp:Button ID="Button1" runat="server" Text="搜尋" CssClass="GreenButton"/>
        <asp:Button ID="Button2" runat="server" Text="新增" OnClick="Insert" CssClass="GreenButton"/>
        <asp:Button ID="Button3" runat="server" Text="存檔" OnClick="Update" CssClass="GreenButton"/>
        <asp:Button ID="Button4" runat="server" Text="下載" OnClick="Download" CssClass="GreenButton"/>
        <div style="display:inline-block;width:275px;text-align:left;">
            <asp:GridView ID="GridView3" runat="server" HorizontalAlign="Center" DataSourceID="SqlDataSource3" GridLines="None" ShowHeader="false" CssClass="GridView3"/>
        </div>
    </asp:Panel>
    <asp:Panel ID="Panel2" runat="server" DefaultButton="Button3" CssClass="Panel2">
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
                <asp:TemplateField HeaderText="日期">
                    <ItemTemplate>
                        <asp:TextBox ID="TextBox2" runat="server" Text='<%# If(IsDate(Eval("日期")), (Year(Eval("日期"))-1911).ToString() & Eval("日期", "{0:.MM.dd}"), "") %>' Maxlength=300 TextMode="MultiLine" CssClass="TextBox TextBox2"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="保證書名稱<br>或存單號碼">
                    <ItemTemplate>
                        <asp:TextBox ID="TextBox3" runat="server" Text='<%# Bind("保證書名稱或存單號碼") %>' Maxlength=300 TextMode="MultiLine" CssClass="TextBox TextBox3"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="收據<br>編號">
                    <ItemTemplate>
                        <asp:TextBox ID="TextBox4" runat="server" Text='<%# Bind("收據編號") %>' Maxlength=300 TextMode="MultiLine" CssClass="TextBox TextBox4"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="(國保)<br>收據<br>編號">
                    <ItemTemplate>
                        <asp:TextBox ID="TextBox5" runat="server" Text='<%# Bind("國保收據編號") %>' Maxlength=300 TextMode="MultiLine" CssClass="TextBox TextBox5"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="戶名">
                    <ItemTemplate>
                        <asp:TextBox ID="TextBox6" runat="server" Text='<%# Bind("戶名") %>' Maxlength=300 TextMode="MultiLine" CssClass="TextBox TextBox6"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="品名">
                    <ItemTemplate>
                        <asp:TextBox ID="TextBox7" runat="server" Text='<%# Bind("品名") %>' Maxlength=300 TextMode="MultiLine" CssClass="TextBox TextBox7"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="案由摘要">
                    <ItemTemplate>
                        <asp:TextBox ID="TextBox8" runat="server" Text='<%# Bind("摘要") %>' Maxlength=500 TextMode="MultiLine" CssClass="TextBox TextBox8"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="單位">
                    <ItemTemplate>
                        <asp:TextBox ID="TextBox9" runat="server" Text='<%# Bind("單位") %>' Maxlength=300 TextMode="MultiLine" CssClass="TextBox TextBox9"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="數量">
                    <ItemTemplate>
                        <asp:TextBox ID="TextBox10" runat="server" Text='<%# Eval("數量", "{0:n0}") %>' Maxlength=13 TextMode="MultiLine" CssClass="TextBox TextBox10"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="金額">
                    <ItemTemplate>
                        <asp:TextBox ID="TextBox11" runat="server" Text='<%# Eval("金額", "{0:n0}") %>' Maxlength=25 TextMode="MultiLine" CssClass="TextBox TextBox11"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="保證書或存單<br>保證期限">
                    <ItemTemplate>
                        <asp:TextBox ID="TextBox12" runat="server" Text='<%# Bind("保證書或存單保證期限") %>' Maxlength=300 TextMode="MultiLine" CssClass="TextBox TextBox12"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="廠商保證<br>責任期限">
                    <ItemTemplate>
                        <asp:TextBox ID="TextBox13" runat="server" Text='<%# Bind("廠商保證責任期限") %>' Maxlength=300 TextMode="MultiLine" CssClass="TextBox TextBox13"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="合約<br>展延<br>情形">
                    <ItemTemplate>
                        <asp:TextBox ID="TextBox14" runat="server" Text='<%# Bind("合約展延情形") %>' Maxlength=300 TextMode="MultiLine" CssClass="TextBox TextBox14"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="保管處">
                    <ItemTemplate>
                        <asp:TextBox ID="TextBox15" runat="server" Text='<%# Bind("保管處") %>' Maxlength=300 TextMode="MultiLine" CssClass="TextBox TextBox15"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="承辦單位">
                    <ItemTemplate>
                        <asp:TextBox ID="TextBox16" runat="server" Text='<%# Bind("承辦單位") %>' Maxlength=300 TextMode="MultiLine" CssClass="TextBox TextBox16"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="承辦人">
                    <ItemTemplate>
                        <asp:TextBox ID="TextBox17" runat="server" Text='<%# Bind("承辦人") %>' Maxlength=300 TextMode="MultiLine" CssClass="TextBox TextBox17"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="備考">
                    <ItemTemplate>
                        <asp:TextBox ID="TextBox18" runat="server" Text='<%# Bind("備考") %>' Maxlength=300 TextMode="MultiLine" CssClass="TextBox TextBox18"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="">
                    <ItemTemplate>
                        <asp:Button ID="Button19" runat="server" Text='<%# If(Eval("已收入"), "支出", "收入") %>' CommandName='<%# If(Eval("已收入"), "支出", "收入") %>' OnClientClick="return confirm('確定?')" CssClass="GreenButton"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="">
                    <ItemTemplate>
                        <asp:Button ID="Button20" runat="server" Text="刪除" CommandName="Delete" OnClientClick="return confirm('確定刪除?')" CssClass="RedButton"/>
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
        <asp:GridView ID="GridView2" runat="server" AutoGenerateColumns="false" CellPadding="1" HorizontalAlign="Center" DataSourceID="SqlDataSource2" DataKeyNames="id" GridLines="None" ShowHeaderWhenEmpty="false" AllowPaging="True" PageSize="20" AllowSorting="True" PagerSettings-PageButtonCount=25 EnableModelValidation="True" AutoGenerateEditButton="false">
            <Columns>
                <asp:TemplateField Visible="false">
                    <ItemTemplate>
                        <asp:Label ID="Label0" runat="server" Text='<%# Bind("id") %>'/>
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="序號">
                    <ItemTemplate>
                        <asp:TextBox ID="TextBox1" runat="server" Text='<%# Bind("序號") %>' Maxlength=300 TextMode="MultiLine" CssClass="TextBox TextBox1_2"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="日期">
                    <ItemTemplate>
                        <asp:TextBox ID="TextBox2" runat="server" Text='<%# If(IsDate(Eval("日期")), (Year(Eval("日期"))-1911).ToString() & Eval("日期", "{0:.MM.dd}"), "") %>' Maxlength=300 TextMode="MultiLine" CssClass="TextBox TextBox2_2"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="收/支">
                    <ItemTemplate>
                        <asp:TextBox ID="TextBox3" runat="server" Text='<%# If(Eval("收支"), "支", "收") %>' Maxlength=300 TextMode="MultiLine" CssClass="TextBox TextBox3_2"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="(國保)<br>收據<br>編號">
                    <ItemTemplate>
                        <asp:TextBox ID="TextBox4" runat="server" Text='<%# Bind("國保收據編號") %>' Maxlength=300 TextMode="MultiLine" CssClass="TextBox TextBox4_2"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="摘要">
                    <ItemTemplate>
                        <asp:TextBox ID="TextBox5" runat="server" Text='<%# Bind("摘要") %>' Maxlength=500 TextMode="MultiLine" CssClass="TextBox TextBox5_2"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="收入金額">
                    <ItemTemplate>
                        <asp:TextBox ID="TextBox6" runat="server" Text='<%# Eval("收入金額", "{0:n0}") %>' Maxlength=25 TextMode="MultiLine" CssClass="TextBox TextBox6_2"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="支出金額">
                    <ItemTemplate>
                        <asp:TextBox ID="TextBox7" runat="server" Text='<%# Eval("支出金額", "{0:n0}") %>' Maxlength=25 TextMode="MultiLine" CssClass="TextBox TextBox7_2"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="餘額">
                    <ItemTemplate>
                        <asp:TextBox ID="TextBox8" runat="server" Text='<%# Eval("餘額", "{0:n0}") %>' Maxlength=25 TextMode="MultiLine" CssClass="TextBox TextBox8_2"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="附票<br>帶期<br>息效">
                    <ItemTemplate>
                        <asp:TextBox ID="TextBox9" runat="server" Text='<%# Bind("戶名") %>' Maxlength=300 TextMode="MultiLine" CssClass="TextBox TextBox9_2"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="備考">
                    <ItemTemplate>
                        <asp:TextBox ID="TextBox10" runat="server" Text='<%# Bind("品名") %>' Maxlength=300 TextMode="MultiLine" CssClass="TextBox TextBox10_2"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="">
                    <ItemTemplate>
                        <asp:Button ID="Button11" runat="server" Text="轉明細表" CommandName="轉明細表" CssClass="GreenButton"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="">
                    <ItemTemplate>
                        <asp:Button ID="Button12" runat="server" Text="支出" CommandName="支出" Enabled='<%# Not Eval("收支") %>' CssClass="GreenButton"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="">
                    <ItemTemplate>
                        <asp:Button ID="Button13" runat="server" Text="下載" CommandName="下載" CssClass="GreenButton"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="">
                    <ItemTemplate>
                        <asp:Button ID="Button14" runat="server" Text="刪除" CommandName="Delete" OnClientClick="return confirm('確定刪除?')" CssClass="RedButton"/>
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
            IF @選項 = '保證書明細表' 
                BEGIN 
                    SELECT * 
                    FROM 保管品明細表 
                    WHERE (種類=0 AND 已支出=0) 
                    AND (''=TRIM(@日期) OR STR(YEAR(日期)-1911) + FORMAT(日期, 'MMdd') LIKE N'%' + REPLACE(REPLACE(REPLACE(@日期, '.', ''), '/', ''), ' ', '') + '%') 
                    AND (''=TRIM(@金額) OR 金額 LIKE N'%' + TRIM(REPLACE(@金額, ',', '')) + '%') 
                    AND ((''=TRIM(@其他)) 
                        OR (保證書名稱或存單號碼 LIKE N'%' + TRIM(@其他) + '%') 
                        OR (收據編號 LIKE N'%' + TRIM(@其他) + '%') 
                        OR (國保收據編號 LIKE N'%' + TRIM(@其他) + '%') 
                        OR (戶名 LIKE N'%' + TRIM(@其他) + '%') 
                        OR (品名 LIKE N'%' + TRIM(@其他) + '%') 
                        OR (摘要 LIKE N'%' + TRIM(@其他) + '%') 
                        OR (保證書或存單保證期限 LIKE N'%' + TRIM(@其他) + '%') 
                        OR (廠商保證責任期限 LIKE N'%' + TRIM(@其他) + '%') 
                        OR (合約展延情形 LIKE N'%' + TRIM(@其他) + '%') 
                        OR (保管處 LIKE N'%' + TRIM(@其他) + '%') 
                        OR (承辦單位 LIKE N'%' + TRIM(@其他) + '%') 
                        OR (承辦人 LIKE N'%' + TRIM(@其他) + '%') 
                        OR (備考 LIKE N'%' + TRIM(@其他) + '%')) 
                    ORDER BY 
                    CASE WHEN 日期 IS NULL THEN 1 ELSE 0 END, 日期, 
                    CASE WHEN 收據編號 = '' THEN 1 ELSE 0 END, 收據編號, 
                    CASE WHEN 國保收據編號 = '' THEN 1 ELSE 0 END, 國保收據編號 
                END 
            ELSE IF @選項 = '定存單明細表' 
                BEGIN 
                    SELECT * 
                    FROM 保管品明細表 
                    WHERE (種類=1 AND 已支出=0) 
                    AND (''=TRIM(@日期) OR STR(YEAR(日期)-1911) + FORMAT(日期, 'MMdd') LIKE N'%' + REPLACE(REPLACE(REPLACE(@日期, '.', ''), '/', ''), ' ', '') + '%') 
                    AND (''=TRIM(@金額) OR 金額 LIKE N'%' + TRIM(REPLACE(@金額, ',', '')) + '%') 
                    AND ((''=TRIM(@其他)) 
                        OR (保證書名稱或存單號碼 LIKE N'%' + TRIM(@其他) + '%') 
                        OR (收據編號 LIKE N'%' + TRIM(@其他) + '%') 
                        OR (國保收據編號 LIKE N'%' + TRIM(@其他) + '%') 
                        OR (戶名 LIKE N'%' + TRIM(@其他) + '%') 
                        OR (品名 LIKE N'%' + TRIM(@其他) + '%') 
                        OR (摘要 LIKE N'%' + TRIM(@其他) + '%') 
                        OR (保證書或存單保證期限 LIKE N'%' + TRIM(@其他) + '%') 
                        OR (廠商保證責任期限 LIKE N'%' + TRIM(@其他) + '%') 
                        OR (合約展延情形 LIKE N'%' + TRIM(@其他) + '%') 
                        OR (保管處 LIKE N'%' + TRIM(@其他) + '%') 
                        OR (承辦單位 LIKE N'%' + TRIM(@其他) + '%') 
                        OR (承辦人 LIKE N'%' + TRIM(@其他) + '%') 
                        OR (備考 LIKE N'%' + TRIM(@其他) + '%')) 
                    ORDER BY 
                    CASE WHEN 日期 IS NULL THEN 1 ELSE 0 END, 日期, 
                    CASE WHEN 收據編號 = '' THEN 1 ELSE 0 END, 收據編號, 
                    CASE WHEN 國保收據編號 = '' THEN 1 ELSE 0 END, 國保收據編號 
                END 
            ELSE 
                BEGIN 
                    SELECT * 
                    FROM 保管品明細表 
                    WHERE 0 = 1 
                END" 
        Insertcommand=" 
            IF @選項 = '保證書明細表' 
                BEGIN 
                    INSERT INTO 保管品明細表 
                    (種類, 已收入, 已支出, 日期, 單位, 數量, 合約展延情形) 
                    VALUES 
                    (0, 0, 0, NULL, N'包', 1, N'無') 
                END 
            ELSE IF @選項 = '定存單明細表' 
                BEGIN 
                    INSERT INTO 保管品明細表 
                    (種類, 已收入, 已支出, 日期, 單位, 數量, 合約展延情形) 
                    VALUES 
                    (1, 0, 0, NULL, N'張', 1, N'無') 
                END" 
        UpdateCommand="" 
        DeleteCommand="
            SET XACT_ABORT ON 
            BEGIN TRANSACTION 
                DELETE FROM 保管品明細表 WHERE id = @id 
                UPDATE 保管品紀錄簿 SET 保管品明細表id = NULL WHERE 保管品明細表id = @id 
            COMMIT TRANSACTION 
            SET XACT_ABORT OFF ">
        <SelectParameters>
            <asp:ControlParameter ControlID="DropDownList1" ConvertEmptyStringToNull="False" Name="選項" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="Input3" ConvertEmptyStringToNull="False" Name="日期" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="Input4" ConvertEmptyStringToNull="False" Name="金額" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="Input5" ConvertEmptyStringToNull="False" Name="其他" PropertyName="Text" Type="String"/>
        </SelectParameters>
        <InsertParameters>
            <asp:ControlParameter ControlID="DropDownList1" ConvertEmptyStringToNull="False" Name="選項" PropertyName="Text" Type="String"/>
        </InsertParameters>
        <UpdateParameters>
        </UpdateParameters>
        <DeleteParameters>
            <asp:Parameter Name="id"/>
        </DeleteParameters>
    </asp:SqlDataSource>
    <asp:SqlDataSource ID="SqlDataSource2" runat="server" ConnectionString='<%$ ConnectionStrings:ApplicationServices%>'
        SelectCommand=" 
            IF @選項 = '保證書紀錄簿' 
                BEGIN 
                    SELECT * 
                    FROM 保管品紀錄簿 
                    WHERE (種類=0) 
                    AND (''=TRIM(@序號) OR 序號 LIKE N'' + SUBSTRING(TRIM(@序號), PATINDEX('%[^0]%', TRIM(@序號)+'.'), LEN(TRIM(@序號))) + '') 
                    AND (''=TRIM(@日期) OR STR(YEAR(日期)-1911) + FORMAT(日期, 'MMdd') LIKE N'%' + REPLACE(REPLACE(REPLACE(@日期, '.', ''), '/', ''), ' ', '') + '%') 
                    AND (''=TRIM(@收支) OR 收支 LIKE N'%' + TRIM(@收支) + '%') 
                    AND ((''=TRIM(@金額)) 
                        OR (收入金額 LIKE N'%' + TRIM(REPLACE(@金額, ',', '')) + '%') 
                        OR (支出金額 LIKE N'%' + TRIM(REPLACE(@金額, ',', '')) + '%') 
                        OR (餘額 LIKE N'%' + TRIM(REPLACE(@金額, ',', '')) + '%')) 
                    AND ((''=TRIM(@其他)) 
                        OR (收支 LIKE N'%' + TRIM(@其他) + '%') 
                        OR (國保收據編號 LIKE N'%' + TRIM(@其他) + '%') 
                        OR (摘要 LIKE N'%' + TRIM(@其他) + '%') 
                        OR (戶名 LIKE N'%' + TRIM(@其他) + '%') 
                        OR (品名 LIKE N'%' + TRIM(@其他) + '%')) 
                    ORDER BY 
                    序號 
                END 
            ELSE IF @選項 = '定存單紀錄簿' 
                BEGIN 
                    SELECT * 
                    FROM 保管品紀錄簿 
                    WHERE (種類=1) 
                    AND (''=TRIM(@序號) OR 序號 LIKE N'' + SUBSTRING(TRIM(@序號), PATINDEX('%[^0]%', TRIM(@序號)+'.'), LEN(TRIM(@序號))) + '') 
                    AND (''=TRIM(@日期) OR STR(YEAR(日期)-1911) + FORMAT(日期, 'MMdd') LIKE N'%' + REPLACE(REPLACE(REPLACE(@日期, '.', ''), '/', ''), ' ', '') + '%') 
                    AND (''=TRIM(@收支) OR 收支 LIKE N'%' + TRIM(@收支) + '%') 
                    AND ((''=TRIM(@金額)) 
                        OR (收入金額 LIKE N'%' + TRIM(REPLACE(@金額, ',', '')) + '%') 
                        OR (支出金額 LIKE N'%' + TRIM(REPLACE(@金額, ',', '')) + '%') 
                        OR (餘額 LIKE N'%' + TRIM(REPLACE(@金額, ',', '')) + '%')) 
                    AND ((''=TRIM(@其他)) 
                        OR (收支 LIKE N'%' + TRIM(@其他) + '%') 
                        OR (國保收據編號 LIKE N'%' + TRIM(@其他) + '%') 
                        OR (摘要 LIKE N'%' + TRIM(@其他) + '%') 
                        OR (戶名 LIKE N'%' + TRIM(@其他) + '%') 
                        OR (品名 LIKE N'%' + TRIM(@其他) + '%')) 
                    ORDER BY 
                    序號 
                END 
            ELSE 
                BEGIN 
                    SELECT * 
                    FROM 保管品紀錄簿 
                    WHERE 0 = 1 
                END" 
        Insertcommand=" 
            IF @選項 = '保證書紀錄簿' 
                BEGIN 
                    INSERT INTO 保管品紀錄簿 
                    (種類, 序號, 收支) 
                    VALUES 
                    (0, (SELECT ISNULL(MAX(序號), 0)+1 FROM 保管品紀錄簿 WHERE 種類 = 0), 0) 
                END 
            ELSE IF @選項 = '定存單紀錄簿' 
                BEGIN 
                    INSERT INTO 保管品紀錄簿 
                    (種類, 序號, 收支) 
                    VALUES 
                    (1, (SELECT ISNULL(MAX(序號), 0)+1 FROM 保管品紀錄簿 WHERE 種類 = 1), 0) 
                END" 
        UpdateCommand="" 
        DeleteCommand="DELETE FROM 保管品紀錄簿 WHERE id=@id">
        <SelectParameters>
            <asp:ControlParameter ControlID="DropDownList1" ConvertEmptyStringToNull="False" Name="選項" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="Input2" ConvertEmptyStringToNull="False" Name="序號" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="Input3" ConvertEmptyStringToNull="False" Name="日期" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="DropDownList2" ConvertEmptyStringToNull="False" Name="收支" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="Input4" ConvertEmptyStringToNull="False" Name="金額" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="Input5" ConvertEmptyStringToNull="False" Name="其他" PropertyName="Text" Type="String"/>
        </SelectParameters>
        <InsertParameters>
            <asp:ControlParameter ControlID="DropDownList1" ConvertEmptyStringToNull="False" Name="選項" PropertyName="Text" Type="String"/>
        </InsertParameters>
        <UpdateParameters>
        </UpdateParameters>
        <DeleteParameters>
            <asp:Parameter Name="id"/>
        </DeleteParameters>
    </asp:SqlDataSource>
    <asp:SqlDataSource ID="SqlDataSource3" runat="server" ConnectionString='<%$ ConnectionStrings:ApplicationServices%>'
        SelectCommand=" 
            IF @選項 = '保證書明細表' 
                BEGIN 
                    SELECT '總數量:' + STR(SUM(數量)) + '包 總金額: ' + REPLACE(CONVERT(varchar(30), CONVERT(money, SUM(金額)), 1), '.00', '元') 
                    FROM 保管品明細表 
                    WHERE (種類=0 AND 已支出=0) 
                    AND (''=TRIM(@日期) OR STR(YEAR(日期)-1911) + FORMAT(日期, 'MMdd') LIKE N'%' + REPLACE(REPLACE(REPLACE(@日期, '.', ''), '/', ''), ' ', '') + '%') 
                    AND (''=TRIM(@金額) OR 金額 LIKE N'%' + TRIM(REPLACE(@金額, ',', '')) + '%') 
                    AND ((''=TRIM(@其他)) 
                        OR (保證書名稱或存單號碼 LIKE N'%' + TRIM(@其他) + '%') 
                        OR (收據編號 LIKE N'%' + TRIM(@其他) + '%') 
                        OR (國保收據編號 LIKE N'%' + TRIM(@其他) + '%') 
                        OR (戶名 LIKE N'%' + TRIM(@其他) + '%') 
                        OR (品名 LIKE N'%' + TRIM(@其他) + '%') 
                        OR (摘要 LIKE N'%' + TRIM(@其他) + '%') 
                        OR (保證書或存單保證期限 LIKE N'%' + TRIM(@其他) + '%') 
                        OR (廠商保證責任期限 LIKE N'%' + TRIM(@其他) + '%') 
                        OR (合約展延情形 LIKE N'%' + TRIM(@其他) + '%') 
                        OR (保管處 LIKE N'%' + TRIM(@其他) + '%') 
                        OR (承辦單位 LIKE N'%' + TRIM(@其他) + '%') 
                        OR (承辦人 LIKE N'%' + TRIM(@其他) + '%') 
                        OR (備考 LIKE N'%' + TRIM(@其他) + '%')) 
                END 
            ELSE IF @選項 = '定存單明細表' 
                BEGIN 
                    SELECT '總數量:' + STR(SUM(數量)) + '張 總金額: ' + REPLACE(CONVERT(varchar(30), CONVERT(money, SUM(金額)), 1), '.00', '元') 
                    FROM 保管品明細表 
                    WHERE (種類=1 AND 已支出=0) 
                    AND (''=TRIM(@日期) OR STR(YEAR(日期)-1911) + FORMAT(日期, 'MMdd') LIKE N'%' + REPLACE(REPLACE(REPLACE(@日期, '.', ''), '/', ''), ' ', '') + '%') 
                    AND (''=TRIM(@金額) OR 金額 LIKE N'%' + TRIM(REPLACE(@金額, ',', '')) + '%') 
                    AND ((''=TRIM(@其他)) 
                        OR (保證書名稱或存單號碼 LIKE N'%' + TRIM(@其他) + '%') 
                        OR (收據編號 LIKE N'%' + TRIM(@其他) + '%') 
                        OR (國保收據編號 LIKE N'%' + TRIM(@其他) + '%') 
                        OR (戶名 LIKE N'%' + TRIM(@其他) + '%') 
                        OR (品名 LIKE N'%' + TRIM(@其他) + '%') 
                        OR (摘要 LIKE N'%' + TRIM(@其他) + '%') 
                        OR (保證書或存單保證期限 LIKE N'%' + TRIM(@其他) + '%') 
                        OR (廠商保證責任期限 LIKE N'%' + TRIM(@其他) + '%') 
                        OR (合約展延情形 LIKE N'%' + TRIM(@其他) + '%') 
                        OR (保管處 LIKE N'%' + TRIM(@其他) + '%') 
                        OR (承辦單位 LIKE N'%' + TRIM(@其他) + '%') 
                        OR (承辦人 LIKE N'%' + TRIM(@其他) + '%') 
                        OR (備考 LIKE N'%' + TRIM(@其他) + '%')) 
                END 
            ELSE 
                BEGIN 
                    SELECT NULL 
                    FROM 保管品明細表 
                    WHERE 0 = 1 
                END">
        <SelectParameters>
            <asp:ControlParameter ControlID="DropDownList1" ConvertEmptyStringToNull="False" Name="選項" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="Input3" ConvertEmptyStringToNull="False" Name="日期" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="Input4" ConvertEmptyStringToNull="False" Name="金額" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="Input5" ConvertEmptyStringToNull="False" Name="其他" PropertyName="Text" Type="String"/>
        </SelectParameters>
    </asp:SqlDataSource>
</asp:Content>


