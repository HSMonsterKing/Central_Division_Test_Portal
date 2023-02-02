<%@ Page Language="VB" MasterPageFile="./MasterPage.master" MaintainScrollPositionOnPostback="true" AutoEventWireup="False" CodeFile="傳票資料.aspx.vb" Inherits="傳票資料" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" runat="Server">
    <asp:ScriptManager ID="ScriptManager1" runat="server" EnablePageMethods = "true" EnableScriptGlobalization="True" EnableScriptLocalization="True"/>
    <script src='js/傳票資料.js'></script>
    <link rel="stylesheet" runat="server" media="screen" href="css\傳票資料.css"/>
    <div>
        <h1>
            <a ID="Title" href="傳票資料.aspx">傳票資料<a>
            <asp:DropDownList ID="DropDownList1" runat="server" AutoPostBack="True" OnSelectedIndexChanged="DropDownList1_SelectedIndexChanged" CssClass="DropDownList">
                <asp:ListItem Text="全" Value="全"></asp:ListItem>
                <asp:ListItem Text="土銀405全" Value="土銀405全"></asp:ListItem>
                <asp:ListItem Text="土銀405匯款" Value="土銀405匯款"></asp:ListItem>
                <asp:ListItem Text="土銀405支票" Value="土銀405支票"></asp:ListItem>
                <asp:ListItem Text="土銀405收入" Value="土銀405收入"></asp:ListItem>
                <asp:ListItem Text="中國信託409全" Value="中國信託409全"></asp:ListItem>
                <asp:ListItem Text="中國信託409收入" Value="中國信託409收入"></asp:ListItem>
                <asp:ListItem Text="中國信託409支出" Value="中國信託409支出"></asp:ListItem>
            </asp:DropDownList>
        </h1>
    </div>
    <asp:Panel ID="Panel1" runat="server" DefaultButton="Button1" CssClass="Panel1">
        <asp:Panel ID="Panel2" runat="server" DefaultButton="Button1" CssClass="Panel2">
            登錄日期<asp:TextBox ID="Calendar1" runat="server" AutoPostBack="True" OnTextChanged="Calendar1_OnTextChanged" CssClass="Input1"/>
            <ajaxToolkit:CalendarExtender ID="CalendarExtender1" runat="server" TargetControlID="Calendar1" format="yyyy/MM/dd"/>
            預付日期<asp:TextBox ID="Calendar2" runat="server" CssClass="Input1"/>
            <ajaxToolkit:CalendarExtender ID="CalendarExtender2" runat="server" TargetControlID="Calendar2" format="yyyy/MM/dd"/>
            登錄序號<asp:TextBox ID="登錄序號" runat="server" CssClass="Input1"/>
            TXT檔名<asp:TextBox ID="TXT檔名" runat="server" CssClass="Input1"/>
        </asp:Panel>
        <asp:Panel ID="Panel3" runat="server" DefaultButton="Button1" CssClass="Panel3">
            支票日期<asp:TextBox ID="Calendar3" runat="server" AutoPostBack="False" CssClass="Input1"/>
            <ajaxToolkit:CalendarExtender ID="CalendarExtender3" runat="server" TargetControlID="Calendar3" format="yyyy/MM/dd"/>
            支票編號<asp:TextBox ID="支票編號" runat="server" CssClass="Input1"/>
        </asp:Panel>
        <asp:Panel ID="Panel7" runat="server" DefaultButton="Button1" CssClass="Panel7">
            每頁筆數<asp:TextBox ID="PageSize" runat="server" Maxlength=7 CssClass="Input2" Text="20"/>
            年<asp:TextBox ID="TextBox1" runat="server" Maxlength=3 CssClass="Input2"/>
            開票日期<asp:TextBox ID="TextBox2" runat="server" Maxlength=7 CssClass="Input2"/>~<asp:TextBox ID="TextBox3" runat="server" Maxlength=7 CssClass="Input2"/>
            傳票號碼<asp:TextBox ID="TextBox4" runat="server" Maxlength=7 CssClass="Input2"/>~<asp:TextBox ID="TextBox5" runat="server" Maxlength=7 CssClass="Input2"/>
            名稱<asp:TextBox ID="名稱" runat="server" Maxlength=300 CssClass="Input2"/>
            登錄序號<asp:TextBox ID="登錄序號s" runat="server" Maxlength=300 CssClass="Input2"/>
            金額<asp:TextBox ID="TextBox6" runat="server" Maxlength=11 CssClass="Input2"/>
        </asp:Panel>
        <asp:Panel ID="Panel4" runat="server" DefaultButton="Button1" CssClass="Panel4">
            匯入帳號<asp:TextBox ID="TextBox7" runat="server" Maxlength=16 CssClass="Input2"/>
        </asp:Panel>
        <asp:Button ID="Button1" runat="server" Text="搜尋" CssClass="GreenButton"/>
        <asp:Button ID="Button2" runat="server" Text="新增" OnClick="Insert" CssClass="GreenButton"/>
        <asp:Button ID="Button3" runat="server" Text="存檔" OnClick="Update" CssClass="GreenButton"/>
        <asp:Panel ID="Panel5" runat="server" DefaultButton="Button1" CssClass="Panel5">
            <asp:Button ID="Button4" runat="server" Text="全選" OnClick="SelectAll" CssClass="GreenButton"/>
            <asp:Button ID="Button5" runat="server" Text="預覽" OnClick="Preview" CssClass="GreenButton"/>
            <asp:Button ID="Button6" runat="server" Text="下載" OnClick="Download" OnClientClick="if(document.getElementById('ContentPlaceHolder1_Button5').getAttribute('disabled')=='disabled'){return true;}else{return confirm('確認無誤?');}" CssClass="GreenButton"/>
            <asp:DropDownList ID="DropDownList2" runat="server" AutoPostBack="true" OnSelectedIndexChanged="DropDownList2_SelectedIndexChanged" DataSourceID="SqlDataSource2" DataTextField="TXT檔名" DataValueField="TXT檔名" CssClass="DropDownList"/>
            <asp:Button ID="Button7" runat="server" Text="刪除" OnClick="Delete" OnClientClick="return confirm('確定刪除?');" CssClass="RedButton"/>
        </asp:Panel>
        <asp:Label ID="Label1" runat="server" Text="" CssClass="GreenLabel"/>
        <asp:Label ID="Label2" runat="server" Text="" CssClass="RedLabel"/>
    </asp:Panel>
    <asp:Panel ID="Panel6" runat="server" DefaultButton="Button3" CssClass="Panel6">
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" CellPadding="1" HorizontalAlign="Center" DataSourceID="SqlDataSource1" DataKeyNames="id" GridLines="None" ShowHeaderWhenEmpty="False" AllowPaging="True" PageSize="20" AllowSorting="False" PagerSettings-PageButtonCount=25 EnableModelValidation="True" AutoGenerateEditButton="False" OnDataBound="GridView1_DataBound">
            <Columns>
                <asp:TemplateField HeaderText="">
                    <ItemTemplate>
                        <asp:Label ID="Label1" runat="server" Text='<%# Container.DataItemIndex + 1 %>' CssClass="Label Label1"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="ID">
                    <ItemTemplate>
                        <asp:Label ID="Label2" runat="server" Text='<%# Bind("id") %>' CssClass="Label Label2"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="傳票送出納檔名" SortExpression="傳票送出納檔名">
                    <ItemTemplate>
                        <asp:TextBox ID="TextBox1" runat="server" Text='<%# Bind("傳票送出納檔名") %>' Maxlength=16 CssClass="TextBox TextBox1"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="摘要說明" SortExpression="摘要說明">
                    <ItemTemplate>
                        <asp:TextBox ID="摘要說明" runat="server" Text='<%# Bind("摘要說明") %>' Maxlength=500 CssClass="TextBox 摘要說明"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="年" SortExpression="年">
                    <ItemTemplate>
                        <asp:TextBox ID="TextBox2" runat="server" Text='<%# Bind("年") %>' Maxlength=3 CssClass="TextBox TextBox2"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="開票日期" SortExpression="開票日期">
                    <ItemTemplate>
                        <asp:TextBox ID="TextBox3" runat="server" Text='<%# Bind("開票日期") %>' Maxlength=7 CssClass="TextBox TextBox3"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="傳票號碼" SortExpression="傳票號碼">
                    <ItemTemplate>
                        <asp:TextBox ID="TextBox4" runat="server" Text='<%# Bind("傳票號碼") %>' Maxlength=7 CssClass="TextBox TextBox4"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="之" SortExpression="之">
                    <ItemTemplate>
                        <asp:TextBox ID="TextBox5" runat="server" Text='<%# Bind("之") %>' Maxlength=4 Enabled="False" CssClass="TextBox TextBox5"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="名稱" SortExpression="名稱">
                    <ItemTemplate>
                        <asp:TextBox ID="TextBox6" runat="server" Text='<%# Bind("名稱") %>' Maxlength=300 CssClass="TextBox TextBox6"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="轉匯款">
                    <ItemTemplate>
                        <asp:Button ID="轉匯款" runat="server" Text="轉匯款" CommandName="Change" OnClientClick="return confirm('確定要轉匯款?')" Enabled='<%# If(Eval("登錄序號").ToString() = "", 1, 0) %>' CssClass="GreenButton"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="登錄序號" SortExpression="登錄序號">
                    <ItemTemplate>
                        <asp:TextBox ID="TextBox7" runat="server" Text='<%# Bind("登錄序號") %>' Maxlength=12 Enabled="False" CssClass="TextBox TextBox7"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="登錄日期" SortExpression="登錄日期">
                    <ItemTemplate>
                        <asp:TextBox ID="TextBox8" runat="server" Text='<%# Bind("登錄日期") %>' Maxlength=7 Enabled="False" CssClass="TextBox TextBox8"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="預付日期" SortExpression="預付日期">
                    <ItemTemplate>
                        <asp:TextBox ID="TextBox9" runat="server" Text='<%# Bind("預付日期") %>' Maxlength=7 Enabled='<%# If(Eval("登錄序號").ToString() = "" And Me.DropDownList1.SelectedValue = "土銀405匯款" And Me.Button5.Text <> "取消", 1, 0) %>' CssClass="TextBox TextBox9"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="收入金額" SortExpression="收入金額">
                    <ItemTemplate>
                        <asp:TextBox ID="TextBox10" runat="server" Text='<%# Bind("收入金額", "{0:n0}") %>' Maxlength=11 CssClass="TextBox TextBox10"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="支出金額" SortExpression="支出金額">
                    <ItemTemplate>
                        <asp:TextBox ID="TextBox11" runat="server" Text='<%# Bind("支出金額", "{0:n0}") %>' Maxlength=11 CssClass="TextBox TextBox11"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="分匯金額">
                    <ItemTemplate>
                        <asp:TextBox ID="分匯金額" runat="server" CssClass="TextBox 分匯金額"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="分匯">
                    <ItemTemplate>
                        <asp:Button ID="分匯" runat="server" Text='<%# If(Me.DropDownList1.SelectedValue = "土銀405匯款", "分匯", "分票") %>' CommandName="Separate" Enabled='<%# If(Eval("登錄序號").ToString() = "", 1, 0) %>' CssClass="GreenButton"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="代碼" SortExpression="代碼">
                    <ItemTemplate>
                        <asp:TextBox ID="TextBox12" runat="server" Text='<%# Bind("收款人代碼") %>' Maxlength=4 CssClass="TextBox TextBox12"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="收款人名稱" SortExpression="收款人名稱">
                    <ItemTemplate>
                        <asp:TextBox ID="TextBox13" runat="server" Text='<%# Bind("收款人名稱") %>' Maxlength=300 CssClass="TextBox TextBox13"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="銀行名稱" SortExpression="銀行名稱">
                    <ItemTemplate>
                        <asp:TextBox ID="TextBox14" runat="server" Text='<%# Bind("匯入銀行名稱") %>' Maxlength=300 CssClass="TextBox TextBox14"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="銀行代碼" SortExpression="銀行代碼">
                    <ItemTemplate>
                        <asp:TextBox ID="TextBox15" runat="server" Text='<%# Bind("匯入銀行代碼") %>' Maxlength=300 CssClass="TextBox TextBox15"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="匯入帳號" SortExpression="匯入帳號">
                    <ItemTemplate>
                        <asp:TextBox ID="TextBox16" runat="server" Text='<%# Bind("匯入帳號") %>' Maxlength=300 CssClass='<%# If(Eval("有效"), "TextBox TextBox16", "TextBox TextBox16 匯款資料有問題") %>'/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="收款人匯款戶名" SortExpression="收款人匯款戶名">
                    <ItemTemplate>
                        <asp:TextBox ID="TextBox17" runat="server" Text='<%# Bind("收款人匯款戶名") %>' Maxlength=300 CssClass='<%# If(Eval("有效"), "TextBox TextBox17", "TextBox TextBox17 匯款資料有問題") %>'/>
                        <ajaxToolkit:AutoCompleteExtender ID="AutoCompleteExtender1" runat="server" TargetControlID="TextBox17" 
                            CompletionSetCount="10000" CompletionInterval="50" EnableCaching="true" MinimumPrefixLength="1" 
                            ServiceMethod="GetMyList" 
                            CompletionListCssClass="CompletionList" 
                            CompletionListItemCssClass="CompletionListItem" 
                            CompletionListHighlightedItemCssClass="CompletionListHighlightedItem"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="統編" SortExpression="收款人統編">
                    <ItemTemplate>
                        <asp:TextBox ID="TextBox18" runat="server" Text='<%# Bind("收款人統編") %>' Maxlength=300 CssClass="TextBox TextBox18"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="EMAIL" SortExpression="收款人EMAIL">
                    <ItemTemplate>
                        <asp:TextBox ID="收款人EMAIL" runat="server" Text='<%# Bind("收款人EMAIL") %>' Maxlength=300 CssClass="TextBox 收款人EMAIL"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="下載">
                    <ItemTemplate>
                        <asp:CheckBox ID="CheckBox1" runat="server" Checked='<%# If(Eval("下載").ToString() = "True", True, False) %>' AutoPostBack="True" OnCheckedChanged="CheckBox1_CheckedChanged" Enabled='<%# If(Eval("登錄序號").ToString() = "", 1, 0) %>'/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="TXT檔名" SortExpression="TXT檔名">
                    <ItemTemplate>
                        <asp:TextBox ID="TextBox19" runat="server" Text='<%# Bind("TXT檔名") %>' Maxlength=11 Enabled="False" CssClass="TextBox TextBox19"></asp:TextBox>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="總金額" SortExpression="總金額">
                    <ItemTemplate>
                        <asp:TextBox ID="TextBox20" runat="server" Text='<%# Bind("總金額", "{0:n0}") %>' Maxlength=11 Enabled="False" CssClass="TextBox TextBox20"></asp:TextBox>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="順序" SortExpression="順序">
                    <ItemTemplate>
                        <asp:TextBox ID="TextBox21" runat="server" Text='<%# Bind("順序") %>' Maxlength=11 CssClass="TextBox TextBox21"></asp:TextBox>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="">
                    <ItemTemplate>
                        <asp:Button ID="清除" runat="server" Text="清除" CommandName="Clean" OnClientClick="return confirm('確定清除以下欄位?\n土銀405匯款:\n　登錄序號\n　登錄日期\n　預付日期\n　TXT檔名\n　總金額\n土銀405支票:\n　支票編號\n　支票日期\n　總金額')" CssClass="RedButton"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="">
                    <ItemTemplate>
                        <asp:Button ID="刪除" runat="server" Text="刪除" CommandName="CustomDelete" OnClientClick="return confirm('確定刪除?')" CssClass="RedButton"/>
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
        <asp:Label ID="Label3" runat="server" Text="" CssClass="Label3"/>
    </asp:Panel>
    <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString='<%$ ConnectionStrings:ApplicationServices%>'
        SelectCommand=" 
            IF @選項 = '全' 
                BEGIN 
                    SELECT 傳票資料.*, '' AS 會計科目及摘要, CASE WHEN (SELECT COUNT(*) FROM 收款人 WHERE 收款人.匯入帳號 = 傳票資料.匯入帳號 AND 收款人.收款人匯款戶名 = 傳票資料.收款人匯款戶名) = 1 THEN 1 ELSE 0 END AS 有效 
                    FROM 傳票資料 
                    WHERE (''=TRIM(@年) OR 傳票資料.年 LIKE TRIM(@年)) 
                    AND (((''=TRIM(@開票日期A) OR 傳票資料.開票日期 LIKE N'%'+TRIM(@開票日期A)) AND (''=TRIM(@開票日期B) OR 傳票資料.開票日期 LIKE N'%'+TRIM(@開票日期B))) 
                        OR (LEN(TRIM(@開票日期A)) = 7 AND LEN(TRIM(@開票日期B)) = 7 AND 傳票資料.開票日期 BETWEEN TRIM(@開票日期A) AND TRIM(@開票日期B)) 
                        OR (LEN(TRIM(@開票日期A)) = 7 AND LEN(TRIM(@開票日期B)) BETWEEN 1 AND 6 AND 傳票資料.開票日期 BETWEEN TRIM(@開票日期A) AND LEFT(TRIM(@開票日期A), 3) + RIGHT('0000' + TRIM(@開票日期B), 4)) 
                        OR (LEN(TRIM(@開票日期A)) BETWEEN 1 AND 6 AND LEN(TRIM(@開票日期B)) = 7 AND 傳票資料.開票日期 BETWEEN LEFT(TRIM(@開票日期B), 3) + RIGHT('0000' + TRIM(@開票日期A), 4) AND TRIM(@開票日期B)) 
                        OR (LEN(TRIM(@開票日期A)) BETWEEN 1 AND 6 AND LEN(TRIM(@開票日期B)) BETWEEN 1 AND 6 AND RIGHT(傳票資料.開票日期, 4) BETWEEN RIGHT('0000' + TRIM(@開票日期A), 4) AND RIGHT('0000' + TRIM(@開票日期B), 4))) 
                    AND (((''=TRIM(@傳票號碼A) OR 傳票資料.傳票號碼 LIKE N'%'+TRIM(@傳票號碼A)) AND (''=TRIM(@傳票號碼B) OR 傳票資料.傳票號碼 LIKE N'%'+TRIM(@傳票號碼B))) 
                        OR (LEN(TRIM(@傳票號碼A)) = 7 AND LEN(TRIM(@傳票號碼B)) = 7 AND 傳票資料.傳票號碼 BETWEEN TRIM(@傳票號碼A) AND TRIM(@傳票號碼B)) 
                        OR (LEN(TRIM(@傳票號碼A)) = 7 AND LEN(TRIM(@傳票號碼B)) BETWEEN 1 AND 6 AND 傳票資料.傳票號碼 BETWEEN TRIM(@傳票號碼A) AND LEFT(TRIM(@傳票號碼A), 1) + RIGHT('000000' + TRIM(@傳票號碼B), 6)) 
                        OR (LEN(TRIM(@傳票號碼A)) BETWEEN 1 AND 6 AND LEN(TRIM(@傳票號碼B)) = 7 AND 傳票資料.傳票號碼 BETWEEN LEFT(TRIM(@傳票號碼B), 1) + RIGHT('000000' + TRIM(@傳票號碼A), 6) AND TRIM(@傳票號碼B)) 
                        OR (LEN(TRIM(@傳票號碼A)) BETWEEN 1 AND 6 AND LEN(TRIM(@傳票號碼B)) BETWEEN 1 AND 6 AND RIGHT(傳票資料.傳票號碼, 6) BETWEEN RIGHT('000000' + TRIM(@傳票號碼A), 6) AND RIGHT('000000' + TRIM(@傳票號碼B), 6))) 
                    AND (''=TRIM(@收入或支出金額) OR 傳票資料.收入金額 LIKE REPLACE(TRIM(@收入或支出金額), ',', '') OR 傳票資料.支出金額 LIKE REPLACE(TRIM(@收入或支出金額), ',', '')) 
                    AND (''=TRIM(@名稱) OR 名稱 LIKE N'%' + TRIM(@名稱) + '%') 
                    AND (''=TRIM(@登錄序號) OR 登錄序號 LIKE N'%' + TRIM(@登錄序號) + '%') 
                    ORDER BY 
                    CASE WHEN 傳票資料.傳票送出納檔名 IS NULL THEN 1 ELSE 0 END, 傳票資料.傳票送出納檔名, 
                    CASE WHEN 傳票資料.傳票號碼 IS NULL THEN 1 ELSE 0 END, 傳票資料.傳票號碼, 
                    傳票資料.之, 
                    傳票資料.支出金額 DESC 
                END 
            ELSE IF @選項 = '土銀405全' 
                BEGIN 
                    SELECT 傳票資料.*, 現金備查簿.會計科目及摘要, CASE WHEN (SELECT COUNT(*) FROM 收款人 WHERE 收款人.匯入帳號 = 傳票資料.匯入帳號 AND 收款人.收款人匯款戶名 = 傳票資料.收款人匯款戶名) = 1 THEN 1 ELSE 0 END AS 有效 
                    FROM 傳票資料 INNER JOIN 現金備查簿 ON 傳票資料.年 = 現金備查簿.年 AND 傳票資料.傳票號碼 = 現金備查簿.傳票號碼 
                    WHERE (''=TRIM(@年) OR 傳票資料.年 LIKE TRIM(@年)) 
                    AND (((''=TRIM(@開票日期A) OR 傳票資料.開票日期 LIKE N'%'+TRIM(@開票日期A)) AND (''=TRIM(@開票日期B) OR 傳票資料.開票日期 LIKE N'%'+TRIM(@開票日期B))) 
                        OR (LEN(TRIM(@開票日期A)) = 7 AND LEN(TRIM(@開票日期B)) = 7 AND 傳票資料.開票日期 BETWEEN TRIM(@開票日期A) AND TRIM(@開票日期B)) 
                        OR (LEN(TRIM(@開票日期A)) = 7 AND LEN(TRIM(@開票日期B)) BETWEEN 1 AND 6 AND 傳票資料.開票日期 BETWEEN TRIM(@開票日期A) AND LEFT(TRIM(@開票日期A), 3) + RIGHT('0000' + TRIM(@開票日期B), 4)) 
                        OR (LEN(TRIM(@開票日期A)) BETWEEN 1 AND 6 AND LEN(TRIM(@開票日期B)) = 7 AND 傳票資料.開票日期 BETWEEN LEFT(TRIM(@開票日期B), 3) + RIGHT('0000' + TRIM(@開票日期A), 4) AND TRIM(@開票日期B)) 
                        OR (LEN(TRIM(@開票日期A)) BETWEEN 1 AND 6 AND LEN(TRIM(@開票日期B)) BETWEEN 1 AND 6 AND RIGHT(傳票資料.開票日期, 4) BETWEEN RIGHT('0000' + TRIM(@開票日期A), 4) AND RIGHT('0000' + TRIM(@開票日期B), 4))) 
                    AND (((''=TRIM(@傳票號碼A) OR 傳票資料.傳票號碼 LIKE N'%'+TRIM(@傳票號碼A)) AND (''=TRIM(@傳票號碼B) OR 傳票資料.傳票號碼 LIKE N'%'+TRIM(@傳票號碼B))) 
                        OR (LEN(TRIM(@傳票號碼A)) = 7 AND LEN(TRIM(@傳票號碼B)) = 7 AND 傳票資料.傳票號碼 BETWEEN TRIM(@傳票號碼A) AND TRIM(@傳票號碼B)) 
                        OR (LEN(TRIM(@傳票號碼A)) = 7 AND LEN(TRIM(@傳票號碼B)) BETWEEN 1 AND 6 AND 傳票資料.傳票號碼 BETWEEN TRIM(@傳票號碼A) AND LEFT(TRIM(@傳票號碼A), 1) + RIGHT('000000' + TRIM(@傳票號碼B), 6)) 
                        OR (LEN(TRIM(@傳票號碼A)) BETWEEN 1 AND 6 AND LEN(TRIM(@傳票號碼B)) = 7 AND 傳票資料.傳票號碼 BETWEEN LEFT(TRIM(@傳票號碼B), 1) + RIGHT('000000' + TRIM(@傳票號碼A), 6) AND TRIM(@傳票號碼B)) 
                        OR (LEN(TRIM(@傳票號碼A)) BETWEEN 1 AND 6 AND LEN(TRIM(@傳票號碼B)) BETWEEN 1 AND 6 AND RIGHT(傳票資料.傳票號碼, 6) BETWEEN RIGHT('000000' + TRIM(@傳票號碼A), 6) AND RIGHT('000000' + TRIM(@傳票號碼B), 6))) 
                    AND (''=TRIM(@收入或支出金額) OR 傳票資料.收入金額 LIKE REPLACE(TRIM(@收入或支出金額), ',', '') OR 傳票資料.支出金額 LIKE REPLACE(TRIM(@收入或支出金額), ',', '')) 
                    AND (''=TRIM(@名稱) OR 名稱 LIKE N'%' + TRIM(@名稱) + '%') 
                    AND (''=TRIM(@登錄序號) OR 登錄序號 LIKE N'%' + TRIM(@登錄序號) + '%') 
                    AND (現金備查簿.收入金額405 > 0 OR 現金備查簿.支出金額405 > 0) 
                    ORDER BY 
                    CASE WHEN 傳票資料.傳票送出納檔名 IS NULL THEN 1 ELSE 0 END, 傳票資料.傳票送出納檔名, 
                    CASE WHEN 傳票資料.傳票號碼 IS NULL THEN 1 ELSE 0 END, 傳票資料.傳票號碼, 
                    傳票資料.之, 
                    傳票資料.支出金額 DESC 
                END 
            ELSE IF @選項 = '土銀405匯款' 
                BEGIN 
                    SELECT 傳票資料.*, 現金備查簿.會計科目及摘要, CASE WHEN (SELECT COUNT(*) FROM 收款人 WHERE 收款人.匯入帳號 = 傳票資料.匯入帳號 AND 收款人.收款人匯款戶名 = 傳票資料.收款人匯款戶名) = 1 THEN 1 ELSE 0 END AS 有效 
                    FROM 傳票資料 INNER JOIN 現金備查簿 ON 傳票資料.年 = 現金備查簿.年 AND 傳票資料.傳票號碼 = 現金備查簿.傳票號碼 
                    WHERE (''=TRIM(@年) OR 傳票資料.年 LIKE TRIM(@年)) 
                    AND (((''=TRIM(@開票日期A) OR 傳票資料.開票日期 LIKE N'%'+TRIM(@開票日期A)) AND (''=TRIM(@開票日期B) OR 傳票資料.開票日期 LIKE N'%'+TRIM(@開票日期B))) 
                        OR (LEN(TRIM(@開票日期A)) = 7 AND LEN(TRIM(@開票日期B)) = 7 AND 傳票資料.開票日期 BETWEEN TRIM(@開票日期A) AND TRIM(@開票日期B)) 
                        OR (LEN(TRIM(@開票日期A)) = 7 AND LEN(TRIM(@開票日期B)) BETWEEN 1 AND 6 AND 傳票資料.開票日期 BETWEEN TRIM(@開票日期A) AND LEFT(TRIM(@開票日期A), 3) + RIGHT('0000' + TRIM(@開票日期B), 4)) 
                        OR (LEN(TRIM(@開票日期A)) BETWEEN 1 AND 6 AND LEN(TRIM(@開票日期B)) = 7 AND 傳票資料.開票日期 BETWEEN LEFT(TRIM(@開票日期B), 3) + RIGHT('0000' + TRIM(@開票日期A), 4) AND TRIM(@開票日期B)) 
                        OR (LEN(TRIM(@開票日期A)) BETWEEN 1 AND 6 AND LEN(TRIM(@開票日期B)) BETWEEN 1 AND 6 AND RIGHT(傳票資料.開票日期, 4) BETWEEN RIGHT('0000' + TRIM(@開票日期A), 4) AND RIGHT('0000' + TRIM(@開票日期B), 4))) 
                    AND (((''=TRIM(@傳票號碼A) OR 傳票資料.傳票號碼 LIKE N'%'+TRIM(@傳票號碼A)) AND (''=TRIM(@傳票號碼B) OR 傳票資料.傳票號碼 LIKE N'%'+TRIM(@傳票號碼B))) 
                        OR (LEN(TRIM(@傳票號碼A)) = 7 AND LEN(TRIM(@傳票號碼B)) = 7 AND 傳票資料.傳票號碼 BETWEEN TRIM(@傳票號碼A) AND TRIM(@傳票號碼B)) 
                        OR (LEN(TRIM(@傳票號碼A)) = 7 AND LEN(TRIM(@傳票號碼B)) BETWEEN 1 AND 6 AND 傳票資料.傳票號碼 BETWEEN TRIM(@傳票號碼A) AND LEFT(TRIM(@傳票號碼A), 1) + RIGHT('000000' + TRIM(@傳票號碼B), 6)) 
                        OR (LEN(TRIM(@傳票號碼A)) BETWEEN 1 AND 6 AND LEN(TRIM(@傳票號碼B)) = 7 AND 傳票資料.傳票號碼 BETWEEN LEFT(TRIM(@傳票號碼B), 1) + RIGHT('000000' + TRIM(@傳票號碼A), 6) AND TRIM(@傳票號碼B)) 
                        OR (LEN(TRIM(@傳票號碼A)) BETWEEN 1 AND 6 AND LEN(TRIM(@傳票號碼B)) BETWEEN 1 AND 6 AND RIGHT(傳票資料.傳票號碼, 6) BETWEEN RIGHT('000000' + TRIM(@傳票號碼A), 6) AND RIGHT('000000' + TRIM(@傳票號碼B), 6))) 
                    AND (''=TRIM(@收入或支出金額) OR 傳票資料.收入金額 LIKE REPLACE(TRIM(@收入或支出金額), ',', '') OR 傳票資料.支出金額 LIKE REPLACE(TRIM(@收入或支出金額), ',', '')) 
                    AND (''=TRIM(@名稱) OR 名稱 LIKE N'%' + TRIM(@名稱) + '%') 
                    AND (''=TRIM(@登錄序號) OR 登錄序號 LIKE N'%' + TRIM(@登錄序號) + '%') 
                    AND (''=TRIM(@匯入帳號) OR 傳票資料.匯入帳號 LIKE N'%'+TRIM(@匯入帳號)+'%') 
                    AND (現金備查簿.收入金額405 > 0 OR 現金備查簿.支出金額405 > 0) 
                    AND (傳票資料.匯入帳號 IS NOT NULL AND 傳票資料.匯入帳號 != '') 
                    AND (@預覽 = '預覽' OR 下載 = 1) 
                    AND (@TXT檔名 LIKE '選擇%' OR TXT檔名 = @TXT檔名) 
                    ORDER BY 
                    CASE WHEN @預覽 <> '預覽' OR @TXT檔名 NOT LIKE '選擇%' THEN 傳票資料.順序 ELSE 0 END, 
                    CASE WHEN 傳票資料.傳票送出納檔名 IS NULL THEN 1 ELSE 0 END, 傳票資料.傳票送出納檔名, 
                    CASE WHEN 傳票資料.傳票號碼 IS NULL THEN 1 ELSE 0 END, 傳票資料.傳票號碼, 
                    傳票資料.之, 
                    傳票資料.支出金額 DESC 
                END 
            ELSE IF @選項 = '土銀405支票' 
                BEGIN 
                    SELECT 傳票資料.*, 現金備查簿.會計科目及摘要, CASE WHEN (SELECT COUNT(*) FROM 收款人 WHERE 收款人.匯入帳號 = 傳票資料.匯入帳號 AND 收款人.收款人匯款戶名 = 傳票資料.收款人匯款戶名) = 1 THEN 1 ELSE 0 END AS 有效 
                    FROM 傳票資料 INNER JOIN 現金備查簿 ON 傳票資料.年 = 現金備查簿.年 AND 傳票資料.傳票號碼 = 現金備查簿.傳票號碼 
                    WHERE (''=TRIM(@年) OR 傳票資料.年 LIKE TRIM(@年)) 
                    AND (((''=TRIM(@開票日期A) OR 傳票資料.開票日期 LIKE N'%'+TRIM(@開票日期A)) AND (''=TRIM(@開票日期B) OR 傳票資料.開票日期 LIKE N'%'+TRIM(@開票日期B))) 
                        OR (LEN(TRIM(@開票日期A)) = 7 AND LEN(TRIM(@開票日期B)) = 7 AND 傳票資料.開票日期 BETWEEN TRIM(@開票日期A) AND TRIM(@開票日期B)) 
                        OR (LEN(TRIM(@開票日期A)) = 7 AND LEN(TRIM(@開票日期B)) BETWEEN 1 AND 6 AND 傳票資料.開票日期 BETWEEN TRIM(@開票日期A) AND LEFT(TRIM(@開票日期A), 3) + RIGHT('0000' + TRIM(@開票日期B), 4)) 
                        OR (LEN(TRIM(@開票日期A)) BETWEEN 1 AND 6 AND LEN(TRIM(@開票日期B)) = 7 AND 傳票資料.開票日期 BETWEEN LEFT(TRIM(@開票日期B), 3) + RIGHT('0000' + TRIM(@開票日期A), 4) AND TRIM(@開票日期B)) 
                        OR (LEN(TRIM(@開票日期A)) BETWEEN 1 AND 6 AND LEN(TRIM(@開票日期B)) BETWEEN 1 AND 6 AND RIGHT(傳票資料.開票日期, 4) BETWEEN RIGHT('0000' + TRIM(@開票日期A), 4) AND RIGHT('0000' + TRIM(@開票日期B), 4))) 
                    AND (((''=TRIM(@傳票號碼A) OR 傳票資料.傳票號碼 LIKE N'%'+TRIM(@傳票號碼A)) AND (''=TRIM(@傳票號碼B) OR 傳票資料.傳票號碼 LIKE N'%'+TRIM(@傳票號碼B))) 
                        OR (LEN(TRIM(@傳票號碼A)) = 7 AND LEN(TRIM(@傳票號碼B)) = 7 AND 傳票資料.傳票號碼 BETWEEN TRIM(@傳票號碼A) AND TRIM(@傳票號碼B)) 
                        OR (LEN(TRIM(@傳票號碼A)) = 7 AND LEN(TRIM(@傳票號碼B)) BETWEEN 1 AND 6 AND 傳票資料.傳票號碼 BETWEEN TRIM(@傳票號碼A) AND LEFT(TRIM(@傳票號碼A), 1) + RIGHT('000000' + TRIM(@傳票號碼B), 6)) 
                        OR (LEN(TRIM(@傳票號碼A)) BETWEEN 1 AND 6 AND LEN(TRIM(@傳票號碼B)) = 7 AND 傳票資料.傳票號碼 BETWEEN LEFT(TRIM(@傳票號碼B), 1) + RIGHT('000000' + TRIM(@傳票號碼A), 6) AND TRIM(@傳票號碼B)) 
                        OR (LEN(TRIM(@傳票號碼A)) BETWEEN 1 AND 6 AND LEN(TRIM(@傳票號碼B)) BETWEEN 1 AND 6 AND RIGHT(傳票資料.傳票號碼, 6) BETWEEN RIGHT('000000' + TRIM(@傳票號碼A), 6) AND RIGHT('000000' + TRIM(@傳票號碼B), 6))) 
                    AND (''=TRIM(@收入或支出金額) OR 傳票資料.收入金額 LIKE REPLACE(TRIM(@收入或支出金額), ',', '') OR 傳票資料.支出金額 LIKE REPLACE(TRIM(@收入或支出金額), ',', '')) 
                    AND (''=TRIM(@名稱) OR 名稱 LIKE N'%' + TRIM(@名稱) + '%') 
                    AND (''=TRIM(@登錄序號) OR 登錄序號 LIKE N'%' + TRIM(@登錄序號) + '%') 
                    AND (現金備查簿.收入金額405 > 0 OR 現金備查簿.支出金額405 > 0) 
                    AND NOT (傳票資料.匯入帳號 IS NOT NULL AND 傳票資料.匯入帳號 != '') 
                    AND (傳票資料.收入金額 = 0 OR 傳票資料.收入金額 IS NULL) 
                    AND (@預覽 = '預覽' OR 下載 = 1) 
                    AND (@TXT檔名 LIKE '選擇%' OR 登錄序號 = @TXT檔名) 
                    ORDER BY 
                    CASE WHEN 傳票資料.傳票送出納檔名 IS NULL THEN 1 ELSE 0 END, 傳票資料.傳票送出納檔名, 
                    CASE WHEN 傳票資料.傳票號碼 IS NULL THEN 1 ELSE 0 END, 傳票資料.傳票號碼, 
                    傳票資料.之, 
                    傳票資料.支出金額 DESC 
                END
            ELSE IF @選項 = '土銀405收入' 
                BEGIN 
                    SELECT 傳票資料.*, 現金備查簿.會計科目及摘要, CASE WHEN (SELECT COUNT(*) FROM 收款人 WHERE 收款人.匯入帳號 = 傳票資料.匯入帳號 AND 收款人.收款人匯款戶名 = 傳票資料.收款人匯款戶名) = 1 THEN 1 ELSE 0 END AS 有效 
                    FROM 傳票資料 INNER JOIN 現金備查簿 ON 傳票資料.年 = 現金備查簿.年 AND 傳票資料.傳票號碼 = 現金備查簿.傳票號碼 
                    WHERE (''=TRIM(@年) OR 傳票資料.年 LIKE TRIM(@年)) 
                    AND (((''=TRIM(@開票日期A) OR 傳票資料.開票日期 LIKE N'%'+TRIM(@開票日期A)) AND (''=TRIM(@開票日期B) OR 傳票資料.開票日期 LIKE N'%'+TRIM(@開票日期B))) 
                        OR (LEN(TRIM(@開票日期A)) = 7 AND LEN(TRIM(@開票日期B)) = 7 AND 傳票資料.開票日期 BETWEEN TRIM(@開票日期A) AND TRIM(@開票日期B)) 
                        OR (LEN(TRIM(@開票日期A)) = 7 AND LEN(TRIM(@開票日期B)) BETWEEN 1 AND 6 AND 傳票資料.開票日期 BETWEEN TRIM(@開票日期A) AND LEFT(TRIM(@開票日期A), 3) + RIGHT('0000' + TRIM(@開票日期B), 4)) 
                        OR (LEN(TRIM(@開票日期A)) BETWEEN 1 AND 6 AND LEN(TRIM(@開票日期B)) = 7 AND 傳票資料.開票日期 BETWEEN LEFT(TRIM(@開票日期B), 3) + RIGHT('0000' + TRIM(@開票日期A), 4) AND TRIM(@開票日期B)) 
                        OR (LEN(TRIM(@開票日期A)) BETWEEN 1 AND 6 AND LEN(TRIM(@開票日期B)) BETWEEN 1 AND 6 AND RIGHT(傳票資料.開票日期, 4) BETWEEN RIGHT('0000' + TRIM(@開票日期A), 4) AND RIGHT('0000' + TRIM(@開票日期B), 4))) 
                    AND (((''=TRIM(@傳票號碼A) OR 傳票資料.傳票號碼 LIKE N'%'+TRIM(@傳票號碼A)) AND (''=TRIM(@傳票號碼B) OR 傳票資料.傳票號碼 LIKE N'%'+TRIM(@傳票號碼B))) 
                        OR (LEN(TRIM(@傳票號碼A)) = 7 AND LEN(TRIM(@傳票號碼B)) = 7 AND 傳票資料.傳票號碼 BETWEEN TRIM(@傳票號碼A) AND TRIM(@傳票號碼B)) 
                        OR (LEN(TRIM(@傳票號碼A)) = 7 AND LEN(TRIM(@傳票號碼B)) BETWEEN 1 AND 6 AND 傳票資料.傳票號碼 BETWEEN TRIM(@傳票號碼A) AND LEFT(TRIM(@傳票號碼A), 1) + RIGHT('000000' + TRIM(@傳票號碼B), 6)) 
                        OR (LEN(TRIM(@傳票號碼A)) BETWEEN 1 AND 6 AND LEN(TRIM(@傳票號碼B)) = 7 AND 傳票資料.傳票號碼 BETWEEN LEFT(TRIM(@傳票號碼B), 1) + RIGHT('000000' + TRIM(@傳票號碼A), 6) AND TRIM(@傳票號碼B)) 
                        OR (LEN(TRIM(@傳票號碼A)) BETWEEN 1 AND 6 AND LEN(TRIM(@傳票號碼B)) BETWEEN 1 AND 6 AND RIGHT(傳票資料.傳票號碼, 6) BETWEEN RIGHT('000000' + TRIM(@傳票號碼A), 6) AND RIGHT('000000' + TRIM(@傳票號碼B), 6))) 
                    AND (''=TRIM(@收入或支出金額) OR 傳票資料.收入金額 LIKE REPLACE(TRIM(@收入或支出金額), ',', '') OR 傳票資料.支出金額 LIKE REPLACE(TRIM(@收入或支出金額), ',', '')) 
                    AND (''=TRIM(@名稱) OR 名稱 LIKE N'%' + TRIM(@名稱) + '%') 
                    AND (''=TRIM(@登錄序號) OR 登錄序號 LIKE N'%' + TRIM(@登錄序號) + '%') 
                    AND (現金備查簿.收入金額405 > 0 OR 現金備查簿.支出金額405 > 0) 
                    AND NOT (傳票資料.匯入帳號 IS NOT NULL AND 傳票資料.匯入帳號 != '') 
                    AND (傳票資料.收入金額 > 0)
                    ORDER BY 
                    CASE WHEN 傳票資料.傳票送出納檔名 IS NULL THEN 1 ELSE 0 END, 傳票資料.傳票送出納檔名, 
                    CASE WHEN 傳票資料.傳票號碼 IS NULL THEN 1 ELSE 0 END, 傳票資料.傳票號碼, 
                    傳票資料.之, 
                    傳票資料.支出金額 DESC 
                END
            ELSE IF @選項 = '中國信託409全' 
                BEGIN 
                    SELECT 傳票資料.*, 現金備查簿.會計科目及摘要, CASE WHEN (SELECT COUNT(*) FROM 收款人 WHERE 收款人.匯入帳號 = 傳票資料.匯入帳號 AND 收款人.收款人匯款戶名 = 傳票資料.收款人匯款戶名) = 1 THEN 1 ELSE 0 END AS 有效 
                    FROM 傳票資料 INNER JOIN 現金備查簿 ON 傳票資料.年 = 現金備查簿.年 AND 傳票資料.傳票號碼 = 現金備查簿.傳票號碼
                    WHERE (''=TRIM(@年) OR 傳票資料.年 LIKE TRIM(@年)) 
                    AND (((''=TRIM(@開票日期A) OR 傳票資料.開票日期 LIKE N'%'+TRIM(@開票日期A)) AND (''=TRIM(@開票日期B) OR 傳票資料.開票日期 LIKE N'%'+TRIM(@開票日期B))) 
                        OR (LEN(TRIM(@開票日期A)) = 7 AND LEN(TRIM(@開票日期B)) = 7 AND 傳票資料.開票日期 BETWEEN TRIM(@開票日期A) AND TRIM(@開票日期B)) 
                        OR (LEN(TRIM(@開票日期A)) = 7 AND LEN(TRIM(@開票日期B)) BETWEEN 1 AND 6 AND 傳票資料.開票日期 BETWEEN TRIM(@開票日期A) AND LEFT(TRIM(@開票日期A), 3) + RIGHT('0000' + TRIM(@開票日期B), 4)) 
                        OR (LEN(TRIM(@開票日期A)) BETWEEN 1 AND 6 AND LEN(TRIM(@開票日期B)) = 7 AND 傳票資料.開票日期 BETWEEN LEFT(TRIM(@開票日期B), 3) + RIGHT('0000' + TRIM(@開票日期A), 4) AND TRIM(@開票日期B)) 
                        OR (LEN(TRIM(@開票日期A)) BETWEEN 1 AND 6 AND LEN(TRIM(@開票日期B)) BETWEEN 1 AND 6 AND RIGHT(傳票資料.開票日期, 4) BETWEEN RIGHT('0000' + TRIM(@開票日期A), 4) AND RIGHT('0000' + TRIM(@開票日期B), 4))) 
                    AND (((''=TRIM(@傳票號碼A) OR 傳票資料.傳票號碼 LIKE N'%'+TRIM(@傳票號碼A)) AND (''=TRIM(@傳票號碼B) OR 傳票資料.傳票號碼 LIKE N'%'+TRIM(@傳票號碼B))) 
                        OR (LEN(TRIM(@傳票號碼A)) = 7 AND LEN(TRIM(@傳票號碼B)) = 7 AND 傳票資料.傳票號碼 BETWEEN TRIM(@傳票號碼A) AND TRIM(@傳票號碼B)) 
                        OR (LEN(TRIM(@傳票號碼A)) = 7 AND LEN(TRIM(@傳票號碼B)) BETWEEN 1 AND 6 AND 傳票資料.傳票號碼 BETWEEN TRIM(@傳票號碼A) AND LEFT(TRIM(@傳票號碼A), 1) + RIGHT('000000' + TRIM(@傳票號碼B), 6)) 
                        OR (LEN(TRIM(@傳票號碼A)) BETWEEN 1 AND 6 AND LEN(TRIM(@傳票號碼B)) = 7 AND 傳票資料.傳票號碼 BETWEEN LEFT(TRIM(@傳票號碼B), 1) + RIGHT('000000' + TRIM(@傳票號碼A), 6) AND TRIM(@傳票號碼B)) 
                        OR (LEN(TRIM(@傳票號碼A)) BETWEEN 1 AND 6 AND LEN(TRIM(@傳票號碼B)) BETWEEN 1 AND 6 AND RIGHT(傳票資料.傳票號碼, 6) BETWEEN RIGHT('000000' + TRIM(@傳票號碼A), 6) AND RIGHT('000000' + TRIM(@傳票號碼B), 6))) 
                    AND (''=TRIM(@收入或支出金額) OR 傳票資料.收入金額 LIKE REPLACE(TRIM(@收入或支出金額), ',', '') OR 傳票資料.支出金額 LIKE REPLACE(TRIM(@收入或支出金額), ',', '')) 
                    AND (''=TRIM(@名稱) OR 名稱 LIKE N'%' + TRIM(@名稱) + '%') 
                    AND (''=TRIM(@登錄序號) OR 登錄序號 LIKE N'%' + TRIM(@登錄序號) + '%') 
                    AND (現金備查簿.收入金額409 > 0 OR 現金備查簿.支出金額409 > 0) 
                    ORDER BY 
                    CASE WHEN 傳票資料.傳票送出納檔名 IS NULL THEN 1 ELSE 0 END, 傳票資料.傳票送出納檔名, 
                    CASE WHEN 傳票資料.傳票號碼 IS NULL THEN 1 ELSE 0 END, 傳票資料.傳票號碼, 
                    傳票資料.之, 
                    傳票資料.支出金額 DESC 
                END
            ELSE IF @選項 = '中國信託409收入' 
                BEGIN 
                    SELECT 傳票資料.*, 現金備查簿.會計科目及摘要, CASE WHEN (SELECT COUNT(*) FROM 收款人 WHERE 收款人.匯入帳號 = 傳票資料.匯入帳號 AND 收款人.收款人匯款戶名 = 傳票資料.收款人匯款戶名) = 1 THEN 1 ELSE 0 END AS 有效 
                    FROM 傳票資料 INNER JOIN 現金備查簿 ON 傳票資料.年 = 現金備查簿.年 AND 傳票資料.傳票號碼 = 現金備查簿.傳票號碼
                    WHERE (''=TRIM(@年) OR 傳票資料.年 LIKE TRIM(@年)) 
                    AND (((''=TRIM(@開票日期A) OR 傳票資料.開票日期 LIKE N'%'+TRIM(@開票日期A)) AND (''=TRIM(@開票日期B) OR 傳票資料.開票日期 LIKE N'%'+TRIM(@開票日期B))) 
                        OR (LEN(TRIM(@開票日期A)) = 7 AND LEN(TRIM(@開票日期B)) = 7 AND 傳票資料.開票日期 BETWEEN TRIM(@開票日期A) AND TRIM(@開票日期B)) 
                        OR (LEN(TRIM(@開票日期A)) = 7 AND LEN(TRIM(@開票日期B)) BETWEEN 1 AND 6 AND 傳票資料.開票日期 BETWEEN TRIM(@開票日期A) AND LEFT(TRIM(@開票日期A), 3) + RIGHT('0000' + TRIM(@開票日期B), 4)) 
                        OR (LEN(TRIM(@開票日期A)) BETWEEN 1 AND 6 AND LEN(TRIM(@開票日期B)) = 7 AND 傳票資料.開票日期 BETWEEN LEFT(TRIM(@開票日期B), 3) + RIGHT('0000' + TRIM(@開票日期A), 4) AND TRIM(@開票日期B)) 
                        OR (LEN(TRIM(@開票日期A)) BETWEEN 1 AND 6 AND LEN(TRIM(@開票日期B)) BETWEEN 1 AND 6 AND RIGHT(傳票資料.開票日期, 4) BETWEEN RIGHT('0000' + TRIM(@開票日期A), 4) AND RIGHT('0000' + TRIM(@開票日期B), 4))) 
                    AND (((''=TRIM(@傳票號碼A) OR 傳票資料.傳票號碼 LIKE N'%'+TRIM(@傳票號碼A)) AND (''=TRIM(@傳票號碼B) OR 傳票資料.傳票號碼 LIKE N'%'+TRIM(@傳票號碼B))) 
                        OR (LEN(TRIM(@傳票號碼A)) = 7 AND LEN(TRIM(@傳票號碼B)) = 7 AND 傳票資料.傳票號碼 BETWEEN TRIM(@傳票號碼A) AND TRIM(@傳票號碼B)) 
                        OR (LEN(TRIM(@傳票號碼A)) = 7 AND LEN(TRIM(@傳票號碼B)) BETWEEN 1 AND 6 AND 傳票資料.傳票號碼 BETWEEN TRIM(@傳票號碼A) AND LEFT(TRIM(@傳票號碼A), 1) + RIGHT('000000' + TRIM(@傳票號碼B), 6)) 
                        OR (LEN(TRIM(@傳票號碼A)) BETWEEN 1 AND 6 AND LEN(TRIM(@傳票號碼B)) = 7 AND 傳票資料.傳票號碼 BETWEEN LEFT(TRIM(@傳票號碼B), 1) + RIGHT('000000' + TRIM(@傳票號碼A), 6) AND TRIM(@傳票號碼B)) 
                        OR (LEN(TRIM(@傳票號碼A)) BETWEEN 1 AND 6 AND LEN(TRIM(@傳票號碼B)) BETWEEN 1 AND 6 AND RIGHT(傳票資料.傳票號碼, 6) BETWEEN RIGHT('000000' + TRIM(@傳票號碼A), 6) AND RIGHT('000000' + TRIM(@傳票號碼B), 6))) 
                    AND (''=TRIM(@收入或支出金額) OR 傳票資料.收入金額 LIKE REPLACE(TRIM(@收入或支出金額), ',', '') OR 傳票資料.支出金額 LIKE REPLACE(TRIM(@收入或支出金額), ',', '')) 
                    AND (''=TRIM(@名稱) OR 名稱 LIKE N'%' + TRIM(@名稱) + '%') 
                    AND (''=TRIM(@登錄序號) OR 登錄序號 LIKE N'%' + TRIM(@登錄序號) + '%') 
                    AND (現金備查簿.收入金額409 > 0) 
                    ORDER BY 
                    CASE WHEN 傳票資料.傳票送出納檔名 IS NULL THEN 1 ELSE 0 END, 傳票資料.傳票送出納檔名, 
                    CASE WHEN 傳票資料.傳票號碼 IS NULL THEN 1 ELSE 0 END, 傳票資料.傳票號碼, 
                    傳票資料.之, 
                    傳票資料.支出金額 DESC 
                END
            ELSE IF @選項 = '中國信託409支出' 
                BEGIN 
                    SELECT 傳票資料.*, 現金備查簿.會計科目及摘要, CASE WHEN (SELECT COUNT(*) FROM 收款人 WHERE 收款人.匯入帳號 = 傳票資料.匯入帳號 AND 收款人.收款人匯款戶名 = 傳票資料.收款人匯款戶名) = 1 THEN 1 ELSE 0 END AS 有效 
                    FROM 傳票資料 INNER JOIN 現金備查簿 ON 傳票資料.年 = 現金備查簿.年 AND 傳票資料.傳票號碼 = 現金備查簿.傳票號碼
                    WHERE (''=TRIM(@年) OR 傳票資料.年 LIKE TRIM(@年)) 
                    AND (((''=TRIM(@開票日期A) OR 傳票資料.開票日期 LIKE N'%'+TRIM(@開票日期A)) AND (''=TRIM(@開票日期B) OR 傳票資料.開票日期 LIKE N'%'+TRIM(@開票日期B))) 
                        OR (LEN(TRIM(@開票日期A)) = 7 AND LEN(TRIM(@開票日期B)) = 7 AND 傳票資料.開票日期 BETWEEN TRIM(@開票日期A) AND TRIM(@開票日期B)) 
                        OR (LEN(TRIM(@開票日期A)) = 7 AND LEN(TRIM(@開票日期B)) BETWEEN 1 AND 6 AND 傳票資料.開票日期 BETWEEN TRIM(@開票日期A) AND LEFT(TRIM(@開票日期A), 3) + RIGHT('0000' + TRIM(@開票日期B), 4)) 
                        OR (LEN(TRIM(@開票日期A)) BETWEEN 1 AND 6 AND LEN(TRIM(@開票日期B)) = 7 AND 傳票資料.開票日期 BETWEEN LEFT(TRIM(@開票日期B), 3) + RIGHT('0000' + TRIM(@開票日期A), 4) AND TRIM(@開票日期B)) 
                        OR (LEN(TRIM(@開票日期A)) BETWEEN 1 AND 6 AND LEN(TRIM(@開票日期B)) BETWEEN 1 AND 6 AND RIGHT(傳票資料.開票日期, 4) BETWEEN RIGHT('0000' + TRIM(@開票日期A), 4) AND RIGHT('0000' + TRIM(@開票日期B), 4))) 
                    AND (((''=TRIM(@傳票號碼A) OR 傳票資料.傳票號碼 LIKE N'%'+TRIM(@傳票號碼A)) AND (''=TRIM(@傳票號碼B) OR 傳票資料.傳票號碼 LIKE N'%'+TRIM(@傳票號碼B))) 
                        OR (LEN(TRIM(@傳票號碼A)) = 7 AND LEN(TRIM(@傳票號碼B)) = 7 AND 傳票資料.傳票號碼 BETWEEN TRIM(@傳票號碼A) AND TRIM(@傳票號碼B)) 
                        OR (LEN(TRIM(@傳票號碼A)) = 7 AND LEN(TRIM(@傳票號碼B)) BETWEEN 1 AND 6 AND 傳票資料.傳票號碼 BETWEEN TRIM(@傳票號碼A) AND LEFT(TRIM(@傳票號碼A), 1) + RIGHT('000000' + TRIM(@傳票號碼B), 6)) 
                        OR (LEN(TRIM(@傳票號碼A)) BETWEEN 1 AND 6 AND LEN(TRIM(@傳票號碼B)) = 7 AND 傳票資料.傳票號碼 BETWEEN LEFT(TRIM(@傳票號碼B), 1) + RIGHT('000000' + TRIM(@傳票號碼A), 6) AND TRIM(@傳票號碼B)) 
                        OR (LEN(TRIM(@傳票號碼A)) BETWEEN 1 AND 6 AND LEN(TRIM(@傳票號碼B)) BETWEEN 1 AND 6 AND RIGHT(傳票資料.傳票號碼, 6) BETWEEN RIGHT('000000' + TRIM(@傳票號碼A), 6) AND RIGHT('000000' + TRIM(@傳票號碼B), 6))) 
                    AND (''=TRIM(@收入或支出金額) OR 傳票資料.收入金額 LIKE REPLACE(TRIM(@收入或支出金額), ',', '') OR 傳票資料.支出金額 LIKE REPLACE(TRIM(@收入或支出金額), ',', '')) 
                    AND (''=TRIM(@名稱) OR 名稱 LIKE N'%' + TRIM(@名稱) + '%') 
                    AND (''=TRIM(@登錄序號) OR 登錄序號 LIKE N'%' + TRIM(@登錄序號) + '%') 
                    AND (現金備查簿.支出金額409 > 0) 
                    AND (@預覽 = '預覽' OR 下載 = 1) 
                    ORDER BY 
                    CASE WHEN 傳票資料.傳票送出納檔名 IS NULL THEN 1 ELSE 0 END, 傳票資料.傳票送出納檔名, 
                    CASE WHEN 傳票資料.傳票號碼 IS NULL THEN 1 ELSE 0 END, 傳票資料.傳票號碼, 
                    傳票資料.之, 
                    傳票資料.支出金額 DESC 
                END" 
        InsertCommand="INSERT INTO 傳票資料 (年) VALUES (NULLIF(N''+@年+'', ''))" 
        DeleteCommand="DELETE FROM 傳票資料 WHERE id = @id">
        <SelectParameters>
            <asp:ControlParameter ControlID="DropDownList1" ConvertEmptyStringToNull="False" Name="選項" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="TextBox1" ConvertEmptyStringToNull="False" Name="年" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="TextBox2" ConvertEmptyStringToNull="False" Name="開票日期A" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="TextBox3" ConvertEmptyStringToNull="False" Name="開票日期B" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="TextBox4" ConvertEmptyStringToNull="False" Name="傳票號碼A" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="TextBox5" ConvertEmptyStringToNull="False" Name="傳票號碼B" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="TextBox6" ConvertEmptyStringToNull="False" Name="收入或支出金額" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="TextBox7" ConvertEmptyStringToNull="False" Name="匯入帳號" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="Button5" ConvertEmptyStringToNull="False" Name="預覽" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="DropDownList2" ConvertEmptyStringToNull="False" Name="TXT檔名" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="名稱" ConvertEmptyStringToNull="False" Name="名稱" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="登錄序號s" ConvertEmptyStringToNull="False" Name="登錄序號" PropertyName="Text" Type="String"/>
        </SelectParameters>
        <InsertParameters>
            <asp:ControlParameter ControlID="TextBox1" ConvertEmptyStringToNull="False" Name="年" PropertyName="Text" Type="String"/>
        </InsertParameters>
        <DeleteParameters>
            <asp:Parameter Name="id" ConvertEmptyStringToNull="False" Type="String"/>
        </DeleteParameters>
    </asp:SqlDataSource>
    <asp:SqlDataSource ID="SqlDataSource2" runat="server" ConnectionString='<%$ ConnectionStrings:ApplicationServices%>'
        SelectCommand=
            "IF @選項 = '土銀405匯款'  
                BEGIN 
                    SELECT '選擇TXT檔名' AS TXT檔名 
                    UNION 
                    SELECT DISTINCT TXT檔名 FROM 傳票資料 WHERE TXT檔名 IS NOT NULL ORDER BY TXT檔名 DESC 
                END
            ELSE IF @選項 = '土銀405支票'  
                BEGIN 
                    SELECT '選擇支票編號' AS TXT檔名 
                    UNION 
                    SELECT DISTINCT 登錄序號 AS TXT檔名 
                    FROM 傳票資料 INNER JOIN 現金備查簿 ON 傳票資料.年 = 現金備查簿.年 AND 傳票資料.傳票號碼 = 現金備查簿.傳票號碼 
                    WHERE 登錄序號 IS NOT NULL 
                    AND (現金備查簿.收入金額405 > 0 OR 現金備查簿.支出金額405 > 0) 
                    AND NOT (傳票資料.匯入帳號 IS NOT NULL AND 傳票資料.匯入帳號 != '') 
                    AND (傳票資料.收入金額 = 0 OR 傳票資料.收入金額 IS NULL) 
                    ORDER BY TXT檔名 DESC 
                END
            ELSE IF @選項 = '中國信託409支出'  
                BEGIN 
                    SELECT '選擇支票編號' AS TXT檔名 
                    UNION 
                    SELECT DISTINCT 登錄序號 AS TXT檔名 
                    FROM 傳票資料 INNER JOIN 現金備查簿 ON 傳票資料.年 = 現金備查簿.年 AND 傳票資料.傳票號碼 = 現金備查簿.傳票號碼 
                    WHERE 登錄序號 IS NOT NULL 
                    AND (現金備查簿.支出金額409 > 0) 
                    ORDER BY TXT檔名 DESC 
                END">
        <SelectParameters>
            <asp:ControlParameter ControlID="DropDownList1" ConvertEmptyStringToNull="False" Name="選項" PropertyName="Text" Type="String"/>
        </SelectParameters>
    </asp:SqlDataSource>
</asp:Content>


