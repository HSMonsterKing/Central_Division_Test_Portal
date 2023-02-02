<%@ Page Language="VB" MasterPageFile="./MasterPage.master" MaintainScrollPositionOnPostback="true" AutoEventWireup="false" CodeFile="主計室審核.aspx.vb" Inherits="主計室審核" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
    <asp:ScriptManager ID="ScriptManager1" runat="server" EnablePageMethods = "true" EnableScriptGlobalization="True" EnableScriptLocalization="True"/>
    <script src='js/主計室審核.js'></script>
    <link rel="stylesheet" runat="server" media="screen" href="css\主計室審核.css"/>
    <div><h1><a id="Title" href="主計室審核.aspx">主計室審核<a></h1></div>
    <asp:Panel ID="Panel1" runat="server" DefaultButton="搜尋" CssClass="Panel1" >
        年<asp:TextBox ID="年" runat="server" Maxlength=3 CssClass="Input2"/>
        種類<asp:DropDownList ID="_種類" runat="server" AutoPostBack="True" CssClass="DropDownList">
            <asp:ListItem Text="A" Value="A"></asp:ListItem>
            <asp:ListItem Text="B" Value="B"></asp:ListItem>
            <asp:ListItem Text="XZ" Value="XZ"></asp:ListItem>
        </asp:DropDownList>
        狀態<asp:DropDownList ID="狀態" runat="server" AutoPostBack="True" CssClass="DropDownList">
            <asp:ListItem Text="未經審核" Value="False"></asp:ListItem>
            <asp:ListItem Text="已經審核" Value="True"></asp:ListItem>
            <asp:ListItem Text="全部" Value=""></asp:ListItem>
        </asp:DropDownList>
        <asp:Button ID="搜尋" runat="server" Text="搜尋" CssClass="GreenButton"/> 
        <asp:Label ID="Label1" runat="server" Text="" CssClass="GreenLabel"/>
        <asp:Label ID="Label2" runat="server" Text="" CssClass="RedLabel"/>
    </asp:Panel>
    <asp:Panel ID="Panel3" runat="server" DefaultButton="" CssClass="Panel3">
        <asp:Button ID="回覆" runat="server" Text="回覆" OnClick="return_" OnClientClick="return confirm('確定回覆，未完成資料將不予處理?')" CssClass="GreenButton"/>
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
                        <asp:Label ID="等帳日期" runat="server" Text='<%# Eval("年","{0:000}")+"/"+Eval("月","{0:00}")+"/"+Eval("日","{0:00}") %>' Maxlength=0 Enabled="False" CssClass="Label 等帳日期"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="號數">
                    <ItemTemplate>
                        <asp:Label ID="號數" runat="server" Text='<%# Eval("號數", "{0:000}") %>' Maxlength=0 Enabled="False" CssClass="Label 號數"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="廠商名稱">
                    <ItemTemplate>
                        <asp:Label ID="廠商名稱" runat="server" Text='<%# Eval("商號") %>' TextMode="MultiLine" Maxlength=0 Enabled="False" CssClass="Label 廠商名稱"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="金額">
                    <ItemTemplate>
                        <asp:Label ID="金額" runat="server" Text='<%# Eval("支出2", "{0:c0}") %>' Maxlength=0 Enabled="False" CssClass="Label 金額"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="摘要">
                    <ItemTemplate>
                        <asp:Label ID="摘要" runat="server" Text='<%# Eval("摘要") %>' TextMode="MultiLine" Maxlength=0 Enabled="False" CssClass="Label 摘要"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="勾選送交主計室">
                    <ItemTemplate>
                    <asp:Label ID="主計室日期" runat="server" Text='<%# Eval("送交主計室日期")%>' Maxlength=0 Enabled="False" CssClass="Label 等帳日期"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="主計室回復情況">
                    <ItemTemplate>
                    <asp:RadioButtonList ID="回覆R" runat="server" AutoPostBack="True" SelectedIndex='<%# If (Eval("回覆").ToString() = "True", 0, 1) %>' RepeatDirection="Horizontal" OnSelectedIndexChanged="回覆R_OnSelectedIndexChanged">
                    <asp:ListItem>完成</asp:ListItem>
                    <asp:ListItem>未完成</asp:ListItem>
                    </asp:RadioButtonList>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                 <asp:TemplateField HeaderText="">
                    <ItemTemplate>
                    <asp:Button ID="駁回" runat="server" Text="駁回" OnClick="駁回_Click" Enabled='<%# If (Eval("回覆").ToString = "False", 1, 0) %>' CssClass="RedButton"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="駁回原因">
                    <ItemTemplate>
                    <asp:TextBox ID="駁回原因" runat="server" Text='<%# Bind("駁回原因") %>' Maxlength=0 Enabled="True" CssClass="TextBox 駁回原因"/>
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
        SelectCommand="SELECT  a.id,a._列,a.年,a.月,a.日,a.號數,a.商號,a.摘要,a.送交主計室日期,a.回覆,a.駁回原因,(CASE WHEN (left(REPLACE(REPLACE(a.摘要,' ',''),CHAR(13)+CHAR(10),''), 10) Like substring(REPLACE(REPLACE(b.摘要,' ',''),CHAR(13)+CHAR(10),''), 3, 10)
				AND a.摘要 not Like '行政訴訟%' ) 
				OR (a.id='417' And b.id='499') 
				THEN (a.支出-b.收入) 
				ELSE a.支出 END) As 支出2 
            FROM 收支備查簿 As a left Join 收支備查簿 As b
            ON (a.號數=b.號數 or b.號數 Is NULL AND (left(REPLACE(REPLACE(a.摘要,' ',''),CHAR(13)+CHAR(10),''), 10) Like substring(REPLACE(REPLACE(b.摘要,' ',''),CHAR(13)+CHAR(10),''), 3, 10)))
			AND b.摘要 Like'_回%'
			AND b._種類 = @_種類 
			AND b.收入 > 0
			WHere a.年 = @年 
            AND a.取號 = 0 
            AND a._種類 = @_種類 
            AND a.支出 > 0
            AND ((a.摘要<>'本月小計' AND a.摘要<>'累計至本月')or a.摘要 is null)
            AND a.過審 = 'True'
            And a.鎖定 = 'True'
            AND a.送交主計室日期 is not NULL
            AND (''=TRIM(@狀態) OR a.回覆 = TRIM(@狀態))
            ORDER BY a.回覆,a._頁, a._列" 
        Insertcommand="" 
        UpdateCommand=""
        DeleteCommand="">
        <SelectParameters>
            <asp:ControlParameter ControlID="年" ConvertEmptyStringToNull="False" Name="年" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="_種類" ConvertEmptyStringToNull="False" Name="_種類" PropertyName="Text" Type="String"/>
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
