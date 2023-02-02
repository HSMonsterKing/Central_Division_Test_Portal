<%@ Page Language="VB" MasterPageFile="./MasterPage.master" MaintainScrollPositionOnPostback="true" AutoEventWireup="false" CodeFile="審核.aspx.vb" Inherits="審核" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
    <asp:ScriptManager ID="ScriptManager1" runat="server" EnablePageMethods = "true" EnableScriptGlobalization="True" EnableScriptLocalization="True"/>
    <script src='js/審核.js'></script>
    <link rel="stylesheet" runat="server" media="screen" href="css\審核.css"/>
    <div><h1><a id="Title" href="審核.aspx">審核<a></h1></div>
    <asp:Panel ID="Panel1" runat="server" DefaultButton="搜尋" CssClass="Panel1">
    年<asp:TextBox ID="年" runat="server" Maxlength=3 CssClass="Input2"/>
        種類<asp:DropDownList ID="_種類" runat="server" AutoPostBack="True" CssClass="DropDownList" OnSelectedIndexChanged="種類_SelectedIndexChanged">
            <asp:ListItem Text="A" Value="A"></asp:ListItem>
            <asp:ListItem Text="B" Value="B"></asp:ListItem>
            <asp:ListItem Text="XZ" Value="XZ"></asp:ListItem>
        </asp:DropDownList>
        <asp:Button ID="搜尋" runat="server" Text="搜尋" CssClass="GreenButton"/>
        <asp:Label ID="Label1" runat="server" Text="" CssClass="GreenLabel"/>
        <asp:Label ID="Label2" runat="server" Text="" CssClass="RedLabel"/>
    </asp:Panel>
    <asp:Panel ID="Panel3" runat="server" CssClass="Panel3">
        <asp:Button ID="通過" runat="server" Text="通過" OnClick="通過_Click" CssClass="GreenButton"/>
        <asp:Button ID="駁回" runat="server" Text="駁回" OnClick="駁回_Click" CssClass="RedButton"/>
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="false" CellPadding="1" HorizontalAlign="Center" DataSourceID="SqlDataSource1" DataKeyNames="id" GridLines="None" ShowHeaderWhenEmpty="false" AllowPaging="True" PageSize="15" AllowSorting="True" PagerSettings-PageButtonCount=25 EnableModelValidation="True" AutoGenerateEditButton="false">
            <Columns>
                <asp:TemplateField HeaderText="id" Visible="false">
                    <ItemTemplate>
                        <asp:TextBox ID="id" runat="server" Text='<%# Bind("id") %>' Maxlength=0 Enabled="False" CssClass="TextBox id"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="" Visible="false">
                    <ItemTemplate>
                        <asp:TextBox ID="_列" runat="server" Text='<%# Bind("_列") %>' Maxlength=0 Enabled="False" CssClass="TextBox _列"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="單位別">
                    <ItemTemplate>
                    <asp:Label ID="單位別" runat="server" Text='<%# Eval("單位別") %>' Maxlength=0 Enabled="True" CssClass="Label 單位別"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="承辦人">
                    <ItemTemplate>
                    <asp:Label ID="承辦人" runat="server" Text='<%# Eval("承辦人") %>' Maxlength=0 Enabled="True" CssClass="Label 承辦人"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="月">
                    <ItemTemplate>
                    <asp:Label ID="月" runat="server" Text='<%# Eval("月") %>' Maxlength=0 Enabled="True" CssClass="Label 月"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="日">
                    <ItemTemplate>
                        <asp:Label ID="日" runat="server" Text='<%# Eval("日") %>' Maxlength=0 Enabled="True" CssClass="Label 日"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="科目">
                    <ItemTemplate>
                    <asp:Label ID="科目" runat="server" Text='<%# Eval("科目") %>' Maxlength=0 Enabled="True" CssClass="Label 科目"/><BR>
                    <asp:Label ID="科目2" runat="server" Text='<%# Eval("科目2") %>' Maxlength=0 Enabled="True" CssClass="Label 科目" visible='<%# If (Eval("科目2").ToString = "", "False", "True") %>'/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="摘要">
                    <ItemTemplate>
                    <asp:Label ID="摘要" runat="server" Text='<%# Eval("摘要") %>' Maxlength=0 Enabled="True" CssClass="Label 摘要"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="姓名">
                    <ItemTemplate>
                        <asp:ImageButton ID="姓名" runat="server" ImageUrl='<%# If (Eval("姓名").ToString = "", "", Eval("姓名")) %>' AutoPostBack="False" OnClientClick="Alert('123456');return false;" CssClass="TextBox 姓名"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="商號">
                    <ItemTemplate>
                    <asp:Label ID="商號" runat="server" Text='<%# Eval("商號") %>' Maxlength=0 Enabled="True" CssClass="Label 商號"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="經手人">
                    <ItemTemplate>
                        <asp:ImageButton ID="經手人" runat="server" ImageUrl='<%# If (Eval("經手人").ToString = "", "", Eval("經手人")) %>' AutoPostBack="False" OnClientClick="Alert('123456');return false;" CssClass="TextBox 經手人"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="種類">
                    <ItemTemplate>
                    <asp:Label ID="種類" runat="server" Text='<%# Eval("種類") %>' Maxlength=0 Enabled="True" CssClass="Label 種類"/>
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
                <asp:TemplateField HeaderText="收入">
                    <ItemTemplate>
                    <asp:Label ID="收入" runat="server" Text='<%# Eval("收入", "{0:c0}") %>' Maxlength=0 Enabled="True" CssClass="Label 收入"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="支出">
                    <ItemTemplate>
                    <asp:Label ID="支出" runat="server" Text='<%# Eval("支出", "{0:c0}") %>' Maxlength=0 Enabled="True" CssClass="Label 支出"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="餘額">
                    <ItemTemplate>
                    <asp:Label ID="餘額" runat="server" Text='<%# Eval("餘額", "{0:c0}") %>' Maxlength=0 Enabled="True" CssClass="Label 餘額"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="備註">
                    <ItemTemplate>
                    <asp:Label ID="備註" runat="server" Text='<%# Eval("備註") %>' Maxlength=0 Enabled="True" CssClass="Label 備註"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="">
                    <ItemTemplate>
                    <asp:CheckBox ID="選取" runat="server" AutoPostBack="True" OnCheckedChanged="選取_CheckedChanged" CssClass="input"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="號數駁回原因">
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
        SelectCommand="
            SELECT * FROM 收支備查簿 
            WHERE 年 = @年 
            AND 取號 = 0 
            AND _種類 = @_種類
            AND 鎖定 = 'True'
            AND 過審 = 'false' 
            ORDER BY _頁, _列" 
        Insertcommand="" 
        UpdateCommand=""
        DeleteCommand="DELETE FROM 收支備查簿 WHERE id=@id">
        <SelectParameters>
            <asp:ControlParameter ControlID="年" ConvertEmptyStringToNull="False" Name="年" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="_種類" ConvertEmptyStringToNull="False" Name="_種類" PropertyName="Text" Type="String"/>
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
