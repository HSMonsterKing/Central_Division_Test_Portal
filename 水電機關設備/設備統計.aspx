<%@ Page Language="VB" MasterPageFile="./MasterPage.master" MaintainScrollPositionOnPostback="true" AutoEventWireup="false" CodeFile="設備統計.aspx.vb" Inherits="設備統計"%>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server" >
    <asp:ScriptManager ID="ScriptManager1" runat="server" EnablePageMethods = "true" EnableScriptGlobalization="True" EnableScriptLocalization="True"/>
    <script src='js/設備統計.js'></script>
    <link rel="stylesheet" runat="server" media="screen" href="css\設備統計.css"/>
    <style>
    </style>
    <div><h1><a id="Title" href="設備統計.aspx">設備統計<a></h1></div>
    <asp:Panel ID="Panel1" runat="server" DefaultButton="搜尋" CssClass="Panel1" >
        <asp:Button ID="搜尋" runat="server" Text="搜尋" CssClass="GreenButton" Visible="False"/>
        <asp:Button ID="新增" runat="server" Text="新增一頁" OnClick="Insert" CssClass="GreenButton"/>
        <asp:Button ID="存檔" runat="server" Text="存檔" OnClick="Update" CssClass="GreenButton"/>
        <asp:Button ID="下載" runat="server" Text="下載" OnClick="Download" CssClass="GreenButton"/>
        <asp:Button ID="年數偵測" runat="server" Text="年數偵測" OnClick="年數偵測_Click" CssClass="GreenButton"/>
        <asp:Button ID="刪除" runat="server" Text="刪除末頁" OnClick="Delete" OnClientClick="return confirm('確定刪除?')" CssClass="RedButton"/>
        狀態:<asp:DropDownList ID="狀態" runat="server" AutoPostBack="True" CssClass="DropDownList">
        <asp:ListItem Text="全部" Value=""></asp:ListItem>
        <asp:ListItem Text="過期" Value="過期"></asp:ListItem>
        <asp:ListItem Text="未過期" Value="未過期"></asp:ListItem>
        </asp:DropDownList>
        過期資料:<asp:CheckBox ID="過期資料" runat="server" AutoPostBack="True" CssClass="input"/>
        <asp:Button ID="測試" runat="server" Text="轉換" CssClass="GreenButton" OnClick="Test" Visible="False"/>
        <asp:Image ID="Image1" runat="server" Text="" CssClass="Image1"/>
        <asp:Label ID="Label1" runat="server" Text="" CssClass="GreenLabel"/>
        <asp:Label ID="Label2" runat="server" Text="" CssClass="RedLabel"/>
    </asp:Panel>
    <asp:Panel ID="Panel3" runat="server" CssClass="Panel3" DefaultButton="存檔">
        登帳截止日期<asp:TextBox ID="登帳截止日期" Text="111/06/30" runat="server" Maxlength=0 CssClass="Input1"/>
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="false" CellPadding="1" HorizontalAlign="Center" DataSourceID="SqlDataSource1" DataKeyNames="id" GridLines="None" ShowHeaderWhenEmpty="false" AllowPaging="True" PageSize="6" AllowSorting="True" PagerSettings-PageButtonCount=25 EnableModelValidation="True" AutoGenerateEditButton="false">
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
                <asp:TemplateField HeaderText="項">
                    <ItemTemplate>
                        <asp:Label ID="項" runat="server" Text='<%# Eval("項") %>' Maxlength=0  CssClass="Label 項"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="財產編號">
                    <ItemTemplate>
                        <asp:TextBox ID="財產編號" runat="server" Text='<%# Bind("財產編號") %>' Maxlength=0 TextMode="MultiLine" Enabled="True" CssClass="TextBox 財產編號"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="財產名稱<br>財產別名">
                    <ItemTemplate>
                        <asp:TextBox ID="財產名稱" runat="server" Text='<%# Bind("財產名稱") %>' Maxlength=0 Enabled="True" placeholder="財產名稱" CssClass="TextBox2 財產名稱"/><BR>
                        <asp:TextBox ID="財產別名" runat="server" Text='<%# Bind("財產別名") %>' Maxlength=0 Enabled="True" placeholder="財產別名" TextMode="MultiLine" CssClass="TextBox2 財產別名"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="分群">
                    <ItemTemplate>
                        <asp:DropDownList ID="分群" runat="server" AutoPostBack="True" CssClass="DropDownList"></asp:DropDownList>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="廠牌<BR>型號">
                    <ItemTemplate>
                        <asp:TextBox ID="廠牌" runat="server" Text='<%# Bind("廠牌") %>' Maxlength=0 Enabled="True" placeholder="廠牌" CssClass="TextBox2 廠牌"/><BR>
                        <asp:TextBox ID="型號" runat="server" Text='<%# Bind("型號") %>' Maxlength=0 Enabled="True" placeholder="型號" CssClass="TextBox2 型號"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="購置日期">
                    <ItemTemplate>
                        <asp:TextBox ID="購置日期" runat="server" Text='<%# If(IsDate(Eval("購置日期")), (Year(Eval("購置日期"))-1911).ToString() & Eval("購置日期", "{0:/MM/dd}"), "") %>' Maxlength=0 Enabled="True" CssClass="TextBox 購置日期"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="單位<BR>數量">
                    <ItemTemplate>
                        <asp:TextBox ID="單位" runat="server" Text='<%# Bind("單位") %>' Maxlength=0 Enabled="True" placeholder="單位" CssClass="TextBox2 單位"/><BR>
                        <asp:TextBox ID="數量" runat="server" Text='<%# Bind("數量", "{0:n0}") %>' Maxlength=0 Enabled="True" placeholder="數量" CssClass="TextBox2 數量"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="年限<BR>年數">
                    <ItemTemplate>
                        <asp:TextBox ID="年限" runat="server" Text='<%# Bind("年限") %>' Maxlength=0 Enabled="True" placeholder="年限" CssClass="TextBox2 年限"/><BR>
                        <asp:TextBox ID="年數" runat="server" Text='<%# If(NOT(Eval("年數")Is DBNull.Value),Eval("年數"),If(IsDate(Eval("購置日期")),
                            If((DateDiff("m",Eval("購置日期"),Today())-If(Day(Eval("購置日期"))>Day(Today()),1,0))\12=0,"",(DateDiff("m",Eval("購置日期"),Today())-If(Day(Eval("購置日期"))>Day(Today()),1,0))\12 & "年") & 
                            If((DateDiff("m",Eval("購置日期"),Today())-If(Day(Eval("購置日期"))>Day(Today()),1,0))-(((DateDiff("m",Eval("購置日期"),Today())-If(Day(Eval("購置日期"))>Day(Today()),1,0))\12)*12)=0,If((DateDiff("m",Eval("購置日期"),Today())-If(Day(Eval("購置日期"))>Day(Today()),1,0))\12=0,"不到1月",""),(DateDiff("m",Eval("購置日期"),Today())-If(Day(Eval("購置日期"))>Day(Today()),1,0))-(((DateDiff("m",Eval("購置日期"),Today())-If(Day(Eval("購置日期"))>Day(Today()),1,0))\12)*12) & "月"), ""))%>' 
                            Maxlength=0 Enabled="True" placeholder="年數" CssClass="TextBox2 年數"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="保管人">
                    <ItemTemplate>
                        <asp:TextBox ID="保管人" runat="server" Text='<%# Bind("保管人") %>' Maxlength=0 Enabled="True" CssClass="TextBox 保管人"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="存置地點">
                    <ItemTemplate>
                        <asp:TextBox ID="存置地點" runat="server" Text='<%# Bind("存置地點") %>' Maxlength=0 TextMode="MultiLine" Enabled="True" CssClass="TextBox 存置地點"/>
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
            FROM 設備
            Where (@狀態='' OR (@狀態='過期' AND (REPLACE(年限,'年','')*12)<=DateDiff(m,購置日期,GETDATE()))OR (@狀態='未過期' AND Not((REPLACE(年限,'年','')*12)<=DateDiff(m,購置日期,GETDATE()))))
            ORDER BY _頁, _列"
        Insertcommand="INSERT INTO 設備 (數量) VALUES (0)"
        UpdateCommand=""
        DeleteCommand="">
        <SelectParameters>
            <asp:ControlParameter ControlID="狀態" ConvertEmptyStringToNull="False" Name="狀態" PropertyName="Text" Type="String"/>
        </SelectParameters>
        <InsertParameters>
        </InsertParameters>
        <UpdateParameters>
        </UpdateParameters>
        <DeleteParameters>
        </DeleteParameters>
    </asp:SqlDataSource>
    <asp:SqlDataSource ID="SqlDataSource2" runat="server" ConnectionString='<%$ ConnectionStrings:ApplicationServices%>'
        SelectCommand="SELECT *
            FROM 設備統計分群
            ORDER BY id"
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
