<%@ Page Language="VB" MasterPageFile="./MasterPage.master" MaintainScrollPositionOnPostback="true" AutoEventWireup="false" CodeFile="零星維修作業.aspx.vb" Inherits="零星維修作業"%>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
    <asp:ScriptManager ID="ScriptManager1" runat="server" EnablePageMethods = "true" EnableScriptGlobalization="True" EnableScriptLocalization="True"/>
    <script src='js/零星維修作業.js'></script>
    <link rel="stylesheet" runat="server" media="screen" href="css\零星維修作業.css"/>
    <style>
    </style>
    <div><h1><a id="Title" href="零星維修作業.aspx">零星維修作業<a></h1></div>
    <asp:Panel ID="Panel1" runat="server" DefaultButton="搜尋" CssClass="Panel1">
        廠商<asp:TextBox ID="廠商" runat="server" Maxlength=3 CssClass="Input1"/>
        <ajaxToolkit:AutoCompleteExtEnder ID="廠商自動" runat="server" TargetControlID="廠商"
                            CompletionSetCount="10000" CompletionInterval="50" EnableCaching="true" MinimumPrefixLength="1"
                            ServiceMethod="GetMyList"
                            CompletionListCssClass="CompletionList"
                            CompletionListItemCssClass="CompletionListItem"
                            CompletionListHighlightedItemCssClass="CompletionListHighlightedItem"/>
        維修日期:<asp:TextBox ID="年" runat="server" Maxlength=3 CssClass="Input2"/>年
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
        <asp:ListItem Text="13" Value="13"></asp:ListItem>
        <asp:ListItem Text="14" Value="14"></asp:ListItem>
        <asp:ListItem Text="15" Value="15"></asp:ListItem>
        <asp:ListItem Text="16" Value="16"></asp:ListItem>
        <asp:ListItem Text="17" Value="17"></asp:ListItem>
        <asp:ListItem Text="18" Value="18"></asp:ListItem>
        <asp:ListItem Text="19" Value="19"></asp:ListItem>
        <asp:ListItem Text="20" Value="20"></asp:ListItem>
        <asp:ListItem Text="21" Value="21"></asp:ListItem>
        <asp:ListItem Text="22" Value="22"></asp:ListItem>
        <asp:ListItem Text="23" Value="23"></asp:ListItem>
        <asp:ListItem Text="24" Value="24"></asp:ListItem>
        <asp:ListItem Text="25" Value="25"></asp:ListItem>
        <asp:ListItem Text="26" Value="26"></asp:ListItem>
        <asp:ListItem Text="27" Value="27"></asp:ListItem>
        <asp:ListItem Text="28" Value="28"></asp:ListItem>
        <asp:ListItem Text="29" Value="29"></asp:ListItem>
        <asp:ListItem Text="30" Value="30"></asp:ListItem>
        <asp:ListItem Text="31" Value="31"></asp:ListItem>
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
        <asp:ListItem Text="13" Value="13"></asp:ListItem>
        <asp:ListItem Text="14" Value="14"></asp:ListItem>
        <asp:ListItem Text="15" Value="15"></asp:ListItem>
        <asp:ListItem Text="16" Value="16"></asp:ListItem>
        <asp:ListItem Text="17" Value="17"></asp:ListItem>
        <asp:ListItem Text="18" Value="18"></asp:ListItem>
        <asp:ListItem Text="19" Value="19"></asp:ListItem>
        <asp:ListItem Text="20" Value="20"></asp:ListItem>
        <asp:ListItem Text="21" Value="21"></asp:ListItem>
        <asp:ListItem Text="22" Value="22"></asp:ListItem>
        <asp:ListItem Text="23" Value="23"></asp:ListItem>
        <asp:ListItem Text="24" Value="24"></asp:ListItem>
        <asp:ListItem Text="25" Value="25"></asp:ListItem>
        <asp:ListItem Text="26" Value="26"></asp:ListItem>
        <asp:ListItem Text="27" Value="27"></asp:ListItem>
        <asp:ListItem Text="28" Value="28"></asp:ListItem>
        <asp:ListItem Text="29" Value="29"></asp:ListItem>
        <asp:ListItem Text="30" Value="30"></asp:ListItem>
        <asp:ListItem Text="31" Value="31"></asp:ListItem>
    </asp:DropDownList>日
        <asp:Button ID="清空" runat="server" Text="清空"  OnClick="Clear_Click"  CssClass="GreenButton"/>
        <asp:Button ID="搜尋" runat="server" Text="搜尋" CssClass="GreenButton"/>
        <asp:Button ID="新增" runat="server" Text="新增一頁" OnClick="Insert" CssClass="GreenButton"/>
        <asp:Button ID="存檔" runat="server" Text="存檔" OnClick="Update" CssClass="GreenButton"/>
        <asp:Button ID="刪除" runat="server" Text="刪除末頁" OnClick="Delete" OnClientClick="return confirm('確定刪除?')" CssClass="RedButton"/>
        <asp:FileUpload ID="FileUpload1" runat="server" AllowMultiple="True" CssClass="UploadButton" />
        <div style="display:inline-block;width:275px;text-align:left;">
            <asp:GridView ID="GridView3" runat="server" HorizontalAlign="Center" DataSourceID="SqlDataSource2" GridLines="None" ShowHeader="false" CssClass="GridView3"/>
        </div>
        <BR>
        <asp:Button ID="測試" runat="server" Text="測試" OnClick="test" CssClass="GreenButton" Visible="false"/>
        <asp:Label ID="Label1" runat="server" Text="" CssClass="GreenLabel"/>
        <asp:Label ID="Label2" runat="server" Text="" CssClass="RedLabel"/>
    </asp:Panel>
    <asp:Panel ID="Panel3" runat="server" CssClass="Panel3" DefaultButton="存檔">
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="false" CellPadding="1" HorizontalAlign="Center" DataSourceID="SqlDataSource1" DataKeyNames="id" GridLines="None" ShowHeaderWhenEmpty="false" AllowPaging="True" PageSize="15" AllowSorting="True" PagerSettings-PageButtonCount=25 EnableModelValidation="True" AutoGenerateEditButton="false">
            <Columns>
            <asp:TemplateField HeaderText="" Visible="false">
                    <ItemTemplate>
                        <asp:TextBox ID="id" runat="server" Text='<%# Bind("id") %>' Maxlength=0 Enabled="False" CssClass="TextBox id"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="編號">
                    <ItemTemplate>
                        <asp:TextBox ID="_列" runat="server" Text='<%# Bind("_列") %>' Maxlength=0 Enabled="True" CssClass="TextBox _列"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="叫修日期">
                    <ItemTemplate>
                        <asp:TextBox ID="叫修日期" runat="server" Text='<%# If(IsDate(Eval("叫修日期")), (Year(Eval("叫修日期"))-1911).ToString() & Eval("叫修日期", "{0:/MM/dd}"), "") %>' Maxlength=0 Enabled="True" CssClass="TextBox 叫修日期"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="維修日期">
                    <ItemTemplate>
                        <asp:TextBox ID="維修日期" runat="server" Text='<%# If(IsDate(Eval("維修日期")), (Year(Eval("維修日期"))-1911).ToString() & Eval("維修日期", "{0:/MM/dd}"), "") %>' Maxlength=0 Enabled="True" CssClass="TextBox 維修日期"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="維修金額">
                    <ItemTemplate>
                        <asp:TextBox ID="維修金額" runat="server" Text='<%# Bind("維修金額", "{0:c0}") %>' Maxlength=0 Enabled="True" CssClass="TextBox 維修金額"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="維修內容">
                    <ItemTemplate>
                        <asp:TextBox ID="維修內容" runat="server" Text='<%# Bind("維修內容") %>' Maxlength=0 Enabled="True" CssClass="TextBox 維修內容"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="ID_廠商" Visible="False">
                    <ItemTemplate>
                        <asp:TextBox ID="ID_廠商" runat="server" Text='<%# Bind("ID_廠商") %>' Maxlength=0 Enabled="True" CssClass="TextBox ID_廠商"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="維護廠商">
                    <ItemTemplate>
                        <asp:TextBox ID="維護廠商" runat="server" Text='<%# Bind("廠商") %>' Maxlength=0 Enabled="True" CssClass="TextBox 維護廠商"/>
                        <ajaxToolkit:AutoCompleteExtEnder ID="維護廠商自動" runat="server" TargetControlID="維護廠商"
                            CompletionSetCount="10000" CompletionInterval="50" EnableCaching="true" MinimumPrefixLength="1"
                            ServiceMethod="GetMyList"
                            CompletionListCssClass="CompletionList"
                            CompletionListItemCssClass="CompletionListItem"
                            CompletionListHighlightedItemCssClass="CompletionListHighlightedItem"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="廠商電話">
                    <ItemTemplate>
                        <asp:TextBox ID="廠商電話" runat="server" Text='<%# Bind("電話") %>' Maxlength=0 Enabled="True" CssClass="TextBox 廠商電話"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="上傳維護紀錄">
                    <ItemTemplate>
                        <asp:Button ID="上傳資料" runat="server" Text="上傳資料" CommandName="上傳資料" Enabled="True" CssClass="GreenButton"/><br>
                        <asp:HyperLink ID="維修紀錄" runat="server" NavigateUrl='<%# If(Eval("維修紀錄").ToString = "", "", "data/零星維修紀錄/" & Eval("維修紀錄")) %>' Text='<%# Eval("維修紀錄") %> ' Visible='<%# If(Eval("維修紀錄").ToString = "", "False", "True") %>' CssClass="檔名"></asp:HyperLink>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="備註">
                    <ItemTemplate>
                        <asp:TextBox ID="備註" runat="server" Text='<%# Bind("備註") %>' Maxlength=0 Enabled="True" CssClass="TextBox 備註"/>
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
            SELECT * FROM 零星維修作業 left Join 廠商資料 on 零星維修作業.ID_廠商 =廠商資料.ID left Join 零星維修作業_維修紀錄 on 零星維修作業.ID_維修紀錄 =零星維修作業_維修紀錄.id_零星維修
            WHERE (''=TRIM(@廠商) OR 廠商資料.廠商 LIKE N'%'+TRIM(@廠商)+'%')
            AND ((''=TRIM(@年) OR ''=TRIM(@月1) OR ''=TRIM(@日1) OR ''=TRIM(@月2) OR ''=TRIM(@日2))
                OR 維修日期 BETWEEN
                    TRY_PARSE(CONVERT(varchar(10),IIF(TRIM(@年)='','2022',TRIM(@年)+1911))+'/'+IIF(TRIM(@月1)='','1',TRIM(@月1))+'/'+IIF(TRIM(@日1)='','1',TRIM(@日1)) AS date)
                    AND
                    TRY_PARSE(CONVERT(varchar(10),IIF(TRIM(@年)='','2022',TRIM(@年)+1911))+'/'+IIF(TRIM(@月2)='','12',TRIM(@月2))+'/'+IIF(TRIM(@日2)='','31',TRIM(@日2)) AS date))
            ORDER BY 零星維修作業._頁,零星維修作業._列"
        Insertcommand=""
        UpdateCommand=""
        DeleteCommand="">
        <SelectParameters>
            <asp:ControlParameter ControlID="廠商" ConvertEmptyStringToNull="False" Name="廠商" PropertyName="Text" Type="String"/>
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
        SelectCommand="SELECT '維護總金額: ' + REPLACE(CONVERT(varchar(30), CONVERT(money, SUM(維修金額)), 1), '.00', '元')
                    FROM 零星維修作業 left Join 廠商資料 on 零星維修作業.ID_廠商 =廠商資料.ID
                    WHERE ((''=TRIM(@年) OR ''=TRIM(@月1) OR ''=TRIM(@日1) OR ''=TRIM(@月2) OR ''=TRIM(@日2))
                    OR 維修日期 BETWEEN
                    TRY_PARSE(CONVERT(varchar(10),IIF(TRIM(@年)='','2022',TRIM(@年)+1911))+'/'+IIF(TRIM(@月1)='','1',TRIM(@月1))+'/'+IIF(TRIM(@日1)='','1',TRIM(@日1)) AS date)
                    AND
                    TRY_PARSE(CONVERT(varchar(10),IIF(TRIM(@年)='','2022',TRIM(@年)+1911))+'/'+IIF(TRIM(@月2)='','12',TRIM(@月2))+'/'+IIF(TRIM(@日2)='','31',TRIM(@日2)) AS date))
                    AND ((''=TRIM(@廠商)) OR 廠商 LIKE N'%'+TRIM(@廠商)+'%') "
        Insertcommand=""
        UpdateCommand=""
        DeleteCommand="">
        <SelectParameters>
            <asp:ControlParameter ControlID="廠商" ConvertEmptyStringToNull="False" Name="廠商" PropertyName="Text" Type="String"/>
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
    <asp:SqlDataSource ID="SqlDataSource3" runat="server" ConnectionString='<%$ ConnectionStrings:ApplicationServices%>'        SelectCommand="SELECT 維修紀錄 FROM 零星維修作業_維修紀錄 where"
        Insertcommand=""
        UpdateCommand=""
        DeleteCommand="">
        <SelectParameters>
            <asp:ControlParameter ControlID="廠商" ConvertEmptyStringToNull="False" Name="廠商" PropertyName="Text" Type="String"/>
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
</asp:Content>
