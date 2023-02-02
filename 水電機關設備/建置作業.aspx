<%@ Page Language="VB" MasterPageFile="./MasterPage.master" MaintainScrollPositionOnPostback="true" AutoEventWireup="false" CodeFile="建置作業.aspx.vb" Inherits="建置作業"%>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server" >
    <asp:ScriptManager ID="ScriptManager1" runat="server" EnablePageMethods = "true" EnableScriptGlobalization="True" EnableScriptLocalization="True"/>
    <script src='js/建置作業.js'></script>
    <link rel="stylesheet" runat="server" media="screen" href="css\建置作業.css"/>
    <style>
    </style>
    <div><h1><a id="Title" href="建置作業.aspx">建置作業<a></h1></div>
    <asp:Panel ID="Panel1" runat="server" DefaultButton="搜尋" CssClass="Panel1" >
        <asp:Button ID="搜尋" runat="server" Text="搜尋" CssClass="GreenButton" Visible="False"/>
        <asp:Button ID="新增" runat="server" Text="新增一頁" OnClick="Insert" CssClass="GreenButton"/>
        <asp:Button ID="存檔" runat="server" Text="存檔" OnClick="Update" CssClass="GreenButton"/>
        <asp:Button ID="刪除" runat="server" Text="刪除末頁" OnClick="Delete" OnClientClick="return confirm('確定刪除?')" CssClass="RedButton"/>
        <asp:Button ID="測試" runat="server" Text="轉換" CssClass="GreenButton" OnClick="Test" Visible="False"/>
        <asp:FileUpload ID="FileUpload1" runat="server" AllowMultiple="True" CssClass="UploadButton"/>
        <asp:Image ID="Image1" runat="server" Text="" CssClass="Image1"/>
        <asp:Label ID="Label1" runat="server" Text="" CssClass="GreenLabel"/>
        <asp:Label ID="Label2" runat="server" Text="" CssClass="RedLabel"/>
    </asp:Panel>
    <asp:Panel ID="Panel3" runat="server" CssClass="Panel3" DefaultButton="存檔">
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="false" CellPadding="1" HorizontalAlign="Center" DataSourceID="SqlDataSource1" DataKeyNames="id" GridLines="None" ShowHeaderWhenEmpty="false" AllowPaging="True" PageSize="15" AllowSorting="True" PagerSettings-PageButtonCount=25 EnableModelValidation="True" AutoGenerateEditButton="false">
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
                        <asp:TextBox ID="_列" runat="server" Text='<%# Bind("_列") %>' Maxlength=0 Enabled="False" CssClass="TextBox _列"/>
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
                        <asp:DropDownList ID="品項" runat="server" Text='<%# Bind("Id_品項") %>' DataSourceID="SqlDataSource2" DataTextField="品項" DataValueField="id" CssClass="DropDownList"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="建置日期">
                    <ItemTemplate>
                        <asp:TextBox ID="建置日期" runat="server" Text='<%# If(IsDate(Eval("建置日期")), (Year(Eval("建置日期"))-1911).ToString() & Eval("建置日期", "{0:/MM/dd}"), "") %>' Maxlength=0 Enabled="True" CssClass="TextBox 建置日期"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="型號">
                    <ItemTemplate>
                        <asp:TextBox ID="型號" runat="server" Text='<%# Bind("型號") %>' Maxlength=0 TextMode="MultiLine" Enabled="True" CssClass="TextBox 型號"/>
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
                <asp:TemplateField HeaderText="維護單位">
                    <ItemTemplate>
                        <asp:TextBox ID="維護單位" runat="server" Text='<%# Bind("維護單位") %>' Maxlength=0 Enabled="True" CssClass="TextBox 維護單位"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="照片">
                    <ItemTemplate>
                        <asp:ImageButton ID="照片" runat="server" ImageUrl='<%# If(Eval("照片縮圖").ToString = "", "", Eval("照片縮圖")) %>' CommandName="照片圖" AutoPostBack="False"  OnClientClick= "" Visible='<%# If(Eval("照片縮圖").ToString = "",False,True) %>' CssClass="照片"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="">
                    <ItemTemplate>
                        <asp:Button ID="上傳資料" runat="server" Text="上傳資料" CommandName="上傳資料" Enabled="True" CssClass="GreenButton"/>
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
            SELECT id,_列,_頁,設備編號,Id_品項,建置日期,型號,存置地點,維護單位,照片縮圖,維護紀錄,備註
            FROM 水電機關設備資料表
            ORDER BY _頁, _列"
        Insertcommand=""
        UpdateCommand=""
        DeleteCommand="DELETE FROM 水電機關設備資料表 WHERE id=@id">
        <SelectParameters>
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
</asp:Content>
