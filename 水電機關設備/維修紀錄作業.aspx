<%@ Page Language="VB" MasterPageFile="./MasterPage.master" MaintainScrollPositionOnPostback="true" AutoEventWireup="false" CodeFile="維修紀錄作業.aspx.vb" Inherits="維修紀錄作業"%>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
    <asp:ScriptManager ID="ScriptManager1" runat="server" EnablePageMethods = "true" EnableScriptGlobalization="True" EnableScriptLocalization="True"/>
    <script src='js/維修紀錄作業.js'></script>
    <link rel="stylesheet" runat="server" media="screen" href="css\維修紀錄作業.css"/>
    <style>
    </style>
    <div><h1><a id="Title" href="維修紀錄作業.aspx">維修紀錄作業<a></h1></div>
    <asp:Panel ID="Panel1" runat="server" DefaultButton="搜尋" CssClass="Panel1">
        <asp:TextBox ID="id" runat="server" Maxlength=3 CssClass="Input2" Visible="false"/>
        <!-- 品項<asp:DropDownList ID="品項" runat="server" DataSourceID="SqlDataSource2" DataTextField="品項" DataValueField="id" Enabled="True" CssClass="DropDownList"/>
        <asp:Button ID="搜尋" runat="server" Text="搜尋" CssClass="GreenButton"/> -->
        <asp:Button ID="新增" runat="server" Text="新增一列" OnClick="Insert" CssClass="GreenButton" Visible="false"/>
        <asp:Button ID="存檔" runat="server" Text="存檔" OnClick="Update" CssClass="GreenButton" Visible="false"/>
        <asp:FileUpload ID="FileUpload1" runat="server" AllowMultiple="True" CssClass="UploadButton" Visible="false"/>
        <asp:Button ID="測試" runat="server" Text="測試" OnClick="test" CssClass="GreenButton" Visible="false"/>
        <asp:Label ID="Label1" runat="server" Text="" CssClass="GreenLabel"/>
        <asp:Label ID="Label2" runat="server" Text="" CssClass="RedLabel"/>
    </asp:Panel>
    <asp:Panel ID="Panel3" runat="server" CssClass="Panel3" DefaultButton="存檔">
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="false" CellPadding="1" HorizontalAlign="Center" DataSourceID="SqlDataSource1" DataKeyNames="id" GridLines="None" ShowHeaderWhenEmpty="false" AllowPaging="True" PageSize="15" AllowSorting="True" PagerSettings-PageButtonCount=25 EnableModelValidation="True" AutoGenerateEditButton="false" OnDataBinding="GridView1_DataBound">
            <Columns>
                <asp:TemplateField HeaderText="id" Visible="false">
                    <ItemTemplate>
                        <asp:TextBox ID="id" runat="server" Text='<%# Bind("id") %>' Maxlength=0 Enabled="False" CssClass="TextBox id"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="編號">
                    <ItemTemplate>
                        <asp:TextBox ID="id_水電" runat="server" Text='<%# Eval("id_水電") %>' Maxlength=0 Enabled="False" CssClass="TextBox 編號"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="設備編號">
                    <ItemTemplate>
                        <asp:TextBox ID="設備編號" runat="server" Text='<%# Eval("設備編號") %>' Maxlength=0 Enabled="False" CssClass="TextBox 設備編號"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="品項">
                    <ItemTemplate>
                        <asp:DropDownList ID="品項" runat="server" Text='<%# Bind("Id_品項") %>' DataSourceID="SqlDataSource2" DataTextField="品項" DataValueField="id" Enabled="False" CssClass="DropDownList"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="維修日期">
                    <ItemTemplate>
                        <asp:TextBox ID="維修日期" runat="server" Text='<%# If(IsDate(Eval("維修日期")), (Year(Eval("維修日期"))-1911).ToString() & Eval("維修日期", "{0:/MM/dd}"), "") %>' Maxlength=0 Enabled="False" CssClass="TextBox 維修日期"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="簽案資料">
                    <ItemTemplate>
                        <asp:HyperLink ID="簽案資料" runat="server" NavigateUrl='<%# "data/簽案資料/" & Eval("簽案資料") %>' Text='<%# Eval("簽案資料") %>' CssClass="簽案資料"></asp:HyperLink>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="" Visible="false">
                    <ItemTemplate>
                        <asp:Button ID="上傳資料" runat="server" Text="上傳資料" CommandName="上傳資料" Visible="false" CssClass="GreenButton"/><br>
                        <asp:Button ID="上傳照片" runat="server" Text="上傳照片" CommandName="上傳照片" Visible="false" OnClientClick="return confirm('確定上傳圖片?')" CssClass="GreenButton"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Left" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="金額">
                    <ItemTemplate>
                        <asp:TextBox ID="金額" runat="server" Text='<%# Bind("金額", "{0:c0}") %>' Maxlength=0 Enabled="False" CssClass="TextBox 金額"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="照片">
                    <ItemTemplate>
                        <asp:ImageButton ID="照片" runat="server" ImageUrl='<%# If(Eval("照片").ToString = "", "", Eval("照片")) %>' CommandName="照片圖" AutoPostBack="False"  OnClientClick= "" Visible='<%# If(Eval("照片").ToString = "", "False", "True") %>' CssClass="label 照片"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="備註">
                    <ItemTemplate>
                        <asp:TextBox ID="備註" runat="server" Text='<%# Bind("備註") %>' Maxlength=0 Enabled="False" CssClass="TextBox 備註"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                    <asp:TemplateField HeaderText="" Visible="false">
                    <ItemTemplate>
                        <asp:Button ID="刪除" runat="server" Text="刪除" CommandName="刪除" OnClientClick="return confirm('確定刪除?')" Visible="false" CssClass="RedButton"/>
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
            SELECT 水電機關設備資料表.*, 維修紀錄表.* FROM 水電機關設備資料表, 維修紀錄表
            WHERE (''=TRIM(@id) OR 水電機關設備資料表.id LIKE TRIM(@id))
            and (''=TRIM(@品項) OR Id_品項 LIKE N'%'+TRIM(@品項)+'%')
            and 水電機關設備資料表.id=維修紀錄表.id_水電
            ORDER BY _頁, _列"
        Insertcommand=""
        UpdateCommand=""
        DeleteCommand="DELETE FROM 水電機關設備資料表 WHERE id=@id">
        <SelectParameters>
            <asp:ControlParameter ControlID="id" ConvertEmptyStringToNull="False" Name="id" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="品項" ConvertEmptyStringToNull="False" Name="品項" PropertyName="Text" Type="String"/>
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
