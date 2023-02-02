<%@ Page Language="VB" MasterPageFile="./MasterPage.master" MaintainScrollPositionOnPostback="true" AutoEventWireup="false" CodeFile="其他作業.aspx.vb" Inherits="其他作業"%>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server" >
    <asp:ScriptManager ID="ScriptManager1" runat="server" EnablePageMethods = "true" EnableScriptGlobalization="True" EnableScriptLocalization="True"/>
    <script src='js/其他作業.js'></script>
    <link rel="stylesheet" runat="server" media="screen" href="css\其他作業.css"/>
    <div><h1><a id="Title" href="其他作業.aspx">其他作業<a></h1></div>
    <asp:Panel ID="Panel1" runat="server" CssClass="Panel1" >
        年<asp:TextBox ID="年" runat="server" Maxlength=3 CssClass="Input2"/>
        <asp:FileUpload ID="FileUpload1" runat="server" AllowMultiple="True" CssClass="UploadButton"/>
        <asp:Button ID="上傳" runat="server" Text="上傳" OnClick="Import" OnClientClick="return confirm('確定上傳?')" CssClass="GreenButton"/>
        <asp:Button ID="測試" runat="server" Text="測試" OnClick="test" CssClass="GreenButton" Visible="false"/>
        <asp:Label ID="Label1" runat="server" Text="" CssClass="GreenLabel"/>
        <asp:Label ID="Label2" runat="server" Text="" CssClass="RedLabel"/>
    </asp:Panel>
    <asp:Panel ID="Panel3" runat="server" CssClass="Panel3" >
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="false" CellPadding="1" HorizontalAlign="Center" DataSourceID="SqlDataSource1" DataKeyNames="id" GridLines="None" ShowHeaderWhenEmpty="false" AllowPaging="True" PageSize="15" AllowSorting="True" PagerSettings-PageButtonCount=25 EnableModelValidation="True" AutoGenerateEditButton="false">
            <Columns>
                <asp:TemplateField HeaderText="id" Visible="false">
                    <ItemTemplate>
                        <asp:Label ID="id" runat="server" Text='<%# Eval("id") %>' Maxlength=0 Enabled="False" CssClass="Label id"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="檔名">
                    <ItemTemplate>
                        <asp:HyperLink ID="檔名" runat="server" NavigateUrl='<%# "./data/" & Eval("檔名") %>' Text='<%# Eval("檔名") %>' CssClass="檔名"></asp:HyperLink>
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
            SELECT * FROM 上傳" 
        Insertcommand="" 
        UpdateCommand=""
        DeleteCommand="DELETE FROM 上傳 WHERE id=@id">
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
</asp:Content>
