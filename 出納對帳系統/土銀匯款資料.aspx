<%@ Page Language="VB" MasterPageFile="./MasterPage.master" MaintainScrollPositionOnPostback="true" AutoEventWireup="False" CodeFile="土銀匯款資料.aspx.vb" Inherits="土銀匯款資料" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
    <link rel="stylesheet" runat="server" media="screen" href="css\土銀匯款資料.css"/>
    <div><h1><a id="Title" href="土銀匯款資料.aspx">土銀匯款資料<a></h1></div>
    <asp:Panel ID="Panel1" runat="server" DefaultButton="Button1" CssClass="Panel1">
        序號<asp:TextBox ID="TextBox1" runat="server" CssClass="Input"/>
        收款人代碼<asp:TextBox ID="TextBox2" runat="server" CssClass="Input"/>
        收款人名稱<asp:TextBox ID="TextBox3" runat="server" CssClass="Input"/>
        匯入銀行代碼<asp:TextBox ID="TextBox4" runat="server" CssClass="Input"/>
        匯入帳號<asp:TextBox ID="TextBox5" runat="server" CssClass="Input"/>
        收款人匯款戶名<asp:TextBox ID="TextBox6" runat="server" CssClass="Input"/>
        收款人統編<asp:TextBox ID="TextBox7" runat="server" CssClass="Input"/>
        <asp:Button ID="Button1" runat="server" Text="搜尋" CssClass="GreenButton"/>
        <asp:Button ID="Button2" runat="server" Text="新增" OnClick="Insert" CssClass="GreenButton"/>
        <asp:Label ID="Label1" runat="server" Text="" CssClass="GreenLabel"/>
        <asp:Label ID="Label2" runat="server" Text="" CssClass="RedLabel"/>
    </asp:Panel>
    <asp:Panel ID="Panel2" runat="server" DefaultButton="Button1" CssClass="Panel2">
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" CellPadding="1" HorizontalAlign="Center" DataSourceID="SqlDataSource1" DataKeyNames="id" GridLines="None" ShowHeaderWhenEmpty="False" AllowPaging="True" PageSize="20" AllowSorting="True" PagerSettings-PageButtonCount=25 EnableModelValidation="True" AutoGenerateEditButton="False" OnRowUpdated="GridView1_RowUpdated">
            <Columns>
                <asp:TemplateField Visible="False">
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
                        <asp:Label ID="Label3" runat="server" Text='<%# Bind("序號") %>' Maxlength=4 CssClass="Label Label3"/>
                    </ItemTemplate>
                    <EditItemTemplate>
                        <asp:TextBox ID="TextBox1" runat="server" Text='<%# Bind("序號") %>' Maxlength=4 CssClass="TextBox TextBox1"/>
                    </EditItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="收款人代碼">
                    <ItemTemplate>
                        <asp:Label ID="Label4" runat="server" Text='<%# Bind("收款人代碼") %>' Maxlength=4 CssClass="Label Label4"/>
                    </ItemTemplate>
                    <EditItemTemplate>
                        <asp:TextBox ID="TextBox2" runat="server" Text='<%# Bind("收款人代碼") %>' Maxlength=4 CssClass="TextBox TextBox2"/>
                    </EditItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="收款人名稱">
                    <ItemTemplate>
                        <asp:Label ID="Label5" runat="server" Text='<%# Bind("收款人名稱") %>' Maxlength=300 CssClass="Label Label5"/>
                    </ItemTemplate>
                    <EditItemTemplate>
                        <asp:TextBox ID="TextBox3" runat="server" Text='<%# Bind("收款人名稱") %>' Maxlength=300 CssClass="TextBox TextBox3"/>
                    </EditItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="銀行代碼">
                    <ItemTemplate>
                        <asp:Label ID="Label6" runat="server" Text='<%# Bind("匯入銀行代碼") %>' Maxlength=7 CssClass="Label Label6"/>
                    </ItemTemplate>
                    <EditItemTemplate>
                        <asp:TextBox ID="TextBox4" runat="server" Text='<%# Bind("匯入銀行代碼") %>' Maxlength=7 CssClass="TextBox TextBox4"/>
                    </EditItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="匯入帳號">
                    <ItemTemplate>
                        <asp:Label ID="Label7" runat="server" Text='<%# Bind("匯入帳號") %>' Maxlength=16 CssClass="Label Label7"/>
                    </ItemTemplate>
                    <EditItemTemplate>
                        <asp:TextBox ID="TextBox5" runat="server" Text='<%# Bind("匯入帳號") %>' Maxlength=16 CssClass="TextBox TextBox5"/>
                    </EditItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="收款人匯款戶名">
                    <ItemTemplate>
                        <asp:Label ID="Label8" runat="server" Text='<%# Bind("收款人匯款戶名") %>' Maxlength=300 CssClass="Label Label8"/>
                    </ItemTemplate>
                    <EditItemTemplate>
                        <asp:TextBox ID="TextBox6" runat="server" Text='<%# Bind("收款人匯款戶名") %>' Maxlength=300 CssClass="TextBox TextBox6"/>
                    </EditItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="收款人統編">
                    <ItemTemplate>
                        <asp:Label ID="Label9" runat="server" Text='<%# Bind("收款人統編") %>' Maxlength=10 CssClass="Label Label9"/>
                    </ItemTemplate>
                    <EditItemTemplate>
                        <asp:TextBox ID="TextBox7" runat="server" Text='<%# Bind("收款人統編") %>' Maxlength=10 CssClass="TextBox TextBox7"/>
                    </EditItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="收款人EMAIL">
                    <ItemTemplate>
                        <asp:Label ID="Label10" runat="server" Text='<%# Bind("收款人EMAIL") %>' Maxlength=300 CssClass="Label Label10"/>
                    </ItemTemplate>
                    <EditItemTemplate>
                        <asp:TextBox ID="TextBox8" runat="server" Text='<%# Bind("收款人EMAIL") %>' Maxlength=300 CssClass="TextBox TextBox8"/>
                    </EditItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="">
                    <ItemTemplate>
                        <asp:Button ID="Button1" runat="server" Text="編輯" CommandName="Edit" CssClass="GreenButton"/>
                    </ItemTemplate>
                    <EditItemTemplate>
                        <asp:Button ID="Button2" runat="server" text="存檔" commandname="Update" CssClass="GreenButton"/>
                        <asp:Button ID="Button3" runat="server" text="取消" commandname="Cancel" CssClass="GreenButton"/>
                        <asp:Button ID="Button4" runat="server" Text="刪除" CommandName="Delete" OnClientClick="return confirm('確定刪除?')" CssClass="RedButton"/>
                    </EditItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Left" CssClass="Item"/>
                </asp:TemplateField>
            </Columns>
            <HeaderStyle BackColor="Green" Font-Bold="True" ForeColor="White" CssClass="Header"/>
            <RowStyle BackColor="#FFFFFF" CssClass="Row"/>
            <AlternatingRowStyle/>
            <SelectedRowStyle/>
            <EditRowStyle CssClass="EditRow"/>
            <PagerStyle BackColor="Green" HorizontalAlign="Center" CssClass="Pager"/>
            <FooterStyle/>
            <PagerSettings  Mode="NumericFirstLast" FirstPageText="<<" PreviousPageText="<" NextPageText=">" LastPageText=">>" />
        </asp:GridView>
    </asp:Panel>
    <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString='<%$ ConnectionStrings:ApplicationServices%>'
        SelectCommand="SELECT * 
        FROM 收款人 
        WHERE (''=TRIM(@_序號) OR 序號 LIKE TRIM(@_序號)) 
        AND (''=TRIM(@_收款人代碼) OR 收款人代碼 LIKE '%'+TRIM(@_收款人代碼)) 
        AND (''=TRIM(@_收款人名稱) OR 收款人名稱 LIKE N'%'+TRIM(@_收款人名稱)+'%') 
        AND (''=TRIM(@_匯入銀行代碼) OR 匯入銀行代碼 LIKE '%'+TRIM(@_匯入銀行代碼)+'%') 
        AND (''=TRIM(@_匯入帳號) OR 匯入帳號 LIKE '%'+TRIM(@_匯入帳號)+'%') 
        AND (''=TRIM(@_收款人匯款戶名) OR 收款人匯款戶名 LIKE N'%'+TRIM(@_收款人匯款戶名)+'%') 
        AND (''=TRIM(@_收款人統編) OR 收款人統編 LIKE '%'+TRIM(@_收款人統編)+'%') 
        ORDER BY CASE WHEN 序號 IS NULL THEN 1 ELSE 0 END, 序號, 收款人代碼" 
        Insertcommand="INSERT INTO 收款人 (序號) VALUES(NULL)" 
        DeleteCommand="DELETE FROM 收款人 WHERE id=@id" 
        UpdateCommand="UPDATE 收款人 SET 序號=@序號, 收款人代碼=@收款人代碼, 收款人名稱=@收款人名稱, 匯入銀行代碼=@匯入銀行代碼, 匯入帳號=@匯入帳號, 收款人匯款戶名=@收款人匯款戶名, 收款人統編=@收款人統編, 收款人EMAIL=@收款人EMAIL WHERE id=@id">
        <SelectParameters>
            <asp:ControlParameter ControlID="TextBox1" ConvertEmptyStringToNull="False" Name="_序號" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="TextBox2" ConvertEmptyStringToNull="False" Name="_收款人代碼" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="TextBox3" ConvertEmptyStringToNull="False" Name="_收款人名稱" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="TextBox4" ConvertEmptyStringToNull="False" Name="_匯入銀行代碼" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="TextBox5" ConvertEmptyStringToNull="False" Name="_匯入帳號" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="TextBox6" ConvertEmptyStringToNull="False" Name="_收款人匯款戶名" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="TextBox7" ConvertEmptyStringToNull="False" Name="_收款人統編" PropertyName="Text" Type="String"/>
        </SelectParameters>
        <DeleteParameters>
            <asp:Parameter Name="id" Type="String"/>
        </DeleteParameters>
        <UpdateParameters>
            <asp:Parameter Name="id" Type="String"/>
            <asp:Parameter Name="序號" Type="String"/>
            <asp:Parameter Name="收款人代碼" Type="String"/>
            <asp:Parameter Name="收款人名稱" Type="String"/>
            <asp:Parameter Name="匯入銀行代碼" Type="String"/>
            <asp:Parameter Name="匯入帳號" Type="String"/>
            <asp:Parameter Name="收款人匯款戶名" Type="String"/>
            <asp:Parameter Name="收款人統編" Type="String"/>
            <asp:Parameter Name="收款人EMAIL" Type="String"/>
        </UpdateParameters>
    </asp:SqlDataSource>
</asp:Content>


