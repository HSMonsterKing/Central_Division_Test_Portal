<%@ Page Title="郵寄種類" Language="VB" MasterPageFile="./MasterPage.master" AutoEventWireup="false" CodeFile="郵寄種類.aspx.vb" Inherits="郵寄種類" %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" Runat="Server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
            <table style="width: 100%;font-size:Large;">
                <tr>
                    <td style="vertical-align: top; width: 100%; font-family: 標楷體; height: 25px; text-align: left">
                        <asp:GridView ID="GridView1" runat="server" style="font-size:Large;" AutoGenerateColumns="False" CellPadding="4"
                            DataSourceID="SqlDataSource1" ForeColor="#333333" GridLines="None" EnableModelValidation="True">
                            <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                            <Columns>
                                <asp:TemplateField HeaderText="排序" SortExpression="排序">
                                    <EditItemTemplate>
                                        <asp:TextBox ID="TextBox1" runat="server" style="font-size:Large;" Text='<%# Bind("排序") %>'></asp:TextBox>
                                    </EditItemTemplate>
                                    <ItemTemplate>
                                        <asp:TextBox ID="TextBox4" runat="server" style="font-size:Large;" Text='<%# Bind("排序") %>' Width="67px" Font-Names="標楷體"></asp:TextBox>
                                        <asp:Label ID="Label1" runat="server" style="font-size:Large;" Text='<%# Eval("id") %>' Visible="False"></asp:Label>
                                    </ItemTemplate>
                                </asp:TemplateField>
                                <asp:TemplateField HeaderText="郵寄種類" SortExpression="郵寄種類">
                                    <EditItemTemplate>
                                        <asp:TextBox ID="TextBox2" runat="server" style="font-size:Large;" Text='<%# Bind("郵寄種類") %>'></asp:TextBox>
                                    </EditItemTemplate>
                                    <ItemTemplate>
                                        <asp:TextBox ID="TextBox5" runat="server" style="font-size:Large;" Text='<%# Bind("郵寄種類") %>' Font-Names="標楷體"></asp:TextBox>
                                    </ItemTemplate>
                                    <HeaderStyle HorizontalAlign="Center" />
                                    <ItemStyle HorizontalAlign="Left" Width="100px" />
                                </asp:TemplateField>
                                <asp:CommandField SelectText="存檔" ShowSelectButton="True" />
                            </Columns>
                            <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                            <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                            <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                            <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                            <EditRowStyle BackColor="#999999" />
                            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                        </asp:GridView>
                    </td>
                </tr>
                <tr>
                    <td style="vertical-align: top; width: 100%; font-family: 標楷體; height: 25px; text-align: left">
                        <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="Data Source=10.52.0.178;Initial Catalog=大宗郵件;User ID=qaz;Password=1qaz@WSX"
                            SelectCommand="SELECT id, 排序, 郵寄種類, 掛號類別 &#13;&#10;FROM 大宗郵件執據_郵寄種類 &#13;&#10;ORDER BY 排序">
                        </asp:SqlDataSource>
                        &nbsp;
                    </td>
                </tr>
            </table>
        </asp:Content>
