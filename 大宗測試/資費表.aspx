<%@ Page Title="資費表" Language="VB" MasterPageFile="./MasterPage.master" AutoEventWireup="false" CodeFile="資費表.aspx.vb" Inherits="資費表" %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" Runat="Server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
            <table style="width: 100%;font-size:Large;">
                <tr>
                    <td style="vertical-align: top; width: 100%; font-family: 標楷體; text-align: left">
                        <asp:GridView ID="GridView1" runat="server" style="font-size:Large;" AutoGenerateColumns="False" DataSourceID="SqlDataSource1" CellPadding="4" ForeColor="#333333" GridLines="None">
                            <Columns>
                                <asp:BoundField DataField="郵寄種類" HeaderText="郵寄種類" SortExpression="郵寄種類" >
                                    <HeaderStyle HorizontalAlign="Center" />
                                    <ItemStyle HorizontalAlign="Left" Width="100px" />
                                </asp:BoundField>
                                <asp:BoundField DataField="重量" HeaderText="重量" SortExpression="重量" >
                                    <HeaderStyle HorizontalAlign="Center" />
                                    <ItemStyle HorizontalAlign="Right" Width="100px" />
                                </asp:BoundField>
                                <asp:TemplateField HeaderText="郵資">
                                    <ItemTemplate>
                                        <asp:Label ID="Label1" runat="server" style="font-size:Large;" Text='<%# Eval("id") %>' Visible="False"></asp:Label>
                                        <asp:TextBox ID="TextBox1" runat="server" style="font-size:Large;" Text='<%# Bind("郵資") %>' Width="70px" Font-Names="標楷體"></asp:TextBox>
                                    </ItemTemplate>
                                    <HeaderStyle HorizontalAlign="Center" />
                                    <ItemStyle HorizontalAlign="Right" Width="100px" />
                                </asp:TemplateField>
                                <asp:CommandField SelectText="存檔" ShowSelectButton="True" />
                            </Columns>
                            <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
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
                    <td style="vertical-align: top; width: 100%; font-family: 標楷體; text-align: left">
                        <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="Data Source=10.52.0.178;Initial Catalog=大宗郵件;User ID=qaz;Password=1qaz@WSX"
                            SelectCommand="SELECT id, 郵寄種類, 郵資, 重量, 序號 &#13;&#10;FROM 大宗郵件執據_資費表 &#13;&#10;ORDER BY 郵寄種類, 重量, 郵資">
                        </asp:SqlDataSource>
                    </td>
                </tr>
            </table>
       </asp:Content>
