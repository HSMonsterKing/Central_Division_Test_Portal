<%@ Page Title="統計" Language="VB" MasterPageFile="./MasterPage.master" AutoEventWireup="false" CodeFile="統計.aspx.vb" Inherits="統計" %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" Runat="Server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">

            <table style="width: 100%;font-size:Large;">
                <tr>
                    <td style="vertical-align: top; width: 100%; font-family: 標楷體; text-align: left">
                        <div style="text-align: left">
                            <table style="width: 100%">
                                <tr>
                                    <td style="vertical-align: top; width: 40%; font-family: 標楷體; text-align: left">
                        寄件日期:<asp:DropDownList ID="DropDownList1" runat="server" style="font-size:Large;height:27px;" AutoPostBack="True" Font-Names="標楷體">
                        </asp:DropDownList>年<asp:DropDownList ID="DropDownList2" runat="server" style="font-size:Large;height:27px;" AutoPostBack="True"
                            Font-Names="標楷體">
                        </asp:DropDownList>月<asp:DropDownList ID="DropDownList3" runat="server" style="font-size:Large;height:27px;" Font-Names="標楷體">
                        </asp:DropDownList>日～<asp:DropDownList ID="DropDownList4" runat="server" style="font-size:Large;height:27px;" AutoPostBack="True"
                            Font-Names="標楷體">
                        </asp:DropDownList>年<asp:DropDownList ID="DropDownList5" runat="server" style="font-size:Large;height:27px;" AutoPostBack="True"
                            Font-Names="標楷體">
                        </asp:DropDownList>月<asp:DropDownList ID="DropDownList6" runat="server" style="font-size:Large;height:27px;" Font-Names="標楷體">
                        </asp:DropDownList>日止<asp:Button ID="Button3" runat="server" style="font-size:Large;" Font-Names="標楷體" Text="統計" />&nbsp;</td>
                                    <td rowspan="4" style="vertical-align: middle; width: 5%; font-family: 標楷體; text-align: right">
                                        備註:</td>
                                    <td style="vertical-align: top; width: 55%; font-family: 標楷體; text-align: left">
                                        <asp:TextBox ID="TextBox1" runat="server" style="font-size:Large;" Font-Names="標楷體" Width="779px"></asp:TextBox></td>
                                </tr>
                                <tr>
                                    <td style="vertical-align: top; width: 40%; font-family: 標楷體; text-align: left">
                        <asp:Label ID="Label1" runat="server" style="font-size:Large;" Font-Names="標楷體" Text="Label"></asp:Label></td>
                                    <td style="vertical-align: top; width: 55%; font-family: 標楷體; text-align: left">
                                        <asp:TextBox ID="TextBox3" runat="server" style="font-size:Large;" Font-Names="標楷體" Width="779px"></asp:TextBox></td>
                                </tr>
                                <tr>
                                    <td style="vertical-align: top; width: 40%; font-family: 標楷體; text-align: left">
                        <asp:Label ID="Label2" runat="server" style="font-size:Large;" Font-Names="標楷體" Text="Label"></asp:Label></td>
                                    <td style="vertical-align: top; width: 55%; font-family: 標楷體; text-align: left">
                                        <asp:TextBox ID="TextBox4" runat="server" style="font-size:Large;" Font-Names="標楷體" Width="779px"></asp:TextBox></td>
                                </tr>
                                <tr>
                                    <td style="vertical-align: top; width: 40%; font-family: 標楷體; text-align: left">
                                    </td>
                                    <td style="vertical-align: top; width: 55%; font-family: 標楷體; text-align: left">
                                        <asp:TextBox ID="TextBox5" runat="server" style="font-size:Large;" Font-Names="標楷體" Width="779px"></asp:TextBox></td>
                                </tr>
                                <tr>
                                    <td style="vertical-align: top; width: 40%; font-family: 標楷體; text-align: left">
                                    </td>
                                    <td style="vertical-align: top; width: 5%; font-family: 標楷體; text-align: left">
                                    </td>
                                    <td style="vertical-align: top; width: 55%; font-family: 標楷體; text-align: left">
                                        <asp:RadioButtonList ID="RadioButtonList1" runat="server" style="font-size:Large;" AutoPostBack="True" DataTextField="名稱"
                                            DataValueField="id" RepeatDirection="Horizontal">
                                            <asp:ListItem Value="1">全部</asp:ListItem>
                                            <asp:ListItem Value="2">中分局本部</asp:ListItem>
                                            <asp:ListItem Value="3">業務科</asp:ListItem>
                                        </asp:RadioButtonList></td>
                                </tr>
                                <tr>
                                    <td style="vertical-align: top; width: 40%; font-family: 標楷體; text-align: left">
                        <asp:GridView ID="GridView1" runat="server" style="font-size:Large;" AutoGenerateColumns="False" DataSourceID="SqlDataSource1" CellPadding="4" EnableModelValidation="True" ForeColor="#333333" GridLines="None">
                            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                            <Columns>
                                <asp:BoundField DataField="郵寄種類" HeaderText="郵寄種類" SortExpression="郵寄種類" >
                                    <HeaderStyle HorizontalAlign="Center" />
                                    <ItemStyle HorizontalAlign="Left" Width="200px" />
                                </asp:BoundField>
                                <asp:BoundField DataField="件數" HeaderText="件數" SortExpression="件數" DataFormatString="{0:N0}" >
                                    <HeaderStyle HorizontalAlign="Center" />
                                    <ItemStyle HorizontalAlign="Right" Width="100px" />
                                </asp:BoundField>
                                <asp:BoundField DataField="郵資" DataFormatString="{0:N0}" HeaderText="郵資" SortExpression="郵資">
                                    <HeaderStyle HorizontalAlign="Center" />
                                    <ItemStyle HorizontalAlign="Right" Width="100px" />
                                </asp:BoundField>
                            </Columns>
                            <EditRowStyle BackColor="#999999" />
                            <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                            <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                            <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                            <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                            <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                        </asp:GridView>
                                    </td>
                                    <td style="vertical-align: top; width: 5%; font-family: 標楷體; text-align: left">
                                    </td>
                                    <td style="vertical-align: top; width: 55%; font-family: 標楷體; text-align: left">
                                        <asp:Button ID="Button4" runat="server" style="font-size:Large;" Font-Names="標楷體" Text="郵務種類日報表" />
                                        <asp:Button ID="Button5" runat="server" style="font-size:Large;" Font-Names="標楷體" Text="每日郵資統計表" />
                                        <asp:Button ID="Button6" runat="server" style="font-size:Large;" Font-Names="標楷體" Text="每日件數統計表" />
                                        <asp:Button ID="Button7" runat="server" style="font-size:Large;" Font-Names="標楷體" Text="每日郵資暨件數統計表" />
                                    </td>
                                </tr>
                                <tr>
                                    <td style="vertical-align: top; width: 40%; font-family: 標楷體; text-align: left">
                        </td>
                                    <td style="vertical-align: top; width: 5%; font-family: 標楷體; text-align: left">
                                    </td>
                                    <td style="vertical-align: top; width: 55%; font-family: 標楷體; text-align: left">
                                    </td>
                                </tr>
                                <tr>
                                    <td style="vertical-align: top; width: 40%; font-family: 標楷體; text-align: left">
                                        &nbsp;</td>
                                    <td style="vertical-align: top; width: 5%; font-family: 標楷體; text-align: left">
                                    </td>
                                    <td style="vertical-align: top; width: 55%; font-family: 標楷體; text-align: left">
                                    </td>
                                </tr>
                            </table>
                        </div>
                    </td>
                </tr>
                <tr>
                    <td style="vertical-align: middle; width: 100%; font-family: 標楷體; text-align: left">
                        &nbsp;
                    </td>
                </tr>
                <tr>
                    <td style="vertical-align: middle; width: 100%; font-family: 標楷體; text-align: left">
                        </td>
                </tr>
                <tr>
                    <td style="vertical-align: middle; width: 100%; font-family: 標楷體; text-align: left">
                        </td>
                </tr>
                <tr>
                    <td style="vertical-align: top; width: 100%; font-family: 標楷體; text-align: left">
                        &nbsp;</td>
                </tr>
                <tr>
                    <td style="vertical-align: top; width: 100%; font-family: 標楷體; text-align: left">
                        <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="Data Source=10.52.0.178;Initial Catalog=大宗郵件;User ID=qaz;Password=1qaz@WSX"
                            SelectCommand="SELECT 郵寄種類,件數,郵資&#13;&#10;FROM 大宗郵件執據_郵寄種類 &#13;&#10;ORDER BY 排序"></asp:SqlDataSource>
                        <asp:TextBox ID="TextBox2" runat="server" style="font-size:Large;" Visible="False"></asp:TextBox></td>
                </tr>
            </table>
      </asp:Content>
