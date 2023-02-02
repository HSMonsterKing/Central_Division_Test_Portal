<%@ Page Title="搜尋" Language="VB" MasterPageFile="./MasterPage.master" AutoEventWireup="false" CodeFile="搜尋.aspx.vb" Inherits="搜尋" %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" Runat="Server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
                <table style="width: 100%;font-size:Large;">
                    <tr>
                        <td style="vertical-align: top; width: 100%; font-family: 標楷體; height: 25px; text-align: left">
                            寄件日期<asp:DropDownList ID="DropDownList1" runat="server" style="font-size:Large;height:27px;" AutoPostBack="True"
                                Font-Names="標楷體" Font-Size="Large">
                            </asp:DropDownList>年<asp:DropDownList ID="DropDownList2" runat="server" style="font-size:Large;height:27px;" AutoPostBack="True"
                                Font-Names="標楷體" Font-Size="Large">
                            </asp:DropDownList>月<asp:DropDownList ID="DropDownList3" runat="server" style="font-size:Large;height:27px;"
                                Font-Names="標楷體" Font-Size="Large">
                            </asp:DropDownList>日～<asp:DropDownList ID="DropDownList4" runat="server" style="font-size:Large;height:27px;" AutoPostBack="True"
                                Font-Names="標楷體" Font-Size="Large">
                            </asp:DropDownList>年<asp:DropDownList ID="DropDownList5" runat="server" style="font-size:Large;height:27px;" AutoPostBack="True"
                                Font-Names="標楷體" Font-Size="Large">
                            </asp:DropDownList>月<asp:DropDownList ID="DropDownList6" runat="server" style="font-size:Large;height:27px;"
                                Font-Names="標楷體" Font-Size="Large">
                            </asp:DropDownList>日</td>
                    </tr>
                    <tr>
                        <td style="vertical-align: top; width: 100%; font-family: 標楷體; height: 25px; text-align: left">
                            請輸入<asp:TextBox ID="TextBox1" runat="server" style="font-size:Large;" Font-Names="標楷體" Width="437px"></asp:TextBox>&nbsp;
                            <asp:Button ID="Button3" runat="server" style="font-size:Large;" Font-Names="標楷體" Text="收件人" />
                            <asp:Button ID="Button4" runat="server" style="font-size:Large;" Font-Names="標楷體" Text="文號" />
                    </tr>
                    <tr>
                        <td style="vertical-align: top; width: 100%; font-family: 標楷體; height: 25px; text-align: left">
                            <asp:GridView ID="GridView1" runat="server" style="font-size:Large;" AutoGenerateColumns="False" CellPadding="4"
                                DataSourceID="SqlDataSource1" ForeColor="#333333" GridLines="None" EnableModelValidation="True">
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <Columns>
                                    <asp:BoundField DataField="日期" HeaderText="日期" ReadOnly="True" SortExpression="日期">
                                        <HeaderStyle HorizontalAlign="Center" />
                                        <ItemStyle HorizontalAlign="Left" Width="70px" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="序號" HeaderText="序號" SortExpression="序號">
                                        <HeaderStyle HorizontalAlign="Center" />
                                        <ItemStyle HorizontalAlign="Left" Width="50px" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="掛號號碼" HeaderText="掛號號碼" SortExpression="掛號號碼">
                                        <HeaderStyle HorizontalAlign="Center" />
                                        <ItemStyle HorizontalAlign="Left" Width="100px" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="掛號類別" HeaderText="掛號類別" SortExpression="掛號類別">
                                    <HeaderStyle HorizontalAlign="Center" />
                                    <ItemStyle HorizontalAlign="Center" Width="150px" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="郵寄" HeaderText="郵寄方式" SortExpression="郵寄">
                                        <HeaderStyle HorizontalAlign="Center" />
                                        <ItemStyle HorizontalAlign="Left" Width="100px" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="文號" HeaderText="文號" SortExpression="文號">
                                        <HeaderStyle HorizontalAlign="Center" />
                                        <ItemStyle HorizontalAlign="Left" Width="200px" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="收件人" HeaderText="收件人" SortExpression="收件人">
                                        <HeaderStyle HorizontalAlign="Center" />
                                        <ItemStyle HorizontalAlign="Left" Width="250px" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="地址" HeaderText="地址" SortExpression="地址">
                                        <HeaderStyle HorizontalAlign="Center" />
                                        <ItemStyle HorizontalAlign="Left" Width="250px" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="備註" HeaderText="備註" SortExpression="備註">
                                        <HeaderStyle HorizontalAlign="Center" />
                                        <ItemStyle HorizontalAlign="Left" Width="200px" />
                                    </asp:BoundField>
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
                        </td>
                    </tr>
                    <tr>
                        <td style="vertical-align: top; width: 100%; font-family: 標楷體; height: 25px; text-align: left">
                            <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="Data Source=10.52.0.178;Initial Catalog=大宗郵件;User ID=qaz;Password=1qaz@WSX"
                                SelectCommand="SELECT 年 + 月 + 日 AS 日期, 序號, 掛號號碼, 收件人, 地址, 文號, 備註, 郵寄 ,掛號類別
                                FROM 郵寄_查詢 
                                where (年 +月+日&gt;=@d1  and  年 +月+日&lt;=@d2  and 收件人 LIKE '%' + @d3 + '%'  and 文號 LIKE '%' + @d4 + '%' )
                                ORDER BY 年 +月+日">
                                <SelectParameters>
                                    <asp:ControlParameter ControlID="TextBox2" Name="d1" PropertyName="Text" Type="String" />
                                    <asp:ControlParameter ControlID="TextBox4" ConvertEmptyStringToNull="False" Name="d3"
                                        PropertyName="Text" Type="String" />
                                    <asp:ControlParameter ControlID="TextBox3" Name="d2" PropertyName="Text" Type="String" />
                                    <asp:ControlParameter ControlID="TextBox5" ConvertEmptyStringToNull="False" Name="d4"
                                        PropertyName="Text" Type="String" />
                                </SelectParameters>
                            </asp:SqlDataSource>
                            <asp:TextBox ID="TextBox2" runat="server" style="font-size:Large;" Visible="False"></asp:TextBox>
                            <asp:TextBox ID="TextBox3" runat="server" style="font-size:Large;" Visible="False"></asp:TextBox>
                            <asp:TextBox ID="TextBox4" runat="server" style="font-size:Large;" Visible="False"></asp:TextBox>
                            <asp:TextBox ID="TextBox5" runat="server" style="font-size:Large;" Visible="False"></asp:TextBox>
                            <asp:TextBox ID="TextBox6" runat="server" style="font-size:Large;" Visible="False"></asp:TextBox></td>
                    </tr>
                </table>
          </asp:Content>
