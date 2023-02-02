<%@ Page Title="電子交換失敗" Language="VB" MasterPageFile="./MasterPage.master" AutoEventWireup="false" CodeFile="電子交換失敗.aspx.vb" Inherits="電子交換失敗" %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" Runat="Server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
    <style type="text/css">
    .auto-style130
    {
        white-space: pre-wrap;
        /*word-wrap: break-word;*/
        /*word-break: break-all;*/
    }
    </style>

    <table style="width: 100%;font-size:Large;font-family:標楷體;">
        <tr>
            <td runat="server" style="vertical-align: top; font-family: 標楷體;">
                <asp:Label ID="Label2" runat="server" style="font-size:Large;" Font-Names="標楷體" ForeColor="Red" Text="Label"></asp:Label></td>
        </tr>
        <tr>
            <td id="TD1" runat="server" style="font-size: large; vertical-align: top;
                font-family: 標楷體;">
                寄件日期<asp:DropDownList ID="DropDownList1" runat="server" style="font-size:Large;height:27px;" AutoPostBack="True" 
                    Font-Names="標楷體">
                </asp:DropDownList>年<asp:DropDownList ID="DropDownList2" runat="server" style="font-size:Large;height:27px;" AutoPostBack="True"
                    Font-Names="標楷體">
                </asp:DropDownList>月<asp:DropDownList ID="DropDownList3" runat="server" style="font-size:Large;height:27px;" AutoPostBack="True"
                    Font-Names="標楷體">
                </asp:DropDownList>日
                批號<asp:DropDownList ID="DropDownList7" runat="server" style="font-size:Large;height:27px;" AutoPostBack="True"
                    Font-Names="標楷體">
                </asp:DropDownList>
                文號<asp:TextBox ID="TextBox1" runat="server" style="font-size:Large;" Font-Names="標楷體"
                    Width="126px"></asp:TextBox>
                <asp:Button ID="Button2" runat="server" style="font-size:Large;" Font-Names="標楷體" Text="郵寄" />
                <asp:Button ID="Button9" runat="server" style="font-size:Large;" Font-Names="標楷體" Text="電子交換" />
                <asp:Button ID="Button12" runat="server" style="font-size:Large;" Font-Names="標楷體" Text="紙本" />
                <asp:Button ID="Button3" runat="server" style="font-size:Large;" Font-Names="標楷體" Text="存檔" />
                </td>
        </tr>
        <tr>
            <td runat="server" style="font-size: large; vertical-align: top; font-family: 標楷體;">
                <asp:TextBox ID="TextBox3" runat="server" style="font-size:Large;" Font-Names="標楷體" Font-Size="Medium" ForeColor="Blue"
                    Height="48px" TextMode="MultiLine" Width=100%></asp:TextBox></td>
        </tr>
        <tr>
            <td style="vertical-align: top; width: 1300px; font-family: 標楷體;">
                <asp:Button ID="Button10" runat="server" style="font-size:Large;" Font-Names="標楷體"
                        Text="全選" />
                <asp:Button ID="Button11" runat="server" style="font-size:Large;" Font-Names="標楷體"
                        Text="全取消" /></td>
        </tr>
        <tr>
            <td style="vertical-align: top; width: 1300px; font-family: 標楷體;">
                <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" CellPadding="4"
                    DataKeyNames="id" DataSourceID="SqlDataSource1" ForeColor="#333333" style="font-size:Large;" GridLines="None" AllowPaging="True" Width="1465px">
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <Columns>
                        <asp:TemplateField HeaderText="Y/N" SortExpression="yn" Visible="false">
                            <EditItemTemplate>
                                <asp:TextBox ID="TextBox8" runat="server" style="font-size:Large;" Text='<%# Bind("yn") %>' Visible="false"></asp:TextBox>
                            </EditItemTemplate>
                            <ItemTemplate>
                                <asp:CheckBox ID="CheckBox1" runat="server" style="font-size:Large;" Checked='<%# Bind("yn") %>' Visible="false"/>
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="業務科">
                            <ItemTemplate>
                                <asp:CheckBox ID="CheckBox2" runat="server" style="font-size:Large;" Checked='<%# Bind("收費小組") %>' />
                            </ItemTemplate>
                            <HeaderStyle HorizontalAlign="Center" />
                            <ItemStyle HorizontalAlign="Center" Width="120px" />
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="收件人" SortExpression="收件人">
                            <EditItemTemplate>
                                <asp:TextBox ID="TextBox6" runat="server" style="font-size:Large;" Text='<%# Bind("收件人") %>'></asp:TextBox>
                            </EditItemTemplate>
                            <ItemTemplate>
                                <asp:TextBox ID="TextBox18" runat="server" style="font-size:Large;" Font-Names="標楷體" Text='<%# Bind("收件人") %>'
                                    Width="360px" TextMode="MultiLine"></asp:TextBox>
                            </ItemTemplate>
                            <HeaderStyle HorizontalAlign="Center" />
                            <ItemStyle HorizontalAlign="Left" Width="300px" />
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="郵遞區號" SortExpression="郵遞區號">
                            <EditItemTemplate>
                                <asp:TextBox ID="TextBox55" runat="server" style="font-size:Large;" Text='<%# Bind("郵遞區號") %>'></asp:TextBox>
                            </EditItemTemplate>
                            <ItemTemplate>
                                <asp:TextBox ID="TextBox56" runat="server" style="font-size:Large;" Font-Names="標楷體" Text='<%# Bind("郵遞區號") %>'
                                    Width="70px" TextMode="MultiLine"></asp:TextBox>
                            </ItemTemplate>
                            <HeaderStyle HorizontalAlign="Center" />
                            <ItemStyle HorizontalAlign="Left" Width="280px" />
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="地址" SortExpression="地址">
                            <EditItemTemplate>
                                <asp:TextBox ID="TextBox7" runat="server" style="font-size:Large;" Text='<%# Bind("地址") %>'></asp:TextBox>
                            </EditItemTemplate>
                            <ItemTemplate>
                                <asp:TextBox ID="TextBox19" runat="server" style="font-size:Large;" Font-Names="標楷體" Text='<%# Bind("地址") %>'
                                    Width="460px" TextMode="MultiLine"></asp:TextBox>
                            </ItemTemplate>
                            <HeaderStyle HorizontalAlign="Center" />
                            <ItemStyle HorizontalAlign="Left" Width="280px" />
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="文號" SortExpression="文號">
                            <EditItemTemplate>
                                <asp:TextBox ID="TextBox8" runat="server" style="font-size:Large;" Text='<%# Bind("文號") %>'></asp:TextBox>
                            </EditItemTemplate>
                            <ItemTemplate>
                                <asp:TextBox ID="TextBox22" runat="server" style="font-size:Large;" Text='<%# Bind("文號") %>' TextMode="MultiLine" Width="245px" Font-Names="標楷體"></asp:TextBox>
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="附件" >
                            <ItemTemplate>
                                <asp:CheckBox ID="CheckBox40" runat="server" style="font-size:Large;" AutoPostBack="True" OnCheckedChanged="GridView1_CheckBox40_CheckedChanged"/>
                            </ItemTemplate>
                            <HeaderStyle HorizontalAlign="Center" />
                            <ItemStyle HorizontalAlign="Center" Width="120px" />
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="備註" SortExpression="備註" visible="false">
                            <EditItemTemplate>
                                <asp:TextBox ID="TextBox5" runat="server" style="font-size:Large;" Text='<%# Bind("備註") %>' visible="false"></asp:TextBox>
                            </EditItemTemplate>
                            <ItemTemplate>
                                <asp:TextBox ID="TextBox9" runat="server" style="font-size:Large;" Font-Names="標楷體" Width="200px" visible="false"></asp:TextBox>
                            </ItemTemplate>
                            <HeaderStyle HorizontalAlign="Center" />
                            <ItemStyle HorizontalAlign="Left" Width="200px" />
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="重量" SortExpression="重量" visible="false">
                            <EditItemTemplate>
                                <asp:TextBox ID="TextBox1" runat="server" style="font-size:Large;" Text='<%# Bind("重量") %>' visible="false"></asp:TextBox>
                            </EditItemTemplate>
                            <ItemTemplate>
                                <asp:TextBox ID="TextBox7" runat="server" style="font-size:Large;" Font-Names="標楷體" Width="50px" visible="false"></asp:TextBox>
                            </ItemTemplate>
                            <HeaderStyle HorizontalAlign="Center" />
                            <ItemStyle HorizontalAlign="Right" Width="50px" />
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="郵資">
                            <ItemTemplate>
                                <asp:DropDownList ID="DropDownList8" runat="server" style="font-size:Large;Large;height:27px;" AutoPostBack="True" Font-Names="標楷體">
                                </asp:DropDownList>
                                <asp:Label ID="Label2" runat="server" style="font-size:Large;" Text='<%# Eval("id") %>' Visible="False"></asp:Label>
                            </ItemTemplate>
                            <HeaderStyle HorizontalAlign="Center" />
                            <ItemStyle HorizontalAlign="Center"/>
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="郵寄種類">
                            <ItemTemplate>
                                <asp:DropDownList ID="DropDownList4" runat="server" style="font-size:Large;Large;height:27px;" AutoPostBack="True" Font-Names="標楷體" OnSelectedIndexChanged = "GridView1_DropDownList4_SelectedIndexChanged">
                                </asp:DropDownList>
                                <asp:Label ID="Label1" runat="server" style="font-size:Large;" Text='<%# Eval("id") %>' Visible="False"></asp:Label>
                            </ItemTemplate>
                            <HeaderStyle HorizontalAlign="Center" />
                            <ItemStyle HorizontalAlign="Left" Width="180px" />
                        </asp:TemplateField>
                        <asp:CommandField SelectText="存檔" ShowSelectButton="True">
                            <ItemStyle HorizontalAlign="Center" Width="40px" />
                        </asp:CommandField>
                        <asp:TemplateField>
                            <ItemTemplate>
                                <asp:Button ID="Button4" runat="server" style="font-size:Large;" CommandName="delete" Font-Names="標楷體" ForeColor="Red"
                                    Text="刪除" Visible="true"/>
                            </ItemTemplate>
                            <ItemStyle HorizontalAlign="Center" Width="40px" />
                        </asp:TemplateField>
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
            <td style="vertical-align: top; width: 1300px; font-family: 標楷體;">
            </td>
        </tr>
        <tr>
            <td style="vertical-align: top; width: 1300px; font-family: 標楷體;">
                <asp:Button ID="Button5" runat="server" Font-Names="標楷體" style="font-size:Large;" Text="存檔" />
                <asp:Button ID="Button13" runat="server" Font-Names="標楷體" style="font-size:Large;" Text="分類排序" />收件人<asp:TextBox
                    ID="TextBox23" runat="server" Font-Names="標楷體" style="font-size:Large;" Width="483px"></asp:TextBox>
                <asp:Button ID="Button14" runat="server" Font-Names="標楷體" style="font-size:Large;" Text="顯示全部" />
                <asp:Button ID="Button15" runat="server" Font-Names="標楷體" style="font-size:Large;" Text="搜尋" /></td>
        </tr>
        <tr>
            <td style="vertical-align: top; width: 1300px; font-family: 標楷體;">
                <asp:GridView ID="GridView2" runat="server" AutoGenerateColumns="False" CellPadding="4"
                    DataKeyNames="id" DataSourceID="SqlDataSource3" ForeColor="#333333" style="font-size:Large;" GridLines="None" AllowPaging="True" PagerSettings-PageButtonCount=100 EnableModelValidation="True">
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <Columns>
                        <asp:TemplateField HeaderText="業務科" >
                            <ItemTemplate>
                                <asp:CheckBox ID="CheckBox3" runat="server" style="font-size:Large;" Checked='<%# Bind("收費小組") %>' />
                            </ItemTemplate>
                            <HeaderStyle HorizontalAlign="Center" />
                            <ItemStyle HorizontalAlign="Center" Width="120px" />
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="序號" SortExpression="序號">
                            <EditItemTemplate>
                                <asp:TextBox ID="TextBox3" runat="server" style="font-size:Large;" Text='<%# Bind("序號") %>'></asp:TextBox>
                            </EditItemTemplate>
                            <ItemTemplate>
                                <asp:TextBox ID="TextBox5" runat="server" style="font-size:Large;" Font-Names="標楷體" Text='<%# Bind("序號") %>'
                                    Width="40px"></asp:TextBox>
                            </ItemTemplate>
                            <HeaderStyle HorizontalAlign="Center" />
                            <ItemStyle HorizontalAlign="Right" Width="40px" />
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="掛號號碼" SortExpression="">
                            <EditItemTemplate>
                                <asp:TextBox ID="TextBox4" runat="server" style="font-size:Large;" Text='<%# Bind("掛號號碼") %>'></asp:TextBox>
                            </EditItemTemplate>
                            <ItemTemplate>
                                <asp:TextBox ID="TextBox25" runat="server" style="font-size:Large;" Font-Names="標楷體" Text='<%# Bind("掛號號碼") %>'
                                    Width="60px"></asp:TextBox>
                                <asp:TextBox ID="TextBox6" runat="server" style="font-size:Large;" Font-Names="標楷體"
                                    Text='<%# Bind("掛號類別") %>' Width="60px" visible="false"></asp:TextBox>
                            </ItemTemplate>
                            <HeaderStyle HorizontalAlign="Center" />
                            <ItemStyle HorizontalAlign="Center" Width="190px" />
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="收件人" SortExpression="收件人">
                            <EditItemTemplate>
                                <asp:TextBox ID="TextBox6" runat="server" style="font-size:Large;" Text='<%# Bind("收件人") %>'></asp:TextBox>
                            </EditItemTemplate>
                            <ItemTemplate>
                                <asp:TextBox ID="TextBox20" runat="server" style="font-size:Large;white-space: pre-wrap;" Font-Names="標楷體" TextMode="MultiLine" class="auto-style130"  Width="360px" Text='<%# Bind("收件人") %>'></asp:TextBox>
                            </ItemTemplate>
                            <HeaderStyle HorizontalAlign="Center" />
                            <ItemStyle HorizontalAlign="Left" Width="300px" />
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="郵遞區號" SortExpression="郵遞區號">
                            <EditItemTemplate>
                                <asp:TextBox ID="TextBox55" runat="server" style="font-size:Large;" Text='<%# Bind("郵遞區號") %>'></asp:TextBox>
                            </EditItemTemplate>
                            <ItemTemplate>
                                <asp:TextBox ID="TextBox56" runat="server" style="font-size:Large;" Font-Names="標楷體" Text='<%# Bind("郵遞區號") %>'
                                    Width="70px" TextMode="MultiLine"></asp:TextBox>
                            </ItemTemplate>
                            <HeaderStyle HorizontalAlign="Center" />
                            <ItemStyle HorizontalAlign="Left" Width="280px" />
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="地址" SortExpression="地址">
                            <EditItemTemplate>
                                <asp:TextBox ID="TextBox7" runat="server" style="font-size:Large;" Text='<%# Bind("地址") %>'></asp:TextBox>
                            </EditItemTemplate>
                            <ItemTemplate>
                                <asp:TextBox ID="TextBox21" runat="server" style="font-size:Large;white-space: pre-wrap;" Font-Names="標楷體" TextMode="MultiLine" Width="460px" Text='<%# Bind("地址") %>'></asp:TextBox>
                            </ItemTemplate>
                            <HeaderStyle HorizontalAlign="Center" />
                            <ItemStyle HorizontalAlign="Left" Width="280px" />
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="文號" SortExpression="文號">
                            <EditItemTemplate>
                                <asp:TextBox ID="TextBox8" runat="server" style="font-size:Large;" Text='<%# Bind("文號") %>'></asp:TextBox>
                            </EditItemTemplate>
                            <ItemTemplate>
                                <asp:TextBox ID="TextBox22" runat="server" style="font-size:Large;white-space: pre-wrap;" Text='<%# Bind("文號") %>' TextMode="MultiLine" Width="345px" Font-Names="標楷體"></asp:TextBox>
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="備註" SortExpression="備註" visible="false">
                            <EditItemTemplate>
                                <asp:TextBox ID="TextBox5" runat="server" style="font-size:Large;" Text='<%# Bind("備註") %>' visible="false"></asp:TextBox>
                            </EditItemTemplate>
                            <ItemTemplate>
                                <asp:TextBox ID="TextBox9" runat="server" style="font-size:Large;" Font-Names="標楷體" Text='<%# Bind("備註") %>'
                                    Width="45px" TextMode="MultiLine" visible="false"></asp:TextBox>
                            </ItemTemplate>
                            <HeaderStyle HorizontalAlign="Center" />
                            <ItemStyle HorizontalAlign="Left" Width="200px" />
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="件數" SortExpression="件數">
                            <EditItemTemplate>
                                <asp:TextBox ID="TextBox9" runat="server" style="font-size:Large;" Text='<%# Bind("件數") %>'></asp:TextBox>
                            </EditItemTemplate>
                            <ItemTemplate>
                                <asp:TextBox ID="TextBox26" runat="server" style="font-size:Large;" Text='<%# Bind("件數") %>' Width="35px" Font-Names="標楷體"></asp:TextBox>
                            </ItemTemplate>
                            <HeaderStyle HorizontalAlign="Center" />
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="重量" SortExpression="重量" visible="false">
                            <EditItemTemplate>
                                <asp:TextBox ID="TextBox1" runat="server" style="font-size:Large;" Text='<%# Bind("重量") %>' visible="false"></asp:TextBox>
                            </EditItemTemplate>
                            <ItemTemplate>
                                <asp:TextBox ID="TextBox7" runat="server" style="font-size:Large;" Font-Names="標楷體" Text='<%# Bind("重量") %>' Width="50px" visible="false"></asp:TextBox>
                            </ItemTemplate>
                            <HeaderStyle HorizontalAlign="Center" />
                            <ItemStyle HorizontalAlign="Right" Width="50px" />
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="郵資">
                            <ItemTemplate>
                                <asp:DropDownList ID="DropDownList8" runat="server" style="font-size:Large;Large;height:27px;" AutoPostBack="True" Font-Names="標楷體">
                                </asp:DropDownList>
                                <asp:Label ID="Label2" runat="server" style="font-size:Large;" Text='<%# Eval("id") %>' Visible="False"></asp:Label>
                            </ItemTemplate>
                            <HeaderStyle HorizontalAlign="Center" />
                            <ItemStyle HorizontalAlign="Left" Width="180px" />
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="郵寄種類">
                            <ItemTemplate>
                                <asp:DropDownList ID="DropDownList4" runat="server" style="font-size:Large;Large;height:27px;" AutoPostBack="True" Font-Names="標楷體" OnSelectedIndexChanged = "GridView2_DropDownList4_SelectedIndexChanged">
                                </asp:DropDownList>
                                <asp:Label ID="Label1" runat="server" style="font-size:Large;" Text='<%# Eval("id") %>' Visible="False"></asp:Label>
                            </ItemTemplate>
                            <HeaderStyle HorizontalAlign="Center" />
                            <ItemStyle HorizontalAlign="Left" Width="1800px" />
                        </asp:TemplateField>
                        <asp:CommandField SelectText="存檔" ShowSelectButton="True">
                            <ItemStyle HorizontalAlign="Center" Width="40px" />
                        </asp:CommandField>
                        <asp:TemplateField>
                            <ItemTemplate>
                                <asp:Button ID="Button4" runat="server" style="font-size:Large;" CommandName="delete" Font-Names="標楷體" ForeColor="Red"
                                    OnClientClick="return confirm('刪除動作無法復原，是否繼續？')" Text="刪除"/>
                            </ItemTemplate>
                            <ItemStyle HorizontalAlign="Center" Width="40px" />
                        </asp:TemplateField>
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
            <td style="vertical-align: top; font-family: 標楷體;">
                <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="Data Source=10.52.0.178;Initial Catalog=大宗郵件;User ID=qaz;Password=1qaz@WSX"
                    DeleteCommand="DELETE FROM 大宗郵件執據_bak WHERE id=@id" SelectCommand="SELECT id, 序號, 掛號號碼, 收件人, 郵遞區號, 地址, 文號, 備註, 重量, 郵資 ,yn,郵寄種類,收費小組 FROM 大宗郵件執據_bak&#13;&#10;where 帳號=@_帳號&#13;&#10;ORDER BY 序號">
                    <SelectParameters>
                        <asp:ControlParameter ControlID="TextBox2" Name="_帳號" PropertyName="Text" Type="String" />
                    </SelectParameters>
                    <DeleteParameters>
                        <asp:Parameter Name="id" />
                    </DeleteParameters>
                </asp:SqlDataSource>
                <asp:SqlDataSource ID="SqlDataSource2" runat="server" ConnectionString="Data Source=10.52.0.178;Initial Catalog=大宗郵件;User ID=qaz;Password=1qaz@WSX"
                    SelectCommand="SELECT 序號, 郵寄種類 FROM 大宗郵件執據_郵寄種類 ORDER BY 排序"></asp:SqlDataSource>
                <asp:SqlDataSource ID="SqlDataSource3" runat="server" ConnectionString="Data Source=10.52.0.178;Initial Catalog=大宗郵件;User ID=qaz;Password=1qaz@WSX"
                    DeleteCommand="DELETE FROM 大宗郵件執據  WHERE id=@id" SelectCommand="SELECT id, 序號, 掛號類別,掛號號碼, 收件人, 郵遞區號, 地址, 文號, 備註, 重量, 郵資, 郵寄種類,件數,收費小組 FROM 大宗郵件執據 where (年=@_年 and 月=@_月 and 日=@_日 and 批號=@_批號) and ((收件人 is null and ''=@_收件人)or 收件人 Like N'%'+@_收件人+'%') ORDER BY 序號">
                    <SelectParameters>
                        <asp:ControlParameter ControlID="DropDownList1" Name="_年" PropertyName="SelectedValue"
                            Type="String" />
                        <asp:ControlParameter ControlID="DropDownList2" Name="_月" PropertyName="SelectedValue"
                            Type="String" />
                        <asp:ControlParameter ControlID="DropDownList3" Name="_日" PropertyName="SelectedValue"
                            Type="String" />
                        <asp:ControlParameter ControlID="DropDownList7" Name="_批號" PropertyName="SelectedValue"
                            Type="String" />
                        <asp:ControlParameter ControlID="TextBox23" ConvertEmptyStringToNull="false" Name="_收件人"
                            PropertyName="Text" Type="String" />
                    </SelectParameters>
                    <DeleteParameters>
                        <asp:Parameter Name="id" />
                    </DeleteParameters>
                </asp:SqlDataSource>
                <asp:TextBox ID="TextBox2" runat="server" style="font-size:Large;" Visible="False"></asp:TextBox></td>
        </tr>
    </table>
</asp:Content>


