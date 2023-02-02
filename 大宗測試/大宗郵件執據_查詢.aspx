<%@ Page Title="大宗郵件執據_查詢" Language="VB" MasterPageFile="./MasterPage.master" AutoEventWireup="false" CodeFile="大宗郵件執據_查詢.aspx.vb" Inherits="大宗郵件執據_查詢" %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" Runat="Server">
    <style type="text/css">
        .auto-style1 {
            width: 1300px;
            height: 28px;
        }
        
        /* Tooltip container */
        .tooltip {
        position: relative;
        display: inline-block;
        }

        /* Tooltip text */
        .tooltip .tooltiptext {
        visibility: hidden;
        background-color: rgba(0,0,0,0.6);
        color: #fff;
        text-align: center;
        padding: 5px;
        border-radius: 6px;
        /* Position the tooltip text - see examples below! */
        position: absolute;
        z-index: 1;
        top: 100%;
        left: 50%;
        margin-left: -20px; /* Use half of the width (120/2 = 60), to center the tooltip */
        }

        /* Show the tooltip text when you mouse over the tooltip container */
        .tooltip:hover .tooltiptext {
        visibility: visible;
        }
    </style>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">

        <table style="width: 100%;font-size:Large;">
            <tr>
                <td colspan="5" style="vertical-align: top; font-family: 標楷體; text-align: left" class="auto-style1">
                    寄件日期<asp:DropDownList ID="DropDownList1" runat="server" style="font-size:Large;height:27px;" AutoPostBack="True"
                        Font-Names="標楷體" Font-Size="Large">
                    </asp:DropDownList>年<asp:DropDownList ID="DropDownList2" runat="server" style="font-size:Large;height:27px;" AutoPostBack="True"
                        Font-Names="標楷體" Font-Size="Large">
                    </asp:DropDownList>月<asp:DropDownList ID="DropDownList3" runat="server" style="font-size:Large;height:27px;" AutoPostBack="True"
                        Font-Names="標楷體" Font-Size="Large">
                    </asp:DropDownList>日
                    郵寄種類<asp:DropDownList ID="DropDownList5" runat="server" style="font-size:Large;height:27px;" AutoPostBack="True"
                        DataTextField="郵寄種類" DataValueField="序號" Font-Names="標楷體" Font-Size="Large">
                    </asp:DropDownList>
                    收件人<asp:TextBox ID="TextBox19" runat="server" style="font-size:Large;" Font-Names="標楷體" Width="300px"></asp:TextBox>
                    文號<asp:TextBox ID="TextBox20" runat="server" style="font-size:Large;" Font-Names="標楷體"></asp:TextBox>
                    <asp:Button ID="Button23" runat="server" style="font-size:Large;" Font-Names="標楷體" Text="查詢" />
                </td>
            </tr>
            <tr>
                <td colspan="5" style="vertical-align: top; width: 1300px; font-family: 標楷體; text-align: left">
                    <asp:Label ID="Label2" runat="server" style="font-size:Large;" Font-Names="標楷體" ForeColor="Teal" Text="Label"></asp:Label></td>
            </tr>
            <tr>
                <td colspan="5" style="vertical-align: top; width: 1300px; font-family: 標楷體; text-align: left">
                    <asp:Button ID="Button10" runat="server" style="font-size:Large;" Font-Names="標楷體" Font-Size="Large"
                        Text="交寄大宗函件執據" />
                    <asp:Button ID="Button50" runat="server" style="font-size:Large;" Font-Names="標楷體" Font-Size="Large"
                        Text="特約郵件郵費單" />
                    <div class="tooltip">
                        <asp:TextBox ID="TextBox52" runat="server" style="font-size:Large;width:25px;" Font-Names="標楷體"></asp:TextBox>
                        <span class="tooltiptext">地址標籤第一頁空下?個，例如上次最後一頁印了3個，請填入3</span>
                    </div>

                    <asp:Button ID="Button51" runat="server" style="font-size:Large;" Font-Names="標楷體" Font-Size="Large"
                        Text="地址標籤" />
                    <asp:Button ID="Button52" runat="server" style="font-size:Large;" Font-Names="標楷體" Font-Size="Large"
                        Text="新地址標籤" />
                    <asp:Button ID="Button11" runat="server" style="font-size:Large;" Font-Names="標楷體" Font-Size="Large"
                        Text="合計" />
                    <asp:Button ID="Button14" runat="server" style="font-size:Large;" Font-Names="標楷體" Font-Size="Large"
                        Text="*合計" />
                    <asp:Button ID="Button13" runat="server" style="font-size:Large;" Font-Names="標楷體" Font-Size="Large" Text="檢查" />
                    <asp:Label ID="Label3" runat="server" style="font-size:Large;" Font-Names="標楷體" ForeColor="Red" Text="Label"></asp:Label>
                </td>
            </tr>
            <tr>
                <td colspan="5" style="vertical-align: top; width: 1300px; font-family: 標楷體; text-align: left">
                    <asp:Button ID="Button2" runat="server" style="font-size:Large;" Font-Names="標楷體" Font-Size="Large" Text="重算序號" visible="false"/>
                    列號<asp:TextBox ID="TextBox30" runat="server" style="font-size:Large;" Font-Names="標楷體" Font-Size="Large" Width="45px"  visible="true"></asp:TextBox>~<asp:TextBox ID="TextBox31" runat="server" style="font-size:Large;" Font-Names="標楷體" Font-Size="Large"  Width="45px"  visible="true"></asp:TextBox>掛號號碼<asp:TextBox ID="TextBox1" runat="server" style="font-size:Large;" Font-Names="標楷體" Font-Size="Large" Width="75px"></asp:TextBox>
                    <asp:TextBox ID="TextBox16" runat="server" style="font-size:Large;" Width="61px"  visible="false">1</asp:TextBox>
                    <asp:Button ID="Button3" runat="server" style="font-size:Large;" Font-Names="標楷體" Font-Size="Large" Text="代入掛號號碼(自動存檔)"/>
                    <asp:Button ID="Button5" runat="server" style="font-size:Large;" Font-Names="標楷體" Font-Size="Large" Text="存檔" visible="false"/>
                    <asp:TextBox ID="TextBox17" runat="server" style="font-size:Large;" Font-Names="標楷體" Font-Size="Large" Width="126px" visible="false"></asp:TextBox>
                    <asp:Button ID="Button12" runat="server" style="font-size:Large;" Font-Names="標楷體" Font-Size="Large" Text="掛號類別" visible="false"/>
                    <asp:TextBox ID="TextBox10" runat="server" style="font-size:Large;" Font-Names="標楷體" Font-Size="Large" Width="81px" visible="false">0</asp:TextBox>
                    <asp:Button ID="Button8" runat="server" style="font-size:Large;" Font-Names="標楷體" Font-Size="Large" Text="重量" visible="false"/>
                    <asp:TextBox ID="TextBox11" runat="server" style="font-size:Large;" Font-Names="標楷體" Font-Size="Large" Width="81px" visible="false">0</asp:TextBox>
                    <asp:Button ID="Button9" runat="server" style="font-size:Large;" Font-Names="標楷體" Font-Size="Large" Text="郵資" visible="false"/>
                </td>
            </tr>
            <tr>
                <td colspan="5" style="vertical-align: top; width: 1300px; font-family: 標楷體; text-align: left">
                </td>
            </tr>
            <tr>
                <td colspan="5" style="vertical-align: top; width: 1300px; font-family: 標楷體; text-align: left">
                    <asp:GridView ID="GridView2" runat="server" style="font-size:Large;" AutoGenerateColumns="False" CellPadding="4"
                        DataKeyNames="id" DataSourceID="SqlDataSource3" ForeColor="#333333" GridLines="None" PagerSettings-PageButtonCount=100 AllowPaging="True" EnableModelValidation="True">
                        <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                        <Columns>                
                            <asp:TemplateField HeaderText="業務科">
                                <ItemTemplate>
                                    <asp:CheckBox ID="CheckBox1" runat="server" style="font-size:Large;" Checked='<%# Bind("收費小組") %>' />
                                </ItemTemplate>
                                <HeaderStyle HorizontalAlign="Center" />
                                <ItemStyle HorizontalAlign="Center" Width="40px" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="列號">   
                                <ItemTemplate>
                                        <%# Container.DataItemIndex + 1 %>   
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="序號" SortExpression="序號" visible="false">
                                <EditItemTemplate>
                                    <asp:TextBox ID="TextBox3" runat="server" style="font-size:Large;" Text='<%# Bind("序號") %>' visible="false"></asp:TextBox>
                                </EditItemTemplate>
                                <ItemTemplate>
                                    <asp:TextBox ID="TextBox5" runat="server" style="font-size:Large;" Font-Names="標楷體" Text='<%# Bind("序號") %>' Width="40px" visible="false"></asp:TextBox>
                                </ItemTemplate>
                                <HeaderStyle HorizontalAlign="Center" />
                                <ItemStyle HorizontalAlign="Right" Width="40px" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="掛號號碼" SortExpression="">
                                <EditItemTemplate>
                                    <asp:TextBox ID="TextBox4" runat="server" style="font-size:Large;" Text='<%# Bind("掛號號碼") %>'></asp:TextBox>
                                </EditItemTemplate>
                                <ItemTemplate>
                                    <asp:TextBox ID="TextBox15" runat="server" style="font-size:Large;" Font-Names="標楷體" Text='<%# Bind("掛號號碼") %>'
                                        Width="65px"></asp:TextBox>
                                    <asp:TextBox ID="TextBox6" runat="server" style="font-size:Large;" Font-Names="標楷體"
                                        Text='<%# Bind("掛號類別") %>' Width="65px" visible="false"></asp:TextBox>
                                </ItemTemplate>
                                <HeaderStyle HorizontalAlign="Center" />
                                <ItemStyle HorizontalAlign="Center" Width="190px" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="收件人" SortExpression="收件人">
                                <EditItemTemplate>
                                    <asp:TextBox ID="TextBox6" runat="server" style="font-size:Large;" Text='<%# Bind("收件人") %>'></asp:TextBox>
                                </EditItemTemplate>
                                <ItemTemplate>
                                    <asp:TextBox ID="TextBox12" runat="server" style="font-size:Large;white-space: pre-wrap;" Font-Names="標楷體" TextMode="MultiLine" Width="340px" Text='<%# Bind("收件人") %>'></asp:TextBox>
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
                                    <asp:TextBox ID="TextBox13" runat="server" style="font-size:Large;white-space: pre-wrap;" Font-Names="標楷體" TextMode="MultiLine" Width="460px" Text='<%# Bind("地址") %>'></asp:TextBox>
                                </ItemTemplate>
                                <HeaderStyle HorizontalAlign="Center" />
                                <ItemStyle HorizontalAlign="Left" Width="250px" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="文號" SortExpression="文號">
                                <EditItemTemplate>
                                    <asp:TextBox ID="TextBox8" runat="server" style="font-size:Large;" Text='<%# Bind("文號") %>'></asp:TextBox>
                                </EditItemTemplate>
                                <ItemTemplate>
                                    <asp:TextBox ID="TextBox14" runat="server" style="font-size:Large;white-space: pre-wrap;" Text='<%# Bind("文號") %>' TextMode="MultiLine" Width="345px" Font-Names="標楷體"></asp:TextBox>
                                </ItemTemplate>
                                <ItemStyle Width="190px" />
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
                                    <asp:TextBox ID="TextBox18" runat="server" style="font-size:Large;" Text='<%# Bind("件數") %>' Width="35px" Font-Names="標楷體"></asp:TextBox>
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
                            <asp:TemplateField HeaderText="郵資" SortExpression="郵資">
                                <EditItemTemplate>
                                    <asp:TextBox ID="TextBox2" runat="server" style="font-size:Large;" Text='<%# Bind("郵資") %>'></asp:TextBox>
                                </EditItemTemplate>
                                <ItemTemplate>
                                    <asp:TextBox ID="TextBox8" runat="server" style="font-size:Large;" Font-Names="標楷體" Text='<%# Bind("郵資") %>'
                                        Width="35px"></asp:TextBox>
                                </ItemTemplate>
                                <HeaderStyle HorizontalAlign="Center" />
                                <ItemStyle HorizontalAlign="Right" Width="50px" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="郵寄種類">
                                <ItemTemplate>
                                    <asp:DropDownList ID="DropDownList4" runat="server" style="font-size:Large;height:27px;" DataSourceID="SqlDataSource2"
                                        DataTextField="郵寄種類" DataValueField="序號" Font-Names="標楷體" SelectedValue='<%# Bind("郵寄種類") %>'>
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
                                        OnClientClick="return confirm('刪除動作無法復原，是否繼續？')" Text="刪除" />
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
                <td style="vertical-align: top; width: 120px; font-family: 標楷體; height: 18px; text-align: right">
                </td>
                <td style="vertical-align: top; width: 300px; font-family: 標楷體; height: 18px; text-align: left">
                </td>
                <td style="vertical-align: top; width: 130px; font-family: 標楷體; height: 18px; text-align: right">
                </td>
                <td style="vertical-align: top; width: 200px; font-family: 標楷體; height: 18px; text-align: left">
                </td>
                <td style="vertical-align: top; width: 100px; font-family: 標楷體; height: 18px; text-align: left">
                </td>
            </tr>
            <tr>
                <td colspan="5" style="vertical-align: top; font-family: 標楷體; text-align: left">
                    <asp:SqlDataSource ID="SqlDataSource2" runat="server" ConnectionString="Data Source=10.52.0.178;Initial Catalog=大宗郵件;User ID=qaz;Password=1qaz@WSX"
                        SelectCommand="SELECT 序號, 郵寄種類 FROM 大宗郵件執據_郵寄種類 ORDER BY 排序"></asp:SqlDataSource>
                    <asp:SqlDataSource ID="SqlDataSource3" runat="server" ConnectionString="Data Source=10.52.0.178;Initial Catalog=大宗郵件;User ID=qaz;Password=1qaz@WSX"
                        DeleteCommand="DELETE FROM 大宗郵件執據  WHERE id=@id"
                        SelectCommand="SELECT *
                        FROM 大宗郵件執據 a INNER JOIN 大宗郵件執據_郵寄種類 b ON a.郵寄種類 = b.序號
                        where a.年=@_年 and a.月=@_月 and a.日=@_日 
                        and (0=@_郵寄種類 or a.郵寄種類=@_郵寄種類) 
                        and ((a.收件人 is null and ''=@_收件人)or(a.收件人 Like N'%'+@_收件人+'%')) 
                        and ((a.文號 is null and ''=@_文號)or(a.文號 LIKE '%'+@_文號+'%')) 
                        ORDER BY b.排序, a.序號">
                        <SelectParameters>
                            <asp:ControlParameter ControlID="DropDownList1" Name="_年" PropertyName="SelectedValue"
                                Type="String" />
                            <asp:ControlParameter ControlID="DropDownList2" Name="_月" PropertyName="SelectedValue"
                                Type="String" />
                            <asp:ControlParameter ControlID="DropDownList3" Name="_日" PropertyName="SelectedValue"
                                Type="String" />
                            <asp:ControlParameter ControlID="DropDownList5" Name="_郵寄種類" PropertyName="SelectedValue"
                                Type="String" ConvertEmptyStringToNull="False" />
                            <asp:ControlParameter ControlID="TextBox19" ConvertEmptyStringToNull="False" Name="_收件人" PropertyName="Text" Type="String" />
                            <asp:ControlParameter ControlID="TextBox20" ConvertEmptyStringToNull="False" Name="_文號" PropertyName="Text" Type="String" />
                        </SelectParameters>
                        <DeleteParameters>
                            <asp:Parameter Name="id" />
                        </DeleteParameters>
                    </asp:SqlDataSource>
                    <asp:TextBox ID="TextBox2" runat="server" style="font-size:Large;" Visible="False"></asp:TextBox></td>
            </tr>
        </table>
    
</asp:Content>