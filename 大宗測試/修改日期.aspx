<%@ Page Title="修改日期" Language="VB" MasterPageFile="./MasterPage.master" AutoEventWireup="false" CodeFile="修改日期.aspx.vb" Inherits="修改日期" %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" Runat="Server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
            <table style="width: 100%;font-size:Large;">
                <tr>
                    <td style="vertical-align: middle; width: 100%; font-family: 標楷體; text-align: left">
                        日期<asp:DropDownList ID="DropDownList1" runat="server" style="font-size:Large;height:27px;" AutoPostBack="True" Font-Names="標楷體"
                            Font-Size="Large">
                        </asp:DropDownList>年<asp:DropDownList ID="DropDownList2" runat="server" style="font-size:Large;height:27px;" AutoPostBack="True"
                            Font-Names="標楷體" Font-Size="Large">
                        </asp:DropDownList>月<asp:DropDownList ID="DropDownList3" runat="server" style="font-size:Large;height:27px;" Font-Names="標楷體"
                            Font-Size="Large">
                        </asp:DropDownList>日
                        序號<asp:TextBox ID="TextBox30" runat="server" style="font-size:Large;" Font-Names="標楷體" Font-Size="Large" Width="45px"  visible="true"></asp:TextBox>~<asp:TextBox ID="TextBox31" runat="server" style="font-size:Large;" Font-Names="標楷體" Font-Size="Large"  Width="45px"  visible="true"></asp:TextBox>
                        改為<asp:DropDownList ID="DropDownList4" runat="server" style="font-size:Large;height:27px;" AutoPostBack="True"
                            Font-Names="標楷體" Font-Size="Large">
                        </asp:DropDownList>年<asp:DropDownList ID="DropDownList5" runat="server" style="font-size:Large;height:27px;" AutoPostBack="True"
                            Font-Names="標楷體" Font-Size="Large">
                        </asp:DropDownList>月<asp:DropDownList ID="DropDownList6" runat="server" style="font-size:Large;height:27px;" Font-Names="標楷體"
                            Font-Size="Large">
                        </asp:DropDownList>日<asp:Button ID="Button3" runat="server" style="font-size:Large;"  Font-Names="標楷體" Text="存檔" /></td>
                    <td style="vertical-align: top; font-family: 標楷體; text-align: left">
                        <asp:TextBox ID="TextBox5" runat="server" style="font-size:Large;" Visible="False"></asp:TextBox></td>
                </tr>
            </table>
</asp:Content>