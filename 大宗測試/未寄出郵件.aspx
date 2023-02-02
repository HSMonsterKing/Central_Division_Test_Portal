<%@ Page Title="未寄出郵件" Language="VB" MasterPageFile="./MasterPage.master" AutoEventWireup="false" CodeFile="未寄出郵件.aspx.vb" Inherits="未寄出郵件" %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" Runat="Server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
    <table style="vertical-align: top; width: 100%; font-family: 標楷體; text-align: left">
        <tr>
            <td style="width: 100%; vertical-align: middle; font-family: 標楷體; text-align: left;">
                <asp:Calendar ID="Calendar1" runat="server" BackColor="White" BorderColor="Black" BorderStyle="Solid" CellSpacing="1" Font-Names="Verdana" Font-Size="9pt" ForeColor="Black" Height="250px" NextPrevFormat="ShortMonth" Width="1056px">
                    <SelectedDayStyle BackColor="#333399" ForeColor="White" />
                    <TodayDayStyle BackColor="#999999" ForeColor="White" />
                    <OtherMonthDayStyle ForeColor="#999999" />
                    <DayStyle BackColor="#CCCCCC" />
                    <NextPrevStyle Font-Bold="True" Font-Size="8pt" ForeColor="White" />
                    <DayHeaderStyle Font-Bold="True" Font-Size="8pt" ForeColor="#333333" Height="8pt" />
                    <TitleStyle BackColor="#333399" BorderStyle="Solid" Font-Bold="True" Font-Size="12pt" ForeColor="White" Height="12pt" />
                </asp:Calendar>
            </td>
        </tr>
        <tr>
            <td style="width: 100%; vertical-align: middle; font-family: 標楷體; text-align: left;">
                <asp:ListBox ID="ListBox1" runat="server" Font-Names="標楷體" Height="500px" Width="1056px" Font-Size="X-Large">
                </asp:ListBox></td>
        </tr>
        <tr>
            <td style="width: 100%; vertical-align: middle; font-family: 標楷體; text-align: left;">
                <asp:ListBox ID="ListBox2" runat="server" Font-Names="標楷體" Height="95px" Width="500px" Font-Size="X-Large" Visible="False">
                </asp:ListBox></td>
        </tr>
    </table>
</asp:Content>
