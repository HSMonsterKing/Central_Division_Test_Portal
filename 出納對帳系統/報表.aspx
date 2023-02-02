<%@ Page Title="出納對帳系統" Language="VB" MasterPageFile="./MasterPage.master" MaintainScrollPositionOnPostback="true" AutoEventWireup="false" CodeFile="報表.aspx.vb" Inherits="報表" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
    <link rel="stylesheet" runat="server" media="screen" href="css\報表.css"/>
    <style>
        #table1, #table1 tr, #table1 td{
            border:1px solid black;
            border-collapse: collapse;
        }
        #table1 td{
            height: 22px;
        }
        #ContentPlaceHolder1_LabelC6, #ContentPlaceHolder1_LabelD6, #ContentPlaceHolder1_LabelE6, #ContentPlaceHolder1_LabelF6, #ContentPlaceHolder1_LabelC7, #ContentPlaceHolder1_LabelD7, #ContentPlaceHolder1_LabelE7, #ContentPlaceHolder1_LabelF7, #ContentPlaceHolder1_LabelC11, #ContentPlaceHolder1_LabelD11, #ContentPlaceHolder1_LabelE11, #ContentPlaceHolder1_LabelF11{
            font-size:Large;
            text-align:right;
            display:block;
            font-family:新細明體;
            padding-left:9px;
            padding-right:9px;
        }
        #ContentPlaceHolder1_LabelC12, #ContentPlaceHolder1_LabelE12, #ContentPlaceHolder1_LabelC13, #ContentPlaceHolder1_LabelE13, #ContentPlaceHolder1_LabelC14, #ContentPlaceHolder1_LabelE14, #ContentPlaceHolder1_LabelC15, #ContentPlaceHolder1_LabelE15{
            font-size:Large;
            font-family:新細明體;
            display:block;
        }
    </style> 
    <div style="width:1600px;"></div>
    <table style="font-size:large;margin:auto;text-align:center;">
        <td style="width:500px;text-align:right;">
        </td>
        <td>
            <table style="font-size:large;margin:auto;text-align:center;">
                <tr>
                    <td>
                    </td>
                    <td style="font-weight:bold;font-size:x-large;">
                        月報表
                    </td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td style="width:113px;text-align:right;">
                    </td>
                    <td runat="server">
                        結帳日期<asp:DropDownList ID="結帳日期a" runat="server" DataSourceID="SqlDataSource2" DataTextField="結帳日期" DataValueField="結帳日期" CssClass="DropDownList"/>~<asp:DropDownList ID="結帳日期b" runat="server" DataSourceID="SqlDataSource2" DataTextField="結帳日期" DataValueField="結帳日期" CssClass="DropDownList"/>
                    </td>
                    <td>
                        <asp:Button ID="下載2" runat="server" style="font-family:標楷體;font-size:Large;" Text="下載" OnClick="Download2"></asp:Button>
                        <asp:Label ID="Label1" runat="server" Text="" CssClass="GreenLabel"/>
                        <asp:Label ID="Label2" runat="server" Text="" CssClass="RedLabel"/>
                    </td>
                </tr>
            </table>
        </td>
        <td style="width:550px;text-align:left;">
        </td>
    </table>
    <asp:SqlDataSource ID="SqlDataSource2" runat="server" ConnectionString='<%$ ConnectionStrings:ApplicationServices%>'
        SelectCommand="SELECT DISTINCT CONVERT(varchar(4), YEAR(結帳日期) - 1911) + FORMAT(結帳日期, '/MM/dd') AS 結帳日期 FROM 日報表 ORDER BY 結帳日期 DESC">
    </asp:SqlDataSource>
    <table style="font-size:large;margin:auto;text-align:center;">
        <td style="width:500px;text-align:right;">
        </td>
        <td>
            <table style="font-size:large;margin:auto;text-align:center;">
                <tr>
                    <td>
                    </td>
                    <td style="font-weight:bold;font-size:x-large;">
                        日報表
                    </td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td style="width:113px;text-align:right;">
                    </td>
                    <td runat="server">
                        <asp:DropDownList ID="DropDownList1" runat="server" autopostback="true" style="font-family:標楷體;font-size:Large;height:25px;" Visible="True"></asp:DropDownList>
                    </td>
                    <td>
                        <asp:Button ID="Button2" runat="server" style="font-family:標楷體;font-size:Large;" Text="下載"></asp:Button>
                        <asp:Button ID="Button3" runat="server" style="font-family:標楷體;font-size:Large;" Text="刪除" OnClientClick="return confirm('確定刪除?')"></asp:Button>
                    </td>
                </tr>
            </table>
        </td>
        <td style="width:500px;text-align:left;">
        </td>
    </table>
    <table id="table1" style="margin:auto;text-align:center;">
        <tr>
            <td>
                摘要
            </td>
            <td style="width: 100px;">
                上日結存
            </td>
            <td style="width: 100px;">
                本日收入
            </td>
            <td style="width: 100px;">
                本日支出
            </td>
            <td style="width: 100px;">
                本日結存
            </td>
        </tr>
        <tr>
            <td>
                土地銀行北台中分行077056000014-中分局405專戶
            </td>
            <td>
                <asp:Label ID="LabelC6" runat="server" Visible="True"></asp:Label>
            </td>
            <td>
                <asp:Label ID="LabelD6" runat="server" Visible="True"></asp:Label>
            </td>
            <td>
                <asp:Label ID="LabelE6" runat="server" Visible="True"></asp:Label>
            </td>
            <td>
                <asp:Label ID="LabelF6" runat="server" Visible="True"></asp:Label>
            </td>
        </tr>
        <tr>
            <td>
                中國信託商業銀行台中分行026350002965-中區強制執行409專戶
            </td>
            <td>
                <asp:Label ID="LabelC7" runat="server" Visible="True"></asp:Label>
            </td>
            <td>
                <asp:Label ID="LabelD7" runat="server" Visible="True"></asp:Label>
            </td>
            <td>
                <asp:Label ID="LabelE7" runat="server" Visible="True"></asp:Label>
            </td>
            <td>
                <asp:Label ID="LabelF7" runat="server" Visible="True"></asp:Label>
            </td>
        </tr>
        <tr>
            <td>
            </td>
            <td>
            </td>
            <td>
            </td>
            <td>
            </td>
            <td>
            </td>
        </tr>
        <tr>
            <td>
            </td>
            <td>
            </td>
            <td>
            </td>
            <td>
            </td>
            <td>
            </td>
        </tr>
        <tr>
            <td>
            </td>
            <td>
            </td>
            <td>
            </td>
            <td>
            </td>
            <td>
            </td>
        </tr>
        <tr>
            <td>
                合計
            </td>
            <td>
                <asp:Label ID="LabelC11" runat="server" Visible="True"></asp:Label>
            </td>
            <td>
                <asp:Label ID="LabelD11" runat="server" Visible="True"></asp:Label>
            </td>
            <td>
                <asp:Label ID="LabelE11" runat="server" Visible="True"></asp:Label>
            </td>
            <td>
                <asp:Label ID="LabelF11" runat="server" Visible="True"></asp:Label>
            </td>
        </tr>
        <tr>
            <td>
                　　　　　　　　　　　　收入傳票　　　　　　　　　　由起
            </td>
            <td>
                <asp:Label ID="LabelC12" runat="server" Visible="True"></asp:Label>
            </td>
            <td>
                至
            </td>
            <td>
                <asp:Label ID="LabelE12" runat="server" Visible="True"></asp:Label>
            </td>
            <td>
                號止
            </td>
        </tr>
        <tr>
            <td>
                　　　　　　　　　　　　支出傳票　　　　　　　　　　由起
            </td>
            <td>
                <asp:Label ID="LabelC13" runat="server" Visible="True"></asp:Label>
            </td>
            <td>
                至
            </td>
            <td>
                <asp:Label ID="LabelE13" runat="server" Visible="True"></asp:Label>
            </td>
            <td>
                號止
            </td>
        </tr>
        <tr>
            <td>
                　　　　　　　　　　　　轉帳傳票　　　　　　　　　　由起
            </td>
            <td>
                <asp:Label ID="LabelC14" runat="server" Visible="True"></asp:Label>
            </td>
            <td>
                至
            </td>
            <td>
                <asp:Label ID="LabelE14" runat="server" Visible="True"></asp:Label>
            </td>
            <td>
                號止
            </td>
        </tr>
        <tr>
            <td>
                　　　　　　　　　　　　分錄傳票　　　　　　　　　　由起
            </td>
            <td>
                <asp:Label ID="LabelC15" runat="server" Visible="True"></asp:Label>
            </td>
            <td>
                至
            </td>
            <td>
                <asp:Label ID="LabelE15" runat="server" Visible="True"></asp:Label>
            </td>
            <td>
                號止
            </td>
        </tr>
    </table>
</asp:Content>


