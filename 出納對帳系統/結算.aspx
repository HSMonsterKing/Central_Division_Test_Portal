<%@ Page Title="出納對帳系統" Language="VB" MasterPageFile="./MasterPage.master" MaintainScrollPositionOnPostback="true" AutoEventWireup="false" CodeFile="結算.aspx.vb" Inherits="結算" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
    <link rel="stylesheet" runat="server" media="screen" href="css\結算.css"/>
    <div style="width:1600px;"></div>
    <table style="font-size:large;margin:auto;text-align:center;">
        <td style="width:500px;text-align:right;">
        </td>
        <td>
            <table style="font-size:large;margin:auto;text-align:center;">
                <tr>
                    <td style="font-weight:bold;font-size:x-large;">
                        結算＆產生日報表
                        <asp:Button ID="Button1" runat="server" style="font-family:標楷體;font-size:Large;color:blue;vertical-align:top;" Text="執行"></asp:Button>
                        <asp:Label ID="debug" runat="server" style="color:green;" visible="true"></asp:Label>
                    </td>
                </tr>
            </table>
        </td>
        <td style="width:500px;text-align:left;">
        </td>
    </table>
    <table style="margin:auto;text-align:center;">
        <tr>
            <td>
                中華民國
            </td>
            <td>
                <asp:DropDownList ID="DropDownList1" runat="server" style="font-size:Large;height:27px;" AutoPostBack="True" Font-Names="標楷體"></asp:DropDownList>
            </td>
            <td>
                年
            </td>
            <td>
                <asp:DropDownList ID="DropDownList2" runat="server" style="font-size:Large;height:27px;" AutoPostBack="True" Font-Names="標楷體"></asp:DropDownList>
            </td>
            <td>
                月
            </td>
            <td>
                <asp:DropDownList ID="DropDownList3" runat="server" style="font-size:Large;height:27px;" AutoPostBack="True" Font-Names="標楷體"></asp:DropDownList>
            </td>
            <td>
                日
            </td>
        </tr>
    </table>
    <table style="margin:auto;text-align:center;">
        <tr>
            <td>
                收入傳票
            </td>
            <td>
                由起
            </td>
            <td>
                <asp:TextBox ID="TextBox1" runat="server" Maxlength=7 style="font-family:新細明體;font-size:large;text-align:left;width:70px;" onkeypress="if(event.keyCode==13){var a=document.getElementById('ContentPlaceHolder1_TextBox3');a.focus();a.selectionStart=a.selectionEnd=a.value.length;return false;}"></asp:TextBox>
            </td>
            <td>
                至
            </td>
            <td>
                <asp:TextBox ID="TextBox2" runat="server" Maxlength=7 style="font-family:新細明體;font-size:large;text-align:left;width:70px;" onkeypress="if(event.keyCode==13){var a=document.getElementById('ContentPlaceHolder1_TextBox4');a.focus();a.selectionStart=a.selectionEnd=a.value.length;return false;}"></asp:TextBox>
            </td>
        </tr>
        <tr>
            <td>
                支出傳票
            </td>
            <td>
                由起
            </td>
            <td>
                <asp:TextBox ID="TextBox3" runat="server" Maxlength=7 style="font-family:新細明體;font-size:large;text-align:left;width:70px;" onkeypress="if(event.keyCode==13){var a=document.getElementById('ContentPlaceHolder1_TextBox5');a.focus();a.selectionStart=a.selectionEnd=a.value.length;return false;}"></asp:TextBox>
            </td>
            <td>
                至
            </td>
            <td>
                <asp:TextBox ID="TextBox4" runat="server" Maxlength=7 style="font-family:新細明體;font-size:large;text-align:left;width:70px;" onkeypress="if(event.keyCode==13){var a=document.getElementById('ContentPlaceHolder1_TextBox6');a.focus();a.selectionStart=a.selectionEnd=a.value.length;return false;}"></asp:TextBox>
            </td>
        </tr>
        <tr>
            <td>
                轉帳傳票
            </td>
            <td>
                由起
            </td>
            <td>
                <asp:TextBox ID="TextBox5" runat="server" Maxlength=7 style="font-family:新細明體;font-size:large;text-align:left;width:70px;" onkeypress="if(event.keyCode==13){var a=document.getElementById('ContentPlaceHolder1_TextBox7');a.focus();a.selectionStart=a.selectionEnd=a.value.length;return false;}"></asp:TextBox>
            </td>
            <td>
                至
            </td>
            <td>
                <asp:TextBox ID="TextBox6" runat="server" Maxlength=7 style="font-family:新細明體;font-size:large;text-align:left;width:70px;" onkeypress="if(event.keyCode==13){var a=document.getElementById('ContentPlaceHolder1_TextBox8');a.focus();a.selectionStart=a.selectionEnd=a.value.length;return false;}"></asp:TextBox>
            </td>
        </tr>
        <tr>
            <td>
                分錄傳票
            </td>
            <td>
                由起
            </td>
            <td>
                <asp:TextBox ID="TextBox7" runat="server" Maxlength=7 style="font-family:新細明體;font-size:large;text-align:left;width:70px;" onkeypress="if(event.keyCode==13){var a=document.getElementById('ContentPlaceHolder1_TextBox2');a.focus();a.selectionStart=a.selectionEnd=a.value.length;return false;}"></asp:TextBox>
            </td>
            <td>
                至
            </td>
            <td>
                <asp:TextBox ID="TextBox8" runat="server" Maxlength=7 style="font-family:新細明體;font-size:large;text-align:left;width:70px;" onkeypress="if(event.keyCode==13){var a=document.getElementById('ContentPlaceHolder1_TextBox1');a.focus();a.selectionStart=a.selectionEnd=a.value.length;return false;}"></asp:TextBox>
            </td>
        </tr>
    </table>
</asp:Content>


