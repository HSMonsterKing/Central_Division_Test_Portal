<%@ Page Title="出納對帳系統" Language="VB" MasterPageFile="./MasterPage.master" MaintainScrollPositionOnPostback="true" AutoEventWireup="false" CodeFile="上傳.aspx.vb" Inherits="上傳" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
    <table style="font-size:large;margin:auto;text-align:center;">
        <tr style="height:30px;">
        </tr>
        <tr>
            <td>
                <asp:FileUpload ID="FileUpload3" runat="server" style="font-size:Large;" Font-Names="標楷體" Width="396px"/>
                <asp:Button ID="Button3" runat="server" style="font-family:標楷體;font-size:Large;" Text="上傳"></asp:Button>
            </td>
        </tr>
        <tr>
            <td>
                此功能會先刪除資料庫裡同一年的傳票資料，並將上傳的檔案存進資料庫，也就是用上傳的檔案取代該年度的所有舊資料，如果上傳的檔案有問題，可能發生錯誤。
            </td>
        </tr>
        <tr style="height:30px;">
        </tr>
        <tr>
            <td>
                <asp:FileUpload ID="FileUpload4" runat="server" style="font-size:Large;" Font-Names="標楷體" Width="396px"/>
                <asp:Button ID="Button4" runat="server" style="font-family:標楷體;font-size:Large;" Text="上傳"></asp:Button>
            </td>
        </tr>
        <tr>
            <td>
                此功能用於上傳傳票送出納檔案中的每筆詳細金額資料，如果上傳的檔案有問題，可能發生錯誤。
            </td>
        </tr>
        <tr style="height:30px;">
        </tr>
        <tr>
            <td>
                <asp:FileUpload ID="FileUpload5" runat="server" style="font-size:Large;" Font-Names="標楷體" Width="396px"/>
                <asp:Button ID="Button5" runat="server" style="font-family:標楷體;font-size:Large;" Text="上傳"></asp:Button>
            </td>
        </tr>
        <tr>
            <td>
                此功能用於上傳客戶基本資料清冊，如果上傳的檔案不對或有問題，可能發生錯誤。
            </td>
        </tr>
        <tr style="height:30px;">
        </tr>
        <tr>
            <td>
                <asp:Button ID="Button6" runat="server" style="font-family:標楷體;font-size:Large;" Text="現金備查簿資料轉換"></asp:Button>
            </td>
        </tr>
        <tr>
            <td>
                此功能用於轉換現金備查簿資料，例如西元轉民國。
            </td>
        </tr>
        <tr>
            <td>
                <asp:Button ID="Button7" runat="server" style="font-family:標楷體;font-size:Large;" Text="現金備查簿資料轉換2"></asp:Button>
            </td>
        </tr>
        <tr>
            <td>
                此功能用於轉換現金備查簿資料。
            </td>
        </tr>
        <tr style="height:30px;">
        </tr>
        <tr>
            <td>
                <asp:FileUpload ID="FileUpload8" runat="server" style="font-size:Large;" Font-Names="標楷體" Width="396px"/>
                <asp:Button ID="Button8" runat="server" style="font-family:標楷體;font-size:Large;" Text="上傳"></asp:Button>
            </td>
        </tr>
        <tr>
            <td>
                此功能用於上傳保管品，如果上傳的檔案不對或有問題，可能發生錯誤。
            </td>
        </tr>
    </table>
    <asp:Label ID="debug" runat="server" style="color:green;" visible="true"></asp:Label>
    <asp:Label ID="debug2" runat="server" style="color:red;" visible="true"></asp:Label>
</asp:Content>


