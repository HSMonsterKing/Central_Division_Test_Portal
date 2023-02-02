<%@ Page Language="VB" MasterPageFile="./MasterPage.master" MaintainScrollPositionOnPostback="true" AutoEventWireup="false" CodeFile="查詢.aspx.vb" Inherits="查詢" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
    <asp:ScriptManager ID="ScriptManager1" runat="server" EnablePageMethods = "true" EnableScriptGlobalization="True" EnableScriptLocalization="True"/>
    <script src='js/查詢.js'></script>
    <link rel="stylesheet" runat="server" media="screen" href="css\查詢.css"/>
    <div><h1><a id="Title" href="查詢.aspx">查詢<a></h1></div>
    <asp:Panel ID="Panel1" runat="server" DefaultButton="搜尋" CssClass="Panel1">
        <asp:TextBox ID="ID" runat="server" Maxlength=12 CssClass="Input2" Visible="false"/>
        <asp:TextBox ID="年" runat="server" Maxlength=3 CssClass="Input2"/>年
        單位別<asp:DropDownList ID="單位別" runat="server" AutoPostBack="True" CssClass="DropDownList" OnSelectedIndexChanged="SelectedIndexChanged_尾頁">
            <asp:ListItem Text="" Value=""></asp:ListItem>
            <asp:ListItem Text="中區交通控制中心" Value="中區交通控制中心"></asp:ListItem>
            <asp:ListItem Text="交通管理科" Value="交通管理科"></asp:ListItem>
            <asp:ListItem Text="工務科" Value="工務科"></asp:ListItem>
            <asp:ListItem Text="分局長室" Value="分局長室"></asp:ListItem>
            <asp:ListItem Text="主計室" Value="主計室"></asp:ListItem>
            <asp:ListItem Text="政風室" Value="政風室"></asp:ListItem>
            <asp:ListItem Text="業務科" Value="業務科"></asp:ListItem>
            <asp:ListItem Text="秘書室" Value="秘書室"></asp:ListItem>
            <asp:ListItem Text="機料及保養場" Value="機料及保養場"></asp:ListItem>
            <asp:ListItem Text="人事室" Value="人事室"></asp:ListItem>
            <asp:ListItem Text="勞安科" Value="勞安科"></asp:ListItem>
            <asp:ListItem Text="5工務段" Value="5工務段"></asp:ListItem>
            <asp:ListItem Text="5服務區" Value="5服務區"></asp:ListItem>
        </asp:DropDownList>
        承辦人<asp:DropDownList ID="承辦人" runat="server" Text='<%# Bind("承辦人") %>' AutoPostBack="True" CssClass="DropDownList" OnSelectedIndexChanged="SelectedIndexChanged_尾頁">
            <asp:ListItem Text="" Value=""></asp:ListItem>
            <asp:ListItem Text="陳春綢" Value="陳春綢"></asp:ListItem>
            <asp:ListItem Text="江嘉珊" Value="江嘉珊"></asp:ListItem>
            <asp:ListItem Text="洪孟恬" Value="洪孟恬"></asp:ListItem>
            <asp:ListItem Text="彭金杏" Value="彭金杏"></asp:ListItem>
            <asp:ListItem Text="柯佳妮" Value="柯佳妮"></asp:ListItem>
            <asp:ListItem Text="藍雅燕" Value="藍雅燕"></asp:ListItem>
        </asp:DropDownList>
        <asp:DropDownList ID="月1" runat="server" AutoPostBack="True" CssClass="DropDownList" OnSelectedIndexChanged="SelectedIndexChanged_尾頁">
            <asp:ListItem Text="" Value=""></asp:ListItem>
            <asp:ListItem Text="1" Value="1" Selected="True"></asp:ListItem>
            <asp:ListItem Text="2" Value="2"></asp:ListItem>
            <asp:ListItem Text="3" Value="3"></asp:ListItem>
            <asp:ListItem Text="4" Value="4"></asp:ListItem>
            <asp:ListItem Text="5" Value="5"></asp:ListItem>
            <asp:ListItem Text="6" Value="6"></asp:ListItem>
            <asp:ListItem Text="7" Value="7"></asp:ListItem>
            <asp:ListItem Text="8" Value="8"></asp:ListItem>
            <asp:ListItem Text="9" Value="9"></asp:ListItem>
            <asp:ListItem Text="10" Value="10"></asp:ListItem>
            <asp:ListItem Text="11" Value="11"></asp:ListItem>
            <asp:ListItem Text="12" Value="12"></asp:ListItem>
        </asp:DropDownList>月
        <asp:DropDownList ID="日1" runat="server" AutoPostBack="True" CssClass="DropDownList" OnSelectedIndexChanged="SelectedIndexChanged_尾頁">
            <asp:ListItem Text="1" Value="1" Selected="True"></asp:ListItem>
            <asp:ListItem Text="2" Value="2"></asp:ListItem>
            <asp:ListItem Text="3" Value="3"></asp:ListItem>
            <asp:ListItem Text="4" Value="4"></asp:ListItem>
            <asp:ListItem Text="5" Value="5"></asp:ListItem>
            <asp:ListItem Text="6" Value="6"></asp:ListItem>
            <asp:ListItem Text="7" Value="7"></asp:ListItem>
            <asp:ListItem Text="8" Value="8"></asp:ListItem>
            <asp:ListItem Text="9" Value="9"></asp:ListItem>
            <asp:ListItem Text="10" Value="10"></asp:ListItem>
            <asp:ListItem Text="11" Value="11"></asp:ListItem>
            <asp:ListItem Text="12" Value="12"></asp:ListItem>
            <asp:ListItem Text="13" Value="13"></asp:ListItem>
            <asp:ListItem Text="14" Value="14"></asp:ListItem>
            <asp:ListItem Text="15" Value="15"></asp:ListItem>
            <asp:ListItem Text="16" Value="16"></asp:ListItem>
            <asp:ListItem Text="17" Value="17"></asp:ListItem>
            <asp:ListItem Text="18" Value="18"></asp:ListItem>
            <asp:ListItem Text="19" Value="19"></asp:ListItem>
            <asp:ListItem Text="20" Value="20"></asp:ListItem>
            <asp:ListItem Text="21" Value="21"></asp:ListItem>
            <asp:ListItem Text="22" Value="22"></asp:ListItem>
            <asp:ListItem Text="23" Value="23"></asp:ListItem>
            <asp:ListItem Text="24" Value="24"></asp:ListItem>
            <asp:ListItem Text="25" Value="25"></asp:ListItem>
            <asp:ListItem Text="26" Value="26"></asp:ListItem>
            <asp:ListItem Text="27" Value="27"></asp:ListItem>
            <asp:ListItem Text="28" Value="28"></asp:ListItem>
            <asp:ListItem Text="29" Value="29"></asp:ListItem>
            <asp:ListItem Text="30" Value="30"></asp:ListItem>
            <asp:ListItem Text="31" Value="31"></asp:ListItem>
            </asp:DropDownList>日~
            <asp:DropDownList ID="月2" runat="server" AutoPostBack="True" CssClass="DropDownList" OnSelectedIndexChanged="SelectedIndexChanged_尾頁">
                <asp:ListItem Text="" Value=""></asp:ListItem>
                <asp:ListItem Text="1" Value="1"></asp:ListItem>
                <asp:ListItem Text="2" Value="2"></asp:ListItem>
                <asp:ListItem Text="3" Value="3"></asp:ListItem>
                <asp:ListItem Text="4" Value="4"></asp:ListItem>
                <asp:ListItem Text="5" Value="5"></asp:ListItem>
                <asp:ListItem Text="6" Value="6"></asp:ListItem>
                <asp:ListItem Text="7" Value="7"></asp:ListItem>
                <asp:ListItem Text="8" Value="8"></asp:ListItem>
                <asp:ListItem Text="9" Value="9"></asp:ListItem>
                <asp:ListItem Text="10" Value="10"></asp:ListItem>
                <asp:ListItem Text="11" Value="11"></asp:ListItem>
                <asp:ListItem Text="12" Value="12" Selected="True"></asp:ListItem>
        </asp:DropDownList>月
        <asp:DropDownList ID="日2" runat="server" AutoPostBack="True" CssClass="DropDownList" OnSelectedIndexChanged="SelectedIndexChanged_尾頁">
                    <asp:ListItem Text="1" Value="1"></asp:ListItem>
                    <asp:ListItem Text="2" Value="2"></asp:ListItem>
                    <asp:ListItem Text="3" Value="3"></asp:ListItem>
                    <asp:ListItem Text="4" Value="4"></asp:ListItem>
                    <asp:ListItem Text="5" Value="5"></asp:ListItem>
                    <asp:ListItem Text="6" Value="6"></asp:ListItem>
                    <asp:ListItem Text="7" Value="7"></asp:ListItem>
                    <asp:ListItem Text="8" Value="8"></asp:ListItem>
                    <asp:ListItem Text="9" Value="9"></asp:ListItem>
                    <asp:ListItem Text="10" Value="10"></asp:ListItem>
                    <asp:ListItem Text="11" Value="11"></asp:ListItem>
                    <asp:ListItem Text="12" Value="12"></asp:ListItem>
                    <asp:ListItem Text="13" Value="13"></asp:ListItem>
                    <asp:ListItem Text="14" Value="14"></asp:ListItem>
                    <asp:ListItem Text="15" Value="15"></asp:ListItem>
                    <asp:ListItem Text="16" Value="16"></asp:ListItem>
                    <asp:ListItem Text="17" Value="17"></asp:ListItem>
                    <asp:ListItem Text="18" Value="18"></asp:ListItem>
                    <asp:ListItem Text="19" Value="19"></asp:ListItem>
                    <asp:ListItem Text="20" Value="20"></asp:ListItem>
                    <asp:ListItem Text="21" Value="21"></asp:ListItem>
                    <asp:ListItem Text="22" Value="22"></asp:ListItem>
                    <asp:ListItem Text="23" Value="23"></asp:ListItem>
                    <asp:ListItem Text="24" Value="24"></asp:ListItem>
                    <asp:ListItem Text="25" Value="25"></asp:ListItem>
                    <asp:ListItem Text="26" Value="26"></asp:ListItem>
                    <asp:ListItem Text="27" Value="27"></asp:ListItem>
                    <asp:ListItem Text="28" Value="28"></asp:ListItem>
                    <asp:ListItem Text="29" Value="29"></asp:ListItem>
                    <asp:ListItem Text="30" Value="30"></asp:ListItem>
                    <asp:ListItem Text="31" Value="31" Selected="True"></asp:ListItem>
            </asp:DropDownList>日
        <BR>
        科目<asp:DropDownList ID="科目" runat="server" AutoPostBack="True" CssClass="DropDownList" OnSelectedIndexChanged="SelectedIndexChanged_尾頁">
                    <asp:ListItem Text="" Value=""></asp:ListItem>
                    <asp:ListItem Text="管總-專業服務費" Value="管總-專業服務費"></asp:ListItem>
                    <asp:ListItem Text="管成-規費" Value="管成-規費"></asp:ListItem>
                    <asp:ListItem Text="管總-用品消耗" Value="管總-用品消耗"></asp:ListItem>
                    <asp:ListItem Text="管總-修理保養及保固費" Value="管總-修理保養及保固費"></asp:ListItem>
                    <asp:ListItem Text="維成-交通及運輸設備修護費" Value="維成-交通及運輸設備修護費"></asp:ListItem>
                    <asp:ListItem Text="管總-郵電費" Value="管總-郵電費"></asp:ListItem>
                    <asp:ListItem Text="管總-水電費" Value="管總-水電費"></asp:ListItem>
                    <asp:ListItem Text="管總-公關費" Value="管總-公關費"></asp:ListItem>
                    <asp:ListItem Text="機料及保養場" Value="機料及保養場"></asp:ListItem>
                    <asp:ListItem Text="管總-福利費" Value="管總-福利費"></asp:ListItem>
                    <asp:ListItem Text="管成-用品消耗" Value="管成-用品消耗"></asp:ListItem>
                    <asp:ListItem Text="管總-印刷裝訂及廣告費" Value="管總-印刷裝訂及廣告費"></asp:ListItem>
                    <asp:ListItem Text="維成-工程及管理諮詢服務費" Value="維成-工程及管理諮詢服務費"></asp:ListItem>
        </asp:DropDownList>
        摘要<asp:TextBox ID="摘要" runat="server" Text='' Maxlength=0 Enabled="True" CssClass="Input1"/>
        <ajaxToolkit:AutoCompleteExtender ID="摘要自動" runat="server" TargetControlID="摘要" 
                            CompletionSetCount="10000" CompletionInterval="50" EnableCaching="true" MinimumPrefixLength="1" 
                            ServiceMethod="GetMyList" 
                            CompletionListCssClass="CompletionList" 
                            CompletionListItemCssClass="CompletionListItem" 
                            CompletionListHighlightedItemCssClass="CompletionListHighlightedItem"/>
        商號<asp:TextBox ID="商號" runat="server" Text='' Maxlength=0 Enabled="True" CssClass="Input2"/>
        <ajaxToolkit:AutoCompleteExtender ID="商號自動" runat="server" TargetControlID="商號" 
                            CompletionSetCount="10000" CompletionInterval="50" EnableCaching="true" MinimumPrefixLength="1" 
                            ServiceMethod="GetMyList" 
                            CompletionListCssClass="CompletionList" 
                            CompletionListItemCssClass="CompletionListItem" 
                            CompletionListHighlightedItemCssClass="CompletionListHighlightedItem"/>
        種類<asp:DropDownList ID="種類" runat="server" AutoPostBack="True" CssClass="DropDownList" OnSelectedIndexChanged="種類_SelectedIndexChanged">
                    <asp:ListItem Text="A" Value="A"></asp:ListItem>
                    <asp:ListItem Text="B" Value="B"></asp:ListItem>
                    <asp:ListItem Text="XZ" Value="XZ"></asp:ListItem>
            </asp:DropDownList>
        號數<asp:TextBox ID="號數1" runat="server" OnTextChanged="號數1_TextChanged" Maxlength=3 CssClass="Input2"/>
        ~號數<asp:TextBox ID="號數2" runat="server" OnTextChanged="號數2_TextChanged" Maxlength=3 CssClass="Input2"/>
        審核狀態<asp:DropDownList ID="審核狀態" runat="server" AutoPostBack="True" CssClass="DropDownList" OnSelectedIndexChanged="SelectedIndexChanged_尾頁">
                    <asp:ListItem Text="" Value=""></asp:ListItem>
                    <asp:ListItem Text="未送審" Value="未送審"></asp:ListItem>
                    <asp:ListItem Text="已送審" Value="已送審"></asp:ListItem>
                    <asp:ListItem Text="駁回" Value="駁回"></asp:ListItem>
                    <asp:ListItem Text="通過" Value="通過"></asp:ListItem>
                    <asp:ListItem Text="送交主計室" Value="送交主計室"></asp:ListItem>
                    <asp:ListItem Text="主計室通過" Value="主計室通過"></asp:ListItem>
                    <asp:ListItem Text="被拿回" Value="被拿回"></asp:ListItem>
        </asp:DropDownList>
        <asp:Button ID="搜尋" runat="server" Text="搜尋" CssClass="GreenButton"/>
        <asp:Button ID="清空" runat="server" Text="清空"  OnClick="Clear_Click"  CssClass="GreenButton"/>
        <asp:Label ID="Label1" runat="server" Text="" CssClass="GreenLabel"/>
        <asp:Label ID="Label2" runat="server" Text="" CssClass="RedLabel"/>
    </asp:Panel>
    <asp:Panel ID="Panel2" runat="server" DefaultButton="P2搜尋" CssClass="Panel2">
        <asp:TextBox ID="P2ID" runat="server" Maxlength=12 CssClass="Input2" Visible="false"/>
        動作日期<asp:TextBox ID="日期1" runat="server" OnTextChanged="日期1_TextChanged" CssClass="Input1" />
        ~<asp:TextBox ID="日期2" runat="server" OnTextChanged="日期2_TextChanged" CssClass="Input1" />
        動作<asp:DropDownList ID="動作" runat="server" AutoPostBack="True" CssClass="DropDownList" >
                    <asp:ListItem Text="" Value=""></asp:ListItem>
                    <asp:ListItem Text="新增" Value="新增"></asp:ListItem>
                    <asp:ListItem Text="修改" Value="修改"></asp:ListItem>
                    <asp:ListItem Text="送審" Value="送審"></asp:ListItem>
                    <asp:ListItem Text="駁回" Value="駁回"></asp:ListItem>
                    <asp:ListItem Text="拿回" Value="拿回"></asp:ListItem>
                    <asp:ListItem Text="通過" Value="通過"></asp:ListItem>
                    <asp:ListItem Text="送交主計室" Value="送交主計室"></asp:ListItem>
                    <asp:ListItem Text="主計室駁回" Value="主計室駁回"></asp:ListItem>
                    <asp:ListItem Text="主計室通過" Value="主計室通過"></asp:ListItem>
        </asp:DropDownList>
        <asp:Button ID="P2搜尋" runat="server" Text="簡易搜尋" OnClick="P2_Select_Click" CssClass="GreenButton"/>
        <asp:Label ID="P2Label1" runat="server" Text="" CssClass="GreenLabel"/>
        <asp:Label ID="P2Label2" runat="server" Text="" CssClass="RedLabel"/>
    </asp:Panel>
    <asp:Panel ID="Panel3" runat="server" CssClass="Panel3" >
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="false" CellPadding="1" HorizontalAlign="Center" DataSourceID="SqlDataSource1" DataKeyNames="id" GridLines="None" ShowHeaderWhenEmpty="false" AllowPaging="True" PageSize="15" AllowSorting="True" PagerSettings-PageButtonCount=25 EnableModelValidation="True" AutoGenerateEditButton="false">
            <Columns>
                <asp:TemplateField HeaderText="id" Visible="false">
                    <ItemTemplate>
                        <asp:Label ID="id" runat="server" Text='<%# Eval("id") %>' Maxlength=0 Enabled="False" CssClass="Label id" />
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="列" Visible="False">
                    <ItemTemplate>
                        <asp:Label ID="_列" runat="server" Text='<%# Eval("_列") %>' Maxlength=0 Enabled="False" CssClass="Label _列"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="單位別">
                    <ItemTemplate>
                    <asp:DropDownList ID="單位別" runat="server" Text='<%# Eval("單位別") %>' AutoPostBack="True" CssClass="DropDownList" Enabled="False">
                        <asp:ListItem Text="" Value=""></asp:ListItem>
                        <asp:ListItem Text="中區交通控制中心" Value="中區交通控制中心"></asp:ListItem>
                        <asp:ListItem Text="交通管理科" Value="交通管理科"></asp:ListItem>
                        <asp:ListItem Text="工務科" Value="工務科"></asp:ListItem>
                        <asp:ListItem Text="分局長室" Value="分局長室"></asp:ListItem>
                        <asp:ListItem Text="主計室" Value="主計室"></asp:ListItem>
                        <asp:ListItem Text="政風室" Value="政風室"></asp:ListItem>
                        <asp:ListItem Text="業務科" Value="業務科"></asp:ListItem>
                        <asp:ListItem Text="秘書室" Value="秘書室"></asp:ListItem>
                        <asp:ListItem Text="機料及保養場" Value="機料及保養場"></asp:ListItem>
                        <asp:ListItem Text="人事室" Value="人事室"></asp:ListItem>
                        <asp:ListItem Text="勞安科" Value="勞安科"></asp:ListItem>
                        <asp:ListItem Text="5工務段" Value="5工務段"></asp:ListItem>
                        <asp:ListItem Text="5服務區" Value="5服務區"></asp:ListItem>
                    </asp:DropDownList>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="承辦人">
                    <ItemTemplate>
                    <asp:DropDownList ID="承辦人" runat="server" Text='<%# Eval("承辦人") %>' AutoPostBack="True" CssClass="DropDownList" Enabled="False">
                        <asp:ListItem Text="" Value=""></asp:ListItem>
                        <asp:ListItem Text="陳春綢" Value="陳春綢"></asp:ListItem>
                        <asp:ListItem Text="江嘉珊" Value="江嘉珊"></asp:ListItem>
                        <asp:ListItem Text="洪孟恬" Value="洪孟恬"></asp:ListItem>
                        <asp:ListItem Text="彭金杏" Value="彭金杏"></asp:ListItem>
                        <asp:ListItem Text="柯佳妮" Value="柯佳妮"></asp:ListItem>
                        <asp:ListItem Text="藍雅燕" Value="藍雅燕"></asp:ListItem>
                    </asp:DropDownList>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="月">
                    <ItemTemplate>
                    <asp:DropDownList ID="月" runat="server" Text='<%# Eval("月") %>' AutoPostBack="True" CssClass="DropDownList" Enabled="False">
                        <asp:ListItem Text="" Value=""></asp:ListItem>
                        <asp:ListItem Text="1" Value="1"></asp:ListItem>
                        <asp:ListItem Text="2" Value="2"></asp:ListItem>
                        <asp:ListItem Text="3" Value="3"></asp:ListItem>
                        <asp:ListItem Text="4" Value="4"></asp:ListItem>
                        <asp:ListItem Text="5" Value="5"></asp:ListItem>
                        <asp:ListItem Text="6" Value="6"></asp:ListItem>
                        <asp:ListItem Text="7" Value="7"></asp:ListItem>
                        <asp:ListItem Text="8" Value="8"></asp:ListItem>
                        <asp:ListItem Text="9" Value="9"></asp:ListItem>
                        <asp:ListItem Text="10" Value="10"></asp:ListItem>
                        <asp:ListItem Text="11" Value="11"></asp:ListItem>
                        <asp:ListItem Text="12" Value="12"></asp:ListItem>
                    </asp:DropDownList>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="日">
                    <ItemTemplate>
                        <asp:DropDownList ID="日" runat="server" Text='<%# Eval("日") %>' AutoPostBack="True" CssClass="DropDownList" Enabled="False">
                        <asp:ListItem Text="" Value=""></asp:ListItem>
                        <asp:ListItem Text="1" Value="1"></asp:ListItem>
                        <asp:ListItem Text="2" Value="2"></asp:ListItem>
                        <asp:ListItem Text="3" Value="3"></asp:ListItem>
                        <asp:ListItem Text="4" Value="4"></asp:ListItem>
                        <asp:ListItem Text="5" Value="5"></asp:ListItem>
                        <asp:ListItem Text="6" Value="6"></asp:ListItem>
                        <asp:ListItem Text="7" Value="7"></asp:ListItem>
                        <asp:ListItem Text="8" Value="8"></asp:ListItem>
                        <asp:ListItem Text="9" Value="9"></asp:ListItem>
                        <asp:ListItem Text="10" Value="10"></asp:ListItem>
                        <asp:ListItem Text="11" Value="11"></asp:ListItem>
                        <asp:ListItem Text="12" Value="12"></asp:ListItem>
                        <asp:ListItem Text="13" Value="13"></asp:ListItem>
                        <asp:ListItem Text="14" Value="14"></asp:ListItem>
                        <asp:ListItem Text="15" Value="15"></asp:ListItem>
                        <asp:ListItem Text="16" Value="16"></asp:ListItem>
                        <asp:ListItem Text="17" Value="17"></asp:ListItem>
                        <asp:ListItem Text="18" Value="18"></asp:ListItem>
                        <asp:ListItem Text="19" Value="19"></asp:ListItem>
                        <asp:ListItem Text="20" Value="20"></asp:ListItem>
                        <asp:ListItem Text="21" Value="21"></asp:ListItem>
                        <asp:ListItem Text="22" Value="22"></asp:ListItem>
                        <asp:ListItem Text="23" Value="23"></asp:ListItem>
                        <asp:ListItem Text="24" Value="24"></asp:ListItem>
                        <asp:ListItem Text="25" Value="25"></asp:ListItem>
                        <asp:ListItem Text="26" Value="26"></asp:ListItem>
                        <asp:ListItem Text="27" Value="27"></asp:ListItem>
                        <asp:ListItem Text="28" Value="28"></asp:ListItem>
                        <asp:ListItem Text="29" Value="29"></asp:ListItem>
                        <asp:ListItem Text="30" Value="30"></asp:ListItem>
                        <asp:ListItem Text="31" Value="31"></asp:ListItem>
                        </asp:DropDownList>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="科目">
                    <ItemTemplate>
                    <asp:DropDownList ID="科目" runat="server" AutoPostBack="True" CssClass="DropDownList" Enabled="False">
                        <asp:ListItem Text="" Value=""></asp:ListItem>
                        <asp:ListItem Text="管總-專業服務費" Value="管總-專業服務費"></asp:ListItem>
                        <asp:ListItem Text="管成-規費" Value="管成-規費"></asp:ListItem>
                        <asp:ListItem Text="管總-用品消耗" Value="管總-用品消耗"></asp:ListItem>
                        <asp:ListItem Text="管總-修理保養及保固費" Value="管總-修理保養及保固費"></asp:ListItem>
                        <asp:ListItem Text="管成-修理保養及保固費" Value="管成-修理保養及保固費"></asp:ListItem>
                        <asp:ListItem Text="維成-交通及運輸設備修護費" Value="維成-交通及運輸設備修護費"></asp:ListItem>
                        <asp:ListItem Text="管總-郵電費" Value="管總-郵電費"></asp:ListItem>
                        <asp:ListItem Text="管總-水電費" Value="管總-水電費"></asp:ListItem>
                        <asp:ListItem Text="管總-公關費" Value="管總-公關費"></asp:ListItem>
                        <asp:ListItem Text="機料及保養場" Value="機料及保養場"></asp:ListItem>
                        <asp:ListItem Text="管總-福利費" Value="管總-福利費"></asp:ListItem>
                        <asp:ListItem Text="管成-用品消耗" Value="管成-用品消耗"></asp:ListItem>
                        <asp:ListItem Text="管總-印刷裝訂及廣告費" Value="管總-印刷裝訂及廣告費"></asp:ListItem>
                        <asp:ListItem Text="維成-工程及管理諮詢服務費" Value="維成-工程及管理諮詢服務費"></asp:ListItem>
                    </asp:DropDownList>
                    <asp:DropDownList ID="科目2" runat="server" AutoPostBack="True" CssClass="DropDownList" Enabled="False" visible='<%# If (Eval("科目2").ToString = "", "False", "True") %>'>
                        <asp:ListItem Text="" Value=""></asp:ListItem>
                        <asp:ListItem Text="管總-專業服務費" Value="管總-專業服務費"></asp:ListItem>
                        <asp:ListItem Text="管成-規費" Value="管成-規費"></asp:ListItem>
                        <asp:ListItem Text="管總-用品消耗" Value="管總-用品消耗"></asp:ListItem>
                        <asp:ListItem Text="管總-修理保養及保固費" Value="管總-修理保養及保固費"></asp:ListItem>
                        <asp:ListItem Text="管成-修理保養及保固費" Value="管成-修理保養及保固費"></asp:ListItem>
                        <asp:ListItem Text="維成-交通及運輸設備修護費" Value="維成-交通及運輸設備修護費"></asp:ListItem>
                        <asp:ListItem Text="管總-郵電費" Value="管總-郵電費"></asp:ListItem>
                        <asp:ListItem Text="管總-水電費" Value="管總-水電費"></asp:ListItem>
                        <asp:ListItem Text="管總-公關費" Value="管總-公關費"></asp:ListItem>
                        <asp:ListItem Text="機料及保養場" Value="機料及保養場"></asp:ListItem>
                        <asp:ListItem Text="管總-福利費" Value="管總-福利費"></asp:ListItem>
                        <asp:ListItem Text="管成-用品消耗" Value="管成-用品消耗"></asp:ListItem>
                        <asp:ListItem Text="管總-印刷裝訂及廣告費" Value="管總-印刷裝訂及廣告費"></asp:ListItem>
                        <asp:ListItem Text="維成-工程及管理諮詢服務費" Value="維成-工程及管理諮詢服務費"></asp:ListItem>
                    </asp:DropDownList>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="摘要">
                    <ItemTemplate>
                        <asp:Label ID="摘要" runat="server" Text='<%# Eval("摘要") %>' Maxlength=0 TextMode="MultiLine" Enabled="False" CssClass="Label 摘要"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="姓名">
                    <ItemTemplate>
                        <asp:ImageButton ID="姓名" runat="server" ImageUrl='<%# If (Eval("姓名").ToString = "", "", Eval("姓名")) %>' AutoPostBack="False" OnClientClick="Alert('123456');return false;" Visible='<%# If (Eval("姓名").ToString = "", "False", "True") %>' CssClass="TextBox 姓名"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="商號">
                    <ItemTemplate>
                        <asp:Label ID="商號" runat="server" Text='<%# Eval("商號") %>' Maxlength=0 TextMode="MultiLine" Enabled="False" CssClass="Label 商號"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="經手人">
                    <ItemTemplate>
                        <asp:ImageButton ID="經手人" runat="server" ImageUrl='<%# If (Eval("經手人").ToString = "", "", Eval("經手人")) %>' AutoPostBack="False" OnClientClick="Alert('123456');return false;" Visible='<%# If (Eval("經手人").ToString = "", "False", "True") %>' CssClass="TextBox 經手人"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="種類">
                    <ItemTemplate>
                        <asp:Label ID="種類" runat="server" Text='<%# Eval("種類") %>' Maxlength=0 Enabled="False" CssClass="Label 種類"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="號數">
                    <ItemTemplate>
                        <asp:Label ID="號數" runat="server" Text='<%# Eval("號數", "{0:000}") %>' Maxlength=0 Enabled="False" CssClass="Label 號數"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="收入">
                    <ItemTemplate>
                        <asp:Label ID="收入" runat="server" Text='<%# Eval("收入", "{0:c0}") %>' Maxlength=0 Enabled="False" CssClass="Label 收入"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="支出">
                    <ItemTemplate>
                        <asp:Label ID="支出" runat="server" Text='<%# Eval("支出", "{0:c0}") %>' Maxlength=0 Enabled="False" CssClass="Label 支出"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="餘額">
                    <ItemTemplate>
                        <asp:Label ID="餘額" runat="server" Text='<%# Eval("餘額", "{0:c0}") %>' Maxlength=0 Enabled="False" CssClass="Label 餘額"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="備註">
                    <ItemTemplate>
                        <asp:Label ID="備註" runat="server" Text='<%# Eval("備註") %>' Maxlength=0 Enabled="False" CssClass="Label 備註"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="狀態">
                    <ItemTemplate>
                    <asp:Label ID="審核狀態" runat="server"  Text='<%# 
                    If (Eval("鎖定").ToString = "True", 
                    If (Eval("過審").ToString = "True", 
                    If (Eval("回覆").ToString = "True", "主計室通過", 
                    If (Eval("送交主計室日期").ToString="", "通過", "送交主計室")),"已送審"), 
                    If (Eval("送出").ToString = "True", 
                    If (Eval("駁回原因").ToString = "拿回", "被拿回", "駁回"), "未送審"))%>' Maxlength=0 CssClass="Label 審核狀態"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="預支日期">
                    <ItemTemplate>
                        <asp:Label ID="預支日期" runat="server" Text='<%# If (IsDate(Eval("預支日期")), (Year(Eval("預支日期"))-1911).ToString() & Eval("預支日期", "{0:/MM/dd}"), "") %>' Maxlength=0 Enabled="False" CssClass="Label 預支日期"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="歸還日期">
                    <ItemTemplate>
                        <asp:Label ID="歸還日期" runat="server" Text='<%# If (IsDate(Eval("歸還日期")), (Year(Eval("歸還日期"))-1911).ToString() & Eval("歸還日期", "{0:/MM/dd}"), "") %>' Maxlength=0 Enabled="False" CssClass="Label 歸還日期"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="駁回原因" >
                    <ItemTemplate>
                        <asp:Label ID="駁回原因" runat="server" Text='<%# Eval("駁回原因") %>' Maxlength=0 Enabled="False" CssClass="Label 駁回原因"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="流程" >
                    <ItemTemplate>
                        <asp:Button ID="流程" runat="server" Text="該筆資料流程"  CommandName="流程" CssClass="GreenButton" />
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
            </Columns>
            <HeaderStyle BackColor="Green" Font-Bold="True" ForeColor="White" CssClass="Header"/>
            <RowStyle BackColor="White" CssClass="Row"/>
            <AlternatingRowStyle/>
            <SelectedRowStyle/>
            <EditRowStyle/>
            <PagerStyle BackColor="Green" HorizontalAlign="Center" CssClass="Pager"/>
            <FooterStyle/>
            <PagerSettings  Mode="NumericFirstLast" FirstPageText="<<" PreviousPageText="<" NextPageText=">" LastPageText=">>" />
        </asp:GridView>
    </asp:Panel>
    <asp:Panel ID="Panel4" runat="server" CssClass="Panel4" Visible="False" >
        <asp:Button ID="返回" runat="server" Text="返回"  OnClick="返回_Click" CssClass="GreenButton" />
        <asp:GridView ID="GridView2" runat="server" AutoGenerateColumns="false" CellPadding="1" HorizontalAlign="Center" DataSourceID="SqlDataSource2" DataKeyNames="id" GridLines="None" ShowHeaderWhenEmpty="false" AllowPaging="True" PageSize="20" AllowSorting="True" PagerSettings-PageButtonCount=25 EnableModelValidation="True" AutoGenerateEditButton="false">
            <Columns>
                <asp:TemplateField HeaderText="資料id" Visible="false">
                    <ItemTemplate>
                        <asp:Label ID="id" runat="server" Text='<%# Eval("ID") %>' Maxlength=0 Enabled="true" CssClass="Label id"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="摘要">
                    <ItemTemplate>
                        <asp:Label ID="摘要" runat="server" Text='<%# Eval("摘要") %>' Maxlength=0 Enabled="true" CssClass="Label 摘要"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="日期">
                    <ItemTemplate>
                        <asp:Label ID="日期" runat="server" Text='<%# If (IsDate(Eval("日期2")), (Year(Eval("日期2"))-1911).ToString() & Eval("日期2", "{0:/MM/dd HH:mm:ss}"), "") %>' Maxlength=0 Enabled="true" CssClass="Label 日期"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="審核" >
                    <ItemTemplate>
                        <asp:Label ID="動作" runat="server" Text='<%# Eval("動作") %>' Maxlength=0 Enabled="False" CssClass="Label 動作"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="命令" >
                    <ItemTemplate>
                        <asp:Label ID="命令" runat="server" Text='<%# Eval("命令") %>' Maxlength=0 Enabled="true" CssClass="Label 命令"/>
                        <asp:Button ID="查詢" runat="server" Text="查詢該筆資料"  CommandName="查詢" Visible='<%# If(Eval("命令").ToString="",0,1) %>' CssClass="GreenButton" />
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="核章" >
                    <ItemTemplate>
                    <asp:ImageButton ID="主任核章" runat="server" ImageUrl='<%# 
                    If (Eval("動作").ToString = "駁回" or Eval("動作").ToString = "通過", "./image/png/主任職章.png",
                    If (Eval("動作").ToString = "主計室通過" or Eval("動作").ToString = "主計室駁回" ,
                    If (Eval("簽章").ToString = "2808", "./image/png/丁燕雪.png",
                    If (Eval("簽章").ToString = "2897", "./image/png/林容如.png",
                    If (Eval("簽章").ToString = "2808_1", "./image/png/主計室職章1.png", 
                    If (Eval("簽章").ToString = "2897_1", "./image/png/主計室職章2.png", ""
                    )))) ,"")
                    )%>' AutoPostBack="False" OnClientClick="Alert('123456');return false;" Visible='<%# If (Eval("動作").ToString = "駁回" Or Eval("動作").ToString = "通過" Or Eval("動作").ToString = "主計室通過" Or Eval("動作").ToString = "主計室駁回", "True", "False") %>' CssClass="TextBox 主任核章"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
            </Columns>
            <HeaderStyle BackColor="Green" Font-Bold="True" ForeColor="White" CssClass="Header"/>
            <RowStyle BackColor="White" CssClass="Row"/>
            <AlternatingRowStyle/>
            <SelectedRowStyle/>
            <EditRowStyle/>
            <PagerStyle BackColor="Green" HorizontalAlign="Center" CssClass="Pager"/>
            <FooterStyle/>
            <PagerSettings  Mode="NumericFirstLast" FirstPageText="<<" PreviousPageText="<" NextPageText=">" LastPageText=">>" />
        </asp:GridView>
    </asp:Panel>
    <asp:Panel ID="Panel5" runat="server" CssClass="Panel5" Visible="False" >
        <asp:Button ID="返回2" runat="server" Text="返回"  OnClick="返回2_Click" CssClass="GreenButton" />
    <table  style="border-collapse:collapse; border:1px solid black;">
    <tr >
        <th style="border:1px solid black;" bgcolor="Green" align='center' width="100" ><asp:Label ID="Label2_1" runat="server" Text='ID' Maxlength=0 Enabled="true" CssClass="Label" ForeColor="White"/></th>
        <th style="border:1px solid black;" bgcolor="Green" align='center' width="800" ><asp:Label ID="Label2_2" runat="server" Text='修改前' Maxlength=0 Enabled="true" CssClass="Label" ForeColor="White"/></th> 
        <th style="border:1px solid black;" bgcolor="Green" align='center' width="800" ><asp:Label ID="Label2_3" runat="server" Text='修改後' Maxlength=0 Enabled="true" CssClass="Label" ForeColor="White"/></th> 
    </tr>
    <tr>
        <td style="border:1px solid black;" bgcolor=#F0F0F0 align='center'><asp:Label ID="I_id" runat="server" Text='' Maxlength=0 Enabled="False" CssClass="Label id"/></td>
        <td style="border:1px solid black;" bgcolor=#F0F0F0 align='center'><asp:Label ID="修改前" runat="server" Text='' Maxlength=0 Enabled="true" CssClass="Label"/>
        <asp:Image ID="姓名前" runat="server"  Enabled="true" CssClass="Label 姓名" ForeColor="White"/>
        <asp:Label ID="修改前1" runat="server" Text='' Maxlength=0 Enabled="true" CssClass="Label"/>
        <asp:Image ID="經手人前" runat="server" Enabled="true" CssClass="Label 經手人" ForeColor="White"/>
        <asp:Label ID="修改前2" runat="server" Text='' Maxlength=0 Enabled="true" CssClass="Label"/></td> 
        <td style="border:1px solid black;" bgcolor=#F0F0F0 align='center'><asp:Label ID="修改後" runat="server" Text='' Maxlength=0 Enabled="true" CssClass="Label"/>
        <asp:Image ID="姓名後" runat="server" Enabled="true" CssClass="Label 姓名" ForeColor="White"/>
        <asp:Label ID="修改後1" runat="server" Text='' Maxlength=0 Enabled="true" CssClass="Label"/>
        <asp:Image ID="經手人後" runat="server" Enabled="true" CssClass="Label 經手人" ForeColor="White"/>
        <asp:Label ID="修改後2" runat="server" Text='' Maxlength=0 Enabled="true" CssClass="Label"/></td> 
    </tr>
    </table>
    </asp:Panel>
    <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString='<%$ ConnectionStrings:ApplicationServices%>'
        SelectCommand="
            SELECT * FROM 收支備查簿
            WHERE 取號 = 0
            AND (''=TRIM(@單位別) OR 單位別 LIKE N'%'+TRIM(@單位別)+'%')
            And (''=TRIM(@承辦人) OR 承辦人 LIKE N'%'+TRIM(@承辦人)+'%')
            AND (''=TRIM(@科目) OR 科目 LIKE N'%'+TRIM(@科目)+'%' OR 科目2 LIKE N'%'+TRIM(@科目)+'%')
            AND (''=TRIM(@摘要) OR 摘要 LIKE N'%'+TRIM(@摘要)+'%')
            AND (''=TRIM(@商號) OR 商號 LIKE N'%'+TRIM(@商號)+'%')
            AND _種類 = @種類
            AND ((''=TRIM(@號數1) OR ''=TRIM(@號數2))
            OR ( 號數 BETWEEN 
            SUBSTRING(TRIM(@號數1), PATINDEX('%[^0]%', TRIM(@號數1)), 3) AND 
            SUBSTRING(TRIM(@號數2), PATINDEX('%[^0]%', TRIM(@號數2)), 3)))
            AND (TRIM(@年)='' OR (CONVERT(date,STR(年+1911)+STR(月)+STR(日)) 
            BETWEEN CONVERT(date,STR(TRIM(@年)+1911)+STR(ISNULL(NULLIF(TRIM(@月1),''),'1'))+STR(ISNULL(NULLIF(TRIM(@日1),''),'1')))
            AND CONVERT(date,STR(@年+1911)+STR(ISNULL(NULLIF(TRIM(@月2),''),'12'))+STR(ISNULL(NULLIF(TRIM(@日2),''),STR(Day(EOMONTH(STR(TRIM(@年)+1911)+'/'+STR(ISNULL(NULLIF(TRIM(@月2),''),'12'))+'/01')))))))
            OR (Not(''=TRIM(@單位別)) OR Not(''=TRIM(@承辦人)) OR Not(''=TRIM(@科目)) OR Not(''=TRIM(@摘要)) OR Not(''=TRIM(@商號)) OR Not(''=TRIM(@號數1)) OR Not(''=TRIM(@號數2)) OR Not(''=TRIM(@審核狀態))) 
			AND (TRIM(@月1)='' OR TRIM(@日1)='' OR TRIM(@月2)='' OR TRIM(@日2)=''))
            AND (''=TRIM(@審核狀態) OR
            IIf (鎖定='True', IIf (過審='True', IIf (回覆 = 'True', '主計室通過', IIf (送交主計室日期 IS NULL, '通過', '送交主計室')), '已送審'), IIf (送出 = 'True', IIf (駁回原因='拿回', '被拿回', '駁回'), '未送審')) LIKE TRIM(@審核狀態))
            ORDER BY _頁, _列"
        Insertcommand="" 
        UpdateCommand=""
        DeleteCommand="">
        <SelectParameters>
            <asp:ControlParameter ControlID="年" ConvertEmptyStringToNull="False" Name="年" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="單位別" ConvertEmptyStringToNull="False" Name="單位別" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="承辦人" ConvertEmptyStringToNull="False" Name="承辦人" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="月1" ConvertEmptyStringToNull="False" Name="月1" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="日1" ConvertEmptyStringToNull="False" Name="日1" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="月2" ConvertEmptyStringToNull="False" Name="月2" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="日2" ConvertEmptyStringToNull="False" Name="日2" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="科目" ConvertEmptyStringToNull="False" Name="科目" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="摘要" ConvertEmptyStringToNull="False" Name="摘要" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="商號" ConvertEmptyStringToNull="False" Name="商號" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="種類" ConvertEmptyStringToNull="False" Name="種類" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="號數1" ConvertEmptyStringToNull="False" Name="號數1" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="號數2" ConvertEmptyStringToNull="False" Name="號數2" PropertyName="Text" Type="String"/> 
            <asp:ControlParameter ControlID="審核狀態" ConvertEmptyStringToNull="False" Name="審核狀態" PropertyName="Text" Type="String"/>
        </SelectParameters>
        <InsertParameters>
        </InsertParameters>
        <UpdateParameters>
        </UpdateParameters>
        <DeleteParameters>
            <asp:Parameter Name="id"/>
        </DeleteParameters>
    </asp:SqlDataSource>
    <asp:SqlDataSource ID="SqlDataSource2" runat="server" ConnectionString='<%$ ConnectionStrings:ApplicationServices%>'
        SelectCommand="SELECT * FROM 日誌,收支備查簿
        Where ((''=TRIM(@id) OR 日誌.id LIKE N'%'+TRIM(@id)+'%')
        And 日誌.id=收支備查簿.id)
        AND (''=TRIM(@動作) OR 動作 LIKE N'%'+TRIM(@動作)+'%')
        AND ((''=TRIM(@日期1) OR ''=TRIM(@日期2))
            OR (日期2 BETWEEN 
            CONVERT(date,(REPLACE(REPLACE(REPLACE(STR(left(@日期1,3)+1911)+substring(@日期1,4,10), '.', ''), '/', ''), ' ', '')))
            AND 
			Dateadd(Day,1,CONVERT(date,(REPLACE(REPLACE(REPLACE(STR(left(@日期2,3)+1911)+substring(@日期2,4,10), '.', ''), '/', ''), ' ', ''))))))
        ORDER BY 日期2 desc"
        Insertcommand="" 
        UpdateCommand=""
        DeleteCommand="">
        <SelectParameters>
            <asp:ControlParameter ControlID="ID" ConvertEmptyStringToNull="False" Name="ID" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="動作" ConvertEmptyStringToNull="False" Name="動作" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="日期1" ConvertEmptyStringToNull="False" Name="日期1" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="日期2" ConvertEmptyStringToNull="False" Name="日期2" PropertyName="Text" Type="String"/>
        </SelectParameters>
        <InsertParameters>
        </InsertParameters>
        <UpdateParameters>
        </UpdateParameters>
        <DeleteParameters>
        </DeleteParameters>
    </asp:SqlDataSource>
</asp:Content>
