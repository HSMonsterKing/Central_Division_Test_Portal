<%@ Page Language="VB" MasterPageFile="./MasterPage.master" MaintainScrollPositionOnPostback="true" AutoEventWireup="false" CodeFile="收支備查簿.aspx.vb" Inherits="收支備查簿"%>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server" >
    <asp:ScriptManager ID="ScriptManager1" runat="server" EnablePageMethods = "true" EnableScriptGlobalization="True" EnableScriptLocalization="True"/>
    <script src='js/收支備查簿.js'></script>
    <link rel="stylesheet" runat="server" media="screen" href="css\收支備查簿.css"/>
    <style>
    </style>
    <div><h1><a id="Title" href="收支備查簿.aspx">收支備查簿<a></h1></div>
    <asp:Panel ID="Panel1" runat="server" DefaultButton="搜尋" CssClass="Panel1" >
        年<asp:TextBox ID="年" runat="server" Maxlength=3 CssClass="Input2"/>
        種類<asp:DropDownList ID="_種類" runat="server" AutoPostBack="True" CssClass="DropDownList" OnSelectedIndexChanged="種類_SelectedIndexChanged">
            <asp:ListItem Text="A" Value="A"></asp:ListItem>
            <asp:ListItem Text="B" Value="B"></asp:ListItem>
            <asp:ListItem Text="XZ" Value="XZ"></asp:ListItem>
        </asp:DropDownList>
        <asp:TextBox ID="頁1" runat="server" OnTextChanged="頁1_TextChanged" Maxlength=3 CssClass="Input2"/>頁
        ~<asp:TextBox ID="頁2" runat="server" OnTextChanged="頁2_TextChanged" Maxlength=3 CssClass="Input2"/>頁
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
        日期<asp:DropDownList ID="月1" runat="server" AutoPostBack="True" CssClass="DropDownList" OnSelectedIndexChanged="SelectedIndexChanged_尾頁">
            <asp:ListItem Text="" Value="" Selected="True"></asp:ListItem>
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
        </asp:DropDownList>月
        <asp:DropDownList ID="日1" runat="server" AutoPostBack="True" CssClass="DropDownList" OnSelectedIndexChanged="SelectedIndexChanged_尾頁">
            <asp:ListItem Text="" Value="" Selected="True"></asp:ListItem>
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
            </asp:DropDownList>日~
            <asp:DropDownList ID="月2" runat="server" AutoPostBack="True" CssClass="DropDownList" OnSelectedIndexChanged="SelectedIndexChanged_尾頁">
                <asp:ListItem Text="" Value="" Selected="True"></asp:ListItem>
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
        </asp:DropDownList>月
        <asp:DropDownList ID="日2" runat="server" AutoPostBack="True" CssClass="DropDownList" OnSelectedIndexChanged="SelectedIndexChanged_尾頁">
                    <asp:ListItem Text="" Value="" Selected="True"></asp:ListItem>
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
                    <asp:ListItem Text="拿回" Value="拿回"></asp:ListItem>
        </asp:DropDownList>
        <BR>
        <asp:Button ID="搜尋" runat="server" Text="搜尋" CssClass="GreenButton"/>
        <asp:Button ID="清空" runat="server" Text="清空" OnClick="Clear_Click" CssClass="GreenButton"/>
        <asp:Button ID="新增" runat="server" Text="新增一頁" OnClick="Insert" CssClass="GreenButton"/>
        <asp:Button ID="存檔" runat="server" Text="存檔" OnClick="Update" CssClass="GreenButton"/>
        <asp:Button ID="下載" runat="server" Text="下載" OnClick="Download" CssClass="GreenButton"/>
        <asp:Button ID="刪除" runat="server" Text="刪除末頁" OnClick="Delete" OnClientClick="return confirm('確定刪除?')" CssClass="RedButton"/>
        <asp:Button ID="測試" runat="server" Text="測試" OnClick="test" CssClass="GreenButton" Visible="False"/>
        <asp:Label ID="Label1" runat="server" Text="" CssClass="GreenLabel"/>
        <asp:Label ID="Label2" runat="server" Text="" CssClass="RedLabel"/>
    </asp:Panel>
  <meta charset="utf-8">
  <title>Signature Pad demo</title>
  <meta name="description" content="Signature Pad - HTML5 canvas based smooth signature drawing using variable width spline interpolation.">
  <meta name="viewport" content="width=device-width, initial-scale=1, minimum-scale=1, maximum-scale=1, user-scalable=no">
  <meta name="apple-mobile-web-app-capable" content="yes">
  <meta name="apple-mobile-web-app-status-bar-style" content="black">
  <script type="text/javascript">
    var _gaq = _gaq || [];
    _gaq.push(['_setAccount', 'UA-39365077-1']);
    _gaq.push(['_trackPageview']);
    (function() {
      var ga = document.createElement('script'); ga.type = 'text/javascript'; ga.async = true;
      ga.src = ('https:' == document.location.protocol ? 'https://ssl' : 'http://www') + '.google-analytics.com/ga.js';
      var s = document.getElementsByTagName('script')[0]; s.parentNode.insertBefore(ga, s);
    })();
  </script>
  <div id="signature-pad" class="signature-pad">
    <div class="signature-pad--body">
      <canvas></canvas>
    </div>
    <div class="signature-pad--footer">
      <div class="description">Sign above</div>
      <div class="signature-pad--actions">
        <div>
          <button type="button" class="button clear" data-action="clear">Clear</button>
          <button type="button" class="button" data-action="change-color">Change color</button>
          <button type="button" class="button" data-action="undo">Undo</button>
        </div>
        <div>
          <button type="button" class="button save" data-action="save-png">簽完名按這後點選要貼上的位置</button>
        </div>
      </div>
    </div>
  </div>
  <script src="/js/s.js"></script>
  <script src="/js/a.js"></script>
  <input id="dataURL" name="dataURL"   />
  <asp:Label ID="Label3" runat="server" Text="" CssClass="RedLabel"/>

    <asp:Panel ID="Panel3" runat="server" CssClass="Panel3" DefaultButton="存檔">
        <asp:Button ID="插入" runat="server" Text="插入資料" OnClick="插入_Click" OnClientClick="return confirm('請先確定是否有選取資料，將插入至資料前一筆，確定插入?')" CssClass="RedButton" style="float:right"/>
        <asp:Button ID="收回" runat="server" Text="拿回單筆號數資料" OnClick="收回_Click" OnClientClick="return confirm('請先確定是否有選取要拿回號數的資料，確定拿回?')" CssClass="RedButton" style="float:right"/>
        <asp:Button ID="刪除該列" runat="server" Text="刪除該列" OnClick="DeleteRow" OnClientClick="return confirm('請先確定是否有選取要刪的資料並將剩餘資料往上遞補，確定刪除?')" CssClass="RedButton" style="float:right"/>
        <asp:Button ID="修改已過審" runat="server" Text="修改主計室通過及送審之資料" OnClick="修改已過審_Click" OnClientClick="return confirm('請先確定是否有選取要修改的資料，確定修改?')" CssClass="RedButton" style="float:right"/>
        <asp:Button ID="交換" runat="server" Text="　交換　" OnClick="交換_Click" OnClientClick="return confirm('請先選取要交換的2筆資料，確定交換?')" CssClass="RedButton" style="float:right"/>
        <asp:Button ID="重置該列" runat="server" Text="重置該列" OnClick="ReSetRow" OnClientClick="return confirm('請先確定是否有選取要重置的資料，確定重置?')" CssClass="RedButton" style="float:right"/>
        <asp:Button ID="送審" runat="server" Text="　送審　" OnClick="SendToDirector" CssClass="GreenButton" style="float:right"/>
        <asp:Button ID="新增科目" runat="server" Text="新增科目" OnClick="新增科目_Click" CssClass="GreenButton" style="float:right"/>
        <asp:Button ID="取號" runat="server" Text="　取號　" OnClick="取號_Click" CssClass="GreenButton" style="float:right"/>
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="false" CellPadding="1" HorizontalAlign="Center" DataSourceID="SqlDataSource1" DataKeyNames="id" GridLines="None" ShowHeaderWhenEmpty="false" AllowPaging="True" PageSize="15" AllowSorting="True" PagerSettings-PageButtonCount=25 EnableModelValidation="True" AutoGenerateEditButton="false" >
            <Columns>
                <asp:TemplateField HeaderText="id" Visible="false">
                    <ItemTemplate>
                        <asp:TextBox ID="id" runat="server" Text='<%# Bind("id") %>' Maxlength=0 Enabled="False" CssClass="TextBox id"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="">
                    <ItemTemplate>
                        <asp:TextBox ID="_列" runat="server" Text='<%# Bind("_列") %>' Maxlength=0 Enabled="False" CssClass="TextBox _列"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="單位別">
                    <ItemTemplate>
                    <asp:DropDownList ID="單位別" runat="server" Text='<%# Bind("單位別") %>' AutoPostBack="True" CssClass="DropDownList">
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
                    <asp:DropDownList ID="承辦人" runat="server" Text='<%# Bind("承辦人") %>' AutoPostBack="True" CssClass="DropDownList">
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
                    <asp:DropDownList ID="月" runat="server" Text='<%# Bind("月") %>' AutoPostBack="True" CssClass="DropDownList" OnSelectedIndexChanged="月_SelectedIndexChanged">
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
                    <asp:DropDownList ID="日" runat="server" Text='<%# Bind("日") %>' AutoPostBack="True" CssClass="DropDownList" OnSelectedIndexChanged="日_SelectedIndexChanged">
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
                <asp:TemplateField HeaderText="科目"><ItemTemplate>
                    <asp:DropDownList ID="科目" runat="server" AutoPostBack="True" CssClass="DropDownList">
                        <asp:ListItem Text="" Value=""></asp:ListItem>
                        <asp:ListItem Text="管總-專業服務費" Value="管總-專業服務費"></asp:ListItem>
                        <asp:ListItem Text="管成-規費" Value="管成-規費"></asp:ListItem>
                        <asp:ListItem Text="管總-用品消耗" Value="管總-用品消耗"></asp:ListItem>
                        <asp:ListItem Text="管總-修理保養及保固費" Value="管總-修理保養及保固費"></asp:ListItem>
                        <asp:ListItem Text="維成-交通及運輸設備修護費" Value="維成-交通及運輸設備修護費"></asp:ListItem>
                        <asp:ListItem Text="管總-郵電費" Value="管總-郵電費"></asp:ListItem>
                        <asp:ListItem Text="管總-水電費" Value="管總-水電費"></asp:ListItem>
                        <asp:ListItem Text="管總-公關費" Value="管總-公關費"></asp:ListItem>
                        <asp:ListItem Text="管成-專業服務費" Value="管成-專業服務費"></asp:ListItem>
                        <asp:ListItem Text="管總-福利費" Value="管總-福利費"></asp:ListItem>
                        <asp:ListItem Text="管成-用品消耗" Value="管成-用品消耗"></asp:ListItem>
                        <asp:ListItem Text="管總-印刷裝訂及廣告費" Value="管總-印刷裝訂及廣告費"></asp:ListItem>
                        <asp:ListItem Text="管成-修理保養及保固費" Value="管成-修理保養及保固費"></asp:ListItem>
                        <asp:ListItem Text="維成-工程及管理諮詢服務費" Value="維成-工程及管理諮詢服務費"></asp:ListItem>
                    </asp:DropDownList>
                    <asp:DropDownList ID="科目2" runat="server" AutoPostBack="True" CssClass="DropDownList" visible='<%# If (Eval("科目2").ToString = "", "False", "True") %>'>
                        <asp:ListItem Text="" Value=""></asp:ListItem>
                        <asp:ListItem Text="管總-專業服務費" Value="管總-專業服務費"></asp:ListItem>
                        <asp:ListItem Text="管成-規費" Value="管成-規費"></asp:ListItem>
                        <asp:ListItem Text="管總-用品消耗" Value="管總-用品消耗"></asp:ListItem>
                        <asp:ListItem Text="管總-修理保養及保固費" Value="管總-修理保養及保固費"></asp:ListItem>
                        <asp:ListItem Text="維成-交通及運輸設備修護費" Value="維成-交通及運輸設備修護費"></asp:ListItem>
                        <asp:ListItem Text="管總-郵電費" Value="管總-郵電費"></asp:ListItem>
                        <asp:ListItem Text="管總-水電費" Value="管總-水電費"></asp:ListItem>
                        <asp:ListItem Text="管總-公關費" Value="管總-公關費"></asp:ListItem>
                        <asp:ListItem Text="管成-專業服務費" Value="管成-專業服務費"></asp:ListItem>
                        <asp:ListItem Text="管總-福利費" Value="管總-福利費"></asp:ListItem>
                        <asp:ListItem Text="管成-用品消耗" Value="管成-用品消耗"></asp:ListItem>
                        <asp:ListItem Text="管總-印刷裝訂及廣告費" Value="管總-印刷裝訂及廣告費"></asp:ListItem>
                        <asp:ListItem Text="管成-修理保養及保固費" Value="管成-修理保養及保固費"></asp:ListItem>
                        <asp:ListItem Text="維成-工程及管理諮詢服務費" Value="維成-工程及管理諮詢服務費"></asp:ListItem>
                    </asp:DropDownList>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="摘要">
                    <ItemTemplate>
                        <asp:TextBox ID="摘要" runat="server" Text='<%# Bind("摘要") %>' TextMode="MultiLine" Maxlength=0 Enabled="True" CssClass="TextBox 摘要"/>
                        <ajaxToolkit:AutoCompleteExtender ID="摘要自動" runat="server" TargetControlID="摘要" 
                            CompletionSetCount="10000" CompletionInterval="50" EnableCaching="true" MinimumPrefixLength="1" 
                            ServiceMethod="GetMyList" 
                            CompletionListCssClass="CompletionList" 
                            CompletionListItemCssClass="CompletionListItem" 
                            CompletionListHighlightedItemCssClass="CompletionListHighlightedItem"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="">
                    <ItemTemplate>
                        <asp:Button ID="本月小計" runat="server" Text="本月小計" CommandName="本月小計" Enabled='true' CssClass="GreenButton" />
                        <asp:Button ID="累計至本月" runat="server" Text="累計至本月" CommandName="累計至本月" Enabled='true' CssClass="GreenButton" />
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Left" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="姓名">
                    <ItemTemplate>
                        <asp:ImageButton ID="姓名" runat="server" ImageUrl='<%# If (Eval("姓名").ToString = "", "", Eval("姓名")) %>' CommandName="簽名圖" AutoPostBack="False" OnClientClick="return confirm('如無資料即刪除，確定?')" AlternateText="按此儲存" CssClass="TextBox 姓名"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="商號">
                    <ItemTemplate>
                        <asp:TextBox ID="商號" runat="server" Text='<%# Bind("商號") %>' TextMode="MultiLine" Maxlength=0 Enabled="True" CssClass="TextBox 商號"/>
                        <ajaxToolkit:AutoCompleteExtender ID="商號自動" runat="server" TargetControlID="商號" 
                            CompletionSetCount="10000" CompletionInterval="50" EnableCaching="true" MinimumPrefixLength="1" 
                            ServiceMethod="GetMyList" 
                            CompletionListCssClass="CompletionList" 
                            CompletionListItemCssClass="CompletionListItem" 
                            CompletionListHighlightedItemCssClass="CompletionListHighlightedItem"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="經手人">
                    <ItemTemplate>
                        <asp:ImageButton ID="經手人" runat="server" ImageUrl='<%# If (Eval("經手人").ToString = "", "", Eval("經手人")) %>' CommandName="經手人圖" AutoPostBack="False" OnClientClick="return confirm('如無資料即刪除，確定?')" AlternateText="按此儲存" CssClass="TextBox 經手人"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="種類">
                    <ItemTemplate>
                        <asp:TextBox ID="種類" runat="server" Text='<%# Bind("種類") %>' Maxlength=0 Enabled="True" CssClass="TextBox 種類"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="號數">
                    <ItemTemplate>
                        <asp:TextBox ID="號數" runat="server" Text='<%# Bind("號數", "{0:000}") %>' Maxlength=0 Enabled="True" CssClass="TextBox 號數"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="取號">
                    <ItemTemplate>
                        <asp:CheckBox ID="取號勾選" runat="server" AutoPostBack="True"   Enabled='<%# If (Eval("號數").ToString = "", 1, 0) %>' CssClass="input"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Left" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="收入">
                    <ItemTemplate>
                        <asp:TextBox ID="收入" runat="server" Text='<%# Bind("收入", "{0:c0}") %>' Maxlength=0 Enabled="True" CssClass="TextBox 收入"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="支出">
                    <ItemTemplate>
                        <asp:TextBox ID="支出" runat="server" Text='<%# Bind("支出", "{0:c0}") %>' Maxlength=0 Enabled="True" CssClass="TextBox 支出"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="餘額">
                    <ItemTemplate>
                        <asp:TextBox ID="餘額" runat="server" Text='<%# Bind("餘額", "{0:c0}") %>' Maxlength=0 Enabled="True" CssClass="TextBox 餘額"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="備註">
                    <ItemTemplate>
                        <asp:TextBox ID="備註" runat="server" Text='<%# Bind("備註") %>' TextMode="MultiLine" Maxlength=0 Enabled="True" CssClass="TextBox 備註"/>
                        <ajaxToolkit:AutoCompleteExtender ID="備註自動" runat="server" TargetControlID="備註" 
                            CompletionSetCount="10000" CompletionInterval="50" EnableCaching="true" MinimumPrefixLength="1" 
                            ServiceMethod="GetMyList" 
                            CompletionListCssClass="CompletionList" 
                            CompletionListItemCssClass="CompletionListItem" 
                            CompletionListHighlightedItemCssClass="CompletionListHighlightedItem"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="選取號數">
                    <ItemTemplate>
                    <asp:CheckBox ID="審核" runat="server" AutoPostBack="True" CommandName="審核勾選" OnCheckedChanged="審核_CheckedChanged" CssClass="input"/><br>
                    <asp:Label ID="審核狀態" runat="server"  Text='<%# 
                    If (Eval("鎖定").ToString = "True", 
                    If (Eval("過審").ToString = "True", 
                    If (Eval("回覆").ToString = "True", "主計室通過", 
                    If (Eval("送交主計室日期").ToString="", "通過", "送交主計室")), "已送審"), 
                    If (Eval("送出").ToString = "True", 
                    If (Eval("駁回原因").ToString = "拿回", "被拿回", "駁回"), "未送審"))%>' Maxlength=0 Enabled="True" CssClass="Label 審核狀態"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="選取">
                    <ItemTemplate>
                    <asp:CheckBox ID="刪除選取" runat="server" AutoPostBack="True" CommandName="刪除選取勾選" CssClass="input"/><br>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="預支日期">
                    <ItemTemplate>
                        <asp:TextBox ID="預支日期" runat="server" Text='<%# If (IsDate(Eval("預支日期")), (Year(Eval("預支日期"))-1911).ToString() & Eval("預支日期", "{0:/MM/dd}"), "") %>' Maxlength=0 Enabled="True" CssClass="TextBox 預支日期" />
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="歸還日期">
                    <ItemTemplate>
                        <asp:TextBox ID="歸還日期" runat="server" Text='<%# If (IsDate(Eval("歸還日期")), (Year(Eval("歸還日期"))-1911).ToString() & Eval("歸還日期", "{0:/MM/dd}"), "") %>' Maxlength=0 Enabled='<%# If (IsDate(Eval("預支日期")), true, false) %>' CssClass="TextBox 歸還日期" />
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="駁回原因">
                    <ItemTemplate>
                    <asp:Label ID="駁回原因" runat="server" Text='<%# Eval("駁回原因") %>' Maxlength=0 Enabled="True" CssClass="Label 駁回原因"/>
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
    <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString='<%$ ConnectionStrings:ApplicationServices%>'
        SelectCommand="
            SELECT * FROM 收支備查簿
            WHERE 年 = @年
            AND _種類 = @_種類
            AND ((''=TRIM(@頁1) OR ''=TRIM(@頁2))
            OR ( _頁 BETWEEN TRIM(@頁1) AND TRIM(@頁2)))
            AND (''=TRIM(@單位別) OR 單位別 LIKE N'%'+TRIM(@單位別)+'%')
            AND (''=TRIM(@承辦人) OR 承辦人 LIKE N'%'+TRIM(@承辦人)+'%')
            AND (((TRIM(@月1)='' OR TRIM(@日1)='') AND (TRIM(@月2)='' OR TRIM(@日2)=''))
            OR (CONVERT(date,STR(年+1911)+STR(月)+STR(日)) 
            BETWEEN CONVERT(date,STR(TRIM(@年)+1911)+STR(ISNULL(NULLIF(TRIM(@月1),''),'1'))+STR(ISNULL(NULLIF(TRIM(@日1),''),'1')))
            AND CONVERT(date,STR(@年+1911)+STR(ISNULL(NULLIF(TRIM(@月2),''),'12'))+STR(ISNULL(NULLIF(TRIM(@日2),''),STR(Day(EOMONTH(STR(TRIM(@年)+1911)+'/'+STR(ISNULL(NULLIF(TRIM(@月2),''),'12'))+'/01'))))))))
            AND (''=TRIM(@科目) OR 科目 LIKE N'%'+TRIM(@科目)+'%')
            AND (''=TRIM(@摘要) OR 摘要 LIKE N'%'+TRIM(@摘要)+'%')
            AND (''=TRIM(@商號) OR 商號 LIKE N'%'+TRIM(@商號)+'%')
            AND ((''=TRIM(@號數1) OR ''=TRIM(@號數2))
            OR ( 號數 BETWEEN 
            SUBSTRING(TRIM(@號數1), PATINDEX('%[^0]%', TRIM(@號數1)), 3) AND 
            SUBSTRING(TRIM(@號數2), PATINDEX('%[^0]%', TRIM(@號數2)), 3)))
            AND 取號 = 0 
            AND (''=TRIM(@審核狀態) OR
            IIf (鎖定='True', IIf (過審='True', IIf (回覆 = 'True', '主計室通過', IIf (送交主計室日期 IS NULL, '通過', '送交主計室')), '已送審'), IIf (送出 = 'True', '駁回', '未送審')) LIKE TRIM(@審核狀態))
            ORDER BY _頁, _列" 
        Insertcommand="" 
        UpdateCommand=""
        DeleteCommand="DELETE FROM 收支備查簿 WHERE id=@id">
        <SelectParameters>
            <asp:ControlParameter ControlID="年" ConvertEmptyStringToNull="False" Name="年" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="_種類" ConvertEmptyStringToNull="False" Name="_種類" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="頁1" ConvertEmptyStringToNull="False" Name="頁1" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="頁2" ConvertEmptyStringToNull="False" Name="頁2" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="單位別" ConvertEmptyStringToNull="False" Name="單位別" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="承辦人" ConvertEmptyStringToNull="False" Name="承辦人" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="月1" ConvertEmptyStringToNull="False" Name="月1" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="日1" ConvertEmptyStringToNull="False" Name="日1" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="月2" ConvertEmptyStringToNull="False" Name="月2" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="日2" ConvertEmptyStringToNull="False" Name="日2" PropertyName="Text" Type="String"/> 
            <asp:ControlParameter ControlID="科目" ConvertEmptyStringToNull="False" Name="科目" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="摘要" ConvertEmptyStringToNull="False" Name="摘要" PropertyName="Text" Type="String"/>
            <asp:ControlParameter ControlID="商號" ConvertEmptyStringToNull="False" Name="商號" PropertyName="Text" Type="String"/>
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
        SelectCommand="
            SELECT * FROM 科目表 Where 科目<>'' OR 科目 IS NOT NULL 
            ORDER BY id" 
        Insertcommand=""
        UpdateCommand=""
        DeleteCommand="">
        <SelectParameters>
        </SelectParameters>
        <InsertParameters>
        </InsertParameters>
        <UpdateParameters>
        </UpdateParameters>
        <DeleteParameters>
        </DeleteParameters>
    </asp:SqlDataSource>
</asp:Content>


