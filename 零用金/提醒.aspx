<%@ Page Language="VB" MasterPageFile="./MasterPage.master" MaintainScrollPositionOnPostback="true" AutoEventWireup="false" CodeFile="提醒.aspx.vb" Inherits="提醒" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
    <asp:ScriptManager ID="ScriptManager1" runat="server" EnablePageMethods = "true" EnableScriptGlobalization="True" EnableScriptLocalization="True"/>
    <script src='js/提醒.js'></script>
    <link rel="stylesheet" runat="server" media="screen" href="css\提醒.css"/>
    <div><h1><a id="Title" href="提醒.aspx">提醒<a></h1></div>
    <asp:Panel ID="Panel1" runat="server" CssClass="Panel1">
        <asp:Label ID="Label1" runat="server" Text="" CssClass="GreenLabel"/>
        <asp:Label ID="Label2" runat="server" Text="" CssClass="RedLabel"/>
        <asp:Label ID="Label3" runat="server" Text="" CssClass="RedLabel2"/>
    </asp:Panel>
    <asp:Panel ID="Panel3" runat="server" CssClass="Panel3" Visible="false">
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="false" CellPadding="1" HorizontalAlign="Center" DataSourceID="SqlDataSource1" DataKeyNames="id" GridLines="None" ShowHeaderWhenEmpty="false" AllowPaging="True" PageSize="15" AllowSorting="True" PagerSettings-PageButtonCount=25 EnableModelValidation="True" AutoGenerateEditButton="false">
            <Columns>
                <asp:TemplateField HeaderText="id" Visible="false">
                    <ItemTemplate>
                        <asp:label width="100" ID="id" runat="server" Text='<%# Eval("ID") %>' Maxlength=0 Enabled="False" CssClass="label id"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="頁">
                    <ItemTemplate>
                        <asp:label width="100" ID="頁" runat="server" Text='<%# Eval("_頁") %>' Maxlength=0 Enabled="True" CssClass="label 頁"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="列">
                    <ItemTemplate>
                        <asp:label width="100" ID="列" runat="server" Text='<%# Eval("_列") %>' Maxlength=0 Enabled="True" CssClass="label 列"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="種類">
                    <ItemTemplate>
                        <asp:label width="100" ID="種類" runat="server" Text='<%# Eval("_種類") %>' Maxlength=0 Enabled="True" CssClass="label 種類"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="號數">
                    <ItemTemplate>
                        <asp:label width="150" ID="號數" runat="server" Text='<%# Eval("號數") %>' Maxlength=0 Enabled="True" CssClass="label 號數"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="狀態">
                    <ItemTemplate>
                        <asp:label width="100" ID="案件" runat="server" Text='<%# If (Eval("鎖定").ToString = "False" And Eval("送出").ToString = "True","遭駁回",If (Eval("預支日期").ToString = "","超過15日，請盡速找經手人簽名" , "請盡速報銷" & If (IsDate(Eval("預支日期")), (Year(Eval("預支日期"))-1911).ToString() & Eval("預支日期", "{0:/MM/dd}"), ""))) %>' Maxlength=0 Enabled="True" CssClass="label 案件2"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                 <asp:TemplateField HeaderText="原因">
                    <ItemTemplate>
                        <asp:label width="100" ID="原因" runat="server" Text='<%# Eval("駁回原因") %>' Maxlength=0 Enabled="True" CssClass="label 駁回原因"/>
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
    <asp:Panel ID="Panel4" runat="server" CssClass="Panel4" Visible="false">
        <asp:GridView ID="GridView2" runat="server" AutoGenerateColumns="false" CellPadding="1" HorizontalAlign="Center" DataSourceID="SqlDataSource2" DataKeyNames="號數,_種類" GridLines="None" ShowHeaderWhenEmpty="false" AllowPaging="True" PageSize="15" AllowSorting="True" PagerSettings-PageButtonCount=25 EnableModelValidation="True" AutoGenerateEditButton="false">
            <Columns>
                <asp:TemplateField HeaderText="id" Visible="false">
                    <ItemTemplate>
                        <asp:label width="100" ID="id" runat="server" Text='' Maxlength=0 Enabled="False" CssClass="label id"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="種類">
                    <ItemTemplate>
                        <asp:label width="150" ID="種類" runat="server" Text='<%# Eval("_種類") %>' Maxlength=0 Enabled="True" CssClass="label 種類"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="號數">
                    <ItemTemplate>
                        <asp:label width="150" ID="號數" runat="server" Text='<%# Eval("號數") %>' Maxlength=0 Enabled="True" CssClass="label 號數"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="狀態">
                    <ItemTemplate>
                        <asp:label width="150" ID="案件" runat="server" Text='<%# If (Eval("鎖定").ToString = "True", "需審理", If (Eval("駁回原因").ToString = "拿回", "被拿回", "遭主計室駁回"))%>' Maxlength=0 Enabled="True" CssClass="label 案件2"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="原因">
                    <ItemTemplate>
                        <asp:label width="100" ID="原因" runat="server" Text='<%# Eval("駁回原因") %>' Maxlength=0 Enabled="True" CssClass="label 駁回原因"/>
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
    <asp:Panel ID="Panel5" runat="server" CssClass="Panel5" Visible="false">
        <asp:GridView ID="GridView3" runat="server" AutoGenerateColumns="false" CellPadding="1" HorizontalAlign="Center" DataSourceID="SqlDataSource3" DataKeyNames="_種類" GridLines="None" ShowHeaderWhenEmpty="false" AllowPaging="True" PageSize="15" AllowSorting="True" PagerSettings-PageButtonCount=25 EnableModelValidation="True" AutoGenerateEditButton="false">
            <Columns>
                <asp:TemplateField HeaderText="id" Visible="false">
                    <ItemTemplate>
                        <asp:label width="100" ID="id" runat="server" Text='' Maxlength=0 Enabled="False" CssClass="label id"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="種類">
                    <ItemTemplate>
                        <asp:label width="150" ID="種類" runat="server" Text='<%# Eval("_種類") %>' Maxlength=0 Enabled="True" CssClass="label 種類"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="狀態">
                    <ItemTemplate>
                        <asp:label width="150" ID="案件" runat="server" Text='需審理' Maxlength=0 Enabled="True" CssClass="label 案件2"/>
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
    <!-- 語法封存 能去除重覆 只取頭ID頁列
        SELECT A.ID,A._頁,A._列,B._種類,B.號數,B.鎖定,B.送出,B.預支日期,B.駁回原因 
        FROM 收支備查簿 AS A Left Join 收支備查簿 AS B
	    ON A.id=B.id AND B.id in (select min(id) From 收支備查簿 group by _種類,號數,鎖定,送出,預支日期,駁回原因 )
        where (B.鎖定='False' And B.送出='True') or ((select datediff(day,getdate(),B.預支日期))<'-1' And B.過審='False')
        order by _種類,號數,_頁, _列
    -->
    <!-- 語法封存 可取出[號數] 最頭頁列 及 最尾頁列
        SELECT A.ID,A._頁,A._列,B._種類,B.號數,B.鎖定,B.送出,B.預支日期,B.駁回原因 
        FROM 收支備查簿 AS A Left Join 收支備查簿 AS B
	    ON A.id=B.id AND 
	    (B.id in (select MIN(id) From 收支備查簿 group by _種類,號數,鎖定,送出,預支日期,駁回原因 ) OR
	    B.id in (select MAX(id) From 收支備查簿 group by _種類,號數,鎖定,送出,預支日期,駁回原因 ))
        where (B.鎖定='False' And B.送出='True') or ((select datediff(day,getdate(),B.預支日期))<'-1' And B.過審='False')
        order by _種類,號數,_頁, _列
    -->
    <!-- 語法封存 原1 未去除重複
        SELECT DISTINCT ID,_頁,_列,_種類,號數,鎖定,送出,預支日期,駁回原因 
        FROM 收支備查簿 
        where (鎖定='False' And 送出='True') or ((select datediff(day,getdate(),預支日期))<'-1' And 過審='False')
        order by _種類,號數,_頁, _列
    -->
    <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString='<%$ ConnectionStrings:ApplicationServices%>'
        SelectCommand="SELECT DISTINCT ID,_頁,_列,_種類,號數,鎖定,送出,預支日期,駁回原因 
        FROM 收支備查簿 
        where (鎖定='False' And 送出='True' AND 駁回原因<>'拿回') 
        Or ((select datediff(day,getdate(),預支日期))<'-1' And (過審='False' AND 歸還日期 IS NULL AND 送出='False')) 
        Or ((datediff(day,getdate(),Trim(str(年+'1911'))+'/'+Trim(str(月))+'/'+Trim(str(日))))<'-15' And (送交主計室日期 IS NULL And 經手人 IS NULL AND 取號=1 AND 支出>0))
        order by _種類,號數,_頁, _列" 
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
    <!-- 語法封存 能去除重覆 只取頭ID
        SELECT A.id,B.號數,B._種類,B.鎖定,B.駁回原因
        FROM 收支備查簿 AS A Left Join 收支備查簿 AS B
	    ON A.id=B.id AND B.id in (select min(id) From 收支備查簿 group by _種類,號數,鎖定,駁回原因 )
        where (B.鎖定='True' And B.送出='True' And B.過審='False') or (B.鎖定='False' And B.過審='True')
        order by 號數
    -->
    <!-- 語法封存 原2 未去除重複
        SELECT DISTINCT id,號數,_種類,鎖定,駁回原因
        FROM 收支備查簿 
        where (鎖定='True' And 送出='True' And 過審='False') or (鎖定='False' And 過審='True')
        order by 號數
    -->
    <asp:SqlDataSource ID="SqlDataSource2" runat="server" ConnectionString='<%$ ConnectionStrings:ApplicationServices%>'
        SelectCommand="SELECT DISTINCT 號數,_種類,鎖定,駁回原因
        FROM 收支備查簿 
        where (鎖定='True' And 送出='True' And 過審='False') or (鎖定='False' And (主計室簽核 IS NOT NULL Or 駁回原因='拿回'))
        order by 號數"
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
    <asp:SqlDataSource ID="SqlDataSource3" runat="server" ConnectionString='<%$ ConnectionStrings:ApplicationServices%>'
        SelectCommand="SELECT DISTINCT _種類
        FROM 收支備查簿 
        where 送交主計室日期 is not NULL And 回覆='false'
        order by _種類" 
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
