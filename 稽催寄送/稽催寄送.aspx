<%@ Page Language="VB" MasterPageFile="./MasterPage.master" MaintainScrollPositionOnPostback="true" AutoEventWireup="false" CodeFile="稽催寄送.aspx.vb" Inherits="稽催寄送"%>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server" >
    <asp:ScriptManager ID="ScriptManager1" runat="server" EnablePageMethods = "true" EnableScriptGlobalization="True" EnableScriptLocalization="True"/>
    <script src='js/稽催寄送.js'></script>
    <link rel="stylesheet" runat="server" media="screen" href="css\稽催寄送.css"/>
    <style>
    </style>
    <div><h1><a id="Title" href="稽催寄送.aspx">稽催寄送<a></h1></div>
    <asp:Panel ID="Panel1" runat="server" CssClass="Panel1" >
        文號:<asp:TextBox ID="文號" runat="server" Maxlength=10 CssClass="Input1"/>
        <asp:Button ID="搜尋" runat="server" Text="搜尋" CssClass="GreenButton"/>
        <asp:Button ID="顯示稽催" runat="server" Text="顯示稽催" OnClick="顯示稽催_Click" CssClass="GreenButton"/>
        <asp:Button ID="顯示內容" runat="server" Text="顯示內容" OnClick="顯示內容_Click" Visible="false"  CssClass="GreenButton"/>
        <!-- 附件:<asp:FileUpload ID="FileUpload1" runat="server" AllowMultiple="True" CssClass="UploadButton"/> -->
        <asp:Button ID="寄送" runat="server" Text="寄送" OnClick="寄送_Click" Visible="false" CssClass="GreenButton"/>
        <asp:Label ID="Label1" runat="server" Text="" CssClass="GreenLabel"/>
        <asp:Label ID="Label2" runat="server" Text="" CssClass="RedLabel"/>
    </asp:Panel>
    <BR>
    <table bgcolor="White">
        <tr>
            <td><asp:Label ID="Label3" runat="server" Text="" CssClass="Label 主旨內容"/></td>
        </tr>
    </table>
    <asp:Panel ID="Panel3" runat="server" CssClass="Panel3" >
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="false" CellPadding="1" HorizontalAlign="Center" DataSourceID="SqlDataSource1" GridLines="None" ShowHeaderWhenEmpty="false" AllowPaging="True" PageSize="10" AllowSorting="True" PagerSettings-PageButtonCount=25 EnableModelValidation="True" AutoGenerateEditButton="false">
            <Columns>
                <asp:TemplateField HeaderText="列">
                    <ItemTemplate>
                        <asp:Label ID="列" runat="server" Text='' Maxlength=0 Enabled="False" CssClass="Label 列"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="承辦人">
                    <ItemTemplate>
                        <asp:Label ID="承辦人" runat="server" Text='<%# Eval("承辦人") %>' Maxlength=0 Enabled="False" CssClass="Label 承辦人"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="文號">
                    <ItemTemplate>
                        <asp:Label ID="文號" runat="server" Text='<%# Eval("文號") %>' Maxlength=0 Enabled="False" CssClass="Label 文號"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="創文日期">
                    <ItemTemplate>
                        <asp:Label ID="創文日期" runat="server" Text='<%# If (IsDate(Eval("創文日期")), (Year(Eval("創文日期"))-1911).ToString() & Eval("創文日期", "{0:/MM/dd}"), "") %>'  Maxlength=0 Enabled="False" CssClass="Label 創文日期"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="限辦日期">
                    <ItemTemplate>
                        <asp:Label ID="限辦日期" runat="server" Text='<%# If (IsDate(Eval("限辦日期")), (Year(Eval("限辦日期"))-1911).ToString() & Eval("限辦日期", "{0:/MM/dd}"), "") %>'  Maxlength=0 Enabled="False" CssClass="Label 限辦日期"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="主旨">
                    <ItemTemplate>
                        <asp:Label ID="主旨" runat="server" Text='<%# Eval("主旨") %>' Maxlength=0 Enabled="False" CssClass="Label 主旨"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="查催次數">
                    <ItemTemplate>
                        <asp:Label ID="查催次數" runat="server" Text='<%# Eval("查催次數") %>' Maxlength=0 Enabled="False" CssClass="Label 查催次數"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="目前流程">
                    <ItemTemplate>
                        <asp:Label ID="目前流程" runat="server" Text='<%# Eval("目前流程") %>' Maxlength=0 TextMode="MultiLine" Enabled="False" CssClass="Label 目前流程"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="目前">
                    <ItemTemplate>
                        <asp:Label ID="目前" runat="server" Text='<%# Eval("目前") %>' Maxlength=0 Enabled="False" CssClass="Label 目前"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="使用日數">
                    <ItemTemplate>
                        <asp:Label ID="使用日數" runat="server" Text='<%# Eval("使用日數") %>' Maxlength=0 Enabled="False" CssClass="Label 使用日數"/>
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
    <asp:Panel ID="Panel4" runat="server" CssClass="Panel4" Visible="false" >
        <asp:GridView ID="GridView2" runat="server" AutoGenerateColumns="false" CellPadding="1" HorizontalAlign="Center" DataSourceID="SqlDataSource2" GridLines="None" ShowHeaderWhenEmpty="false" AllowPaging="True" PageSize="10" AllowSorting="True" PagerSettings-PageButtonCount=25 EnableModelValidation="True" AutoGenerateEditButton="false">
            <Columns>
                <asp:TemplateField HeaderText="列">
                    <ItemTemplate>
                        <asp:Label ID="列" runat="server" Text='' Maxlength=0 Enabled="False" CssClass="Label 列"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="承辦人">
                    <ItemTemplate>
                        <asp:Label ID="承辦人" runat="server" Text='<%# Eval("承辦人") %>' Maxlength=0 Enabled="False" CssClass="Label 承辦人"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="文號">
                    <ItemTemplate>
                        <asp:Label ID="文號" runat="server" Text='<%# Eval("文號") %>' Maxlength=0 Enabled="False" CssClass="Label 文號"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="EMAIL">
                    <ItemTemplate>
                        <asp:Label ID="EMAIL" runat="server" Text='<%# Eval("EMAIL") %>' Maxlength=0 Enabled="False" CssClass="Label EMAIL"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="主管">
                    <ItemTemplate>
                        <asp:Label ID="主管" runat="server" Text='<%# Eval("主管") %>' Maxlength=0 Enabled="False" CssClass="Label 主管"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="主管_EMAIL">
                    <ItemTemplate>
                        <asp:Label ID="主管_EMAIL" runat="server" Text='<%# Eval("主管_EMAIL") %>' Maxlength=0 Enabled="False" CssClass="Label 主管_EMAIL"/>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center"/>
                    <ItemStyle HorizontalAlign="Center" CssClass="Item"/>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="使用日數">
                    <ItemTemplate>
                        <asp:Label ID="使用日數" runat="server" Text='<%# Eval("使用日數") %>' Maxlength=0 Enabled="False" CssClass="Label 使用日數"/>
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
            select Distinct 
            COM_DATA.STR_DATA as 承辦人
            ,case when CUR_FLOW.RECV_DOC_NO IS NULL Then CUR_FLOW.CREATE_DOC_NO Else CUR_FLOW.RECV_DOC_NO END as 文號 
            ,ADMIN_DOC.EXAMINE_DATE as 創文日期
            ,ADMIN_DOC.DUE_DATE as 限辦日期
            ,case when CUR_FLOW.FROM_SUBJECT IS NULL Then CUR_FLOW.CREATE_SUBJECT Else CUR_FLOW.FROM_SUBJECT END as 主旨
            ,ADMIN_DOC.EXAMINE_NO as 查催次數
            ,MY_ACTION_NAME as 目前流程
            , a2.STR_DATA as 目前
            ,ADMIN_DOC.USING_DAY as 使用日數
            from CUR_FLOW
            full JOIN CREATE_DOC ON CREATE_DOC.FLOW_ID=CUR_FLOW.MAIN_FLOW_ID
            left JOIN ADMIN_DOC ON CUR_FLOW.MAIN_FLOW_ID=ADMIN_DOC.FLOW_ID
            left JOIN COM_DATA ON COM_DATA.ID=CUR_FLOW.CHARGE_USER_ID
            left JOIN COM_DATA as a2 ON a2.ID=ADMIN_DOC.CURRENT_MAN
            left JOIN DOC_LIST ON CUR_FLOW.MAIN_FLOW_ID=DOC_LIST.FLOW_ID
            Where 
            DATEADD (DAY,0-(ADMIN_DOC.USING_DAY)+7,convert(varchar, getdate(), 111))>=ADMIN_DOC.EXAMINE_DATE
            AND 
            ADMIN_DOC.USING_DAY
            >=6
            AND CUR_FLOW.ID=MAIN_FLOW_ID
            AND (''=TRIM(@文號) OR (case when CUR_FLOW.RECV_DOC_NO IS NULL Then CUR_FLOW.CREATE_DOC_NO Else CUR_FLOW.RECV_DOC_NO END LIKE N'%'+TRIM(@文號)+'%'))"
        Insertcommand="" 
        UpdateCommand=""
        DeleteCommand="">
        <SelectParameters>
            <asp:ControlParameter ControlID="文號" ConvertEmptyStringToNull="False" Name="文號" PropertyName="Text" Type="String"/>
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
        select Distinct 
        COM_DATA.STR_DATA as 承辦人
        ,CFW_kw.dbo.USERS.EMAIL as EMAIL
        ,case when CUR_FLOW.RECV_DOC_NO IS NULL Then CUR_FLOW.CREATE_DOC_NO Else CUR_FLOW.RECV_DOC_NO END as 文號 
        ,ADMIN_DOC.USING_DAY as 使用日數
        ,sir.NAME as 主管
        ,sir.EMAIL as 主管_EMAIL
        from CUR_FLOW
        left JOIN ADMIN_DOC ON CUR_FLOW.MAIN_FLOW_ID=ADMIN_DOC.FLOW_ID
        left JOIN COM_DATA ON COM_DATA.ID=CUR_FLOW.CHARGE_USER_ID
        left JOIN CFW_kw.dbo.USERS ON COM_DATA.STR_DATA=CFW_kw.dbo.USERS.NAME
        left JOIN CFW_kw.dbo.USERS as sir ON CFW_kw.dbo.USERS.DEPT_ID=sir.DEPT_ID
        Where
        ADMIN_DOC.USING_DAY>=6
        AND CUR_FLOW.ID=MAIN_FLOW_ID
        AND sir.DEC_LEVEL='50'
        AND sir.STATUS='0'
        " 
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
            <asp:Parameter Name="id"/>
        </DeleteParameters>
    </asp:SqlDataSource>
</asp:Content>


