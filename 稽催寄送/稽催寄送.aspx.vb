Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel.XlDirection
Imports Microsoft.Office.Interop.Excel.XlPageBreak
Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.VisualBasic.Logging
' Imports Microsoft.Office.Interop.Outlook.O1Mail
' Imports OutLook = Microsoft.Office.Interop.Outlook
Imports System.IO
Imports System.Math
Imports System.Diagnostics
Imports System.Data.Sql
Imports System.Data.SqlTypes
Imports System.Data.SqlClient
Imports System.Configuration
Imports System.Text.RegularExpressions
Imports System.Web.UI.WebControls
Imports System.Drawing
Partial Class 稽催寄送
    Inherits System.Web.UI.Page
    Dim CFW_wf2 As String = System.Web.Configuration.WebConfigurationManager.ConnectionStrings("ApplicationServices").ConnectionString
    'Dim con_14 As String = System.Web.Configuration.WebConfigurationManager.ConnectionStrings("ApplicationServices2").ConnectionString
    Dim data As New SqlDataSource
    Dim data_dv As Data.DataView
    Dim data_dv1 As Data.DataView
    Dim data_dv2 As Data.DataView
    Dim data_dv3 As Data.DataView
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        data.ConnectionString = CFW_wf2
        If Not Page.IsPostBack Then
        Else
            Me.Label1.Text = ""
            Me.Label2.Text = ""
        End If 
    End Sub
    Protected Sub 顯示稽催_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        If Panel3.Visible=true Then
            Panel3.Visible=false
            Panel4.Visible=true
            顯示內容.Visible=true
        Else
            Panel3.Visible=true
            Panel4.Visible=false
            顯示內容.Visible=false
            寄送.Visible=false
            Label3.Text=""
        End If
    End Sub
    Protected Sub 顯示內容_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        If Panel4.Visible=true Then
            Panel4.Visible=false
            寄送.Visible=true
            Dim p_Email As string
            Dim 文號 As String
            Dim 單位 As String
            Dim 承辦人 As String
            Dim 辦理天數 As String
            Dim 承辦人_EMAIL As String
            Dim 限辦日期 As string
            data.SelectCommAnd = "select Distinct  " & _
                "case when CUR_FLOW.RECV_DOC_NO IS NULL Then CUR_FLOW.CREATE_DOC_NO Else CUR_FLOW.RECV_DOC_NO END as 文號  " & _
                ",CFW_kw.dbo.DEPT.NAME As 單位 " & _
                ",COM_DATA.STR_DATA As 承辦人 " & _
                ",ADMIN_DOC.USING_DAY As 辦理天數 " & _
                ",CUR_FLOW.DOC_CAT As CAT " & _
                ",ADMIN_DOC.DUE_DATE as 限辦日期 " & _
                "from CUR_FLOW " & _
                "left JOIN COM_DATA ON COM_DATA.ID=CUR_FLOW.CHARGE_USER_ID " & _
                "left JOIN ADMIN_DOC ON CUR_FLOW.MAIN_FLOW_ID=ADMIN_DOC.FLOW_ID " & _
                "left JOIN CFW_kw.dbo.USERS ON COM_DATA.STR_DATA=CFW_kw.dbo.USERS.NAME " & _
                "left JOIN CFW_kw.dbo.DEPT ON CFW_kw.dbo.USERS.DEPT_ID=CFW_kw.dbo.DEPT.ID " & _
                "Where " & _
                "ADMIN_DOC.USING_DAY>=6 " & _
                "AND CUR_FLOW.ID=MAIN_FLOW_ID"
            data_dv2 = data.Select(New DataSourceSelectArguments)'取主旨
            data.SelectCommAnd = "select Distinct  " & _
                "CFW_kw.dbo.DEPT.NAME As 單位 " & _
                "from CUR_FLOW " & _
                "left JOIN COM_DATA ON COM_DATA.ID=CUR_FLOW.CHARGE_USER_ID " & _
                "left JOIN ADMIN_DOC ON CUR_FLOW.MAIN_FLOW_ID=ADMIN_DOC.FLOW_ID " & _
                "left JOIN CFW_kw.dbo.USERS ON COM_DATA.STR_DATA=CFW_kw.dbo.USERS.NAME " & _
                "left JOIN CFW_kw.dbo.DEPT ON CFW_kw.dbo.USERS.DEPT_ID=CFW_kw.dbo.DEPT.ID " & _
                "Where " & _
                "ADMIN_DOC.USING_DAY>=6 " & _
                "AND CUR_FLOW.ID=MAIN_FLOW_ID"
            data_dv3 = data.Select(New DataSourceSelectArguments)'取單位
            p_Email="一、檢陳貴單位辦理6日(含)以上未結公文稽催表(如附件)，敬請督導催辦所屬同仁各案流程可使用時間，於期限屆滿以前督促辦結，避免逾期，以增進公文處理時效。<BR><BR>"& _
            "二、為精進本分局為民服務品質，人民陳情案限辦日期雖為30日(含假日)，惟依規定未能申請展期，請各單位承辦人於15日內辦結，以提升本分局處理速度及為民服務品質。<BR><BR>"& _
            "文號&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;單位&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;承辦人&nbsp;&nbsp;&nbsp;&nbsp;辦理天數 (備註)<BR>"
            For i=0 to data_dv2.Count - 1'取承辦人Email，以下有問題
                文號=data_dv2(i)("文號").ToString()
                單位=data_dv2(i)("單位").ToString()
                承辦人=data_dv2(i)("承辦人").ToString()
                辦理天數=data_dv2(i)("辦理天數").ToString()
                限辦日期 = data_dv2(i)("限辦日期").ToString()
                p_Email+= (i+1) & "," & 文號 & "&nbsp;&nbsp;&nbsp;&nbsp;" & 單位 & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & 承辦人 & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & 辦理天數 & "天"
                ' Label1.Text+=(i+1) & "," & 文號 & "    " & 單位 & "     " & 承辦人 & "     " & 辦理天數 & "<BR>"
                If data_dv2(i)("CAT").ToString()="5"
                    If IsDate(限辦日期)
                        限辦日期 = ToTaiwanCalendar(限辦日期)
                    END If
                    p_Email+="(人民陳情案，" & 限辦日期 & "到期)"
                    'Label1.Text+="執行IF<BR>"
                    'Label1.Text+=data_dv2(i)("辦理天數").ToString("yyyy/MM/dd") & "<BR>"
                End If
                p_Email+="<BR>"
            Next
            p_Email+="<BR>此致<BR>"
            For i=0 to data_dv3.Count - 1'取承辦人單位
                單位=data_dv3(i)("單位").ToString()
                p_Email+= 單位 &"<BR>"
            Next
            p_Email+="<BR>秘書室 敬啟<BR>"& _
            "********************************<BR><BR>"& _
            "高速公路局中區養護工程分局<BR>"& _
            "秘書室 助理員 楊嘉惠  敬啟<BR>"& _
            "電話:04-22529181#2903<BR>"& _
            "E-Mail:<a href='mailto:chiahui@freeway.gov.tw' target='_blank'>chiahui@freeway.gov.tw</a><BR><BR>"& _
            "*******************************"
            ' Me.table1.visibility="visible"
            Label3.Text=p_Email
        Else
            Panel4.Visible=true
            寄送.Visible=false
            Label3.Text=""
            ' Me.table1.visibility="hidden"
        End If
    End Sub
    Protected Sub 寄送_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        If Label3.Text<>""
            sendEMailThroughOUTLOOK()
        End if
    End Sub
    Public Sub sendEMailThroughOUTLOOK()
        '客戶端outlook
        'Dim 主管 As String
        Dim 主管_Email As String
        Dim p_Email As String
        'Dim 2901_Email As String="hollow@freeway.gov.tw"'副本寄送給主任
        '取陳辦人
        data.ConnectionString = CFW_wf2
        data.SelectCommAnd = "select Distinct " & _
            "sir.EMAIL as 主管_EMAIL " & _
            "from CUR_FLOW " & _
            "left JOIN COM_DATA ON COM_DATA.ID=CUR_FLOW.CHARGE_USER_ID " & _
            "left JOIN ADMIN_DOC ON CUR_FLOW.MAIN_FLOW_ID=ADMIN_DOC.FLOW_ID " & _
            "left JOIN CFW_kw.dbo.USERS ON COM_DATA.STR_DATA=CFW_kw.dbo.USERS.NAME " & _
            "left JOIN CFW_kw.dbo.USERS as sir ON CFW_kw.dbo.USERS.DEPT_ID=sir.DEPT_ID " & _
            "Where " & _
            "ADMIN_DOC.USING_DAY>=6 " & _
            "AND CUR_FLOW.ID=MAIN_FLOW_ID " & _
            "AND (sir.DEC_LEVEL='50' OR sir.DEC_LEVEL='60') " & _
			"AND sir.STATUS='0'"
        data_dv = data.Select(New DataSourceSelectArguments)'取主管Email
        data.SelectCommAnd = "select Distinct  " & _
            "CFW_kw.dbo.USERS.EMAIL As EMAIL  " & _
            "from CUR_FLOW " & _
            "left JOIN COM_DATA ON COM_DATA.ID=CUR_FLOW.CHARGE_USER_ID " & _
            "left JOIN ADMIN_DOC ON CUR_FLOW.MAIN_FLOW_ID=ADMIN_DOC.FLOW_ID " & _
            "left JOIN CFW_kw.dbo.USERS ON COM_DATA.STR_DATA=CFW_kw.dbo.USERS.NAME " & _
            "Where " & _
            "ADMIN_DOC.USING_DAY>=6 " & _
            "AND CUR_FLOW.ID=MAIN_FLOW_ID"
        data_dv1 = data.Select(New DataSourceSelectArguments)'取承辦人Email
        ' data.SelectCommAnd = "select Distinct  " & _
        '     "case when CUR_FLOW.RECV_DOC_NO IS NULL Then CUR_FLOW.CREATE_DOC_NO Else CUR_FLOW.RECV_DOC_NO END as 文號  " & _
        '     ",CFW_kw.dbo.DEPT.NAME As 單位 " & _
        '     ",COM_DATA.STR_DATA As 承辦人 " & _
        '     ",ADMIN_DOC.USING_DAY As 辦理天數 " & _
        '     ",CUR_FLOW.DOC_CAT As CAT " & _
        '     ",ADMIN_DOC.DUE_DATE as 限辦日期 " & _
        '     "from CUR_FLOW " & _
        '     "left JOIN COM_DATA ON COM_DATA.ID=CUR_FLOW.CHARGE_USER_ID " & _
        '     "left JOIN ADMIN_DOC ON CUR_FLOW.MAIN_FLOW_ID=ADMIN_DOC.FLOW_ID " & _
        '     "left JOIN CFW_kw.dbo.USERS ON COM_DATA.STR_DATA=CFW_kw.dbo.USERS.NAME " & _
        '     "left JOIN CFW_kw.dbo.DEPT ON CFW_kw.dbo.USERS.DEPT_ID=CFW_kw.dbo.DEPT.ID " & _
        '     "Where " & _
        '     "ADMIN_DOC.USING_DAY>=6 " & _
        '     "AND CUR_FLOW.ID=MAIN_FLOW_ID"
        ' data_dv2 = data.Select(New DataSourceSelectArguments)'取主旨
        ' data.SelectCommAnd = "select Distinct  " & _
        '     "CFW_kw.dbo.DEPT.NAME As 單位 " & _
        '     "from CUR_FLOW " & _
        '     "left JOIN COM_DATA ON COM_DATA.ID=CUR_FLOW.CHARGE_USER_ID " & _
        '     "left JOIN ADMIN_DOC ON CUR_FLOW.MAIN_FLOW_ID=ADMIN_DOC.FLOW_ID " & _
        '     "left JOIN CFW_kw.dbo.USERS ON COM_DATA.STR_DATA=CFW_kw.dbo.USERS.NAME " & _
        '     "left JOIN CFW_kw.dbo.DEPT ON CFW_kw.dbo.USERS.DEPT_ID=CFW_kw.dbo.DEPT.ID " & _
        '     "Where " & _
        '     "ADMIN_DOC.USING_DAY>=6 " & _
        '     "AND CUR_FLOW.ID=MAIN_FLOW_ID"
        ' data_dv3 = data.Select(New DataSourceSelectArguments)'取單位
        For i=0 to data_dv.Count - 1'取主管Email
            主管_EMAIL=data_dv(i)("主管_EMAIL").ToString()
            If i=0 Then
                p_Email="mailto:" & 主管_EMAIL
            Else
                p_Email+=","&主管_EMAIL
            End If
        Next
        ' For Each PostedFile As HttpPostedFile In FileUpload1.PostedFiles'上傳附件
        '     Dim MyGUID As String = Guid.NewGuid().ToString("N")
        '     Dim Myfiles As String = MapPath(".\data\Temp\") & MyGUID
        '     PostedFile.SaveAs(Myfiles)
        '     Try
        '         File.Copy(Myfiles, MapPath(".\data\報表作業檔案\") & PostedFile.FileName, False)
        '     Catch
        '     End Try
        '     data.InsertCommand = _
        '         "IF NOT EXISTS(SELECT * FROM 報表作業上傳 WHERE 檔名 = N'" & PostedFile.FileName & "') " & _
        '         "BEGIN " & _
        '         "INSERT INTO 報表作業上傳 (檔名) " & _
        '         "VALUES('" & PostedFile.FileName & "')" & _
        '         "END"
        '     data.Insert()
        '     System.IO.File.Delete(Myfiles)
        '     Me.GridView1.PageIndex = Int32.MaxValue
        ' Next
        ' For i=0 to data_dv1.Count - 1'取承辦人Email
        '     承辦人_EMAIL=data_dv1(i)("EMAIL").ToString()
        '     p_Email+="," & 承辦人_EMAIL
        ' Next
        'Label1.Text=p_Email
        'p_Email=承辦人_EMAIL & "?subject:(非社交演練)檢陳公文辦理6日(含)以上未結案稽催表(如附件)＆body=" & (i+1) & 文號 & 單位 & 辦理天數
        ' Dim Test As string="mailto:stanleyoreo@yahoo.com.tw ?cc=stanleyoreo@gmail.com&subject=(非社交演練)檢陳公文辦理6日(含)以上未結案稽催表(如附件)&body=Hi"
        ' Response.Redirect(Test)
        '1/12未來修正為正式寄信
        'Response.write("<script language=javascript>mailto:stanleyoreo@yahoo.com.tw?subject=(非社交演練)檢陳公文辦理6日(含)以上未結案稽催表(如附件)＆body=Hi);</script>")
        p_Email+="?cc=hollow@freeway.gov.tw&subject=(非社交演練)檢陳公文辦理6日(含)以上未結案稽催表(如附件)"
        Response.Redirect(p_Email)
        
        ' Dim REW As string = "<script type=application/ld+json>{ " & _
        '       "@context: http://schema.org/extensions, " & _
        '       "@type: MessageCard, " & _
        '       "hideOriginalBody: true, " & _
        '       "title: 請盡速辦理, " & _
        '       "sections: [{ " & _
        '         "text: Please review the expense report below., " & _
        '         "facts: [{ " & _
        '           "name: ID, " & _
        '           "value: 98432019 " & _
        '         "}, { " & _
        '           "name: Amount, " & _
        '           "value: 83.27 USD " & _
        '         "}, { " & _
        '           "name: Submitter, " & _
        '           "value: Kathrine Joseph " & _
        '         "}, { " & _
        '           "name: Description, " & _
        '           "value: Dinner with client " & _
        '         "}] " & _
        '       "}], " & _
        '       "potentialAction: [{ " & _
        '         "@type: HttpPost, " & _
        '         "name: Approve, " & _
        '         "target:  " & _
        '       "}, { " & _
        '         "@type: OpenUri, " & _
        '         "name: View Expense, " & _
        '         "targets: [ { os: default,  " & _
        '         "uri: https://expense.contoso.com/view?id=98432019} ] " & _
        '       "}] " & _
        '     "} " & _
        '     "</script>"
        ' Response.write(REW)
            
        ' Try
        '     ' Create the Outlook application.似乎遠端電腦未安裝OUTLOOK所以無法執行
        '     ' Dim oApp as Object
        '     ' oMsg=CreateObject("Outlook.Application")
        '     ' Dim oApp As New Outlook.Application()
        '     Dim oApp As New Outlook.ApplicationClass()
        '     ' Create a new mail item.
        '     Dim oMsg As New Outlook.MailItem()
        '     oMsg =DirectCast(oApp.CreateItem(Outlook.OlItemType.olMailItem), Outlook.MailItem)
        '     ' Set HTMLBody.
        '     'add the body of the email
        '     oMsg.HTMLBody ="請盡速辦理"'預定放承辦人機璀等相關資料
        '     'Add an attachment.附件
        '     Dim sDisplayName As String ="MyAttachment"
        '     Dim iPosition As Integer =CInt(oMsg.Body.Length) + 1
        '     Dim iAttachType As Integer =CInt(Outlook.OlAttachmentType.olByValue)
        '     'now attached the file
        '     Dim oAttach As Outlook.Attachment = oMsg.Attachments.Add("C:\\fileName.jpg", iAttachType, iPosition, sDisplayName)
        '     'Subject line
        '     oMsg.Subject ="稽催"
        '     ' Add a recipient.放承辦人E_MALL
        '     Dim oRecips As Outlook.Recipients =DirectCast(oMsg.Recipients, Outlook.Recipients)
        '     ' Change the recipient in the next line if necessary.
        '     Dim oRecip As Outlook.Recipient =DirectCast(oRecips.Add("E_MAIL"), Outlook.Recipient)
        '     oRecip.Resolve()
        '     ' Send.
        '     oMsg.Send()
        '     ' Clean up.
        '     oRecip =Nothing
        '     oRecips =Nothing
        '     oMsg =Nothing
        '     oApp =Nothing
        ' 'end of try block
        ' Catch exAs Exception
        ' End Try
        ' 'end of catch
    End Sub
    Protected Sub GridView1_DataBound(ByVal sender As Object, ByVal e As System.EventArgs) HAndles GridView1.DataBound
        For i= 0 to Me.GridView1.Rows.Count - 1
                CType(Me.GridView1.Rows(i).FindControl("列"), Label).Text=(i+1)
            Next
    End Sub
    Protected Sub GridView2_DataBound(ByVal sender As Object, ByVal e As System.EventArgs) HAndles GridView2.DataBound
        For i= 0 to Me.GridView2.Rows.Count - 1
                CType(Me.GridView2.Rows(i).FindControl("列"), Label).Text=(i+1)
            Next
    End Sub
End Class