Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel.XlDirection
Imports Microsoft.VisualBasic.Logging
Imports System.IO
Imports System.Drawing
Imports System.Diagnostics
Imports System.Data.Sql
Imports System.Data.SqlTypes
Imports System.Data.SqlClient
Imports System.Configuration
Imports System.Text.RegularExpressions
Partial Class MasterPage
    Inherits System.Web.UI.MasterPage
    Dim con_14 As String = System.Web.Configuration.WebConfigurationManager.ConnectionStrings("ApplicationServices").ConnectionString
    Dim data As New SqlDataSource
    Dim data_dv As Data.DataView
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        data.ConnectionString = con_14
        If Not Page.IsPostBack Then
        End If 
        提醒l.text=""
        帳號名.text=Session("姓名")
        Dim 權限() As object={查詢,日誌,提醒,登出,修改密碼,更新資料}
        Dim 權限2902() As object={取號,收支備查簿,常用清單,科目清單,分配明細表,核銷支出明細備查簿,其他作業,留言板,查核資料}
        If not(Session("Uid") is Nothing)
            For i=0 to 權限.count-1
                權限(i).Visible=true
            Next
            登入.Visible=false
            data.SelectCommand = "SELECT DATEDIFF(day,(SELECT psdtime FROM 帳號 where 帳號='" & session("Uid") & "'),GETDATE())"
            data_dv = data.Select(New DataSourceSelectArguments)
            If  data_dv(0)(0) > 90 And System.IO.Path.GetFileName(Request.PhysicalPath)<>"修改密碼.aspx"
                'Response.Write("<Script language='JavaScript'>alert('密碼已經超過3個月未修改，請修改密碼！');location.href('./修改密碼.aspx');</Script>")'盡量別使用Response.Write，會破壞排版
                Response.Redirect("修改密碼.aspx")
            End If
            '加入權限判斷
            If  (Session("atype") <> "all" And System.IO.Path.GetFileName(Request.PhysicalPath)="建立帳號.aspx") OR 
            (not(Session("atype") = "IsAccountantLogin" Or Session("atype") = "all") And System.IO.Path.GetFileName(Request.PhysicalPath)="主計室審核.aspx") OR
            (not(Session("atype") = "IsDirectorLogin" Or Session("atype") = "all") And System.IO.Path.GetFileName(Request.PhysicalPath)="審核.aspx")
                Response.write("<script language=javascript>history.go(-1);</script>)")
            End If
            For i=0 to 權限2902.count-1
                If  (not(Session("atype") = "IsUserLogin" Or Session("atype") = "all") And System.IO.Path.GetFileName(Request.PhysicalPath)= 權限2902(i).id & ".aspx")
                    Response.write("<script language=javascript>history.go(-1);</script>)")
                End If
            Next
            
        End If
        If Session("atype") = "IsUserLogin" Or Session("atype") = "all"
            For i=0 to 權限2902.count-1
                權限2902(i).Visible=true
            Next
            data.ConnectionString = con_14
            data.SelectCommand = "SELECT ID FROM 收支備查簿 where (鎖定='False' And 送出='True' AND 駁回原因<>'拿回') or ((select datediff(day,getdate(),預支日期))<=2 And (過審='False' AND 歸還日期 IS NULL AND 送出='False'))"
            data_dv = data.Select(New DataSourceSelectArguments)
            If data_dv.Count > 0
                提醒l.text=提醒l.text & "(" & data_dv.Count & ")"
            End If 
        End If 
        If Session("atype") = "IsDirectorLogin" Or Session("atype") = "all"
            審核.Visible=true
            data.ConnectionString = con_14
            data.SelectCommand = "SELECT Distinct 號數 FROM 收支備查簿 where (鎖定='True' And 送出='True' And 過審='False') or (鎖定='False' And (過審='True' Or 駁回原因='拿回'))"
            data_dv = data.Select(New DataSourceSelectArguments)
            If data_dv.Count > 0
                提醒l.text=提醒l.text & "(" & data_dv.Count & ")"
            End If 
        End If 
        If Session("atype") = "IsAccountantLogin" Or Session("atype") = "all"
            主計室審核.Visible=true
            data.ConnectionString = con_14
            data.SelectCommand = "SELECT _種類 FROM 收支備查簿 where 送交主計室日期 is not NULL And 回覆='false'"
            data_dv = data.Select(New DataSourceSelectArguments)
            If data_dv.Count > 0
                提醒l.text=提醒l.text & "(" & data_dv.Count & ")"
            End If
        End If
        If Session("atype") = "all"
            建立帳號.Visible=true
        End If
        If (System.IO.Path.GetFileName(Request.PhysicalPath)<>"登入.aspx" And Session("Uid") is Nothing)  'OR (System.IO.Path.GetFileName(Request.PhysicalPath)="建立帳號.aspx" And Session("atype")<>"all")
            Response.Redirect("Default.aspx")
        End If 
    End Sub
    Protected Sub Download(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim MyGUID As String = Guid.NewGuid().ToString("N")
        Dim MyExcel As String = MapPath(".\Excel\Temp\") & MyGUID & ".xls"
        System.IO.File.Copy(MapPath(".\Excel\零用金查核表.xls"), MyExcel)
        Response.Clear()
        Response.ClearHeaders()
        Response.Buffer = True
        Response.ContentType = "application/octet-stream"
        Dim downloadfilename = "零用金查核表.xls"
        Response.AddHeader("Content-Disposition", "attachment;FileName=" + Uri.EscapeDataString(downloadfilename))
        Response.WriteFile(MyExcel)
        System.Web.HttpContext.Current.ApplicationInstance.CompleteRequest()
        Response.Flush()
        System.IO.File.Delete(MyExcel)
        Response.End()
    End Sub
    Protected Sub Download2(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim MyGUID As String = Guid.NewGuid().ToString("N")
        Dim MyWord As String = MapPath(".\Excel\Temp\") & MyGUID & ".docx"
        System.IO.File.Copy(MapPath(".\Excel\新年度零用金申請通知-110年.docx"), MyWord)
        Response.Clear()
        Response.ClearHeaders()
        Response.Buffer = True
        Response.ContentType = "application/octet-stream"
        dim downloadfilename2 = "新年度零用金申請通知-110年.docx"
        Response.AddHeader("Content-Disposition", "attachment;FileName=" + Uri.EscapeDataString(downloadfilename2))
        Response.WriteFile(MyWord)
        System.Web.HttpContext.Current.ApplicationInstance.CompleteRequest()
        Response.Flush()
        System.IO.File.Delete(MyWord)
        Response.End()
    End Sub
    protected Sub 登出M_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Session("Uid")=Nothing
        Session("type") = Nothing
        Session("atype") = Nothing
        Session("姓名") = Nothing
        Response.Redirect("Default.aspx")
    End Sub 
    protected Sub 回到上一頁_click(ByVal sender As Object, ByVal e As System.EventArgs)
    Response.write("<script language=javascript>history.go(-2);</script>)")
    End Sub 
End Class
