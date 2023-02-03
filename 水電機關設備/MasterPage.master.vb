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
        Dim 權限() As object={更新資料,修改密碼,登出}
        Dim 權限林() As object={建置作業,品項資料,報表作業,查詢作業,零星維修作業,維護契約,例行故障維修,廠商資料,設備統計,設備統計分群,濾心更換週期,水質檢驗週期,濾心日誌,留言板}
        IF Session("水_atype")="all"
            建立帳號.Visible=True
        End If
        If Session("水_Uid") = true
            帳號名.text=Session("水_帳號名")
            for i=0 to 權限.count-1
                權限(i).Visible=True
            Next
            If Session("水_atype")="IsUserLogin" OR Session("水_atype")="IsDirectorLogin" OR Session("水_atype")="all"
                for i=0 to 權限林.count-1
                    權限林(i).Visible=True
                Next
            End if
            If Session("水_atype")="all"
                測試.Visible=True
            End if
            登入.Visible=false
            data.SelectCommand = "SELECT DATEDIFF(day,(SELECT psdtime FROM 帳號 where 帳號='" & Session("水_Uid") & "'),GETDATE())"
            data_dv = data.Select(New DataSourceSelectArguments)
            If  data_dv(0)(0) > 180 And System.IO.Path.GetFileName(Request.PhysicalPath)<>"修改密碼.aspx"
                'Response.Write("<Script language='JavaScript'>alert('密碼已經超過3個月未修改，請修改密碼！');location.href('./修改密碼.aspx');</Script>")'盡量別使用Response.Write，會破壞排版
                Response.Redirect("修改密碼.aspx")
            End If
            for i=0 to 權限.count-1
                If System.IO.Path.GetFileName(Request.PhysicalPath)="建立帳號.aspx" And 帳號名.text<>"除錯模式"
                    Response.write("<script language=javascript>history.go(-1);</script>)")
                End If
            Next
        End If
        If (System.IO.Path.GetFileName(Request.PhysicalPath)<>"登入.aspx" And Session("水_Uid")=Nothing) 
            Response.Redirect("Default.aspx")
        End If
        
    End Sub
    protected Sub 登出M_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Session("水_Uid")=Nothing
        Session("水_atype") = Nothing
        Response.Redirect("Default.aspx")
    End Sub
    protected Sub 回到上一頁_click(ByVal sender As Object, ByVal e As System.EventArgs)
    Response.write("<script language=javascript>history.go(-2);</script>)")
    End Sub
End Class
