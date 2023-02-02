Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel.XlDirection
Imports Microsoft.Office.Interop.Excel.XlPageBreak
Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.VisualBasic.Logging
Imports System.IO
Imports System.Collections.Generic
Imports System.Data.Odbc
Imports System.Math
Imports System.Diagnostics
Imports System.Data.Sql
Imports System.Data.SqlTypes
Imports System.Data.SqlClient
Imports System.Configuration
Imports System.Text.RegularExpressions
Imports System.Web.UI.WebControls
Imports System.Drawing
Imports System.Linq
Imports System.Security.Cryptography
Partial Class 登入
    Inherits System.Web.UI.Page
    Dim con_14 As String = System.Web.Configuration.WebConfigurationManager.ConnectionStrings("ApplicationServices").ConnectionString
    Dim data As New SqlDataSource
    Dim data_dv As Data.DataView
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        data.ConnectionString = con_14
        Session("title") = "登入"
        If not(Session("Uid") is Nothing)
        帳號.text=Session("Uid")
        登入.Visible=false
        登出.Visible=true
        End If 
    End Sub
    protected Sub 登入_Click(ByVal sender As Object, ByVal e As System.EventArgs)'密碼錯誤，會顯示錯誤
        If 帳號.Text<>"" And 密碼.Text<>""
            dim sql As String  = "SELECT * FROM [帳號] WHERE (帳號='" & 帳號.Text & "' AND  密碼='" & Psd() & "')"
            data.SelectCommand = sql
            data_dv = data.Select(New DataSourceSelectArguments)
            If not(data_dv is Nothing) AND (data_dv.count>0)
                Dim type As String =  data_dv(0)("type").ToString()
                Session("type") = type
                Session("atype") =  data_dv(0)("atype").ToString()
                dim last As String = "UPDATE 帳號 SET lastli = '" & DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") & "' WHERE 帳號 ='" & 帳號.Text & "'"
                data.UpdateCommand =last
                data.Update()
                Session("Uid") = 帳號.Text
                Session("姓名") = data_dv(0)("姓名").ToString()
                ' 設定 session 變數'Session["名稱"] = 值
                If type = "1"'
                    Session("atype") = "IsUserLogin"'芳
                Else If(type = "2")
                    Session("atype") = "IsDirectorLogin"'主任
                Else If(type = "3")
                    Session("atype") = "IsAccountantLogin"'主計
                Else If(type = "4")'zxcv
                    Session("atype") = "all"'除錯
                End If 
                Response.Redirect("提醒.aspx")
            Else
                Label1.Text = "帳號密碼錯誤" ' 錯誤訊息
            End If 
        Else
        Label1.Text = "帳號密碼未填寫"
        End If 
    End Sub
    Function Psd() As string
        dim SHA256 As sha256 = new SHA256CryptoServiceProvider()'建立一個SHA256
        dim source As byte() = Encoding.Default.GetBytes(密碼.Text)'將字串轉為Byte[]
        dim crypto As byte()  = sha256.ComputeHash(source)'進行SHA256加密
        dim result As string = Convert.ToBase64String(crypto)'把加密後的字串從Byte[]轉為字串
        return result
    End Function
    protected Sub 密碼轉換_Click(ByVal sender As Object, ByVal e As System.EventArgs)
      帳號.Text = Psd()
    End Sub 
    protected Sub 登出_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Session("Uid")=Nothing
        Session("type") = Nothing
        Session("atype") = Nothing
        Session("姓名") = Nothing
        Response.Redirect("Default.aspx")
    End Sub 
End Class