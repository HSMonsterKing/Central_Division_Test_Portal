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
        Session("水_title") = "登入"
        If not(Session("水_Uid") is Nothing)
            帳號.text=Session("水_Uid")
            登入.Visible=false
            登出.Visible=true
        Else
            Dim _IP As String = Request.ServerVariables("REMOTE_HOST")
            Dim r1, r2, r3 As Integer
            If Len(Trim(_IP)) <= 0 Then
                _IP = System.Security.Principal.WindowsIdentity.GetCurrent().Name
            End If
            r1 = Len(_IP)
            For i As Integer = 1 To r1
                If Mid(_IP, i, 1) = "\" Then
                    r2 = i
                End If
            Next
            r3 = r1 - r2
            _IP = Right(_IP, r3)
            '自動登入
            Dim sql as String  = "SELECT IP FROM 帳號 Where IP='" & _IP & "'"
            data.SelectCommand = sql
            data_dv = data.Select(New DataSourceSelectArguments)
            If data_dv.count>0
                sql = "SELECT * FROM 帳號 Where IP='" & _IP & "'"
                data.SelectCommand = sql
                data_dv = data.Select(New DataSourceSelectArguments)
                帳號.Text=data_dv(0)("帳號").ToString()
                If not(data_dv is Nothing) AND (data_dv.count>0)
                    Dim type As String = data_dv(0)("type").ToString()
                    Dim last as String = "UPDATE 帳號 SET lastli = '" & DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") & "' WHERE 帳號 ='" & 帳號.Text & "'"
                    data.UpdateCommand =last
                    data.Update()
                    ' 設定 session 變數
                    ' Session["名稱"] = 值
                    Session("水_Uid") = 帳號.Text
                    Session("水_atype") = data_dv(0)("atype").ToString()
                    Session("水_帳號名") = data_dv(0)("姓名").ToString()
                    data.SelectCommand = "SELECT DATEDIFF(day,(SELECT psdtime FROM 帳號 where 帳號='" & Session("水_Uid") & "'),GETDATE())"
                    data_dv = data.Select(New DataSourceSelectArguments)
                    If  data_dv(0)(0) > 180
                        'Response.Write("<Script language='JavaScript'>alert('密碼已經超過3個月未修改，請修改密碼！');location.href('./修改密碼.aspx');</Script>")'盡量別使用Response.Write，會破壞排版
                        Response.Redirect("修改密碼.aspx")
                    Else
                        If Session("水_atype")="IsSRLogin"
                             Response.Redirect("招標列印.aspx")
                        Else
                            Response.Redirect("建置作業.aspx")
                        End If
                    End If
                End If
            End If
        End If
    End Sub
    protected Sub 登入_Click(ByVal sender As Object, ByVal e As System.EventArgs)'密碼錯誤，會顯示錯誤
        'cn = new OdbcConnection("Dsn=SRB;uid=lusu666;pwd=zxcvbASDFG")
        If 帳號.Text<>"" and 密碼.Text<>""
            Dim sql as String  = "SELECT * FROM 帳號 WHERE (帳號='" & 帳號.Text & "' AND  密碼='" & Psd() & "')"
            data.SelectCommand = sql
            data_dv = data.Select(New DataSourceSelectArguments)
            If not(data_dv is Nothing) AND (data_dv.count>0)
                Dim type As String = data_dv(0)("type").ToString()
                Dim last as String = "UPDATE 帳號 SET lastli = '" & DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") & "' WHERE 帳號 ='" & 帳號.Text & "'"
                data.UpdateCommand =last
                data.Update()
                ' 設定 session 變數
                ' Session["名稱"] = 值
                Session("水_Uid") = 帳號.Text
                Session("水_atype") = data_dv(0)("atype").ToString()
                If Session("水_Uid")=true
                    Session("水_帳號名") = data_dv(0)("姓名").ToString()
                    data.SelectCommand = "SELECT DATEDIFF(day,(SELECT psdtime FROM 帳號 where 帳號='" & Session("水_Uid") & "'),GETDATE())"
                    data_dv = data.Select(New DataSourceSelectArguments)
                    If  data_dv(0)(0) > 180
                        'Response.Write("<Script language='JavaScript'>alert('密碼已經超過3個月未修改，請修改密碼！');location.href('./修改密碼.aspx');</Script>")'盡量別使用Response.Write，會破壞排版
                        Response.Redirect("修改密碼.aspx")
                    Else
                        If Session("水_atype")="IsSRLogin"
                             Response.Redirect("招標列印.aspx")
                        Else
                            Response.Redirect("建置作業.aspx")
                        End If
                    End If
                End If
            Else
                Label2.Text = "帳號密碼錯誤" ' 錯誤訊息
            End If
        Else
            Label2.Text = "帳號密碼未填寫"
        End If
    End Sub
    Function Psd() As string
        Dim SHA256 as sha256 = new SHA256CryptoServiceProvider()'建立一個SHA256
        Dim source as byte() = Encoding.Default.GetBytes(密碼.Text)'將字串轉為Byte[]
        Dim crypto as byte()  = sha256.ComputeHash(source)'進行SHA256加密
        Dim result as string = Convert.ToBase64String(crypto)'把加密後的字串從Byte[]轉為字串
        return result
    End Function
    protected Sub 密碼轉換_Click(ByVal sender As Object, ByVal e As System.EventArgs)
      帳號.Text = Psd()
    End Sub
    protected Sub 登出_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Session("水_IsUserLogin") = Nothing
        Session("水_IsDirectorLogin") = Nothing
        Session("水_Uid")=Nothing
        Session("水_atype") = Nothing
        Response.Redirect("Default.aspx")
    End Sub
End Class