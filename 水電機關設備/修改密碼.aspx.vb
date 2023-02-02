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
Partial Class 修改密碼
    Inherits System.Web.UI.Page
    Dim con_14 As String = System.Web.Configuration.WebConfigurationManager.ConnectionStrings("ApplicationServices").ConnectionString
    Dim data As New SqlDataSource
    Dim data_dv As Data.DataView
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        data.ConnectionString = con_14
        Session("水_title") = "修改密碼"
        If not(Session("水_Uid") is Nothing)
            帳號.text=Session("水_Uid")
        End If
        If Session("水_Uid")=true'3/1 已改進警告完不會跳網頁
            data.SelectCommand = "SELECT DATEDIFF(day,(SELECT psdtime FROM 帳號 where 帳號='" & Session("水_Uid") & "'),GETDATE())"
            data_dv = data.Select(New DataSourceSelectArguments)
            If  data_dv(0)(0) > 180
                RedLabel.text="密碼已經超過6個月未修改，請修改密碼！"
            End If
        End If
    End Sub
    protected Sub 修改密碼_Click(ByVal sender As Object, ByVal e As System.EventArgs)'密碼錯誤，會顯示錯誤
        If (帳號.Text<>"" and 密碼.Text<>"" and 密碼.Text = 確認密碼.Text)
            Dim sql as String  = "SELECT * FROM [帳號] WHERE (帳號='" & 帳號.Text & "')"
            data.SelectCommand = sql
            data_dv = data.Select(New DataSourceSelectArguments)
            If not(data_dv is Nothing) AND (data_dv.count>0)
                Dim type As String =  data_dv(0)("type").ToString()
                Dim last as String = "UPDATE 帳號 SET lastli = '" & DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") & "',密碼='" & Psd() & "',psdtime= '" & DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") & "' WHERE 帳號 ='" & 帳號.Text & "'"
                data.UpdateCommand =last
                data.Update()
                Session.Clear()
                ' 設定 session 變數
                ' Session["名稱"] = 值
            End If
            Response.Redirect("登入.aspx")
        Else
            Label1.Text = "帳號密碼未填寫或確認密碼與密碼不相同"
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
End Class