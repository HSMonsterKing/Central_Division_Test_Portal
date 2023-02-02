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
Partial Class 建立帳號
    Inherits System.Web.UI.Page
    Dim con_14 As String = System.Web.Configuration.WebConfigurationManager.ConnectionStrings("ApplicationServices").ConnectionString
    Dim data As New SqlDataSource
    Dim data_dv As Data.DataView
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        data.ConnectionString = con_14
        Session("水_title") = "建立帳號"
        If Session("水_Uid") is Nothing
            Response.Redirect("登入.aspx")
        End If
    End Sub
    protected Sub 建立帳號_Click(ByVal sender As Object, ByVal e As System.EventArgs)'密碼錯誤，會顯示錯誤
        If (帳號.Text<>"" And 密碼.Text<>"" And 密碼.Text = 確認密碼.Text)
            Dim sql as String  = "SELECT * FROM [帳號] WHERE (帳號='" & 帳號.Text & "')"
            data.SelectCommand = sql
            data_dv = data.Select(New DataSourceSelectArguments)
            If not(data_dv is Nothing) AND (data_dv.count=0)
                data.InsertCommand = "Insert Into [帳號] (帳號,密碼,姓名) Values ('" & 帳號.Text & "','" & Psd() & "','" & 姓名.Text & "')"
                data.Insert()
                Label1.Text = "新增成功"
                ' 設定 session 變數
                ' Session["名稱"] = 值
            Else
                Label1.Text = "此帳號已存在"
            End If
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
      密碼轉換後.Text = Psd()
    End Sub
End Class