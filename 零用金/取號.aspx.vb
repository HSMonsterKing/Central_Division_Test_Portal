Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel.XlDirection
Imports Microsoft.Office.Interop.Excel.XlPageBreak
Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.VisualBasic.Logging
Imports System.IO
Imports System.Math
Imports System.Drawing
Imports System.Diagnostics
Imports System.Data.Sql
Imports System.Data.SqlTypes
Imports System.Data.SqlClient
Imports System.Configuration
Imports System.Text.RegularExpressions
Partial Class 取號
    Inherits System.Web.UI.Page
    Dim con_14 As String = System.Web.Configuration.WebConfigurationManager.ConnectionStrings("ApplicationServices").ConnectionString
    Dim data As New SqlDataSource
    Dim data_dv As Data.DataView
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        data.ConnectionString = con_14
        If Not Page.IsPostBack Then
            Me.年.Text = (DateTime.Now.Year - 1911).ToString()
            Me.GridView1.PageIndex = Int32.MaxValue
        Else
            Me.Label1.Text = ""
            Me.Label2.Text = ""
        End If 
    End Sub
    Protected Sub Insert(ByVal sender As Object, ByVal e As System.EventArgs)
        Update(sender, e)
        Me.SqlDataSource1.Insert()
        Me.GridView1.PageIndex = Int32.MaxValue
        Me.GridView1.DataBind()
        Label1.Text="新增成功"
    End Sub
    Protected Sub Update(ByVal sender As Object, ByVal e As System.EventArgs)
        'TODO:轉換成整數會失敗
        For i = 0 To Me.GridView1.Rows.Count - 1
            Dim id As String = CType(Me.GridView1.Rows(i).FindControl("id"), TextBox).Text
            Dim 單位別 As String = CType(Me.GridView1.Rows(i).FindControl("單位別"),DropDownList).Text
            Dim 承辦人 As String = CType(Me.GridView1.Rows(i).FindControl("承辦人"), DropDownList).Text
            Dim 種類 As String = CType(Me.GridView1.Rows(i).FindControl("種類"), TextBox).Text
            Dim 號數 As String = CType(Me.GridView1.Rows(i).FindControl("號數"), TextBox).Text
            Dim 備註 As String = CType(Me.GridView1.Rows(i).FindControl("備註"), TextBox).Text
            data.UpdateCommand = "UPDATE 收支備查簿 SET " & _
            "單位別 = NULLIF(N'" & 單位別 & "', ''), " & _
            "承辦人 = NULLIF(N'" & 承辦人 & "', ''), " & _
            "種類 = NULLIF(N'" & 種類 & "', ''), " & _
            "號數 = NULLIF(N'" & 號數 & "', ''), " & _
            "備註 = NULLIF(N'" & 備註 & "', '') " & _
            "WHERE id = '" & id & "'"
            data.Update()
        Next
        Label1.Text="已存檔成功"
    End Sub
    Protected Sub GridView1_RowCommand(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles GridView1.RowCommand
        If e.CommandName = "取號"
            Update(sender, e)
            Dim i As Long = e.CommandSource.NamingContainer.RowIndex
            Dim id As String = CType(Me.GridView1.Rows(i).FindControl("id"), TextBox).Text
            Dim 年 As String = Me.年.Text
            Dim 種類 As String = Me._種類.Text
            data.UpdateCommand = "UPDATE 收支備查簿 SET " & _
            "種類 = NULLIF(N'" & 種類 & "', ''),號數 = (SELECT ISNULL(MAX(號數) + 1, 1) FROM 收支備查簿 WHERE 年 = " & 年 & " AND _種類 = '" & 種類 & "') WHERE id=" & id
            data.Update()
            Me.GridView1.DataBind()
            Label1.Text="已取號成功"
        End If 
    End Sub
    <System.Web.Script.Services.ScriptMethod(), System.Web.Services.WebMethod()>
    Public Shared Function GetMyList(ByVal prefixText As String, ByVal count As Integer)
        Dim con_14 As String = System.Web.Configuration.WebConfigurationManager.ConnectionStrings("ApplicationServices").ConnectionString
        Dim data As New SqlDataSource
        Dim data_dv As Data.DataView
        Dim MyList As New List(Of String)
        data.ConnectionString = con_14
        data.SelectCommand = "SELECT TOP " & count & " 常用文字 FROM 常用清單 WHERE 常用文字 LIKE '%" & prefixText & "%' ORDER BY CASE WHEN 常用文字 IS NULL THEN 1 ELSE 0 END, 常用文字"
        data_dv = data.Select(New DataSourceSelectArguments)
        For i As Long = 0 To data_dv.Count() - 1
            MyList.Add(data_dv(i)(0).ToString())
        Next
        Return MyList
    End Function
End Class