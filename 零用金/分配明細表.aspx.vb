Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel.XlDirection
Imports Microsoft.Office.Interop.Excel.XlPageBreak
Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.VisualBasic.Logging
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
Partial Class 分配明細表
    Inherits System.Web.UI.Page
    Dim con_14 As String = System.Web.Configuration.WebConfigurationManager.ConnectionStrings("ApplicationServices").ConnectionString
    Dim data As New SqlDataSource
    Dim data_dv As Data.DataView
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        data.ConnectionString = con_14
        If Not Page.IsPostBack Then
            Me.年.Text = (DateTime.Now.Year - 1911).ToString()
            Me.月.Text = DateTime.Now.month.ToString()
            Me.日.Text = DateTime.Now.day.ToString()
            Me.GridView1.PageIndex = Int32.MaxValue
        Else
            Me.Label1.Text = ""
            Me.Label2.Text = ""
        End If 
    End Sub
    Protected Sub Insert(ByVal sender As Object, ByVal e As System.EventArgs)
        Update(sender, e)
        Dim 年 As String = Me.年.Text
        Me.SqlDataSource1.Insert()
        Me.GridView1.PageIndex = Int32.MaxValue
        Me.GridView1.DataBind()
        Label1.Text="新增成功"
    End Sub
    Protected Sub Update(ByVal sender As Object, ByVal e As System.EventArgs)
        'TODO:轉換成整數會失敗
        For i = 0 To Me.GridView1.Rows.Count - 1
            Dim id As String = CType(Me.GridView1.Rows(i).FindControl("id"), TextBox).Text
            Dim 零用金使用單位 As String = CType(Me.GridView1.Rows(i).FindControl("零用金使用單位"), TextBox).Text
            Dim 金額 As String = CType(Me.GridView1.Rows(i).FindControl("金額"), TextBox).Text
            金額=金額.Replace(",", "").Replace("N", "").Replace("T", "").Replace("$", "")
            Dim 累計 As String = CType(Me.GridView1.Rows(i).FindControl("累計"), Label).Text
            data.UpdateCommand = "UPDATE 分配明細表 SET " & _
            "零用金使用單位 = NULLIF(N'" & 零用金使用單位 & "', ''), " & _
            "金額 = NULLIF(N'" & 金額 & "', '') " & _
            "WHERE id = '" & id & "'"
            data.Update()
        Next
        Me.GridView1.DataBind()
        '重算餘額
        data.UpdateCommand = _
            "WITH CTE AS " & _
            "(SELECT *, " & _
                "(CASE WHEN ISNULL(金額,0) = 0 THEN 累計 ELSE 0 END)" & _
                " + " & _
                "(SUM(金額) OVER (ORDER BY id " & _
                "ROWS BETWEEN UNBOUNDED PRECEDING AND CURRENT ROW))" & _
                "AS RunningTotal " & _
            "FROM 分配明細表) " & _
            "UPDATE CTE SET 累計 = RunningTotal"
            data.Update()
        Label1.Text="存檔成功"
    End Sub
    Protected Sub Delete(ByVal sender As Object, ByVal e As System.EventArgs)
        Update(sender, e)
        Me.GridView1.PageIndex = Int32.MaxValue
        Me.GridView1.DataBind()
        dim i As Int32 = Me.GridView1.Rows.Count - 1
        Dim id As String = CType(Me.GridView1.Rows(i).FindControl("id"), TextBox).Text
        data.DeleteCommand = "DELETE FROM 分配明細表 " & _
            "WHERE id = '" & id & "'"
        data.Delete()
        Me.GridView1.DataBind()
        Label1.Text="刪除成功"
    End Sub
End Class