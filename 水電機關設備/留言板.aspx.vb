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
Partial Class 留言板
    Inherits System.Web.UI.Page
    Dim con_14 As String = System.Web.Configuration.WebConfigurationManager.ConnectionStrings("ApplicationServices").ConnectionString
    Dim data As New SqlDataSource
    Dim data_dv As Data.DataView
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        data.ConnectionString = con_14
        If Not Page.IsPostBack Then
            Me.GridView1.PageIndex = Int32.MaxValue
        Else
            Me.Label1.Text = ""
            Me.Label2.Text = ""
        End If 
    End Sub
    Protected Sub Insert(ByVal sender As Object, ByVal e As System.EventArgs)
        Update(sender, e)
            dim insert1 As string
            data.InsertCommand = _
                "INSERT INTO 留言 " & _
                "(編號,留言) " & _
                "VALUES " & _
                "((SELECT ISNULL(MAX(編號) + 1, 1) FROM 留言),NULL)"
            data.Insert()
        Me.GridView1.PageIndex = Int32.MaxValue
        Me.GridView1.DataBind()
        Label1.Text="已新增成功"
    End Sub
    Protected Sub Update(ByVal sender As Object, ByVal e As System.EventArgs)'更新
        'TODO:轉換成整數會失敗
        For i = 0 To Me.GridView1.Rows.Count - 1
            Dim id As String = CType(Me.GridView1.Rows(i).FindControl("id"), TextBox).Text
            Dim 留言 As String = CType(Me.GridView1.Rows(i).FindControl("留言板"), TextBox).Text
             dim Update1 As string ="UPDATE 留言 SET " & _
            "留言 = NULLIF(N'" & 留言 & "', '') " & _
            "WHERE id = '" & id & "'"
            data.UpdateCommand = Update1
            data.Update()
        Next
        Me.GridView1.DataBind()
        Label1.Text="已存檔成功"
    End Sub
    Protected Sub test(ByVal sender As Object, ByVal e As System.EventArgs)
    End Sub
    Protected Sub Delete(ByVal sender As Object, ByVal e As System.EventArgs)'刪除
        Update(sender, e)
        Me.GridView1.PageIndex = Int32.MaxValue
        Me.GridView1.DataBind()
        dim delete1 As string=""
        dim i As Int32 = Me.GridView1.Rows.Count - 1
        Dim id As String = CType(Me.GridView1.Rows(i).FindControl("id"), TextBox).Text
        delete1="DELETE FROM 留言 " & _
        "WHERE id = '" & id & "'"
        data.DeleteCommand =delete1
        data.Delete()
        Me.GridView1.DataBind()
        Label1.Text="已刪除成功"
    End Sub
End Class