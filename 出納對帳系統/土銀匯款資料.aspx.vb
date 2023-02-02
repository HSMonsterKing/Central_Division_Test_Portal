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
Partial Class 土銀匯款資料
    Inherits System.Web.UI.Page
    Dim con_14 As String = System.Web.Configuration.WebConfigurationManager.ConnectionStrings("ApplicationServices").ConnectionString
    Dim data As New SqlDataSource
    Dim data_dv As Data.DataView
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        data.ConnectionString = con_14
        Me.Label1.Text = ""
        Me.Label2.Text = ""
        If Not Page.IsPostBack Then
            Me.GridView1.Columns(4).Visible = "False"
            Me.GridView1.PageIndex = Int32.MaxValue
        End If
    End Sub
    Protected Sub Insert(ByVal sender As Object, ByVal e As System.EventArgs)
        Me.SqlDataSource1.Insert()
        Me.GridView1.PageIndex = Int32.MaxValue
    End Sub
    Protected Sub GridView1_RowUpdated(ByVal sender As Object, ByVal e As GridViewUpdatedEventArgs)
        If e.Exception Is Nothing
            Me.Label1.Text = "成功"
            Me.Label2.Text = ""
            Me.GridView1.DataBind()
        Else
            Me.Label1.Text = ""
            Me.Label2.Text = "失敗"
        End If
    End Sub
End Class