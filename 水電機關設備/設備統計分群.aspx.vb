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
Imports System.Drawing.Imaging
Imports System.Data.OleDb
Imports System.Drawing.Drawing2D
Partial Class 設備統計分群
    Inherits System.Web.UI.Page
    Dim con_14 As String = System.Web.Configuration.WebConfigurationManager.ConnectionStrings("ApplicationServices").ConnectionString
    Dim data As New SqlDataSource
    Dim data_dv As Data.DataView
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        data.ConnectionString = con_14
        If Not Page.IsPostBack Then
            Me.GridView1.PageIndex = Int32.MaxValue
            If Session("水_Uid")="3855"
                測試.Visible=True
            End If
        Else
            Me.Label1.Text = ""
            Me.Label2.Text = ""
        End If
    End Sub
    Protected Sub Insert(ByVal sender As Object, ByVal e As System.EventArgs)
        Update(sender, e)
        Me.SqlDataSource1.Insert()
        Me.GridView1.PageIndex = Int32.MaxValue
    End Sub
    Protected Sub Update(ByVal sender As Object, ByVal e As System.EventArgs)
        'TODO:轉換成整數會失敗
        For i = 0 To Me.GridView1.Rows.Count - 1
            Dim id As String = CType(Me.GridView1.Rows(i).FindControl("id"), TextBox).Text
            Dim 分群 As String = CType(Me.GridView1.Rows(i).FindControl("分群"),TextBox).Text
            Dim Update1 as string ="UPDATE 設備統計分群 SET " & _
            "分群 = NULLIF(N'" & 分群 & "', '') " & _
            "WHERE id = '" & id & "'"
            data.UpdateCommand = Update1
            data.Update()
            '將設備編號產生並排序
            ' data.UpdateCommand = "WITH CTE AS (Select *,Row_Number() OVER(Partition by Id_品項 order by ID ) AS '序號' From 水電機關設備資料表)" & _
            ' "UPDATE CTE SET 設備編號 = 序號"
            data.Update()
        Next
        Me.GridView1.DataBind()
    End Sub
    Protected Sub Delete(ByVal sender As Object, ByVal e As System.EventArgs)
    End Sub
    Protected Sub Test(ByVal sender As Object, ByVal e As System.EventArgs)
    End Sub
    Protected Sub GridView1_RowCommand(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles GridView1.RowCommand
        If e.CommandName = "刪除"
            '連同維護紀錄作業的相關資料一併刪除
            Update(sender, e)
            Dim i As Long = e.CommandSource.NamingContainer.RowIndex
            Dim id As String = CType(Me.GridView1.Rows(i).FindControl("id"), TextBox).Text
            data.DeleteCommand = "Delete 設備統計分群 Where id='" & id & "'"
            data.Delete()
            ' data.UpdateCommand = "Update 設備 Set " & _
            ' "項 = NULL," & _
            ' "財產編號 = NULL," & _
            ' "財產名稱 = NULL," & _
            ' "財產別名 = NULL," & _
            ' "廠牌 = NULL," & _
            ' "型號 = NULL," & _
            ' "購置日期 = NULL," & _
            ' "單位 = NULL," & _
            ' "數量 = NULL," & _
            ' "保管人 = NULL," & _
            ' "存置地點 = NULL " & _
            ' "WHERE id='" & id & "'"
            ' data.Update()
            Me.GridView1.DataBind()
        End If
    End Sub 
End Class