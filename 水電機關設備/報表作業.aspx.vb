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
Partial Class 報表作業
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
    Protected Sub Import(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not FileUpload1.HasFile
            Label2.text="如要上傳資料，請先按'選擇檔案'"
            Exit Sub
        End If
        For Each PostedFile As HttpPostedFile In FileUpload1.PostedFiles
            Dim MyGUID As String = Guid.NewGuid().ToString("N")
            Dim Myfiles As String = MapPath(".\data\Temp\") & MyGUID
            PostedFile.SaveAs(Myfiles)
            Try
                File.Copy(Myfiles, MapPath(".\data\報表作業檔案\") & PostedFile.FileName, False)
            Catch
            End Try
            data.InsertCommand = _
                "IF NOT EXISTS(SELECT * FROM 報表作業上傳 WHERE 檔名 = N'" & PostedFile.FileName & "') " & _
                "BEGIN " & _
                "INSERT INTO 報表作業上傳 (檔名) " & _
                "VALUES('" & PostedFile.FileName & "')" & _
                "END"
            data.Insert()
            System.IO.File.Delete(Myfiles)
            Me.GridView1.PageIndex = Int32.MaxValue
        Next
        Me.GridView1.DataBind()
    End Sub
    Protected Sub test(ByVal sender As Object, ByVal e As System.EventArgs)
    End Sub
    Protected Sub GridView1_RowCommand(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles GridView1.RowCommand
        If e.CommandName = "刪除"
            Dim i As Long = e.CommandSource.NamingContainer.RowIndex
            Dim id As String = CType(Me.GridView1.Rows(i).FindControl("id"), TextBox).Text
            Dim 檔名 As String = CType(Me.GridView1.Rows(i).FindControl("檔名"), HyperLink).Text
            data.DeleteCommand = "DELETE FROM 報表作業上傳 WHERE id=" & id
            data.Delete()
            If 檔名<>""
                data.ConnectionString = con_14
                data.SelectCommand = "SELECT * FROM 報表作業上傳 Where 檔名 =N'" & 檔名 &"'"
                data_dv = data.Select(New DataSourceSelectArguments)
                If data_dv.count<1
                    System.IO.File.Delete(MapPath(".\data\報表作業檔案\") & 檔名)
                End If
            End If
            Me.GridView1.DataBind()
        End If
    End Sub
End Class