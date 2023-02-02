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
Partial Class 查核資料
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
        Me.SqlDataSource1.Insert()
        Me.GridView1.PageIndex = Int32.MaxValue
        Me.GridView1.DataBind()
        Label1.Text="新增成功"
    End Sub
    Protected Sub Update(ByVal sender As Object, ByVal e As System.EventArgs)'更新
        'TODO:轉換成整數會失敗
        For i = 0 To Me.GridView1.Rows.Count - 1
            Dim id As String = CType(Me.GridView1.Rows(i).FindControl("id"), TextBox).Text
            Dim 查核類型 As String = CType(Me.GridView1.Rows(i).FindControl("查核類型"), DropDownList).Text
            Dim 查核時間 As String = nothing
            If CType(Me.GridView1.Rows(i).FindControl("查核時間"), TextBox).text<>nothing
                查核時間 = CType(Me.GridView1.Rows(i).FindControl("查核時間"), TextBox).Text
                查核時間 = taiwancalendarto(查核時間)
            End If 
            Dim 查核地點 As String = CType(Me.GridView1.Rows(i).FindControl("查核地點"), TextBox).Text
            Dim 查核人員 As String = CType(Me.GridView1.Rows(i).FindControl("查核人員"), TextBox).Text
            Dim 備註 As String = CType(Me.GridView1.Rows(i).FindControl("備註"), TextBox).Text
            data.UpdateCommand = "UPDATE 查核資料表 SET " & _
            "查核類型 = NULLIF(N'" & 查核類型 & "', ''), " & _
            "查核時間 = (CASE WHEN ISDATE(NULLIF(N'" & 查核時間 & "', ''))=1 Then NULLIF(N'" & 查核時間 & "', '') End ), " & _
            "查核地點 = NULLIF(N'" & 查核地點 & "', ''), " & _
            "查核人員 = NULLIF(N'" & 查核人員 & "', ''), " & _
            "備註 = NULLIF(N'" & 備註 & "', '') " & _
            "WHERE id = '" & id & "'"
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
        If Panel3.Visible=True
            dim i As Int32 = Me.GridView1.Rows.Count - 1
            Dim id As String = CType(Me.GridView1.Rows(i).FindControl("id"), TextBox).Text
            data.DeleteCommand = "DELETE FROM 查核資料表 " & _
            "WHERE id = '" & id & "'"
            data.Delete()
            data.DeleteCommand = "DELETE FROM 查核資料上傳 " & _
            "WHERE id_查核 = '" & id & "'"
            data.Delete()
        Else
            dim i As Int32 = Me.GridView2.Rows.Count - 1
            Dim id As String = CType(Me.GridView2.Rows(i).FindControl("id"), TextBox).Text
            Dim id_查核 As String = CType(Me.GridView2.Rows(i).FindControl("id_查核"), TextBox).Text
            If Me.GridView2.Rows.count = 1
                data.UpdateCommand = "UPDATE 查核資料表 SET " & _
                    "查核資料 = 'False' " & _
                    "WHERE id = '" & id_查核 & "'"
                data.Update()
            End If 
            data.DeleteCommand = "DELETE FROM 查核資料上傳 " & _
            "WHERE id = '" & id & "'"
            data.Delete()
        End If 
        Me.GridView1.DataBind()
        Me.GridView2.DataBind()
        Label1.Text="刪除成功"
    End Sub
    Protected Sub GridView1_RowCommand(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles GridView1.RowCommand
        If e.CommandName = "上傳資料"'上傳資料按紐
            Dim i As Long = e.CommandSource.NamingContainer.RowIndex
            Dim id As String = CType(Me.GridView1.Rows(i).FindControl("id"), TextBox).Text
            If Not FileUpload1.HasFile
                label2.Text="請先按下[選擇檔案]按鈕，再選擇要上傳的檔案"
                Exit Sub
            End If 
            For Each PostedFile As HttpPostedFile In FileUpload1.PostedFiles
                Dim MyGUID As String = Guid.NewGuid().ToString("N")
                Dim Myfiles As String = MapPath(".\data\Temp\") & MyGUID
                PostedFile.SaveAs(Myfiles)
                Try
                    File.Copy(Myfiles, MapPath(".\data\查核資料\") & PostedFile.FileName, False)
                Catch
                End Try
                data.InsertCommand = _
                    "IF NOT EXISTS(SELECT * FROM 查核資料上傳 WHERE 查核資料 = N'" & PostedFile.FileName & "'and id_查核=" & id & ") "  & _
                    "BEGIN " & _
                    "INSERT INTO 查核資料上傳 " & _
                    "(id_查核, 查核資料) " & _
                    "VALUES " & _
                    "(N'" & id & "',NULLIF(N'" & PostedFile.FileName & "', ''))" & _
                    "END"
                data.Insert()
                data.UpdateCommand = _
                    "IF EXISTS(SELECT * FROM 查核資料表 WHERE id=" & id & " And 查核資料='False') "  & _
                    "BEGIN " & _
                    "UPDATE 查核資料表 SET " & _
                    "查核資料 = 'True' " & _
                    "WHERE id = '" & id & "'" & _
                    "END"
                data.Update()
                System.IO.File.Delete(Myfiles)
            Next
            Me.GridView1.DataBind()
            Label1.Text="已上傳成功"
        ElseIf e.CommandName = "select_click"
            Dim i As Long = e.CommandSource.NamingContainer.RowIndex
            Dim id As String = CType(Me.GridView1.Rows(i).FindControl("id"), TextBox).Text
            Panel3.Visible=False
            Panel4.Visible=true
            SqlDataSource2.SelectCommand="SELECT * FROM 查核資料上傳 where id_查核 = " & id
        End If 
    End Sub
    Protected Sub 返回_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Panel3.Visible=true
        Panel4.Visible=False
    End Sub
End Class