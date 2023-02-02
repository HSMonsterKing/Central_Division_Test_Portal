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
Partial Class 維護契約
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
    Protected Sub Insert(ByVal sender As Object, ByVal e As System.EventArgs)'維修金額預設值寫在資料庫
        Update(sender, e)
        For i = 1 To 15
            data.InsertCommand = _
            "INSERT INTO 維護契約表 " & _
            "( _頁, _列) " & _
            "VALUES " & _
            "(" & (Me.GridView1.PageCount + 1).ToString() & ", " & i & ")"
            data.Insert()
        Next
        Me.GridView1.PageIndex = Int32.MaxValue
        Me.GridView1.DataBind()
    End Sub
    Protected Sub Update(ByVal sender As Object, ByVal e As System.EventArgs)
        For i = 0 To Me.GridView1.Rows.Count - 1
            Dim id As String = CType(Me.GridView1.Rows(i).FindControl("id"), TextBox).Text
            Dim 契約名稱 As String = CType(Me.GridView1.Rows(i).FindControl("契約名稱"), TextBox).Text
            Dim ID_廠商 As String = CType(Me.GridView1.Rows(i).FindControl("ID_廠商"), TextBox).Text
            Dim 維護廠商 As String = CType(Me.GridView1.Rows(i).FindControl("維護廠商"), TextBox).Text
            Dim 廠商電話 As String = CType(Me.GridView1.Rows(i).FindControl("廠商電話"), TextBox).Text
            Dim 維護內容 As String = CType(Me.GridView1.Rows(i).FindControl("維護內容"), TextBox).Text
            Dim 維護頻率 As String = CType(Me.GridView1.Rows(i).FindControl("維護頻率"), DropDownList).Text
            Dim 備註 As String = CType(Me.GridView1.Rows(i).FindControl("備註"), TextBox).Text
            '輸入廠商時能自動代出廠商全名、廠商電話，假如沒有此廠商且不為空，則新增，，否則更新，但電話不能為空
            data.UpdateCommand = "set transaction isolation level serializable; "& _
                "begin tran " & _
                    "If exists (select * from 廠商資料 with(xlock) where 廠商 = '" & 維護廠商 & "') "& _
                    "begin If '" & 廠商電話 & "'IS NOT NULL AND '" & 廠商電話 & "'<>'' " & _
                        "begin " & _
                            "update 廠商資料 set 電話 = '" & 廠商電話 & "' Where 廠商 = '" & 維護廠商 & "' "& _
                        "End " & _
                    "End " & _
                    "Else If '" & 維護廠商 & "'IS NOT NULL AND '" & 維護廠商 & "'<>'' " & _
                        "begin " & _
                            "insert 廠商資料 values ('" & 維護廠商 & "','" & 廠商電話 & "') " & _
                        "End " & _
                "commit"
            data.Update()
            data.UpdateCommand = "UPDATE 維護契約表 SET " & _
            "契約名稱 = NULLIF(N'" & 契約名稱 & "', ''), " & _
            "ID_廠商 = NULLIF((Select ID from 廠商資料 Where 廠商='" & 維護廠商 &"'), '')," & _
            "維護內容 = NULLIF(N'" & 維護內容 & "', ''), " & _
            "維護頻率 = NULLIF(N'" & 維護頻率 & "', ''), " & _
            "備註 = NULLIF(N'" & 備註 & "', '') " & _
            "WHERE id = '" & id & "'"
            data.Update()
        Next
        Me.GridView1.DataBind()
    End Sub
    Protected Sub Delete(ByVal sender As Object, ByVal e As System.EventArgs)
        Me.GridView1.PageIndex = Int32.MaxValue
        Me.GridView1.DataBind()
        Dim delete1 as string=""
        For i = 0 To Me.GridView1.Rows.Count - 1
            Dim id As String = CType(Me.GridView1.Rows(i).FindControl("id"), TextBox).Text
            delete1="DELETE FROM 維護契約表 " & _
            "WHERE id = '" & id & "'"
            data.deleteCommand =delete1
            data.delete()
        Next
        Me.GridView1.DataBind()
    End Sub
    Protected Sub test(ByVal sender As Object, ByVal e As System.EventArgs)
    End Sub
    Protected Sub 返回_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Panel3.Visible=true
        Panel4.Visible=False
        Me.ID_維修契約.Text=""
    End Sub
    <System.Web.Script.Services.ScriptMethod(), System.Web.Services.WebMethod()>
    Public Shared Function GetMyList(ByVal prefixText As String, ByVal count As Integer)'向全表單分享GetMyList這個方法
        Dim con_14 As String = System.Web.Configuration.WebConfigurationManager.ConnectionStrings("ApplicationServices").ConnectionString
        Dim data As New SqlDataSource
        Dim data_dv As Data.DataView
        Dim MyList As New List(Of String)
        data.ConnectionString = con_14
        data.SelectCommand = "SELECT TOP " & count & " 廠商 FROM 廠商資料 WHERE 廠商 LIKE '%" & prefixText & "%' ORDER BY CASE WHEN 廠商 IS NULL THEN 1 ELSE 0 END, 廠商"
        data_dv = data.Select(New DataSourceSelectArguments)
        For i As Long = 0 To data_dv.Count() - 1
            MyList.Add(data_dv(i)(0).ToString())
        Next
        Return MyList
    End Function
    Protected Sub GridView1_RowCommand(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles GridView1.RowCommand
        If e.CommandName = "上傳資料"'上傳資料按紐
            Dim i As Long = e.CommandSource.NamingContainer.RowIndex
            Dim id As String = CType(Me.GridView1.Rows(i).FindControl("id"), TextBox).Text
            Dim 契約名稱 As String = CType(Me.GridView1.Rows(i).FindControl("契約名稱"), TextBox).Text
            If Not Me.FileUpload1.HasFile
                Label2.text="如要上傳資料，請先按'選擇檔案'"
                Exit Sub
            End If
            For Each PostedFile As HttpPostedFile In Me.FileUpload1.PostedFiles
                Dim MyGUID As String = Guid.NewGuid().ToString("N")
                Dim Myfiles As String = MapPath(".\data\Temp\") & MyGUID
                PostedFile.SaveAs(Myfiles)
                Try
                    File.Copy(Myfiles, MapPath(".\data\維修契約維護紀錄\") & PostedFile.FileName, False)
                Catch
                End Try
                '假如維護契約有編號，則新增編號和檔案，否則增加新的檔案，
                data.InsertCommand = _
                    "IF NOT EXISTS(SELECT * FROM 維護契約表 left Join 維護契約表_維護紀錄 on 維護契約表.ID_維護紀錄 = 維護契約表_維護紀錄.ID_維修契約 WHERE 維護契約_維修紀錄 = N'" & PostedFile.FileName & "' AND 維護契約表.id= '" & id & "' ) " & _
	                "BEGIN " & _
	                "Insert into 維護契約表_維護紀錄 " & _
	                "(ID_維修契約,維護契約_維修紀錄) " & _
	                "VALUES " & _
	                "("& id &", N'" & PostedFile.FileName & "') " & _
                    "" & _
	                "UPDATE 維護契約表 SET " & _
	                "ID_維護紀錄 = (SELECT DISTINCT 維護契約表_維護紀錄.ID_維修契約 FROM 維護契約表_維護紀錄 WHERE 維護契約_維修紀錄 = N'" & PostedFile.FileName & "' AND ID_維修契約=" & id & ") " & _
	                "WHERE id = '" & id & "'" & _
	                "END "
                data.Insert()
                System.IO.File.Delete(Myfiles)
            Next
            Me.GridView1.DataBind()
            Dim ID_維護紀錄 As String = CType(Me.GridView1.Rows(i).FindControl("ID_維護紀錄"), TextBox).Text
            Panel3.Visible=False
            Panel4.Visible=true
            Me.ID_維修契約.Text=ID_維護紀錄
            Label4_1.Text=契約名稱
            Me.GridView2.DataBind()
        ElseIf e.CommandName = "維護紀錄"'顯示該筆資料的維護紀錄
            '取ID
            Dim i As Long = e.CommandSource.NamingContainer.RowIndex
            Dim ID_維護紀錄 As String = CType(Me.GridView1.Rows(i).FindControl("ID_維護紀錄"), TextBox).Text
            Dim 契約名稱 As String = CType(Me.GridView1.Rows(i).FindControl("契約名稱"), TextBox).Text
            Panel3.Visible=False
            Panel4.Visible=true
            Me.ID_維修契約.Text=ID_維護紀錄
            Label4_1.Text=契約名稱
            Me.GridView2.DataBind()
        ElseIf e.CommandName = "刪除"
            Update(sender,e)
            Dim i As Long = e.CommandSource.NamingContainer.RowIndex
            Dim id As String = CType(Me.GridView1.Rows(i).FindControl("id"), TextBox).Text
            data.UpdateCommand = _
                    "UPDATE 維護契約表 SET " & _
	                "契約名稱=NULL,ID_廠商=NULL,維護內容=NULL,維護頻率=NULL,ID_維護紀錄=NULL,維護紀錄=NULL,備註=NULL " & _
	                "WHERE id = '" & id &"' "
            data.Update()
            data.ConnectionString = con_14
            data.SelectCommand = "SELECT 維護契約_維修紀錄 FROM 維護契約表_維護紀錄 Where ID_維修契約=" & id
            data_dv = data.Select(New DataSourceSelectArguments)
            Dim data_dv2 As Data.DataView
            data.DeleteCommand = "Delete From 維護契約表_維護紀錄 Where ID_維修契約=" & id
            data.Delete()
                for j  As Long = 0 To data_dv.Count() - 1
                    data.SelectCommand = "SELECT * FROM 維護契約表_維護紀錄 Where 維護契約_維修紀錄 =N'" & data_dv(j)("維護契約_維修紀錄").ToString() &"'"
                    data_dv2 = data.Select(New DataSourceSelectArguments)
                    If data_dv2.count<1
                        System.IO.File.Delete(MapPath(".\data\維修契約維護紀錄\") & data_dv(j)("維護契約_維修紀錄").ToString())
                    End If
                Next
            Me.GridView1.DataBind()
            Me.GridView2.DataBind()
        End If
    End sub
    Protected Sub GridView2_RowCommand(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles GridView2.RowCommand
        If e.CommandName = "刪除"
            Dim i As Long = e.CommandSource.NamingContainer.RowIndex
            Dim id As String = CType(Me.GridView2.Rows(i).FindControl("id"), TextBox).Text
            Dim 維護契約_維修紀錄 As String = CType(Me.GridView2.Rows(i).FindControl("檔名"), HyperLink).Text
            '重置收支備查簿
            data.SelectCommand = "SELECT count(維護契約_維修紀錄) As 資料數 FROM 維護契約表_維護紀錄 Where 維護契約_維修紀錄 = '" & 維護契約_維修紀錄 & "' Group by 維護契約_維修紀錄"
            data_dv = data.Select(New DataSourceSelectArguments)
            data.DeleteCommand = "Delete From 維護契約表_維護紀錄 Where id=" & id
            data.Delete()
            If data_dv(0)("資料數").ToString()=1'除非只剩一筆，否則不刪除文件檔案，因為有別的資料參照
                System.IO.File.Delete(MapPath(".\data\維修契約維護紀錄\") & 維護契約_維修紀錄)
            End If
            Me.GridView2.DataBind()
        End If
    End sub
End Class