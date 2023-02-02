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
Partial Class 零星維修作業
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
    Protected Sub Insert(ByVal sender As Object, ByVal e As System.EventArgs)'維修金額預設值寫在資料庫
        Update(sender, e)
        For i = 1 To 15
            data.InsertCommand = _
            "INSERT INTO 零星維修作業 " & _
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
            Dim 列 As String = CType(Me.GridView1.Rows(i).FindControl("_列"), TextBox).Text
            Dim 叫修日期 As String = nothing
            If CType(Me.GridView1.Rows(i).FindControl("叫修日期"), TextBox).text<>""
                叫修日期 = CType(Me.GridView1.Rows(i).FindControl("叫修日期"), TextBox).text
                叫修日期 = taiwancalendarto(叫修日期)
            End If
            Dim 維修日期 As String = nothing
            If CType(Me.GridView1.Rows(i).FindControl("維修日期"), TextBox).text<>""
                維修日期 = CType(Me.GridView1.Rows(i).FindControl("維修日期"), TextBox).text
                維修日期 = taiwancalendarto(維修日期)
            End If
            Dim 維修金額 As String = Mid(CType(Me.GridView1.Rows(i).FindControl("維修金額"), TextBox).Text,4)
            維修金額=維修金額.Replace(",", "")
            Dim 維修內容 As String = CType(Me.GridView1.Rows(i).FindControl("維修內容"), TextBox).Text
            Dim ID_廠商 As String = CType(Me.GridView1.Rows(i).FindControl("ID_廠商"), TextBox).Text
            Dim 維護廠商 As String = CType(Me.GridView1.Rows(i).FindControl("維護廠商"), TextBox).Text
            Dim 廠商電話 As String = CType(Me.GridView1.Rows(i).FindControl("廠商電話"), TextBox).Text
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
            data.UpdateCommand = "UPDATE 零星維修作業 SET " & _
            "_列 = ISNULL(NULLIF(N'" & 列 & "', ''),_列), " & _
            "叫修日期 = IIF(ISDATE(TRIM(N'" & 叫修日期 & "'))=1,TRIM(N'" & 叫修日期 & "'),NULL), " & _
            "維修日期 = IIF(ISDATE(TRIM(N'" & 維修日期 & "'))=1,TRIM(N'" & 維修日期 & "'),NULL), " & _
            "維修金額 = REPLACE(ISNULL(NULLIF('" & 維修金額 & "', ''),'0'), ',', ''), " & _
            "維修內容 = NULLIF(N'" & 維修內容 & "', ''), " & _
            "ID_廠商 = NULLIF((Select ID from 廠商資料 Where 廠商='" & 維護廠商 &"'), '')," & _
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
            delete1="DELETE FROM 零星維修作業 " & _
            "WHERE id = '" & id & "'"
            data.deleteCommand =delete1
            data.delete()
        Next
        Me.GridView1.DataBind()
    End Sub
    Protected Sub test(ByVal sender As Object, ByVal e As System.EventArgs)
    End Sub
    Protected Sub 年_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles 年.TextChanged
        月1_SelectedIndexChanged(sender,e)
        月2_SelectedIndexChanged(sender,e)
    End Sub
    Protected Sub 月1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles 月1.SelectedIndexChanged
        GetDay(月1,日1)
    End Sub
    Protected Sub 月2_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles 月2.SelectedIndexChanged
        GetDay(月2,日2)
    End sub
    Protected Sub GridView1_RowCommand(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles GridView1.RowCommand
        If e.CommandName = "上傳資料"'上傳資料按紐
            Dim i As Long = e.CommandSource.NamingContainer.RowIndex
            Dim id As String = CType(Me.GridView1.Rows(i).FindControl("id"), TextBox).Text
            If Not Me.FileUpload1.HasFile
                Label2.text="如要上傳資料，請先按'選擇檔案'"
                Exit Sub
            End If
            For Each PostedFile As HttpPostedFile In Me.FileUpload1.PostedFiles
                Dim MyGUID As String = Guid.NewGuid().ToString("N")
                Dim Myfiles As String = MapPath(".\data\Temp\") & MyGUID
                PostedFile.SaveAs(Myfiles)
                Try
                    File.Copy(Myfiles, MapPath(".\data\零星維修紀錄\") & PostedFile.FileName, False)
                Catch
                End Try
                ' data.ConnectionString = con_14
                ' data.SelectCommand = "SELECT * FROM 零星維修作業 Where id =" & id
                ' data_dv = data.Select(New DataSourceSelectArguments)
                Dim 維修紀錄 As String = CType(Me.GridView1.Rows(i).FindControl("維修紀錄"), HyperLink).Text
                ' Dim data_dv2 As Data.DataView
                ' If 維修紀錄<>""
                '     data.SelectCommand = "SELECT count(維修紀錄) As 資料數 FROM 零星維修作業 Where 維修紀錄 = '" & 維修紀錄 & "' Group by 維修紀錄"
                '     data_dv2 = data.Select(New DataSourceSelectArguments)
                ' End If
                data.InsertCommand = _
                    "IF NOT EXISTS(SELECT * FROM 零星維修作業 WHERE 維修紀錄 = N'" & PostedFile.FileName & "'and id=" & id & ") "  & _
                    "BEGIN " & _
                    "UPDATE 零星維修作業 SET " & _
                    "維修紀錄 = NULLIF(N'" & PostedFile.FileName & "', '')" & _
                    "WHERE id = '" & id & "'" & _
                    "END"
                data.Insert()
                System.IO.File.Delete(Myfiles)
                ' If 維修紀錄<>""
                '     If (data_dv2.count>0)
                '         If data_dv2(0)("資料數").ToString()=1'除非只剩一筆，否則不刪除文件檔案，因為有別的資料參照
                '             For i2 As Long = 0 To data_dv.Count() - 1
                '                 System.IO.File.Delete(MapPath(".\data\零星維修紀錄\") & 維修紀錄)
                '             Next
                '         End If
                '     End If
                ' End If
                '
                CheckFileEmpty(維修紀錄)
            Next
            Me.GridView1.DataBind()
        ElseIf e.CommandName = "刪除"
            Update(sender,e)
            Dim i As Long = e.CommandSource.NamingContainer.RowIndex
            Dim id As String = CType(Me.GridView1.Rows(i).FindControl("id"), TextBox).Text
            Dim 維修紀錄 As String = CType(Me.GridView1.Rows(i).FindControl("維修紀錄"), HyperLink).Text
            ' data.ConnectionString = con_14
            ' data.SelectCommand = "SELECT * FROM 零星維修作業 Where id =" & id
            ' data_dv = data.Select(New DataSourceSelectArguments)
            ' Dim data_dv2 As Data.DataView
            ' data.SelectCommand = "SELECT count(維修紀錄) As 資料數 FROM 零星維修作業 Where 維修紀錄 = '" & 維修紀錄 & "' Group by 維修紀錄"
            ' data_dv2 = data.Select(New DataSourceSelectArguments)
            ' If (data_dv2.count>0)
            '     If data_dv2(0)("資料數").ToString()=1'除非只剩一筆，否則不刪除文件檔案，因為有別的資料參照
            '         For i2 As Long = 0 To data_dv.Count() - 1
            '             System.IO.File.Delete(MapPath(".\data\零星維修紀錄\") & 維修紀錄)
            '         Next
            '     End If
            ' End If
            data.UpdateCommand = _
                    "UPDATE 零星維修作業 SET " & _
	                "ID_維修品項=NULL,叫修日期=NULL,維修日期=NULL,維修內容=NULL,維修金額=0,ID_廠商=NULL,維修紀錄=NULL,ID_維修紀錄=NULL,備註=NULL " & _
	                "WHERE id = '" & id &"' "
            data.Update()
            CheckFileEmpty(維修紀錄)
            Me.GridView1.DataBind()
        End If
    End sub
    Protected Sub Clear_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Response.Redirect(Request.Url.ToString())
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
    Protected Sub CheckFileEmpty(ByVal File As String)'無資料時，被拒絕存取
        '查詢是否有資料
        If File<>""
            data.ConnectionString = con_14
            data.SelectCommand = "SELECT * FROM 零星維修作業 Where 維修紀錄 =N'" & File &"'"
            data_dv = data.Select(New DataSourceSelectArguments)
            If data_dv.count<1
                System.IO.File.Delete(MapPath(".\data\零星維修紀錄\") & File)
            End If
        End If
    End Sub
    Public Sub GetDay(ByVal month As Object,ByVal day As Object)'以月取日，收尋，日可不留白
        If month.text<>""
            Dim currentdate = day.SelectedValue
            day.Items.Clear()
            day.Items.Add("")
            day.Items(0).Value = ""
            Dim DIMonth As int32
            If Me.年.text<>""
               DIMonth = DateTime.DaysInMonth((CLng(Me.年.text)+1911), CLng(month.SelectedValue))
            Else
               DIMonth = DateTime.DaysInMonth(CLng(DateTime.Now.Year.ToString()), CLng(month.SelectedValue))
            End If
            For i = 1 To DIMonth
                    day.Items.Add((i).ToString("0"))
                    day.Items(i).Value = (i).ToString("0")
            Next
            If day.Items.IndexOf(day.Items.FindByValue(currentdate)) = -1
                day.SelectedIndex = day.Items.Count - 1
            Else
                day.SelectedIndex = day.Items.IndexOf(day.Items.FindByValue(currentdate))
            End If
        End If
    End Sub
End Class