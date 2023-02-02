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
Partial Class 更新資料
    Inherits System.Web.UI.Page
    Dim con_14 As String = System.Web.Configuration.WebConfigurationManager.ConnectionStrings("ApplicationServices").ConnectionString
    Dim data As New SqlDataSource
    Dim data_dv As Data.DataView
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        data.ConnectionString = con_14
        If Not Page.IsPostBack Then
            Me.GridView1.PageIndex = 0
            PermissionOn()
        Else
            Me.Label1.Text = ""
            Me.Label2.Text = ""
        End If
    End Sub
    Protected Sub Insert(ByVal sender As Object, ByVal e As System.EventArgs)
        Update(sender, e)
        dim insert1 As string
        data.InsertCommand = _
            "INSERT INTO 更新資料表 " & _
            "(更新日期,更新版本) " & _
            "VALUES " & _
            "('" & (DateTime.Now).ToString("yyyy/MM/dd") & "','" & (DateTime.Now).ToString("yyyyMMdd") & ".ver')"
        data.Insert()
        Me.GridView1.PageIndex = 0
        Me.GridView1.DataBind()
    End Sub
    Protected Sub Update(ByVal sender As Object, ByVal e As System.EventArgs)
    For i = 0 To Me.GridView1.Rows.Count - 1
            Dim id As String = CType(Me.GridView1.Rows(i).FindControl("id"), TextBox).Text
            Dim 更新日期 As String = nothing
            If CType(Me.GridView1.Rows(i).FindControl("更新日期"), TextBox).text<>""
                更新日期 = CType(Me.GridView1.Rows(i).FindControl("更新日期"), TextBox).text
                更新日期 = taiwancalendarto(更新日期)
            End If
            Dim 更新版本 As String = CType(Me.GridView1.Rows(i).FindControl("更新版本"), TextBox).Text
            Dim 內容 As String = CType(Me.GridView1.Rows(i).FindControl("內容"), TextBox).Text
            內容=內容.Replace("'", "''")
            Dim Update1 As string ="UPDATE 更新資料表 SET " & _
            "更新日期 = IIF(ISDATE(TRIM(N'" & 更新日期 & "'))=1,TRIM(N'" & 更新日期 & "'),NULL), " & _
            "更新版本 = NULLIF(N'" & 更新版本 & "', ''), " & _
            "內容 = NULLIF(N'" & 內容 & "', '') " & _
            "WHERE id = '" & id & "'"
            data.UpdateCommand = Update1
            data.Update()
        Next
        Me.GridView1.DataBind()
    End Sub
    Protected Sub Delete(ByVal sender As Object, ByVal e As System.EventArgs)
        Update(sender, e)
        Me.GridView1.DataBind()
        dim delete1 As string=""
        dim i As Int32 = Me.GridView1.Rows.Count - 1
        Dim id As String = CType(Me.GridView1.Rows(i).FindControl("id"), TextBox).Text
        delete1="DELETE FROM 更新資料表 " & _
        "WHERE id = '" & id & "'"
        data.DeleteCommand =delete1
        data.Delete()
        Me.GridView1.DataBind()
    End Sub
    Protected Sub GridView1_DataBound(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.DataBound
        PermissionOn()
    End Sub
    Protected Sub PermissionOn()
        If session("水_Uid")="3855"
                新增.Visible = true
                存檔.Visible = true
                For i = 0 To Me.GridView1.Rows.Count - 1
                    CType(Me.GridView1.Rows(i).FindControl("更新日期"), TextBox).Enabled=true
                    CType(Me.GridView1.Rows(i).FindControl("更新版本"), TextBox).Enabled=true
                    CType(Me.GridView1.Rows(i).FindControl("內容"), TextBox).Enabled=true
                    Me.GridView1.columns(4).Visible = True
                Next
        End If
    End Sub
    Protected Sub GridView1_RowCommand(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles GridView1.RowCommand
    If e.CommandName = "刪除"
            '連同維護紀錄作業的相關資料一併刪除
            Update(sender, e)
            Dim i As Long = e.CommandSource.NamingContainer.RowIndex
            Dim id As String = CType(Me.GridView1.Rows(i).FindControl("id"), TextBox).Text
            data.DeleteCommand = "DELETE FROM 更新資料表 WHERE ID=" & id
            data.Delete()
            Me.GridView1.DataBind()
        End If
    End Sub
    Protected Sub test(ByVal sender As Object, ByVal e As System.EventArgs)
        Label1.Text=Me.更新版本.Items(Me.更新版本.Items.Count-1).Value
    End Sub
    Protected Sub 年_DataBound(ByVal sender As Object, ByVal e As System.EventArgs) Handles 年.DataBound
        Me.年.Items.Add("")
        Me.年.Items(Me.年.Items.Count-1).Value = ""
        Me.年.SelectedValue=""
    End Sub
    Protected Sub 更新版本_DataBound(ByVal sender As Object, ByVal e As System.EventArgs) Handles 更新版本.DataBound
        Me.更新版本.Items.Add("")
        Me.更新版本.Items(Me.更新版本.Items.Count-1).Value = ""
        Me.更新版本.SelectedValue=""
    End Sub
    Protected Sub 年_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles 年.SelectedIndexChanged
        月1_SelectedIndexChanged(sender,e)
        月2_SelectedIndexChanged(sender,e)
    End Sub
    Protected Sub 月1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles 月1.SelectedIndexChanged
        GetDay(月1,日1)
    End Sub
    Protected Sub 月2_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles 月2.SelectedIndexChanged
        GetDay(月2,日2)
    End sub
    Public Sub GetDay(ByVal month As Object,ByVal day As Object)'以月取日，收尋，日可不留白
        If month.text<>""
            Dim currentdate = day.SelectedValue
            day.Items.Clear()
            day.Items.Add("")
            day.Items(0).Value = ""
            Dim DIMonth As int32
            '年為DropDownList .TEXT為Value值
            If Me.年.text<>""
               DIMonth = DateTime.DaysInMonth((CLng(Me.年.text)), CLng(month.SelectedValue))
            Else
               DIMonth = DateTime.DaysInMonth((CLng("110") + 1911), CLng(month.SelectedValue))
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