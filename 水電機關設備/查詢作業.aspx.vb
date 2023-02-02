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
Partial Class 查詢作業
    Inherits System.Web.UI.Page
    Dim con_14 As String = System.Web.Configuration.WebConfigurationManager.ConnectionStrings("ApplicationServices").ConnectionString
    Dim data As New SqlDataSource
    Dim data_dv As Data.DataView
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        data.ConnectionString = con_14
        If Not Page.IsPostBack Then
            Me.年.SelectedValue="2021"
            Me.GridView1.PageIndex = Int32.MaxValue
            月1.SelectedValue=(DateTime.Now.AddMonths(-1).Month).ToString()
            月1_SelectedIndexChanged(sender,e)
            日1.SelectedValue=(DateTime.Now.AddMonths(-1).Day).ToString()
            月2.SelectedValue=(DateTime.Now.Month).ToString()
            月2_SelectedIndexChanged(sender,e)
            日2.SelectedValue=(DateTime.Now.Day).ToString()
        Else
            Me.Label1.Text = ""
            Me.Label2.Text = ""
        End If
    End Sub
    Protected Sub test(ByVal sender As Object, ByVal e As System.EventArgs)
    End Sub
    Protected Sub GridView1_RowCommand(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles GridView1.RowCommand
        If e.CommandName = "維護紀錄"
            Dim i As Long = e.CommandSource.NamingContainer.RowIndex
            Dim id As String = CType(Me.GridView1.Rows(i).FindControl("id"), label).Text
            Dim 品項 As String = CType(Me.GridView1.Rows(i).FindControl("品項"), DropDownList).Text
            Session("水_id")=id
            Session("水_品項")=品項
            Session("水_編輯權限")=Nothing
            Response.Redirect("維修紀錄作業.aspx")
        ElseIf e.CommandName = "照片圖"'點擊後放大 解:另外設定頁面用以顯示圖片
            Dim i As Long = e.CommandSource.NamingContainer.RowIndex
            Dim id As String = CType(Me.GridView1.Rows(i).FindControl("id"), Label).Text
            data.ConnectionString = con_14
            data.SelectCommAnd = "SELECT id,壓縮照片 FROM 水電機關設備資料表 Where id="& id'全輸出，不輸出無編號
            data_dv = data.Select(New DataSourceSelectArguments)
            Session("水_照片")=data_dv(0)("壓縮照片").ToString
            Response.Redirect("照片.aspx")
        End If
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
    Protected Sub Clear_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Response.Redirect(Request.Url.ToString())
    End Sub
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