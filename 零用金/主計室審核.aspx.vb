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
Partial Class 主計室審核
    Inherits System.Web.UI.Page
    Dim con_14 As String = System.Web.Configuration.WebConfigurationManager.ConnectionStrings("ApplicationServices").ConnectionString
    Dim data As New SqlDataSource
    Dim data_dv As Data.DataView
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        data.ConnectionString = con_14
        If Not Page.IsPostBack Then
            Me.年.Text = (DateTime.Now.Year - 1911).ToString()
            Me.GridView1.PageIndex =0
        Else
            Me.Label1.Text = ""
            Me.Label2.Text = ""
        End If 
    End Sub
    Protected Sub return_(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim 作用 As boolean = False
        Dim 種類 As String = Me._種類.Text
        For i=0 to Me.GridView1.Rows.Count - 1
            Dim id As String = CType(Me.GridView1.Rows(i).FindControl("id"), Label).Text
            Dim 號數 As String = CType(Me.GridView1.Rows(i).FindControl("號數"), Label).Text
            If Session("Uid")<>Nothing
                Dim Uid As String = Session("Uid")
                If CType(Me.GridView1.Rows(i).FindControl("回覆R"), RadioButtonList).SelectedIndex=0
                    作用=True
                    If Uid="2897"
                    Else
                    Uid="2808"
                    End If
                    '--
                    '更新日誌
                    Dim data_dv2 As Data.DataView
                    data.ConnectionString = con_14
                    data.SelectCommAnd = "Select id From 收支備查簿 " & _
                        "Where 送交主計室日期 IS NOT NULL And 回覆 = 'false' AND 號數 = '" & 號數 & "'AND _種類 = '" & 種類 & "'"
                    data_dv2 = data.Select(New DataSourceSelectArguments)
                    For k=0 to data_dv2.count-1
                        data.insertCommand = _
                        "INSERT INTO 日誌 " & _
                        "(id, 動作, 命令,日期,日期2,簽章) " & _
                        "VALUES " & _
                        "(N'" & data_dv2(k)("id").ToString() & "', N'主計室通過', '', N'" & DateTime.now.tostring() & "' , '" & DateTime.now.tostring("yyyy-MM-dd HH:mm:ss") & "',N'" & Uid & "')"
                        data.insert()
                    Next
                    '送交資料
                    data.UpdateCommand = "UPDATE 收支備查簿 SET " & _
                        "回覆 = 'True',主計室簽核=NULLIF(N'" & Uid & "', '') " & _
                        "Where 送交主計室日期 IS NOT NULL And 回覆 = 'false' AND 號數 = '" & 號數 & "'AND _種類 = '" & 種類 & "'"
                    data.Update()
                End If 
            Else'有可能這行不會執行
                Me.Label2.Text="請重新登入"
            End If 
        Next
        If 作用=True
            Label1.Text="回覆通過成功"
            Me.GridView1.DataBind()
        Else
            Me.Label2.Text="要通過的資料請先勾選[完成]，在按下[回覆]"
        End If
    End Sub
    protected Sub 駁回_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim 種類 As String = Me._種類.Text
        Dim button As button = sender '下面三段為取審核_CheckedChanged在GridView的位置
        Dim row As GridViewRow = button.NamingContainer
        Dim index As Integer = row.RowIndex
        Dim i As Long = index
        Dim 號數 As String = CType(Me.GridView1.Rows(i).FindControl("號數"), Label).Text
        Dim 駁回原因 As String 
        For j=0 to Me.GridView1.Rows.Count - 1
            Dim id As String = CType(Me.GridView1.Rows(j).FindControl("id"), Label).Text
            If CType(Me.GridView1.Rows(i).FindControl("回覆R"),RadioButtonList).SelectedIndex=1 And 號數= CType(Me.GridView1.Rows(j).FindControl("號數"), Label).Text
                駁回原因 = CType(Me.GridView1.Rows(j).FindControl("駁回原因"), TextBox).Text
                Dim Uid As String = Session("Uid")
                If uid="2897"
                Else
                    uid="2808"
                End If 
                '如有跨頁應該會出問題
                '--
                '更新日誌
                Dim data_dv2 As Data.DataView
                data.ConnectionString = con_14
                data.SelectCommAnd = "Select id From 收支備查簿 " & _
                "Where 送交主計室日期 IS NOT NULL And 回覆 = 'false' AND 號數 = '" & 號數 & "'AND _種類 = '" & 種類 & "'"
                data_dv2 = data.Select(New DataSourceSelectArguments)
                For k=0 to data_dv2.count-1
                    data.insertCommand = _
                    "INSERT INTO 日誌 " & _
                    "(id, 動作, 命令,日期,日期2,簽章) " & _
                    "VALUES " & _
                    "(N'" & data_dv2(k)("id").ToString() & "', N'主計室駁回', '', N'" & DateTime.now.tostring() & "' , '" & DateTime.now.tostring("yyyy-MM-dd HH:mm:ss") & "',N'" & Uid & "')"
                    data.insert()
                Next
                '送交資料
                data.UpdateCommand = "UPDATE 收支備查簿 SET " & _
                    "送交主計室日期 = NULL,鎖定 = 'False',過審 = 'False',主計室簽核=NULLIF(N'" &  Uid & "', ''),駁回原因 = NULLIF('" & 駁回原因 & "', '')" & _
                    "Where 回覆 = 'false' AND 號數 = '" & 號數 & "'AND _種類 = '" & 種類 & "'"
                data.Update()
            End If 
        Next
        data.UpdateCommand = "UPDATE 日誌 Set 原因 = (Select 駁回原因 From 收支備查簿 Where 收支備查簿.id=日誌.id) Where 動作='主計室駁回' " & _ 
        "AND 日期2 Between DATEADD(second,-2,'" & DateTime.now.tostring("yyyy-MM-dd HH:mm:ss") & "') AND DATEADD(second,1,'" & DateTime.now.tostring("yyyy-MM-dd HH:mm:ss") & "')"
        data.Update()
        Label1.Text="駁回成功"
        Me.GridView1.DataBind()
    End Sub
    protected Sub 回覆R_OnSelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)'自動將相同號數勾取，並將另一個取消勾取
        '下面三段為取審核_CheckedChanged在GridView的位置
        Dim row As GridViewRow = sender.NamingContainer
        Dim index As Integer = row.RowIndex
        Dim i As Long = index
        Dim 號數 As String = CType(Me.GridView1.Rows(i).FindControl("號數"), Label).Text
        If 號數<>""
            For j = 0 to Me.GridView1.Rows.Count - 1
                If CType(Me.GridView1.Rows(j).FindControl("號數"), Label).Text=號數
                    CType(Me.GridView1.Rows(j).FindControl("回覆R"), RadioButtonList).SelectedIndex=CType(Me.GridView1.Rows(i).FindControl("回覆R"), RadioButtonList).SelectedIndex
                End If 
            Next
        End If
    End Sub
End Class