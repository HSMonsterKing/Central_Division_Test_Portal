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
Partial Class 審核
    Inherits System.Web.UI.Page
    Dim con_14 As String = System.Web.Configuration.WebConfigurationManager.ConnectionStrings("ApplicationServices").ConnectionString
    Dim data As New SqlDataSource
    Dim data_dv As Data.DataView
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        data.ConnectionString = con_14
        If Not Page.IsPostBack Then
            Me.年.Text = (DateTime.Now.Year - 1911).ToString()
            Me.GridView1.PageIndex = 0'Int32.MaxValue
        Else
            Me.Label1.Text = ""
            Me.Label2.Text = ""
        End If 
    End Sub
    Protected Sub 種類_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)'當種類變成A才顯示收入
        If (Me._種類.SelectedValue="A") Then
            Me.GridView1.columns(6).Visible = True
            Me.GridView1.columns(13).Visible = True
            Me.GridView1.columns(15).Visible = True
        Else
            Me.GridView1.columns(6).Visible = False
            Me.GridView1.columns(13).Visible = False
            Me.GridView1.columns(15).Visible = False
        End If 
    End Sub
    protected Sub 通過_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim 作用 As boolean = False
        Dim 年 As String = Me.年.Text
        Dim 種類 As String = Me._種類.Text
        data.ConnectionString = con_14
        For i=0 to Me.GridView1.Rows.Count - 1
            Dim id As String = CType(Me.GridView1.Rows(i).FindControl("id"), TextBox).Text
            Dim 號數 As String = CType(Me.GridView1.Rows(i).FindControl("號數"), Label).Text
            If CType(Me.GridView1.Rows(i).FindControl("選取"), CheckBox).Checked=True
                作用=True
                '更新日誌
                Dim data_dv2 As Data.DataView
                data.ConnectionString = con_14
                data.SelectCommAnd = "Select id From 收支備查簿 Where 鎖定 = 'True' AND 送出 = 'True' AND 過審 = 'False' AND 主計室簽核 IS NULL AND 號數 = '" & 號數 & "'And 取號='False' AND _種類 = '" & Me._種類.Text & "'"
                data_dv2 = data.Select(New DataSourceSelectArguments)
                For k=0 to data_dv2.count-1
                    data.InsertCommand = _
                    "INSERT INTO 日誌 " & _
                    "(id, 動作,日期,日期2) " & _
                    "VALUES " & _
                    "(N'" & data_dv2(k)("id").ToString() & "', N'通過' , N'" & DateTime.now.tostring() & "' , '" & DateTime.now.tostring("yyyy-MM-dd HH:mm:ss") & "')"
                    data.insert()
                Next
                '送交資料
                data.UpdateCommand = "UPDATE 收支備查簿 SET " & _
                "過審 = 'True',駁回原因 = NULL " & _
                "Where 號數 = '" & 號數 & "' And _種類 = '" & Me._種類.Text & "'"
                data.Update()
            End If 
        Next
        If 作用=False
            label2.Text="請先選取欲通過之資料"
        Else
            Me.GridView1.DataBind()
            Label1.Text="已回覆通過"
        End If
    End Sub
    protected Sub 駁回_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim 作用 As boolean = False
        Dim 年 As String = Me.年.Text
        Dim 種類 As String = Me._種類.Text
        Dim 駁回原因 As String
        For i=0 to Me.GridView1.Rows.Count - 1
            Dim id As String = CType(Me.GridView1.Rows(i).FindControl("id"), TextBox).Text
            Dim 號數 As String = CType(Me.GridView1.Rows(i).FindControl("號數"), Label).Text
            If  CType(Me.GridView1.Rows(i).FindControl("駁回原因"), TextBox).Text<>""
                駁回原因=CType(Me.GridView1.Rows(i).FindControl("駁回原因"), TextBox).Text
            End If
            If CType(Me.GridView1.Rows(i).FindControl("選取"), CheckBox).Checked=True
                作用=True
                '更新日誌
                Dim data_dv2 As Data.DataView
                data.ConnectionString = con_14
                data.SelectCommAnd = "Select id From 收支備查簿 " & _
                    "Where 鎖定 = 'True' AND 送出 = 'True' AND 過審 = 'False' AND 主計室簽核 IS NULL AND 號數 = '" & 號數 & "'And 取號='False' AND _種類 = '" & Me._種類.Text & "'"
                data_dv2 = data.Select(New DataSourceSelectArguments)
                For k=0 to data_dv2.count-1
                    data.insertCommand = _
                    "INSERT INTO 日誌 " & _
                    "(id, 動作,日期,日期2) " & _
                    "VALUES " & _
                    "(N'" & data_dv2(k)("id").ToString() & "', N'駁回', N'" & DateTime.now.tostring() & "' , '" & DateTime.now.tostring("yyyy-MM-dd HH:mm:ss") & "')"
                    data.insert()
                Next
                '送交資料
                data.UpdateCommand = "UPDATE 收支備查簿 SET " & _
                    "鎖定 = 'False',駁回原因 = NULLIF('"& 駁回原因 &"','') " & _
                    "WHERE 號數='" & 號數 & "' AND _種類 = '" & Me._種類.Text & "' AND 送出 = 'True' AND 過審 = 'False' AND 主計室簽核 IS NULL "
                data.Update()
            End If 
        Next
        If 作用=False
            label2.Text="請先選取欲駁回之資料"
        Else
            data.UpdateCommand = "UPDATE 日誌 Set 原因 = (Select 駁回原因 From 收支備查簿 Where 收支備查簿.id=日誌.id) " & _ 
            "Where 動作='駁回' AND 日期2 Between DATEADD(second,-2,'" & DateTime.now.tostring("yyyy-MM-dd HH:mm:ss") & "') AND DATEADD(second,1,'" & DateTime.now.tostring("yyyy-MM-dd HH:mm:ss") & "')"
            data.Update()
            Me.GridView1.DataBind()
            Label1.Text="已駁回成功"
        End If
    End Sub 
    Function BASE64_TO_IMG(ByVal BASE_STR As String) As Drawing.Image'下載資料轉圖片
        Try
            Dim IMG As Drawing.Image
            If BASE_STR <> Nothing 'And Strings.Right(BASE_STR, 1) = "=" 
                Dim BYT As Byte() = Convert.FromBase64String(BASE_STR.Remove(0,22))
                Dim MS As New IO.MemoryStream(BYT)
                IMG = Drawing.Image.FromStream(MS)
                Return IMG
            Else
                Return Nothing
            End If 
        Catch ex As Exception
            Return Nothing
        End Try
    End Function 
    Protected Sub 選取_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim checkbox As CheckBox = sender '下面三段為取審核_CheckedChanged在GridView的位置
        Dim row As GridViewRow = checkbox.NamingContainer
        Dim index As Integer = row.RowIndex
        Dim i As Long = index
        If CType(Me.GridView1.Rows(i).FindControl("號數"), Label).Text<>""
            Dim 號數 As String = CType(Me.GridView1.Rows(i).FindControl("號數"), Label).Text
                For j = 0 to Me.GridView1.Rows.Count - 1
                    If CType(Me.GridView1.Rows(j).FindControl("號數"), Label).Text=號數
                        CType(Me.GridView1.Rows(j).FindControl("選取"), CheckBox).Checked=CType(Me.GridView1.Rows(i).FindControl("選取"), CheckBox).Checked
                    End If 
                Next
        End If 
    End Sub
End Class