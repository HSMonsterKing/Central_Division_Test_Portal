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
Partial Class 查詢
    Inherits System.Web.UI.Page
    Dim con_14 As String = System.Web.Configuration.WebConfigurationManager.ConnectionStrings("ApplicationServices").ConnectionString
    Dim data As New SqlDataSource
    Dim data_dv As Data.DataView
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        data.ConnectionString = con_14
        If Not Page.IsPostBack Then
            Me.年.Text = (DateTime.Now.Year - 1911).ToString()
            Me.GridView1.PageIndex = Int32.MaxValue
            月1.SelectedValue=(DateTime.Now.AddMonths(-1).Month).ToString()
            月1_SelectedIndexChanged(sender,e)
            日1.SelectedValue=(DateTime.Now.AddMonths(-1).Day).ToString()
            月2.SelectedValue=(DateTime.Now.Month).ToString()
            月2_SelectedIndexChanged(sender,e)
            日2.SelectedValue=(DateTime.Now.Day).ToString()
            data.SelectCommand = "select 科目 from 科目表 Where 科目<>'' OR 科目 IS NOT NULL  order by 科目"
            data_dv = data.Select(New DataSourceSelectArguments)
            科目.Items.Clear()
            科目.Items.Add("")
            科目.Items(0).Value = ""
            For j = 0 To data_dv.Count - 1
                Dim 科目名稱 As String = data_dv(j)(0)
                科目.Items.Add(科目名稱)
                科目.Items(j+1).Value = 科目名稱
            Next
        Else
            Me.Label1.Text = ""
            Me.Label2.Text = ""
        End If 
    End Sub
    Protected Sub GridView1_RowCommand(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles GridView1.RowCommand
        If e.CommandName = "流程"'顯示該筆資料的日誌
            '取ID
            Dim i As Long = e.CommandSource.NamingContainer.RowIndex
            Dim id As String = CType(Me.GridView1.Rows(i).FindControl("id"), Label).Text
            Panel3.Visible=False
            Panel4.Visible=true
            Me.ID.text = id
        End If 
    End Sub
    Protected Sub GridView2_RowCommand(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles GridView2.RowCommand
        If e.CommandName = "查詢"
            Dim i As Long = e.CommandSource.NamingContainer.RowIndex
            dim 命令 As string =CType(Me.GridView2.Rows(i).FindControl("命令"), Label).text
            dim idm As string=Mid(命令,8)
            Panel4.Visible=False
            Panel5.Visible=true
            I_id.text=idm
            If left(命令,2)="修改"
                data.ConnectionString = con_14
                data.SelectCommand = "SELECT * FROM 修改資料 where id=" & idm
                data_dv = data.Select(New DataSourceSelectArguments)
                dim 修改前s As string=""
                dim 修改前s1 As string=""
                dim 修改前s2 As string=""
                dim 修改後s As string=""
                dim 修改後s1 As string=""
                dim 修改後s2 As string=""
                For i = 0 To data_dv.Count - 1
                    Dim id As String = data_dv(i)("id").ToString()
                    Dim id_收 As String = data_dv(i)("id_收").ToString()
                    Dim 單位別 As String = data_dv(i)("單位別").ToString()
                    Dim 單位別_改 As String = data_dv(i)("單位別_改").ToString()
                    Dim 承辦人 As String = data_dv(i)("承辦人").ToString()
                    Dim 承辦人_改 As String = data_dv(i)("承辦人_改").ToString()
                    Dim 月 As String = data_dv(i)("月").ToString()
                    Dim 月_改 As String = data_dv(i)("月_改").ToString()
                    Dim 日 As String = data_dv(i)("日").ToString()
                    Dim 日_改 As String = data_dv(i)("日_改").ToString()
                    Dim 科目 As String = data_dv(i)("科目").ToString()
                    Dim 科目_改 As String = data_dv(i)("科目_改").ToString()
                    Dim 摘要 As String = data_dv(i)("摘要").ToString()
                    Dim 摘要_改 As String = data_dv(i)("摘要_改").ToString()
                    Dim 姓名 As String = data_dv(i)("姓名").ToString()
                    Dim 姓名_改 As String = data_dv(i)("姓名_改").ToString()
                    Dim 商號 As String = data_dv(i)("商號").ToString()
                    Dim 商號_改 As String = data_dv(i)("商號_改").ToString()
                    Dim 經手人 As string = data_dv(i)("經手人").ToString()
                    Dim 經手人_改 As string = data_dv(i)("經手人_改").ToString()
                    Dim 種類 As String = data_dv(i)("種類").ToString()
                    Dim 種類_改 As String = data_dv(i)("種類_改").ToString()
                    Dim 號數 As String = data_dv(i)("號數").ToString()
                    Dim 號數_改 As String = data_dv(i)("號數_改").ToString()
                    Dim 收入 As String = data_dv(i)("收入").ToString()
                    Dim 收入_改 As String = data_dv(i)("收入_改").ToString()
                    Dim 支出 As String = data_dv(i)("支出").ToString()
                    Dim 支出_改 As String = data_dv(i)("支出_改").ToString()
                    Dim 備註 As String = data_dv(i)("備註").ToString()
                    Dim 備註_改 As String = data_dv(i)("備註_改").ToString()
                    If 單位別<>單位別_改
                        修改前s=修改前s & ",單位別為[" & 單位別 & "]"
                        修改後s=修改後s & ",單位別為[" & 單位別_改 & "]"
                    End If 
                    If 承辦人<>承辦人_改
                        修改前s=修改前s & ",承辦人為[" & 承辦人 & "]"
                        修改後s=修改後s & ",承辦人為[" & 承辦人_改 & "]"
                    End If 
                    If 月<>月_改
                        修改前s=修改前s & ",月為[" & 月 & "]"
                        修改後s=修改後s & ",月為[" & 月_改 & "]"
                    End If 
                    If 日<>日_改
                        修改前s=修改前s & ",日為[" & 日 & "]"
                        修改後s=修改後s & ",日為[" & 日_改 & "]"
                    End If 
                    If 科目<>科目_改
                        修改前s=修改前s & ",科目為[" & 科目 & "]"
                        修改後s=修改後s & ",科目為[" & 科目_改 & "]"
                    End If 
                    If 摘要<>摘要_改
                        修改前s=修改前s & ",摘要為[" & 摘要 & "]"
                        修改後s=修改後s & ",摘要為[" & 摘要_改 & "]"
                    End If 
                    If 姓名<>姓名_改
                        修改前s=修改前s & ",姓名為["
                        姓名前.ImageUrl=姓名
                        修改前s1=修改前s1 & "]"
                        修改後s=修改後s & ",姓名為["
                        姓名後.ImageUrl=姓名_改
                        修改後s1=修改後s1 & "]"
                    Else
                        姓名前.ImageUrl=""
                        姓名後.ImageUrl=""
                    End If 
                    If 商號<>商號_改
                        修改前s1=修改前s1 & ",商號為[" & 商號 & "]"
                        修改後s1=修改後s1 & ",商號為[" & 商號_改 & "]"
                    End If 
                    If 經手人<>經手人_改
                        修改前s1=修改前s1 & ",經手人為["
                        經手人前.ImageUrl=經手人
                        修改前s2=修改前s2 & "]"
                        修改後s1=修改後s1 & ",經手人為["
                        經手人後.ImageUrl=經手人_改
                        修改後s2=修改後s2 & "]"
                    Else
                        經手人前.ImageUrl=""
                        經手人後.ImageUrl=""
                    End If 
                    If 種類<>種類_改
                        修改前s2=修改前s2 & ",種類為[" & 種類 & "]"
                        修改後s2=修改後s2 & ",種類為[" & 種類_改 & "]"
                    End If 
                    If 號數<>號數_改
                        修改前s2=修改前s2 & ",號數為[" & 號數 & "]"
                        修改後s2=修改後s2 & ",號數為[" & 號數_改 & "]"
                    End If 
                    If 收入<>收入_改
                        修改前s2=修改前s2 & ",收入為[" & 收入 & "]"
                        修改後s2=修改後s2 & ",收入為[" & 收入_改 & "]"
                    End If 
                    If 支出<>支出_改
                        修改前s2=修改前s2 & ",支出為[" & 支出 & "]"
                        修改後s2=修改後s2 & ",支出為[" & 支出_改 & "]"
                    End If 
                    If 備註<>備註_改
                        修改前s2=修改前s2 & ",備註為[" & 備註 & "]"
                        修改後s2=修改後s2 & ",備註為[" & 備註_改 & "]"
                    End If 
                Next
                修改前.text=修改前s
                修改前1.text=修改前s1
                修改前2.text=修改前s2
                If 修改前s<>""
                    修改前.text=Mid(修改前s,2)
                ElseIf 修改前s1<>""
                    修改前1.text=Mid(修改前s1,2)
                ElseIf 修改前s2<>""
                    修改前2.text=Mid(修改前s2,2)
                End If
                修改後.text=修改後s
                修改後1.text=修改後s1
                修改後2.text=修改後s2
                If 修改後s<>""
                    修改後.text=Mid(修改後s,2)
                ElseIf 修改後s1<>""
                    修改後1.text=Mid(修改後s1,2)
                ElseIf 修改後s2<>""
                    修改後2.text=Mid(修改後s2,2)
                End If
            End If 
        End If 
    End Sub
    Protected Sub GridView1_DataBound(ByVal sender As Object, ByVal e As System.EventArgs) HAndles GridView1.DataBound
        GenerateDropdownlist()
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
        Me.GridView1.PageIndex = Int32.MaxValue
    End Sub
    Protected Sub SelectedIndexChanged_尾頁(ByVal sender As Object, ByVal e As System.EventArgs)
        Me.GridView1.PageIndex = Int32.MaxValue
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
    Protected Sub 返回_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Panel3.Visible=true
        Me.Panel3.DataBind()
        Panel4.Visible=False
    End Sub
    Protected Sub Clear_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Response.Redirect(Request.Url.ToString())
    End Sub
    Protected Sub P2_Select_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Me.ID.text=""
        Panel3.Visible=False
        Panel4.Visible=true
        Me.GridView2.DataBind()
    End Sub
    Protected Sub 號數1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        If Me.號數2.text="" 
            Me.號數2.text= Me.號數1.text
        End If 
    End Sub
    Protected Sub 號數2_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        If Me.號數1.text="" 
            Me.號數1.text= Me.號數2.text
        End If 
    End Sub
    Protected Sub 日期1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        If Me.日期2.text="" 
        Me.日期2.text=(DateTime.Now.Year - 1911).ToString()+"/"+(DateTime.Now.Month).ToString("00")+"/"+(DateTime.Now.Day).ToString("00")
        End If 
    End Sub
    Protected Sub 日期2_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        If Me.日期1.text="" 
        Me.日期1.text=(DateTime.Now.Year - 1911).ToString()+"/"+(DateTime.Now.Month).ToString("00")+"/"+(DateTime.Now.Day).ToString("00")
        End If 
    End Sub
    Protected Sub 返回2_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Panel4.Visible=true
        Panel5.Visible=False
    End Sub
    'Shared之前，需要下行程式碼
    <System.Web.Script.Services.ScriptMethod(), System.Web.Services.WebMethod()>
    Public Shared Function GetMyList(ByVal prefixText As String, ByVal count As Integer)'常用文字
        Dim con_14 As String = System.Web.Configuration.WebConfigurationManager.ConnectionStrings("ApplicationServices").ConnectionString
        Dim data As New SqlDataSource
        Dim data_dv As Data.DataView
        Dim MyList As New List(Of String)
        data.ConnectionString = con_14
        data.SelectCommAnd = "SELECT TOP " & count & " 常用文字 FROM 常用清單 WHERE 常用文字 LIKE '%" & prefixText & "%' ORDER BY CASE WHEN 常用文字 IS NULL THEN 1 ELSE 0 END, 常用文字"
        data_dv = data.Select(New DataSourceSelectArguments)
        For i As Long = 0 To data_dv.Count() - 1
            MyList.Add(data_dv(i)(0).ToString())
        Next
        Return MyList
    End Function
    Public Sub GenerateDropdownlist()
        data.SelectCommand = "select 科目 from 科目表 Where 科目<>'' OR 科目 IS NOT NULL order by 科目"
        data_dv = data.Select(New DataSourceSelectArguments)
        For i = 0 To Me.GridView1.Rows.Count - 1
            Dim id As String = CType(Me.GridView1.Rows(i).FindControl("id"), Label).Text
            data.SelectCommand = "select * from 收支備查簿 Where id='"& id &"'"
            Dim data_dv2 As Data.DataView = data.Select(New DataSourceSelectArguments)
            Dim 科目 As DropDownList = CType(Me.GridView1.Rows(i).FindControl("科目"), DropDownList)
            Dim 科目2 As DropDownList = CType(Me.GridView1.Rows(i).FindControl("科目2"), DropDownList)
            Dim 科目S As String = data_dv2(0)("科目").ToString()
            Dim 科目S2 As String = data_dv2(0)("科目2").ToString()
            科目.Items.Clear()
            科目2.Items.Clear()
            科目.Items.Add("")
            科目.Items(0).Value = ""
            科目2.Items.Add("")
            科目2.Items(0).Value = ""
            For j = 0 To data_dv.Count - 1
                Dim 科目名稱 As String = data_dv(j)(0)
                科目.Items.Add(科目名稱)
                科目.Items(j+1).Value = 科目名稱
                科目2.Items.Add(科目名稱)
                科目2.Items(j+1).Value = 科目名稱
            Next
            科目.SelectedIndex=科目.Items.IndexOf(科目.Items.FindByValue(科目S))
            科目2.SelectedIndex=科目2.Items.IndexOf(科目2.Items.FindByValue(科目S2))
        Next
    End Sub
    Public Sub GetDay(ByVal month As Object,ByVal day As Object)'以月取日，收尋，日可不留白
        If month.SelectedValue<>"" And Me.年.text<>""
            Dim currentdate = day.SelectedValue
            day.Items.Clear()
            day.Items.Add("")
            day.Items(0).Value = ""
            For i = 1 To DateTime.DaysInMonth((CLng(Me.年.text) + 1911), CLng(month.SelectedValue))
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