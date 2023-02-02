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
Partial Class 科目清單
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
        Label1.Text="已新增成功"
    End Sub
    Protected Sub Update(ByVal sender As Object, ByVal e As System.EventArgs)
        'TODO:轉換成整數會失敗
        For i = 0 To Me.GridView1.Rows.Count - 1
            Dim id As String = CType(Me.GridView1.Rows(i).FindControl("id"), TextBox).Text
            Dim 科目 As String = CType(Me.GridView1.Rows(i).FindControl("科目"), TextBox).Text
            data.UpdateCommand = "UPDATE 科目表 SET " & _
            "科目 = NULLIF(N'" & 科目 & "', '')" & _
            "WHERE id = '" & id & "'"
            data.Update()
        Next
        Label1.Text="已存檔成功"
    End Sub
    Protected Sub GridView1_RowCommand(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles GridView1.RowCommand
        If e.CommandName = "刪除"'牽動各個功能，謹慎編寫
            Update(sender, e)
            Dim i As Long = e.CommandSource.NamingContainer.RowIndex
            Dim id As String = CType(Me.GridView1.Rows(i).FindControl("id"), TextBox).Text
            Dim 科目 As String = CType(Me.GridView1.Rows(i).FindControl("科目"), TextBox).Text
            '新增修改資料的預先載入
            Dim sql As string = "SELECT * FROM [收支備查簿] where 科目='" & 科目 & "' OR 科目2 = '" & 科目 & "'"
            data.SelectCommAnd = sql
            data_dv = data.Select(New DataSourceSelectArguments)
            If data_dv.count>0
                Dim id_改 As String = data_dv(0)("ID").ToString()
                Dim 單位別_前 As String = data_dv(0)("單位別").ToString()
                Dim 承辦人_前 As String = data_dv(0)("承辦人").ToString()
                Dim 月_前 As String = data_dv(0)("月").ToString()
                Dim 日_前 As String = data_dv(0)("日").ToString()
                Dim 科目_前 As String = data_dv(0)("科目").ToString()
                If data_dv(0)("科目2").ToString()<>""
                    科目_前 = data_dv(0)("科目").ToString() & ";" & data_dv(0)("科目2").ToString()
                End If
                Dim 摘要_前 As String = data_dv(0)("摘要").ToString()
                Dim 姓名_前 As String = data_dv(0)("姓名").ToString()
                Dim 商號_前 As String = data_dv(0)("商號").ToString()
                Dim 經手人_前 As string = data_dv(0)("經手人").ToString()
                Dim 種類_前 As String = data_dv(0)("種類").ToString()
                Dim 號數_前 As String = data_dv(0)("號數").ToString()
                Dim 收入_前 As String = data_dv(0)("收入").ToString()
                Dim 支出_前 As String = data_dv(0)("支出").ToString()
                Dim 備註_前 As String = data_dv(0)("備註").ToString()
                Dim 送出 As String = data_dv(0)("送出").ToString()
                '新增修改資料的預先載入結束，更新科目
                data.UpdateCommand = "UPDATE 收支備查簿 SET " & _
                "科目 = NULL" & _
                "WHERE 科目 = '" & 科目 & "'"
                data.Update()
                data.UpdateCommand = "UPDATE 收支備查簿 SET " & _
                "科目2 = NULL" & _
                "WHERE 科目2 = '" & 科目 & "'"
                data.Update()
                '更新科目結束，判斷是否有收支備查簿被修正
                If 送出="True"'駁回新增修改紀錄，駁回為 鎖定false,送出true，要新增至修改資料
                    sql  = "SELECT * FROM [收支備查簿] where id='" & id_改 & "'"
                    data.SelectCommAnd = sql
                    data_dv = data.Select(New DataSourceSelectArguments)
                    Dim 單位別_後 As String = data_dv(0)("單位別").ToString()
                    Dim 承辦人_後 As String = data_dv(0)("承辦人").ToString()
                    Dim 月_後 As String = data_dv(0)("月").ToString()
                    Dim 日_後 As String = data_dv(0)("日").ToString()
                    Dim 科目_後 As String = data_dv(0)("科目").ToString()
                    If data_dv(0)("科目2").ToString()<>""
                        科目_後 = data_dv(0)("科目").ToString() & ";" & data_dv(0)("科目2").ToString()
                    End If
                    Dim 摘要_後 As String = data_dv(0)("摘要").ToString()
                    Dim 姓名_後 As String = data_dv(0)("姓名").ToString()
                    Dim 商號_後 As String = data_dv(0)("商號").ToString()
                    Dim 經手人_後 As string = data_dv(0)("經手人").ToString()
                    Dim 種類_後 As String = data_dv(0)("種類").ToString()
                    Dim 號數_後 As String = data_dv(0)("號數").ToString()
                    Dim 收入_後 As String = data_dv(0)("收入").ToString()
                    Dim 支出_後 As String = data_dv(0)("支出").ToString()
                    Dim 備註_後 As String = data_dv(0)("備註").ToString()
                    '判斷資料是否有修改
                    If 號數_後<>""'?
                        號數_後=CType(號數_後,Int32)
                        號數_後=CType(號數_後,string)
                    End If 
                    If 單位別_後<>單位別_前 or 承辦人_後<>承辦人_前 or 月_後<>月_前 or 日_後<>日_前 or 科目_後<>科目_前 or 摘要_後<>摘要_前 or 姓名_後<>姓名_前 or 商號_後<>商號_前 or 經手人_後<>經手人_前 or 種類_後<>種類_前 or 號數_後<>號數_前 or 收入_後<>收入_前 or 支出_後<>支出_前 or 備註_後<>備註_前
                        '有BUG，但原因無法得知，故增加原因，原因:B、XZ沒有收入欄位，""<>"0"，已經修正
                        ' Label1.text=""
                        ' If 單位別_後<>單位別_前
                        '     Label1.text+="單位別已改變:" & 單位別_後 & "->" & 單位別_前
                        ' End If
                        ' If 承辦人_後<>承辦人_前
                        '     Label1.text+="承辦人已改變:" & 承辦人_後 & "->" & 承辦人_前
                        ' End If
                        ' If 月_後<>月_前
                        '     Label1.text+="月已改變:" & 月_後 & "->" & 月_前
                        ' End If
                        ' If 日_後<>日_前
                        '     Label1.text+="日已改變:" & 日_後 & "->" & 日_前
                        ' End If
                        ' If 科目_後<>科目_前
                        '     Label1.text+="科目已改變:" & 科目_後 & "->" & 科目_前
                        ' End If
                        ' If 摘要_後<>摘要_前
                        '     Label1.text+="摘要已改變:" & 摘要_後 & "->" & 摘要_前
                        ' End If
                        ' If 姓名_後<>姓名_前
                        '     Label1.text+="姓名已改變:" & 姓名_後 & "->" & 姓名_前
                        ' End If
                        ' If 商號_後<>商號_前
                        '     Label1.text+="商號已改變:" & 商號_後 & "->" & 商號_前
                        ' End If
                        ' If 經手人_後<>經手人_前
                        '     Label1.text+="經手人已改變:" & 經手人_後 & "->" & 經手人_前
                        ' End If
                        ' If 種類_後<>種類_前
                        '     Label1.text+="種類已改變:" & 種類_後 & "->" & 種類_前
                        ' End If
                        ' If 號數_後<>號數_前
                        '     Label1.text+="號數已改變:" & 號數_後 & "->" & 號數_前
                        ' End If
                        ' If 收入_後<>收入_前
                        '     Label1.text+="收入已改變:" & 收入_後 & "->" & 收入_前
                        ' End If
                        ' If 支出_後<>支出_前
                        '     Label1.text+="支出已改變:" & 支出_後 & "->" & 支出_前
                        ' End If
                        ' If 備註_後<>備註_前
                        '     Label1.text+="備註已改變:" & 備註_後 & "->" & 備註_前
                        ' End If
                        '程式正式執行
                        '新增修改資料
                        dim date1 As string = DateTime.now.tostring()
                        dim date2 As string = DateTime.now.tostring("yyyy-MM-dd HH:mm:ss")
                        dim insert1 As string = _
                        "INSERT INTO 修改資料 " & _
                        "(id_收,單位別,單位別_改,承辦人,承辦人_改,月,月_改,日,日_改,科目,科目_改,摘要,摘要_改,姓名,姓名_改,商號,商號_改,經手人,經手人_改,種類,種類_改,號數,號數_改,收入,收入_改,支出,支出_改,備註,備註_改,date) " & _
                        "VALUES " & _
                        "(" & id_改 & ",NULLIF(N'" & 單位別_前 & "', ''),NULLIF(N'" & 單位別_後 & "', ''),NULLIF(N'" & _
                        承辦人_前 & "', ''),NULLIF(N'" & 承辦人_後 & "', ''),NULLIF('" & _
                        月_前 & "', ''),NULLIF('" & 月_後 & "', ''),NULLIF('" & _
                        日_前 & "', ''),NULLIF('" & 日_後 & "', ''),NULLIF(N'" & _
                        科目_前 & "', ''),NULLIF(N'" & 科目_後 & "', ''),NULLIF(N'" & _
                        摘要_前 & "', ''),NULLIF(N'" & 摘要_後 & "', ''),NULLIF(N'" & _
                        姓名_前 & "', ''),NULLIF(N'" & 姓名_後 & "', ''),NULLIF(N'" & _ 
                        商號_前 & "', ''),NULLIF(N'" & 商號_後 & "', ''),NULLIF(N'" & _
                        經手人_前 & "', ''),NULLIF(N'" & 經手人_後 & "', ''),NULLIF(N'" & _
                        種類_前 & "', ''),NULLIF(N'" & 種類_後 & "', ''),NULLIF('" & _
                        號數_前 & "', ''),NULLIF('" & 號數_後 & "', ''),NULLIF('" & _
                        收入_前 & "', ''),NULLIF('" & 收入_後 & "', ''),NULLIF('" & _
                        支出_前 & "', ''),NULLIF('" & 支出_後 & "', ''),NULLIF(N'" & _
                        備註_前 & "', ''),NULLIF(N'" & 備註_後 & "', ''),NULLIF(N'" & _
                        date1 &  "')"
                        data.insertCommAnd = insert1
                        data.insert()
                        '查詢修改資料表的ID，並把他給日誌
                        sql = "SELECT * FROM [修改資料] where id_收=" & id_改 & "And date = '" & date1 & "'"
                        data.SelectCommAnd = sql
                        data_dv = data.Select(New DataSourceSelectArguments)
                        dim id_2 As string =data_dv(0)("id").ToString()
                        data.insertCommAnd = _
                        "INSERT INTO 日誌 " & _
                        "(id, 動作, 命令,日期,日期2) " & _
                        "VALUES " & _
                        "(N'" & id_改 & "', N'修改', N'修改資料id="& id_2 &"', N'" & date1 & "', '" & date2 & "')"
                        data.insert()
                    End If 
                End If 
            End If 
            '新增修改結束，刪除科目表之科目
            data.DeleteCommand = "DELETE FROM 科目表 WHERE id=" & id
            data.Delete()
            Me.GridView1.DataBind()
            Label1.Text="已刪除成功"
        End If 
    End Sub
End Class