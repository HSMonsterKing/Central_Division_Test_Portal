Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel.XlDirection
Imports Microsoft.VisualBasic.Logging
Imports System.IO
Imports System.Drawing
Imports System.Diagnostics
Imports System.Data.Sql
Imports System.Data.SqlTypes
Imports System.Data.SqlClient
Imports System.Configuration
Imports System.Text.RegularExpressions
Partial Class 分錄
    Inherits System.Web.UI.Page
    Dim con_14 As String = System.Web.Configuration.WebConfigurationManager.ConnectionStrings("ApplicationServices").ConnectionString
    Dim data As New SqlDataSource
    Dim data_dv As Data.DataView
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        data.ConnectionString = con_14
        Me.GridView1.PageSize = Me.PageSize.Text
        If Not Page.IsPostBack Then
            Me.GridView1.Columns(8).Visible = False
            Me.GridView1.Columns(9).Visible = False
            Me.GridView1.Columns(10).HeaderText = "金額"
            Me.GridView1.Columns(11).Visible = False
            Me.GridView1.Columns(12).Visible = False
            Me.GridView1.Columns(13).Visible = False
            Me.GridView1.Columns(14).Visible = False
            Me.GridView1.Columns(15).Visible = False
            Me.GridView1.Columns(16).Visible = False
            Me.GridView1.Columns(18).Visible = False
            
            Me.TextBox1.Text = (DateTime.Now.Year - 1911).ToString()
            Me.GridView1.PageIndex = Int32.MaxValue
        Else
            Me.Label1.Text = ""
            Me.Label2.Text = ""
        End If
    End Sub
    Protected Sub Insert(ByVal sender As Object, ByVal e As System.EventArgs)
        Me.SqlDataSource1.Insert()
        Me.GridView1.PageIndex = Int32.MaxValue
    End Sub
    Protected Sub Update(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim Result As Long = 1
        For i = 0 To Me.GridView1.Rows.Count - 1
            Try
                Dim Label1 As String = CType(Me.GridView1.Rows(i).FindControl("Label1"), Label).Text
                Dim 開票日期 As String = CType(Me.GridView1.Rows(i).FindControl("開票日期"), TextBox).Text
                Dim TextBox1 As String = CType(Me.GridView1.Rows(i).FindControl("TextBox1"), TextBox).Text
                Dim TextBox2 As String = CType(Me.GridView1.Rows(i).FindControl("TextBox2"), TextBox).Text
                Dim TextBox3 As String = CType(Me.GridView1.Rows(i).FindControl("TextBox3"), TextBox).Text
                Dim TextBox4 As String = CType(Me.GridView1.Rows(i).FindControl("TextBox4"), TextBox).Text
                Dim TextBox5 As String = CType(Me.GridView1.Rows(i).FindControl("TextBox5"), TextBox).Text
                Dim TextBox6 As String = CType(Me.GridView1.Rows(i).FindControl("TextBox6"), TextBox).Text
                Dim TextBox7 As String = CType(Me.GridView1.Rows(i).FindControl("TextBox7"), TextBox).Text
                Dim TextBox8 As String = CType(Me.GridView1.Rows(i).FindControl("TextBox8"), TextBox).Text
                Dim TextBox9 As String = CType(Me.GridView1.Rows(i).FindControl("TextBox9"), TextBox).Text
                Dim TextBox10 As String = CType(Me.GridView1.Rows(i).FindControl("TextBox10"), TextBox).Text
                Dim TextBox11 As String = CType(Me.GridView1.Rows(i).FindControl("TextBox11"), TextBox).Text
                Dim TextBox12 As String = CType(Me.GridView1.Rows(i).FindControl("TextBox12"), TextBox).Text
                
                TextBox6 = TextBox6.Replace(",", "")
                TextBox7 = TextBox7.Replace(",", "")
                TextBox8 = TextBox8.Replace(",", "")
                TextBox9 = TextBox9.Replace(",", "")
                TextBox10 = TextBox10.Replace(",", "")
                TextBox11 = TextBox11.Replace(",", "")
                
                開票日期 = 開票日期.Replace("'", "")
                TextBox1 = TextBox1.Replace("'", "")
                TextBox2 = TextBox2.Replace("'", "")
                TextBox3 = TextBox3.Replace("'", "")
                TextBox4 = TextBox4.Replace("'", "")
                TextBox5 = TextBox5.Replace("'", "")
                TextBox6 = TextBox6.Replace("'", "")
                TextBox7 = TextBox7.Replace("'", "")
                TextBox8 = TextBox8.Replace("'", "")
                TextBox9 = TextBox9.Replace("'", "")
                TextBox10 = TextBox10.Replace("'", "")
                TextBox11 = TextBox11.Replace("'", "")
                TextBox12 = TextBox12.Replace("'", "")
                
                開票日期 = "N'" & 開票日期 & "'"
                TextBox1 = "N'" & TextBox1 & "'"
                TextBox2 = "N'" & TextBox2 & "'"
                TextBox3 = "N'" & TextBox3 & "'"
                TextBox4 = "N'" & TextBox4 & "'"
                TextBox5 = "N'" & TextBox5 & "'"
                TextBox6 = "N'" & TextBox6 & "'"
                TextBox7 = "N'" & TextBox7 & "'"
                TextBox8 = "N'" & TextBox8 & "'"
                TextBox9 = "N'" & TextBox9 & "'"
                TextBox10 = "N'" & TextBox10 & "'"
                TextBox11 = "N'" & TextBox11 & "'"
                TextBox12 = "N'" & TextBox12 & "'"
                
                開票日期 = If(Me.GridView1.Columns(2).Visible, 開票日期, "NULL")
                TextBox1 = If(Me.GridView1.Columns(5).Visible, TextBox1, "NULL")
                TextBox2 = If(Me.GridView1.Columns(6).Visible, TextBox2, "NULL")
                TextBox3 = If(Me.GridView1.Columns(7).Visible, TextBox3, "NULL")
                TextBox4 = If(Me.GridView1.Columns(8).Visible, TextBox4, "NULL")
                TextBox5 = If(Me.GridView1.Columns(9).Visible, TextBox5, "NULL")
                TextBox6 = If(Me.GridView1.Columns(10).Visible, TextBox6, "NULL")
                TextBox7 = If(Me.GridView1.Columns(11).Visible, TextBox7, "NULL")
                TextBox8 = If(Me.GridView1.Columns(12).Visible, TextBox8, "NULL")
                TextBox9 = If(Me.GridView1.Columns(14).Visible, TextBox9, "NULL")
                TextBox10 = If(Me.GridView1.Columns(15).Visible, TextBox10, "NULL")
                TextBox11 = If(Me.GridView1.Columns(16).Visible, TextBox11, "NULL")
                TextBox12 = If(Me.GridView1.Columns(17).Visible, TextBox12, "NULL")
                
                data.UpdateCommand = "UPDATE 分錄 SET " & _
                "開票日期 = NULLIF(ISNULL(" & 開票日期 & ", 開票日期),''), " & _
                "種類 = NULLIF(ISNULL(" & TextBox1 & ", 種類),''), " & _
                "傳票號碼 = NULLIF(ISNULL(" & TextBox2 & ", 傳票號碼),''), " & _
                "會計科目及摘要 = NULLIF(ISNULL(" & TextBox3 & ", 會計科目及摘要),''), " & _
                "付款日 = NULLIF(ISNULL(" & TextBox4 & ", 付款日),''), " & _
                "支票編號 = NULLIF(ISNULL(" & TextBox5 & ", 支票編號),''), " & _
                "收入金額405 = NULLIF(ISNULL(" & TextBox6 & ", 收入金額405),''), " & _
                "支出金額405 = NULLIF(ISNULL(" & TextBox7 & ", 支出金額405),''), " & _
                "餘額405 = NULLIF(ISNULL(" & TextBox8 & ", 餘額405),''), " & _
                "收入金額409 = NULLIF(ISNULL(" & TextBox9 & ", 收入金額409),''), " & _
                "支出金額409 = NULLIF(ISNULL(" & TextBox10 & ", 支出金額409),''), " & _
                "餘額409 = NULLIF(ISNULL(" & TextBox11 & ", 餘額409),''), " & _
                "廠商及備註 = NULLIF(ISNULL(" & TextBox12 & ", 廠商及備註),'') " & _
                "WHERE id = '" & Label1 & "'"
                data.Update()
            Catch
                Result = 2
            End Try
            data.Update()
        Next
        If Result = 1
            Me.Label1.Text = "成功"
            Me.Label2.Text = ""
            Me.GridView1.DataBind()
        Else If Result = 2
            Me.Label1.Text = ""
            Me.Label2.Text = "失敗"
        End If
    End Sub
    Protected Sub Download(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim MyGUID As String = Guid.NewGuid().ToString("N")
        Dim MyExcel As String = MapPath(".\Excel\Temp\") & MyGUID & ".xls"
        System.IO.File.Copy(MapPath(".\Excel\分錄.xls"), MyExcel)
        Dim xlApp As New Excel.ApplicationClass()
        xlApp.DisplayAlerts = False
        xlApp.ScreenUpdating = false
        xlApp.EnableEvents = false
        Dim xlWorkBook As Excel.Workbook = xlApp.Workbooks.Open(MyExcel, 0, False, 5, "", "", True, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", False, False, 0, True, 1, 0)
        Dim xlWorkSheet As Excel.Worksheet
        
        xlWorkSheet = CType(xlWorkBook.Sheets("分錄"), Excel.Worksheet)
        xlWorkSheet.Activate()
        Dim _年 As String = Me.TextBox1.Text
        xlWorkSheet.PageSetup.CenterHeader = Regex.Replace(xlWorkSheet.PageSetup.CenterHeader, "[0-9]{3}年度", _年 & "年度")
        'data.SelectCommand = _
        '"select 開票日期, 序號, 結帳日期, 種類, 傳票號碼, 會計科目及摘要, 收入金額405, 廠商及備註 " & _
        '"from 分錄 where 年 = '" & _年 & "' order by case when 序號 is null then 1 else 0 end, 序號, 傳票號碼"
        'data_dv = data.Select(New DataSourceSelectArguments)
        data_dv = Me.SqlDataSource1.Select(New DataSourceSelectArguments)
        Dim arr(data_dv.Count, 13) As Object
        For i = 0 To data_dv.Count - 1
            Dim _開票日期 As String = data_dv(i)("開票日期").ToString()
            Dim _序號 As String = data_dv(i)("序號").ToString()
            Dim _結帳日期 As String = data_dv(i)("結帳日期").ToString()
            Dim _種類 As String = data_dv(i)("種類").ToString()
            Dim _傳票號碼 As String = data_dv(i)("傳票號碼").ToString()
            Dim _會計科目及摘要 As String = data_dv(i)("會計科目及摘要").ToString()
            Dim _收入金額405 As String = data_dv(i)("收入金額405").ToString()
            Dim _廠商及備註 As String = data_dv(i)("廠商及備註").ToString()
            
            _結帳日期 = totaiwancalendar(_結帳日期)
            
            arr(i, 0) = _開票日期
            arr(i, 1) = _序號
            arr(i, 2) = _結帳日期
            arr(i, 3) = _種類
            arr(i, 4) = _傳票號碼
            arr(i, 5) = _會計科目及摘要
            arr(i, 6) = _收入金額405
            arr(i, 7) = _廠商及備註
        Next
        
        xlWorkSheet.Range(xlWorkSheet.Cells(2, 1), xlWorkSheet.Cells(2 + data_dv.Count - 1, 13)).Value = arr
        
        If _年 <> "109"
            data.SelectCommand = "select top 1 餘額405, 餘額409 from 分錄 where 年 = '" & (CLng(_年) - 1).ToString() & "' order by 序號 desc"
            data_dv = data.Select(New DataSourceSelectArguments)
            If data_dv.Count > 0
                xlWorkSheet.Cells(3, 9) = data_dv(0)(0)
                xlWorkSheet.Cells(3, 12) = data_dv(0)(1)
            End If
        End If
        xlWorkSheet.Range("E:E").WrapText = False
        xlWorkSheet.Range("M:M").WrapText = False
        
        xlWorkBook.Save()
        xlWorkBook.Close()
        xlApp.Quit()
        ReleaseObject(xlWorkSheet)
        ReleaseObject(xlWorkBook)
        ReleaseObject(xlApp)
        
        Response.Clear()
        Response.ClearHeaders()
        Response.Buffer = True
        Response.ContentType = "application/octet-stream"
        Dim downloadfilename = "分錄.xls"
        Response.AddHeader("Content-Disposition", "attachment;FileName=" + Uri.EscapeDataString(downloadfilename))
        Response.WriteFile(MyExcel)
        System.Web.HttpContext.Current.ApplicationInstance.CompleteRequest()
        Response.Flush()
        System.IO.File.Delete(MyExcel)
        Response.End()
    End Sub
    Protected Sub GridView1_RowCommand(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles GridView1.RowCommand
        If e.CommandName = "Download"
            Dim MyGUID As String = Guid.NewGuid().ToString("N")
            Dim MyExcel As String = MapPath(".\Excel\Temp\") & MyGUID & ".xls"
            System.IO.File.Copy(MapPath(".\Excel\分錄.xls"), MyExcel)
            Dim xlApp As New Excel.ApplicationClass()
            xlApp.DisplayAlerts = False
            xlApp.ScreenUpdating = false
            xlApp.EnableEvents = false
            Dim xlWorkBook As Excel.Workbook = xlApp.Workbooks.Open(MyExcel, 0, False, 5, "", "", True, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", False, False, 0, True, 1, 0)
            Dim xlWorkSheet As Excel.Worksheet
            
            xlWorkSheet = CType(xlWorkBook.Sheets("支出憑證黏存單"), Excel.Worksheet)
            xlWorkSheet.Activate()
            
            xlWorkSheet.Cells(3, 1).Value = "所屬年度：" & Me.TextBox1.Text
            Dim CurrentRow = Me.GridView1.Rows(e.CommandSource.NamingContainer.RowIndex)
            xlWorkSheet.Cells(8, 16).Value = CType(CurrentRow.FindControl("TextBox3"), TextBox).Text
            Dim a As String = CType(CurrentRow.FindControl("TextBox7"), TextBox).Text
            Dim b As String = CType(CurrentRow.FindControl("TextBox10"), TextBox).Text
            a = Trim(a)
            b = Trim(b)
            a = a.Replace(",", "")
            b = b.Replace(",", "")
            If a = "" Or a = "0"
                a = b
            End If
            For i = 0 To 9
                If a.Length > 0
                    xlWorkSheet.Cells(8, 13 - i).Value = a.Substring(a.Length - 1)
                    a = a.Substring(0, a.Length - 1)
                End If
            Next
            
            xlWorkBook.Save()
            xlWorkBook.Close()
            xlApp.Quit()
            ReleaseObject(xlWorkSheet)
            ReleaseObject(xlWorkBook)
            ReleaseObject(xlApp)
            
            Response.Clear()
            Response.ClearHeaders()
            Response.Buffer = True
            Response.ContentType = "application/octet-stream"
            Dim downloadfilename = "支出憑證黏存單 " & CType(CurrentRow.FindControl("TextBox2"), TextBox).Text & ".xls"
            Response.AddHeader("Content-Disposition", "attachment;FileName=" + Uri.EscapeDataString(downloadfilename))
            Response.WriteFile(MyExcel)
            System.Web.HttpContext.Current.ApplicationInstance.CompleteRequest()
            Response.Flush()
            System.IO.File.Delete(MyExcel)
            Response.End()
        Else If e.CommandName = "轉保證書明細表"
            Update(sender, e)
            
            Dim CurrentRow = Me.GridView1.Rows(e.CommandSource.NamingContainer.RowIndex)
            Dim Label1 As String = CType(CurrentRow.FindControl("Label1"), Label).Text
            
            data.UpdateCommand = _
                "INSERT INTO 保管品明細表 " & _
                "(種類, 已收入, 已支出, 日期, 國保收據編號, 摘要, 單位, 數量, 金額, 合約展延情形) " & _
                "VALUES " & _
                "(0, 0, 0, " & _
                "(SELECT " & _
                    "TRY_PARSE( " & _
                        "STR(CAST(SUBSTRING(開票日期, 1, LEN(開票日期) - 4) AS int) + 1911) " & _
                        "+ '/' + SUBSTRING(開票日期, LEN(開票日期) - 3, 2) " & _
                        "+ '/' + SUBSTRING(開票日期, LEN(開票日期) - 1, 2) " & _
                    "AS date) " & _
                "FROM 分錄 WHERE id = N'" & Label1 & "'), " & _
                "(SELECT ISNULL(廠商及備註, '') FROM 分錄 WHERE id = N'" & Label1 & "'), " & _
                "(SELECT ISNULL(會計科目及摘要, '') FROM 分錄 WHERE id = N'" & Label1 & "'), N'包', 1, " & _
                "(SELECT ISNULL(收入金額405, 0) FROM 分錄 WHERE id = N'" & Label1 & "'), N'無') "
            data.Update()
        Else If e.CommandName = "轉定存單明細表"
            Update(sender, e)
            
            Dim CurrentRow = Me.GridView1.Rows(e.CommandSource.NamingContainer.RowIndex)
            Dim Label1 As String = CType(CurrentRow.FindControl("Label1"), Label).Text
            
            data.UpdateCommand = _
                "INSERT INTO 保管品明細表 " & _
                "(種類, 已收入, 已支出, 日期, 國保收據編號, 摘要, 單位, 數量, 金額, 合約展延情形) " & _
                "VALUES " & _
                "(1, 0, 0, " & _
                "(SELECT " & _
                    "TRY_PARSE( " & _
                        "STR(CAST(SUBSTRING(開票日期, 1, LEN(開票日期) - 4) AS int) + 1911) " & _
                        "+ '/' + SUBSTRING(開票日期, LEN(開票日期) - 3, 2) " & _
                        "+ '/' + SUBSTRING(開票日期, LEN(開票日期) - 1, 2) " & _
                    "AS date) " & _
                "FROM 分錄 WHERE id = N'" & Label1 & "'), " & _
                "(SELECT ISNULL(廠商及備註, '') FROM 分錄 WHERE id = N'" & Label1 & "'), " & _
                "(SELECT ISNULL(會計科目及摘要, '') FROM 分錄 WHERE id = N'" & Label1 & "'), N'張', 1, " & _
                "(SELECT ISNULL(收入金額405, 0) FROM 分錄 WHERE id = N'" & Label1 & "'), N'無') "
            data.Update()
        ' Else If e.CommandName = "轉保證書紀錄簿"
        '     Update(sender, e)
            
        '     Dim CurrentRow = Me.GridView1.Rows(e.CommandSource.NamingContainer.RowIndex)
        '     Dim Label1 As String = CType(CurrentRow.FindControl("Label1"), Label).Text
            
        '     data.UpdateCommand = _
        '         "INSERT INTO 保管品紀錄簿 " & _
        '         "(種類, 序號, 日期, 收支, 國保收據編號, 摘要, 收入金額, 支出金額, 餘額) " & _
        '         "(SELECT 0, " & _
        '             "(SELECT ISNULL(MAX(序號), 0)+1 FROM 保管品紀錄簿 WHERE 種類 = 0), " & _
        '             "TRY_PARSE( " & _
        '                 "STR(CAST(SUBSTRING(開票日期, 1, LEN(開票日期) - 4) AS int) + 1911) " & _
        '                 "+ '/' + SUBSTRING(開票日期, LEN(開票日期) - 3, 2) " & _
        '                 "+ '/' + SUBSTRING(開票日期, LEN(開票日期) - 1, 2) " & _
        '             "AS date), " & _
        '             "0, " & _
        '             "ISNULL(廠商及備註, ''), " & _
        '             "ISNULL(會計科目及摘要, ''), " & _
        '             "ISNULL(收入金額405, 0), " & _
        '             "0, " & _
        '             "(SELECT ISNULL((SELECT 餘額 FROM 保管品紀錄簿 WHERE 序號 = (SELECT ISNULL(MAX(序號), 0) FROM 保管品紀錄簿 WHERE 種類 = 0) AND 種類 = 0), 0)) + ISNULL(收入金額405, 0) " & _
        '         "FROM 分錄 WHERE id = N'" & Label1 & "')"
        '     data.Update()
        ' Else If e.CommandName = "轉定存單紀錄簿"
        '     Update(sender, e)
            
        '     Dim CurrentRow = Me.GridView1.Rows(e.CommandSource.NamingContainer.RowIndex)
        '     Dim Label1 As String = CType(CurrentRow.FindControl("Label1"), Label).Text
            
        '     data.UpdateCommand = _
        '         "INSERT INTO 保管品紀錄簿 " & _
        '         "(種類, 序號, 日期, 收支, 國保收據編號, 摘要, 收入金額, 支出金額, 餘額) " & _
        '         "(SELECT 1, " & _
        '             "(SELECT ISNULL(MAX(序號), 0)+1 FROM 保管品紀錄簿 WHERE 種類 = 1), " & _
        '             "TRY_PARSE( " & _
        '                 "STR(CAST(SUBSTRING(開票日期, 1, LEN(開票日期) - 4) AS int) + 1911) " & _
        '                 "+ '/' + SUBSTRING(開票日期, LEN(開票日期) - 3, 2) " & _
        '                 "+ '/' + SUBSTRING(開票日期, LEN(開票日期) - 1, 2) " & _
        '             "AS date), " & _
        '             "0, " & _
        '             "ISNULL(廠商及備註, ''), " & _
        '             "ISNULL(會計科目及摘要, ''), " & _
        '             "ISNULL(收入金額405, 0), " & _
        '             "0, " & _
        '             "(SELECT ISNULL((SELECT 餘額 FROM 保管品紀錄簿 WHERE 序號 = (SELECT ISNULL(MAX(序號), 0) FROM 保管品紀錄簿 WHERE 種類 = 1) AND 種類 = 1), 0)) + ISNULL(收入金額405, 0) " & _
        '         "FROM 分錄 WHERE id = N'" & Label1 & "')"
        '     data.Update()
        End If
    End Sub
End Class
