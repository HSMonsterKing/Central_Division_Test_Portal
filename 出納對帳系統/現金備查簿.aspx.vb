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
Partial Class 現金備查簿
    Inherits System.Web.UI.Page
    Dim con_14 As String = System.Web.Configuration.WebConfigurationManager.ConnectionStrings("ApplicationServices").ConnectionString
    Dim data As New SqlDataSource
    Dim data_dv As Data.DataView
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        data.ConnectionString = con_14
        Me.GridView1.PageSize = Me.PageSize.Text
        If Not Page.IsPostBack Then
            Me.TextBox1.Text = (DateTime.Now.Year - 1911).ToString()
            Me.GridView1.PageIndex = Int32.MaxValue
        Else
            Me.Label1.Text = ""
            Me.Label2.Text = ""
        End If
        
        'Me.Label1.Text = System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription
        'Me.Label1.Text = (From num In {"國字"} Select num)(0)
        'Me.Label1.Text = (From num In {"國字"} Select System.Data.Objects.SqlClient.SqlFunctions.DataLength(num))(0)
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
                
                TextBox1 = If(Me.GridView1.Columns(4).Visible, TextBox1, "NULL")
                TextBox2 = If(Me.GridView1.Columns(5).Visible, TextBox2, "NULL")
                TextBox3 = If(Me.GridView1.Columns(6).Visible, TextBox3, "NULL")
                TextBox4 = If(Me.GridView1.Columns(7).Visible, TextBox4, "NULL")
                TextBox5 = If(Me.GridView1.Columns(8).Visible, TextBox5, "NULL")
                TextBox6 = If(Me.GridView1.Columns(9).Visible, TextBox6, "NULL")
                TextBox7 = If(Me.GridView1.Columns(10).Visible, TextBox7, "NULL")
                TextBox8 = If(Me.GridView1.Columns(11).Visible, TextBox8, "NULL")
                TextBox9 = If(Me.GridView1.Columns(12).Visible, TextBox9, "NULL")
                TextBox10 = If(Me.GridView1.Columns(13).Visible, TextBox10, "NULL")
                TextBox11 = If(Me.GridView1.Columns(14).Visible, TextBox11, "NULL")
                TextBox12 = If(Me.GridView1.Columns(15).Visible, TextBox12, "NULL")
                
                data.UpdateCommand = "UPDATE 現金備查簿 SET " & _
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
        System.IO.File.Copy(MapPath(".\Excel\現金備查簿.xls"), MyExcel)
        Dim xlApp As New Excel.ApplicationClass()
        xlApp.DisplayAlerts = False
        xlApp.ScreenUpdating = false
        xlApp.EnableEvents = false
        Dim xlWorkBook As Excel.Workbook = xlApp.Workbooks.Open(MyExcel, 0, False, 5, "", "", True, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", False, False, 0, True, 1, 0)
        Dim xlWorkSheet As Excel.Worksheet
        
        xlWorkSheet = CType(xlWorkBook.Sheets("備查簿"), Excel.Worksheet)
        xlWorkSheet.Activate()
        Dim _年 As String = Me.TextBox1.Text
        xlWorkSheet.PageSetup.CenterHeader = Regex.Replace(xlWorkSheet.PageSetup.CenterHeader, "[0-9]{3}年度", _年 & "年度")
        data.SelectCommand = "select * from 現金備查簿 where 年 = '" & _年 & "' order by case when 序號 is null then 1 else 0 end, 序號, 傳票號碼"
        data_dv = data.Select(New DataSourceSelectArguments)
        'data_dv = Me.SqlDataSource1.Select(New DataSourceSelectArguments)
        Dim arr(data_dv.Count, 13) As Object
        For i = 0 To data_dv.Count - 1
            Dim _序號 As String = data_dv(i)("序號").ToString()
            Dim _結帳日期 As String = data_dv(i)("結帳日期").ToString()
            Dim _種類 As String = data_dv(i)("種類").ToString()
            Dim _傳票號碼 As String = data_dv(i)("傳票號碼").ToString()
            Dim _會計科目及摘要 As String = data_dv(i)("會計科目及摘要").ToString()
            Dim _支票編號 As String = data_dv(i)("支票編號").ToString()
            Dim _收入金額405 As String = data_dv(i)("收入金額405").ToString()
            Dim _支出金額405 As String = data_dv(i)("支出金額405").ToString()
            Dim _餘額405 As String = data_dv(i)("餘額405").ToString()
            Dim _收入金額409 As String = data_dv(i)("收入金額409").ToString()
            Dim _支出金額409 As String = data_dv(i)("支出金額409").ToString()
            Dim _餘額409 As String = data_dv(i)("餘額409").ToString()
            Dim _廠商及備註 As String = data_dv(i)("廠商及備註").ToString()
            
            _結帳日期 = totaiwancalendar(_結帳日期)
            
            arr(i, 0) = _序號
            arr(i, 1) = _結帳日期
            arr(i, 2) = _種類
            arr(i, 3) = _傳票號碼
            arr(i, 4) = _會計科目及摘要
            arr(i, 5) = _支票編號.Replace(Chr(13), "").Trim(Chr(10))
            arr(i, 6) = _收入金額405
            arr(i, 7) = _支出金額405
            arr(i, 8) = _餘額405
            arr(i, 9) = _收入金額409
            arr(i, 10) = _支出金額409
            arr(i, 11) = _餘額409
            arr(i, 12) = _廠商及備註
        Next
        
        xlWorkSheet.Range(xlWorkSheet.Cells(4, 1), xlWorkSheet.Cells(4 + data_dv.Count - 1, 13)).Value = arr
        
        If _年 <> "109"
            data.SelectCommand = "select top 1 餘額405, 餘額409 from 現金備查簿 where 年 = '" & (CLng(_年) - 1).ToString() & "' order by 序號 desc"
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
        Dim downloadfilename = "現金備查簿" & _年 & "年" & Now.ToString(" (MM月dd日HH點mm分下載)") & ".xls"
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
            System.IO.File.Copy(MapPath(".\Excel\現金備查簿.xls"), MyExcel)
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
        End If
    End Sub
    '換頁前先自動存檔，未完成如果存檔失敗則不換頁，另外會跑很慢。
    'Protected Sub GridView1_PageIndexChanging(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.PageIndexChanging
    '    Update(sender, e)
    'End Sub
    <System.Web.Script.Services.ScriptMethod(), System.Web.Services.WebMethod()>
    Public Shared Function GetMyList(ByVal prefixText As String, ByVal count As Integer)
        Dim con_14 As String = System.Web.Configuration.WebConfigurationManager.ConnectionStrings("ApplicationServices").ConnectionString
        Dim data As New SqlDataSource
        Dim data_dv As Data.DataView
        Dim MyList As New List(Of String)
        data.ConnectionString = con_14
        data.SelectCommand = _
            "WITH CTE AS " & _
            "(SELECT DISTINCT TOP " & count & " 廠商及備註 FROM 現金備查簿 WHERE 廠商及備註 LIKE N'%" & prefixText & "%') " & _
            "SELECT * FROM CTE " & _
            "ORDER BY " & _
            "CASE WHEN (廠商及備註 LIKE N'" & prefixText & "%') THEN 0 ELSE 1 END, " & _
            "廠商及備註"
        data_dv = data.Select(New DataSourceSelectArguments)
        For i As Long = 0 To data_dv.Count() - 1
            MyList.Add(data_dv(i)(0).ToString())
        Next
        Return MyList
    End Function
End Class