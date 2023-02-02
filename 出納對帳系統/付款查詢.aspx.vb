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
Partial Class 付款查詢
    Inherits System.Web.UI.Page
    Dim con_14 As String = System.Web.Configuration.WebConfigurationManager.ConnectionStrings("ApplicationServices").ConnectionString
    Dim data As New SqlDataSource
    Dim data_dv As Data.DataView
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        data.ConnectionString = con_14
        Me.GridView1.PageSize = Me.每頁筆數.Text
        
        If Not Page.IsPostBack Then
            Try
                Me.年A.SelectedValue = Request.Cookies("yearA").Value
            Catch ex As Exception
            End Try
            Try
                Me.年B.SelectedValue = Request.Cookies("yearB").Value
            Catch ex As Exception
            End Try
            
            選項_SelectedIndexChanged(sender, e)
        Else
        End If
    End Sub
    Protected Sub 選項_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Select Case Me.選項.SelectedValue
            Case "405"
            Case "409"
        End Select
    End Sub
    Protected Sub 年A_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim 年A As New HttpCookie("yearA")
        年A.Value = Me.年A.SelectedValue
        年A.Expires = DateTime.MaxValue
        Response.Cookies.Set(年A)
    End Sub
    Protected Sub 年B_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim 年B As New HttpCookie("yearB")
        年B.Value = Me.年B.SelectedValue
        年B.Expires = DateTime.MaxValue
        Response.Cookies.Set(年B)
    End Sub
    Protected Sub Download(ByVal sender As Object, ByVal e As System.EventArgs)
        Select Case Me.選項.SelectedValue
            Case "405"
                Dim _GUID As String = Guid.NewGuid().ToString("N")
                Dim MyExcel As String = MapPath(".\Excel\Temp\") & _GUID & ".xls"
                System.IO.File.Copy(MapPath(".\Excel\年度報表.xls"), MyExcel)
                Dim xlApp As New Excel.ApplicationClass()
                xlApp.DisplayAlerts = False
                xlApp.ScreenUpdating = false
                xlApp.EnableEvents = false
                Dim xlWorkBook As Excel.Workbook = xlApp.Workbooks.Open(MyExcel, 0, False, 5, "", "", True, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", False, False, 0, True, 1, 0)
                Dim xlWorkSheet As Excel.Worksheet
                
                xlWorkSheet = CType(xlWorkBook.Sheets("年度報表"), Excel.Worksheet)
                'data.SelectCommand = "" & _
                '    "SELECT DENSE_RANK() OVER (ORDER BY 名稱) AS 序號, 名稱, 傳票號碼, 預付日期 as 付款日, 支出金額, CASE WHEN (匯入帳號 IS NULL OR 匯入帳號 = '') THEN 登錄序號 ELSE 匯入帳號 END AS 匯入帳號, 匯入銀行名稱, 摘要說明 " & _
                '    "FROM 傳票資料 " & _
                '    "WHERE (年 = '" & Me.年A.SelectedValue & "') " & _
                '    "AND (支出金額 > 0 AND 支出金額 IS NOT NULL) " & _
                '    "ORDER BY " & _
                '    "名稱, 傳票號碼, 付款日, 匯入銀行名稱, 匯入帳號, 支出金額"
                data.SelectCommand = _
                    "SELECT " & _
                        "DENSE_RANK() OVER (ORDER BY 傳票資料.名稱) AS 序號, " & _
                        "傳票資料.傳票號碼, " & _
                        "現金備查簿.會計科目及摘要 AS 摘要, " & _
                        "傳票資料.預付日期 AS 付款日, " & _
                        "傳票資料.支出金額 AS 金額, " & _
                        "傳票資料.登錄序號 AS 支票編號, " & _
                        "傳票資料.名稱, " & _
                        "傳票資料.匯入銀行名稱 AS 銀行名稱, " & _
                        "傳票資料.匯入帳號 AS 帳號 " & _
                    "FROM 傳票資料 INNER JOIN 現金備查簿 ON 傳票資料.年 = 現金備查簿.年 AND 傳票資料.傳票號碼 = 現金備查簿.傳票號碼 " & _
                    "WHERE (傳票資料.年 = '" & Me.年A.SelectedValue & "') " & _
                        "AND (現金備查簿.收入金額405 > 0 OR 現金備查簿.支出金額405 > 0) " & _
                        "AND (傳票資料.支出金額 > 0 AND 傳票資料.支出金額 IS NOT NULL) " & _
                    "ORDER BY " & _
                        "序號, " & _
                        "CASE WHEN 傳票資料.傳票號碼 IS NULL THEN 1 ELSE 0 END, 傳票資料.傳票號碼, " & _
                        "CASE WHEN 傳票資料.預付日期 IS NULL THEN 1 ELSE 0 END, 傳票資料.預付日期, " & _
                        "CASE WHEN 傳票資料.名稱 IS NULL THEN 1 ELSE 0 END, 傳票資料.名稱"
                data_dv = data.Select(New DataSourceSelectArguments)
                Dim arr(data_dv.Count, 9) As Object
                For i = 0 To data_dv.Count - 1
                    Dim _序號 As String = data_dv(i)("序號").ToString()
                    Dim _名稱 As String = data_dv(i)("名稱").ToString()
                    Dim _銀行名稱 As String = data_dv(i)("銀行名稱").ToString()
                    Dim _帳號 As String = data_dv(i)("帳號").ToString()
                    Dim _支票編號 As String = data_dv(i)("支票編號").ToString()
                    Dim _金額 As String = data_dv(i)("金額").ToString()
                    Dim _付款日 As String = data_dv(i)("付款日").ToString()
                    _付款日 = totaiwancalendar(_付款日)
                    Dim _摘要 As String = data_dv(i)("摘要").ToString()
                    Dim _傳票號碼 As String = data_dv(i)("傳票號碼").ToString()
                    
                    arr(i, 0) = _序號'.Replace(Chr(13), "").Trim(Chr(10))
                    arr(i, 1) = _名稱
                    arr(i, 2) = _銀行名稱
                    arr(i, 3) = _帳號
                    arr(i, 4) = _支票編號
                    arr(i, 5) = _金額
                    arr(i, 6) = _付款日
                    arr(i, 7) = _摘要
                    arr(i, 8) = _傳票號碼
                Next
                xlWorkSheet.Range(xlWorkSheet.Cells(2, 1), xlWorkSheet.Cells(2 + data_dv.Count - 1, 9)).Value = arr
                xlWorkSheet.Range(xlWorkSheet.Cells(2, 1), xlWorkSheet.Cells(2 + data_dv.Count - 1, 9)).RowHeight = 15
                
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
                Dim downloadfilename
                downloadfilename = "405付款" & Me.年A.SelectedValue & "年度報表.xls"
                Response.AddHeader("Content-Disposition", "attachment;FileName=" + Uri.EscapeDataString(downloadfilename))
                Response.WriteFile(MyExcel)
                System.Web.HttpContext.Current.ApplicationInstance.CompleteRequest()
                Response.Flush()
                System.IO.File.Delete(MyExcel)
                Response.End()
            Case "409"
                Dim _GUID As String = Guid.NewGuid().ToString("N")
                Dim MyExcel As String = MapPath(".\Excel\Temp\") & _GUID & ".xls"
                System.IO.File.Copy(MapPath(".\Excel\年度報表.xls"), MyExcel)
                Dim xlApp As New Excel.ApplicationClass()
                xlApp.DisplayAlerts = False
                xlApp.ScreenUpdating = false
                xlApp.EnableEvents = false
                Dim xlWorkBook As Excel.Workbook = xlApp.Workbooks.Open(MyExcel, 0, False, 5, "", "", True, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", False, False, 0, True, 1, 0)
                Dim xlWorkSheet As Excel.Worksheet
                
                xlWorkSheet = CType(xlWorkBook.Sheets("年度報表"), Excel.Worksheet)
                data.SelectCommand = _
                    "SELECT " & _
                        "DENSE_RANK() OVER (ORDER BY 傳票資料.名稱) AS 序號, " & _
                        "傳票資料.傳票號碼, " & _
                        "現金備查簿.會計科目及摘要 AS 摘要, " & _
                        "傳票資料.預付日期 AS 付款日, " & _
                        "傳票資料.支出金額 AS 金額, " & _
                        "傳票資料.登錄序號 AS 支票編號, " & _
                        "傳票資料.名稱, " & _
                        "傳票資料.匯入銀行名稱 AS 銀行名稱, " & _
                        "傳票資料.匯入帳號 AS 帳號 " & _
                    "FROM 傳票資料 INNER JOIN 現金備查簿 ON 傳票資料.年 = 現金備查簿.年 AND 傳票資料.傳票號碼 = 現金備查簿.傳票號碼 " & _
                    "WHERE (傳票資料.年 = '" & Me.年A.SelectedValue & "') " & _
                        "AND (現金備查簿.支出金額409 > 0) " & _
                    "ORDER BY " & _
                        "序號, " & _
                        "CASE WHEN 傳票資料.傳票號碼 IS NULL THEN 1 ELSE 0 END, 傳票資料.傳票號碼, " & _
                        "CASE WHEN 傳票資料.預付日期 IS NULL THEN 1 ELSE 0 END, 傳票資料.預付日期, " & _
                        "CASE WHEN 傳票資料.名稱 IS NULL THEN 1 ELSE 0 END, 傳票資料.名稱"
                data_dv = data.Select(New DataSourceSelectArguments)
                Dim arr(data_dv.Count, 9) As Object
                For i = 0 To data_dv.Count - 1
                    Dim _序號 As String = data_dv(i)("序號").ToString()
                    Dim _名稱 As String = data_dv(i)("名稱").ToString()
                    Dim _銀行名稱 As String = data_dv(i)("銀行名稱").ToString()
                    Dim _帳號 As String = data_dv(i)("帳號").ToString()
                    Dim _支票編號 As String = data_dv(i)("支票編號").ToString()
                    Dim _金額 As String = data_dv(i)("金額").ToString()
                    Dim _付款日 As String = data_dv(i)("付款日").ToString()
                    _付款日 = totaiwancalendar(_付款日)
                    Dim _摘要 As String = data_dv(i)("摘要").ToString()
                    Dim _傳票號碼 As String = data_dv(i)("傳票號碼").ToString()
                    
                    arr(i, 0) = _序號'.Replace(Chr(13), "").Trim(Chr(10))
                    arr(i, 1) = _名稱
                    arr(i, 2) = _銀行名稱
                    arr(i, 3) = _帳號
                    arr(i, 4) = _支票編號
                    arr(i, 5) = _金額
                    arr(i, 6) = _付款日
                    arr(i, 7) = _摘要
                    arr(i, 8) = _傳票號碼
                Next
                xlWorkSheet.Range(xlWorkSheet.Cells(2, 1), xlWorkSheet.Cells(2 + data_dv.Count - 1, 9)).Value = arr
                xlWorkSheet.Range(xlWorkSheet.Cells(2, 1), xlWorkSheet.Cells(2 + data_dv.Count - 1, 9)).RowHeight = 15
                
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
                Dim downloadfilename
                downloadfilename = "409付款" & Me.年A.SelectedValue & "年度報表.xls"
                Response.AddHeader("Content-Disposition", "attachment;FileName=" + Uri.EscapeDataString(downloadfilename))
                Response.WriteFile(MyExcel)
                System.Web.HttpContext.Current.ApplicationInstance.CompleteRequest()
                Response.Flush()
                System.IO.File.Delete(MyExcel)
                Response.End()
        End Select
    End Sub
    Protected Sub GridView1_RowCommand(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles GridView1.RowCommand
        Select Case Me.選項.SelectedValue
            Case "405"
            Case "409"
        End Select
    End Sub
    <System.Web.Script.Services.ScriptMethod(), System.Web.Services.WebMethod()>
    Public Shared Function GetMyList(ByVal prefixText As String, ByVal count As Integer)
        Dim con_14 As String = System.Web.Configuration.WebConfigurationManager.ConnectionStrings("ApplicationServices").ConnectionString
        Dim data As New SqlDataSource
        Dim data_dv As Data.DataView
        Dim MyList As New List(Of String)
        data.ConnectionString = con_14
        data.SelectCommand = _
            "WITH CTE AS " & _
            "(SELECT DISTINCT TOP " & count & " 名稱 FROM 傳票資料 WHERE 名稱 LIKE N'%" & prefixText & "%') " & _
            "SELECT * FROM CTE " & _
            "ORDER BY " & _
            "CASE WHEN (名稱 LIKE N'" & prefixText & "%') THEN 0 ELSE 1 END, " & _
            "名稱"
        data_dv = data.Select(New DataSourceSelectArguments)
        For i As Long = 0 To data_dv.Count() - 1
            MyList.Add(data_dv(i)(0).ToString())
        Next
        Return MyList
    End Function
End Class
