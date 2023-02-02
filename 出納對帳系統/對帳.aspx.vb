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
Partial Class 對帳
    Inherits System.Web.UI.Page
    Dim con_14 As String = System.Web.Configuration.WebConfigurationManager.ConnectionStrings("ApplicationServices").ConnectionString
    Dim data As New SqlDataSource
    Dim data_dv As Data.DataView
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        data.ConnectionString = con_14
        Me.Label1.Text = ""
        Me.Label2.Text = ""
        If Not Page.IsPostBack Then
            'Me.DropDownList1.Items.Add("405")
            Me.DropDownList1.Items.Add("409")
            DropDownList1_SelectedIndexChanged(sender, e)
        End If
    End Sub
    Protected Sub DropDownList1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
    End Sub
    Protected Sub CheckBox1_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim i As Long = sender.NamingContainer.RowIndex
        Dim id As String = CType(Me.GridView1.Rows(i).FindControl("Label2"), Label).Text
        Dim Checked As Boolean = CType(Me.GridView1.Rows(i).FindControl("CheckBox1"), CheckBox).Checked
        data.UpdateCommand = "UPDATE 對帳 SET 選取 = CAST('" & Checked & "' AS bit) WHERE id = '" & id & "'"
        data.Update()
    End Sub
    Class 調結表一行
        Public 年 As String = ""
        Public 月 As String = ""
        Public 日 As String = ""
        Public 傳票號碼 As String = ""
        Public 支票編號 As String = ""
        Public 小計 As String = ""
        Public 備註 As String = ""
        Public 位置 As String = ""
    End Class
    Protected Sub Download(ByVal sender As Object, ByVal e As System.EventArgs)
        Select Case Me.DropDownList1.SelectedValue
            Case "409"
                '使用者操作錯誤
                data.SelectCommand = _
                "SELECT ISNULL(檔名, '') AS 檔名 FROM 對帳 WHERE 選取 = 1 AND 備註 = '調結表'"
                data_dv = data.Select(New DataSourceSelectArguments)
                If data_dv.Count <> 1
                    ScriptManager.RegisterStartupScript(Me, Page.GetType, "s1", "setTimeout(function(){alert('請選擇一個調結表。');}, 50);", True)
                    Exit Sub
                End If
                data.SelectCommand = _
                "SELECT ISNULL(檔名, '') AS 檔名 FROM 對帳 WHERE 選取 = 1 AND 備註 = '專案代收查詢'"
                data_dv = data.Select(New DataSourceSelectArguments)
                If data_dv.Count = 0
                    ScriptManager.RegisterStartupScript(Me, Page.GetType, "s1", "setTimeout(function(){alert('請選擇專案代收查詢。');}, 50);", True)
                    Exit Sub
                End If
                data.SelectCommand = _
                "SELECT ISNULL(檔名, '') AS 檔名 FROM 對帳 WHERE 選取 = 1 AND 備註 = '帳戶明細'"
                data_dv = data.Select(New DataSourceSelectArguments)
                If data_dv.Count = 0
                    ScriptManager.RegisterStartupScript(Me, Page.GetType, "s1", "setTimeout(function(){alert('請選擇帳戶明細。');}, 50);", True)
                    Exit Sub
                End If
                
                '產生新的銀行存透明細帳
                Dim MyGUID As String = ""
                Dim MyExcel1 As String = ""
                Dim MyExcel2 As String = ""
                Dim xlApp As New Excel.ApplicationClass()
                xlApp.DisplayAlerts = False
                xlApp.ScreenUpdating = false
                xlApp.EnableEvents = false
                Dim xlWorkbook1 As Excel.Workbook = Nothing
                Dim xlWorkbook2 As Excel.Workbook = Nothing
                Dim xlWorksheet1 As Excel.Worksheet = Nothing
                Dim xlWorksheet2 As Excel.Worksheet = Nothing
                
                MyGUID = Guid.NewGuid().ToString("N")
                MyExcel1 = MapPath(".\Excel\Temp\") & MyGUID & ".xlsx"
                System.IO.File.Copy(MapPath(".\Excel\銀行存透明細帳.xlsx"), MyExcel1)
                xlWorkbook1 = xlApp.Workbooks.Open(MyExcel1, 0, False, 5, "", "", True, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", False, False, 0, True, 1, 0)
                
                '宣告dict
                Dim dict As New Dictionary(Of String, Object)
                
                '把明細帳放入dict
                data.SelectCommand = _
                "SELECT 傳票資料.開票日期, 傳票資料.傳票號碼, 傳票資料.之, 傳票資料.摘要說明, 傳票資料.名稱, 傳票資料.收入金額, 傳票資料.支出金額, 傳票資料.登錄序號, " & _
                "NULL AS 備註欄, CASE WHEN 現金備查簿.支票編號 Like '%不予支付%' THEN '※帳務調整，不予支付' ELSE NULL END AS 位置 " & _
                "FROM 傳票資料 INNER JOIN 現金備查簿 " & _
                "ON 傳票資料.年 = 現金備查簿.年 AND 傳票資料.傳票號碼 = 現金備查簿.傳票號碼 " & _
                "WHERE 現金備查簿.結帳日期 BETWEEN '" & taiwancalendarto(Me.結帳日期a.SelectedValue) & "' AND '" & taiwancalendarto(Me.結帳日期b.SelectedValue) & "' " & _
                "AND (現金備查簿.收入金額409 > 0 OR 現金備查簿.支出金額409 > 0) " & _
                "ORDER BY 現金備查簿.結帳日期, 傳票資料.傳票號碼, 傳票資料.之"
                data_dv = data.Select(New DataSourceSelectArguments)
                
                dict.Add("明細帳", New Object(data_dv.Count - 1, 9){})
                
                For i = 0 To data_dv.Count - 1
                    For j = 0 To 9
                        dict("明細帳")(i, j) = data_dv(i)(j).ToString()
                    Next
                Next
                
                xlWorksheet1 = xlWorkbook1.Worksheets("明細帳")
                xlWorksheet1.Range(xlWorksheet1.Cells(2, 1), xlWorksheet1.Cells(dict("明細帳").GetLength(0) + 1, dict("明細帳").GetLength(1))).Value = dict("明細帳")
                
                dict("明細帳") = xlWorksheet1.Range(xlWorksheet1.Cells(1, 1), xlWorksheet1.Cells(xlWorksheet1.Range("B1").End(xlDown).Row, 11)).Value
                
                '把舊調結表放入dict
                data.SelectCommand = _
                "SELECT ISNULL(檔名, '') AS 檔名 FROM 對帳 WHERE 選取 = 1 AND 備註 = '調結表'"
                data_dv = data.Select(New DataSourceSelectArguments)
                
                For i = 0 To data_dv.Count - 1
                    MyGUID = Guid.NewGuid().ToString("N")
                    MyExcel2 = MapPath(".\Excel\Temp\") & MyGUID & ".xls"
                    System.IO.File.Copy(MapPath(".\Excel\對帳\" & data_dv(i)("檔名")), MyExcel2)
                    xlWorkbook2 = xlApp.Workbooks.Open(MyExcel2, 0, False, 5, "", "", True, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", False, False, 0, True, 1, 0)
                    xlWorksheet2 = xlWorkbook2.ActiveSheet
                    xlWorksheet1 = xlWorkbook1.Worksheets(1)
                    xlWorksheet2.Copy(After := xlWorksheet1)
                    xlWorksheet1 = xlWorkbook1.Worksheets(2)
                    xlWorksheet1.Name = "舊調結表"
                    dict.Add("舊調結表", xlWorksheet1.Range(xlWorksheet1.Cells(1, 1), xlWorksheet1.Cells((xlWorksheet1.Range("A65536").End(xlUp).Row \ 34) * 34 + 34, 11)).Value)
                    xlWorkbook2.Close(SaveChanges := False)
                    System.IO.File.Delete(MyExcel2)
                Next
                
                '把專案代收查詢放入dict
                data.SelectCommand = _
                "SELECT ISNULL(檔名, '') AS 檔名 FROM 對帳 WHERE 選取 = 1 AND 備註 = '專案代收查詢' ORDER BY 起"
                data_dv = data.Select(New DataSourceSelectArguments)
                For i = 0 To data_dv.Count - 1
                    MyGUID = Guid.NewGuid().ToString("N")
                    MyExcel2 = MapPath(".\Excel\Temp\") & MyGUID & ".xls"
                    System.IO.File.Copy(MapPath(".\Excel\對帳\" & data_dv(i)("檔名")), MyExcel2)
                    xlWorkbook2 = xlApp.Workbooks.Open(MyExcel2, 0, False, 5, "", "", True, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", False, False, 0, True, 1, 0)
                    xlWorksheet1 = xlWorkbook1.Worksheets("專案代收查詢")
                    xlWorksheet2 = xlWorkbook2.ActiveSheet
                    Dim temp As Object
                    If dict.ContainsKey("專案代收查詢")
                        temp = xlWorksheet2.Range(xlWorksheet2.Cells(20, 1), xlWorksheet2.Cells(xlWorksheet2.Range("A19").End(xlDown).Row, 19)).Value
                        xlWorksheet1.Range(xlWorksheet1.Cells(xlWorksheet1.Range("A1").End(xlDown).Row + 1, 1), xlWorksheet1.Cells(xlWorksheet1.Range("A1").End(xlDown).Row + temp.GetLength(0), temp.GetLength(1))).Value = temp
                        dict("專案代收查詢") = xlWorksheet1.Range(xlWorksheet1.Cells(1, 1), xlWorksheet1.Cells(xlWorksheet1.Range("A1").End(xlDown).Row, 19)).Value
                    Else
                        temp = xlWorksheet2.Range(xlWorksheet2.Cells(19, 1), xlWorksheet2.Cells(xlWorksheet2.Range("A19").End(xlDown).Row, 19)).Value
                        xlWorksheet1.Range(xlWorksheet1.Cells(1, 1), xlWorksheet1.Cells(temp.GetLength(0), temp.GetLength(1))).Value = temp
                        dict.Add("專案代收查詢", xlWorksheet1.Range(xlWorksheet1.Cells(1, 1), xlWorksheet1.Cells(xlWorksheet1.Range("A1").End(xlDown).Row, 19)).Value)
                    End If
                    xlWorkbook2.Close(SaveChanges := False)
                    System.IO.File.Delete(MyExcel2)
                Next
                
                '把帳戶明細放入dict
                data.SelectCommand = _
                "SELECT ISNULL(檔名, '') AS 檔名 FROM 對帳 WHERE 選取 = 1 AND 備註 = '帳戶明細' ORDER BY 起"
                data_dv = data.Select(New DataSourceSelectArguments)
                For i = 0 To data_dv.Count - 1
                    MyGUID = Guid.NewGuid().ToString("N")
                    MyExcel2 = MapPath(".\Excel\Temp\") & MyGUID & ".xls"
                    System.IO.File.Copy(MapPath(".\Excel\對帳\" & data_dv(i)("檔名")), MyExcel2)
                    xlWorkbook2 = xlApp.Workbooks.Open(MyExcel2, 0, False, 5, "", "", True, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", False, False, 0, True, 1, 0)
                    xlWorksheet1 = xlWorkbook1.Worksheets("帳戶明細")
                    xlWorksheet2 = xlWorkbook2.ActiveSheet
                    Dim temp As Object
                    If dict.ContainsKey("帳戶明細")
                        temp = xlWorksheet2.Range(xlWorksheet2.Cells(19, 1), xlWorksheet2.Cells(xlWorksheet2.Range("A18").End(xlDown).Row, 13)).Value
                        xlWorksheet1.Range(xlWorksheet1.Cells(xlWorksheet1.Range("A1").End(xlDown).Row + 1, 1), xlWorksheet1.Cells(xlWorksheet1.Range("A1").End(xlDown).Row + temp.GetLength(0), temp.GetLength(1))).Value = temp
                        dict("帳戶明細") = xlWorksheet1.Range(xlWorksheet1.Cells(1, 1), xlWorksheet1.Cells(xlWorksheet1.Range("A1").End(xlDown).Row, 13)).Value
                    Else
                        temp = xlWorksheet2.Range(xlWorksheet2.Cells(18, 1), xlWorksheet2.Cells(xlWorksheet2.Range("A18").End(xlDown).Row, 13)).Value
                        xlWorksheet1.Range(xlWorksheet1.Cells(1, 1), xlWorksheet1.Cells(temp.GetLength(0), temp.GetLength(1))).Value = temp
                        dict.Add("帳戶明細", xlWorksheet1.Range(xlWorksheet1.Cells(1, 1), xlWorksheet1.Cells(xlWorksheet1.Range("A1").End(xlDown).Row, 13)).Value)
                    End If
                    xlWorkbook2.Close(SaveChanges := False)
                    System.IO.File.Delete(MyExcel2)
                Next
                
                '對dict的資料進行必要型別轉換等處理
                For i = 2 To dict("帳戶明細").GetLength(0)
                    dict("帳戶明細")(i, 4) = dict("帳戶明細")(i, 4).ToString()
                    dict("帳戶明細")(i, 5) = dict("帳戶明細")(i, 5).ToString()
                    dict("帳戶明細")(i, 10) = If(dict("帳戶明細")(i, 10) = Nothing, "", dict("帳戶明細")(i, 10).ToString())
                Next
                For i = 2 To dict("明細帳").GetLength(0)
                    dict("明細帳")(i, 1) = If(dict("明細帳")(i, 1) = Nothing, "       ", dict("明細帳")(i, 1).ToString())
                    dict("明細帳")(i, 2) = dict("明細帳")(i, 2).ToString()
                    dict("明細帳")(i, 5) = If(dict("明細帳")(i, 5) = Nothing, "", dict("明細帳")(i, 5).ToString())
                Next
                For i = 8 To dict("舊調結表").GetLength(0)
                    dict("舊調結表")(i, 9) = If(dict("舊調結表")(i, 9) = Nothing, "", dict("舊調結表")(i, 9).ToString())
                Next
                
                '處理帳戶明細代收總和沖正問題
                For i = dict("帳戶明細").GetLength(0) To 2 Step -1
                    If dict("帳戶明細")(i, 12) = Nothing
                        If dict("帳戶明細")(i, 3).Contains("代收總")
                            dict("帳戶明細")(i, 12) = "※不對帳(代收總)"
                        Else If dict("帳戶明細")(i, 3).Contains("沖正")
                            dict("帳戶明細")(i, 12) = "※不對帳(沖正)"
                            For j = i - 1 To 2 Step -1
                                If dict("帳戶明細")(i, 4) = dict("帳戶明細")(j, 5) And dict("帳戶明細")(i, 10) = dict("帳戶明細")(j, 10)
                                    dict("帳戶明細")(j, 12) = "※不對帳(已沖正)"
                                    Exit For
                                End If
                            Next
                        End If
                    End If
                Next
                
                '宣告變數
                Dim A As Object = Nothing
                Dim B As Object = Nothing
                Dim C As Object = Nothing
                Dim N As Object = Nothing
                Dim mode As Long = -1
                
                '比對前先處理資料
                dict.Add("明細帳(收入)", New LinkedList(Of Object))
                dict.Add("明細帳(支出)", New LinkedList(Of Object))
                For i = 2 To dict("明細帳").GetLength(0)
                    If dict("明細帳")(i, 10) = Nothing
                        If dict("明細帳")(i, 6) > 0
                            Dim 金額0 As Long = CLng(dict("明細帳")(i, 6))
                            Dim 列號1 As Long = i
                            Dim 原始日期 As Object = dict("明細帳")(i, 1)
                            Dim 日期2 As Date = Nothing
                            DateTime.TryParse(If(原始日期 = Nothing, Nothing, (CLng(原始日期.SubString(0, 3)) + 1911).ToString() & "/" & 原始日期.SubString(3, 2) & "/" & 原始日期.SubString(5, 2)), 日期2)
                            dict("明細帳(收入)").AddLast(New Object(2){金額0, 列號1, 日期2})
                        Else If dict("明細帳")(i, 7) > 0
                            Dim 金額0 As Long = CLng(dict("明細帳")(i, 7))
                            Dim 列號1 As Long = i
                            Dim 原始日期 As Object = dict("明細帳")(i, 1)
                            Dim 日期2 As Date = Nothing
                            DateTime.TryParse(If(原始日期 = Nothing, Nothing, (CLng(原始日期.SubString(0, 3)) + 1911).ToString() & "/" & 原始日期.SubString(3, 2) & "/" & 原始日期.SubString(5, 2)), 日期2)
                            dict("明細帳(支出)").AddLast(New Object(2){金額0, 列號1, 日期2})
                        End If
                    End If
                Next
                dict.Add("舊調結表(減：本基金已登帳而銀行未登帳之支出金額)", New LinkedList(Of Object))
                dict.Add("舊調結表(加：本基金已登帳而銀行未登帳之收入金額)", New LinkedList(Of Object))
                dict.Add("舊調結表(減：本基金未登帳而銀行已登帳之收入金額)", New LinkedList(Of Object))
                For i = 8 To dict("舊調結表").GetLength(0)
                    If dict("舊調結表")(i, 10) = Nothing
                        If dict("舊調結表")(i, 1) <> Nothing
                            Select Case dict("舊調結表")(i, 1)
                                Case "減：本基金已登帳而銀行未登帳之支出金額"
                                    mode = 1
                                Case "加：本基金已登帳而銀行未登帳之收入金額"
                                    mode = 2
                                Case "減：本基金未登帳而銀行已登帳之收入金額"
                                    mode = 3
                                Case Else
                                    mode = -1
                            End Select
                        End If
                        If dict("舊調結表")(i, 7) <> Nothing
                            Select Case mode
                                Case 1
                                    Dim 金額0 As Long = CLng(dict("舊調結表")(i, 7))
                                    Dim 列號1 As Long = i
                                    Dim 原始日期 As Object = If(dict("舊調結表")(i, 2) Is Nothing, Nothing, CLng(dict("舊調結表")(i, 2)).ToString("000") & CLng(dict("舊調結表")(i, 3)).ToString("00") & CLng(dict("舊調結表")(i, 4)).ToString("00"))
                                    Dim 日期2 As Date = Nothing
                                    DateTime.TryParse(If(原始日期 = Nothing, Nothing, (CLng(原始日期.SubString(0, 3)) + 1911).ToString() & "/" & 原始日期.SubString(3, 2) & "/" & 原始日期.SubString(5, 2)), 日期2)
                                    dict("舊調結表(減：本基金已登帳而銀行未登帳之支出金額)").AddLast(New Object(2){金額0, 列號1, 日期2})
                                Case 2
                                    Dim 金額0 As Long = CLng(dict("舊調結表")(i, 7))
                                    Dim 列號1 As Long = i
                                    Dim 原始日期 As Object = If(dict("舊調結表")(i, 2) Is Nothing, Nothing, CLng(dict("舊調結表")(i, 2)).ToString("000") & CLng(dict("舊調結表")(i, 3)).ToString("00") & CLng(dict("舊調結表")(i, 4)).ToString("00"))
                                    Dim 日期2 As Date = Nothing
                                    DateTime.TryParse(If(原始日期 = Nothing, Nothing, (CLng(原始日期.SubString(0, 3)) + 1911).ToString() & "/" & 原始日期.SubString(3, 2) & "/" & 原始日期.SubString(5, 2)), 日期2)
                                    dict("舊調結表(加：本基金已登帳而銀行未登帳之收入金額)").AddLast(New Object(2){金額0, 列號1, 日期2})
                                Case 3
                                    Dim 金額0 As Long = CLng(dict("舊調結表")(i, 7))
                                    Dim 列號1 As Long = i
                                    Dim 原始日期 As Object = If(dict("舊調結表")(i, 2) Is Nothing, Nothing, CLng(dict("舊調結表")(i, 2)).ToString("000") & CLng(dict("舊調結表")(i, 3)).ToString("00") & CLng(dict("舊調結表")(i, 4)).ToString("00"))
                                    Dim 日期2 As Date = Nothing
                                    DateTime.TryParse(If(原始日期 = Nothing, Nothing, (CLng(原始日期.SubString(0, 3)) + 1911).ToString() & "/" & 原始日期.SubString(3, 2) & "/" & 原始日期.SubString(5, 2)), 日期2)
                                    dict("舊調結表(減：本基金未登帳而銀行已登帳之收入金額)").AddLast(New Object(2){金額0, 列號1, 日期2})
                            End Select
                        End If
                    End If
                    i = If((i Mod 34) = 28, i + 12, i)
                Next
                dict.Add("專案代收查詢(收入)", New LinkedList(Of Object))
                For i = 2 To dict("專案代收查詢").GetLength(0)
                    If dict("專案代收查詢")(i, 18) = Nothing
                        Dim 金額0 As Long = CLng(dict("專案代收查詢")(i, 4))
                        Dim 列號1 As Long = i
                        Dim 原始日期 As Object = dict("專案代收查詢")(i, 1)
                        Dim 日期2 As Date = Nothing
                        DateTime.TryParse(If(原始日期 = Nothing, Nothing, 原始日期), 日期2)
                        dict("專案代收查詢(收入)").AddLast(New Object(2){金額0, 列號1, 日期2})
                    End If
                Next
                dict.Add("帳戶明細(收入)", New LinkedList(Of Object))
                dict.Add("帳戶明細(支出)", New LinkedList(Of Object))
                For i = 2 To dict("帳戶明細").GetLength(0)
                    If dict("帳戶明細")(i, 12) = Nothing
                        If dict("帳戶明細")(i, 5) <> "-"
                            Dim 金額0 As Long = CLng(dict("帳戶明細")(i, 5))
                            Dim 列號1 As Long = i
                            Dim 原始日期 As Object = dict("帳戶明細")(i, 1)
                            Dim 日期2 As Date = Nothing
                            DateTime.TryParse(If(原始日期 = Nothing, Nothing, 原始日期), 日期2)
                            dict("帳戶明細(收入)").AddLast(New Object(2){金額0, 列號1, 日期2})
                        Else If dict("帳戶明細")(i, 4) <> "-"
                            Dim 金額0 As Long = CLng(dict("帳戶明細")(i, 4))
                            Dim 列號1 As Long = i
                            Dim 原始日期 As Object = dict("帳戶明細")(i, 1)
                            Dim 日期2 As Date = Nothing
                            DateTime.TryParse(If(原始日期 = Nothing, Nothing, 原始日期), 日期2)
                            dict("帳戶明細(支出)").AddLast(New Object(2){金額0, 列號1, 日期2})
                        End If
                    End If
                Next
                
                '紀錄金額出現次數，判斷有沒有重複
                dict.Add("金額數(收入)", New Dictionary(Of Long, Long))
                dict.Add("金額數(支出)", New Dictionary(Of Long, Long))
                A = dict("明細帳(收入)").First
                While A IsNot Nothing
                    If dict("金額數(收入)").ContainsKey(A.Value(0))
                        dict("金額數(收入)")(A.Value(0)) = dict("金額數(收入)")(A.Value(0)) + 1
                    Else
                        dict("金額數(收入)").Add(A.Value(0), 1)
                    End If
                    A = A.Next
                End While
                A = dict("專案代收查詢(收入)").First
                While A IsNot Nothing
                    If dict("金額數(收入)").ContainsKey(A.Value(0))
                        dict("金額數(收入)")(A.Value(0)) = dict("金額數(收入)")(A.Value(0)) - 1
                    Else
                        dict("金額數(收入)").Add(A.Value(0), -1)
                    End If
                    A = A.Next
                End While
                A = dict("帳戶明細(收入)").First
                While A IsNot Nothing
                    If dict("金額數(收入)").ContainsKey(A.Value(0))
                        dict("金額數(收入)")(A.Value(0)) = dict("金額數(收入)")(A.Value(0)) - 1
                    Else
                        dict("金額數(收入)").Add(A.Value(0), -1)
                    End If
                    A = A.Next
                End While
                A = dict("明細帳(支出)").First
                While A IsNot Nothing
                    If dict("金額數(支出)").ContainsKey(A.Value(0))
                        dict("金額數(支出)")(A.Value(0)) = dict("金額數(支出)")(A.Value(0)) + 1
                    Else
                        dict("金額數(支出)").Add(A.Value(0), 1)
                    End If
                    A = A.Next
                End While
                A = dict("帳戶明細(支出)").First
                While A IsNot Nothing
                    If dict("金額數(支出)").ContainsKey(A.Value(0))
                        dict("金額數(支出)")(A.Value(0)) = dict("金額數(支出)")(A.Value(0)) - 1
                    Else
                        dict("金額數(支出)").Add(A.Value(0), -1)
                    End If
                    A = A.Next
                End While
                A = dict("舊調結表(減：本基金已登帳而銀行未登帳之支出金額)").First
                While A IsNot Nothing
                    If dict("金額數(支出)").ContainsKey(A.Value(0))
                        dict("金額數(支出)")(A.Value(0)) = dict("金額數(支出)")(A.Value(0)) + 1
                    Else
                        dict("金額數(支出)").Add(A.Value(0), 1)
                    End If
                    A = A.Next
                End While
                A = dict("舊調結表(加：本基金已登帳而銀行未登帳之收入金額)").First
                While A IsNot Nothing
                    If dict("金額數(收入)").ContainsKey(A.Value(0))
                        dict("金額數(收入)")(A.Value(0)) = dict("金額數(收入)")(A.Value(0)) + 1
                    Else
                        dict("金額數(收入)").Add(A.Value(0), 1)
                    End If
                    A = A.Next
                End While
                A = dict("舊調結表(減：本基金未登帳而銀行已登帳之收入金額)").First
                While A IsNot Nothing
                    If dict("金額數(收入)").ContainsKey(A.Value(0))
                        dict("金額數(收入)")(A.Value(0)) = dict("金額數(收入)")(A.Value(0)) - 1
                    Else
                        dict("金額數(收入)").Add(A.Value(0), -1)
                    End If
                    A = A.Next
                End While
                
                '開始一對一的對帳
                Dim 結帳日期a As DateTime = Convert.ToDateTime(taiwancalendarto(Me.結帳日期a.SelectedValue))
                Dim 閾值 As New DateTime(結帳日期a.Year, 結帳日期a.Month, 21)
                A = dict("舊調結表(減：本基金已登帳而銀行未登帳之支出金額)").First
                While A IsNot Nothing
                    N = A.Next
                    If dict("金額數(支出)")(A.Value(0)) = 0 Or A.Value(2) < 閾值
                        B = dict("帳戶明細(支出)").First
                        While B IsNot Nothing
                            If A.Value(0) = B.Value(0)
                                dict("舊調結表")(A.Value(1), 10) = "帳戶明細"
                                dict("舊調結表")(A.Value(1), 11) = B.Value(1)
                                dict("帳戶明細")(B.Value(1), 12) = "舊調結表"
                                dict("帳戶明細")(B.Value(1), 13) = A.Value(1)
                                dict("舊調結表(減：本基金已登帳而銀行未登帳之支出金額)").Remove(A)
                                dict("帳戶明細(支出)").Remove(B)
                                Exit While
                            End If
                            B = B.Next
                        End While
                    End If
                    A = N
                End While
                A = dict("舊調結表(加：本基金已登帳而銀行未登帳之收入金額)").First
                While A IsNot Nothing
                    N = A.Next
                    If dict("金額數(收入)")(A.Value(0)) = 0 Or A.Value(2) < 閾值
                        B = dict("帳戶明細(收入)").First
                        While B IsNot Nothing
                            If A.Value(0) = B.Value(0)
                                dict("舊調結表")(A.Value(1), 10) = "帳戶明細"
                                dict("舊調結表")(A.Value(1), 11) = B.Value(1)
                                dict("帳戶明細")(B.Value(1), 12) = "舊調結表"
                                dict("帳戶明細")(B.Value(1), 13) = A.Value(1)
                                dict("舊調結表(加：本基金已登帳而銀行未登帳之收入金額)").Remove(A)
                                dict("帳戶明細(收入)").Remove(B)
                                Exit While
                            End If
                            B = B.Next
                        End While
                    End If
                    A = N
                End While
                A = dict("舊調結表(減：本基金未登帳而銀行已登帳之收入金額)").First
                While A IsNot Nothing
                    N = A.Next
                    If dict("金額數(收入)")(A.Value(0)) = 0 Or A.Value(2) < 閾值
                        B = dict("明細帳(收入)").First
                        While B IsNot Nothing
                            If A.Value(0) = B.Value(0)
                                dict("舊調結表")(A.Value(1), 10) = "明細帳"
                                dict("舊調結表")(A.Value(1), 11) = B.Value(1)
                                dict("明細帳")(B.Value(1), 10) = "舊調結表"
                                dict("明細帳")(B.Value(1), 11) = A.Value(1)
                                dict("舊調結表(減：本基金未登帳而銀行已登帳之收入金額)").Remove(A)
                                dict("明細帳(收入)").Remove(B)
                                Exit While
                            End If
                            B = B.Next
                        End While
                    End If
                    A = N
                End While
                A = dict("明細帳(收入)").First
                While A IsNot Nothing
                    N = A.Next
                    If dict("金額數(收入)")(A.Value(0)) = 0 Or A.Value(2) < 閾值
                        B = dict("專案代收查詢(收入)").First
                        While B IsNot Nothing
                            If A.Value(2) < B.Value(2)
                                Exit While
                            End If
                            If A.Value(0) = B.Value(0)
                                dict("明細帳")(A.Value(1), 10) = "專案代收查詢"
                                dict("明細帳")(A.Value(1), 11) = B.Value(1)
                                dict("專案代收查詢")(B.Value(1), 18) = "明細帳"
                                dict("專案代收查詢")(B.Value(1), 19) = A.Value(1)
                                dict("明細帳(收入)").Remove(A)
                                dict("專案代收查詢(收入)").Remove(B)
                                Exit While
                            End If
                            B = B.Next
                        End While
                    End If
                    A = N
                End While
                A = dict("明細帳(收入)").First
                While A IsNot Nothing
                    N = A.Next
                    If dict("金額數(收入)")(A.Value(0)) = 0 Or A.Value(2) < 閾值
                        B = dict("帳戶明細(收入)").First
                        While B IsNot Nothing
                            If A.Value(0) = B.Value(0)
                                dict("明細帳")(A.Value(1), 10) = "帳戶明細"
                                dict("明細帳")(A.Value(1), 11) = B.Value(1)
                                dict("帳戶明細")(B.Value(1), 12) = "明細帳"
                                dict("帳戶明細")(B.Value(1), 13) = A.Value(1)
                                dict("明細帳(收入)").Remove(A)
                                dict("帳戶明細(收入)").Remove(B)
                                Exit While
                            End If
                            B = B.Next
                        End While
                    End If
                    A = N
                End While
                A = dict("明細帳(支出)").First
                While A IsNot Nothing
                    N = A.Next
                    If dict("金額數(支出)")(A.Value(0)) = 0 Or A.Value(2) < 閾值
                        B = dict("帳戶明細(支出)").First
                        While B IsNot Nothing
                            If A.Value(0) = B.Value(0)
                                dict("明細帳")(A.Value(1), 10) = "帳戶明細"
                                dict("明細帳")(A.Value(1), 11) = B.Value(1)
                                dict("帳戶明細")(B.Value(1), 12) = "明細帳"
                                dict("帳戶明細")(B.Value(1), 13) = A.Value(1)
                                dict("明細帳(支出)").Remove(A)
                                dict("帳戶明細(支出)").Remove(B)
                                Exit While
                            End If
                            B = B.Next
                        End While
                    End If
                    A = N
                End While
                
                '開始產生調結表
                Dim 頁數 As Long = _
                    dict("舊調結表(減：本基金已登帳而銀行未登帳之支出金額)").Count + _
                    dict("舊調結表(加：本基金已登帳而銀行未登帳之收入金額)").Count + _
                    dict("舊調結表(減：本基金未登帳而銀行已登帳之收入金額)").Count + _
                    dict("明細帳(收入)").Count + _
                    dict("明細帳(支出)").Count + _
                    dict("專案代收查詢(收入)").Count + _
                    dict("帳戶明細(收入)").Count + _
                    dict("帳戶明細(支出)").Count
                頁數 = (頁數 + 23) \ 22
                xlWorksheet1 = xlWorkbook1.Worksheets("調結表")
                xlWorksheet1.PageSetup.PrintArea = xlWorksheet1.Range(xlWorksheet1.Cells(1, 1), xlWorksheet1.Cells(34 * 頁數, 9)).Address
                For i = 2 To 頁數
                    xlWorksheet1.Range(xlWorksheet1.Cells(34 * i - 33, 1), xlWorksheet1.Cells(34 * i, 9)).Value(11) = xlWorksheet1.Range(xlWorksheet1.Cells(1, 1), xlWorksheet1.Cells(34, 9)).Value(11)
                    xlWorksheet1.Range(xlWorksheet1.Cells(34 * i - 27, 1), xlWorksheet1.Cells(34 * i - 6, 9)).RowHeight = 30.25
                    xlWorksheet1.Rows(34 * i - 33).PageBreak = xlPageBreakManual
                Next
                
                '把調結表放入dict
                dict.Add("調結表", xlWorksheet1.Range(xlWorksheet1.Cells(1, 1), xlWorksheet1.Cells(34 * 頁數, 11)).Value)
                Dim 日期 As Date = Convert.ToDateTime(taiwancalendarto(Me.結帳日期a.SelectedValue)).AddDays(12)
                日期 = 日期.AddDays(DateTime.DaysInMonth(日期.Year, 日期.Month) - 日期.Day)
                For i = 1 To 頁數
                    dict("調結表")(34 * i - 30, 2) = "".PadLeft(99) & "中華民國  " & (日期.Year - 1911).ToString() & "  年  " & 日期.Month.ToString() & "  月  " & 日期.Day.ToString() & "  日"
                    dict("調結表")(34 * i - 31, 9) = "全  " & 頁數.ToString() & "  頁第  " & i.ToString() & "  頁"
                    dict("調結表")(34 * i - 32, 1) = "  銀行名稱:中國信託銀行台中分行"
                    dict("調結表")(34 * i - 30, 1) = "  帳戶號碼:026350002965"
                Next
                dict("調結表")(7, 1) = "銀行對帳單結存金額"
                dict("調結表")(34 * 頁數 - 6, 1) = "本基金銀行存款帳面餘額"
                dict("調結表")(34 * 頁數 - 5, 1) = "製表"
                dict("調結表")(34 * 頁數 - 5, 2) = "秘書室"
                dict("調結表")(34 * 頁數 - 5, 6) = "   主辦會計人員"
                dict("調結表")(34 * 頁數 - 5, 9) = " 機關長官"
                Dim 列數 As Long = 8
                A = dict("舊調結表(減：本基金已登帳而銀行未登帳之支出金額)").First
                While A IsNot Nothing
                    dict("調結表")(列數, 1) = "減：本基金已登帳而銀行未登帳之支出金額"
                    For Each i In {2, 3, 4, 5, 6, 9}
                        dict("調結表")(列數, i) = dict("舊調結表")(A.Value(1), i)
                    Next
                    dict("調結表")(列數, 7) = A.Value(0)
                    dict("調結表")(列數, 10) = "舊調結表"
                    dict("調結表")(列數, 11) = A.Value(1)
                    dict("舊調結表")(A.Value(1), 10) = "⚠調結表"
                    dict("舊調結表")(A.Value(1), 11) = 列數
                    A = A.Next
                    列數 = If((列數 Mod 34) = 28, 列數 + 13, 列數 + 1)
                End While
                A = dict("明細帳(支出)").First
                While A IsNot Nothing
                    dict("調結表")(列數, 1) = "減：本基金已登帳而銀行未登帳之支出金額"
                    dict("調結表")(列數, 2) = dict("明細帳")(A.Value(1), 1).SubString(0, 3)
                    dict("調結表")(列數, 3) = dict("明細帳")(A.Value(1), 1).SubString(3, 2)
                    dict("調結表")(列數, 4) = dict("明細帳")(A.Value(1), 1).SubString(5, 2)
                    dict("調結表")(列數, 5) = "#" & dict("明細帳")(A.Value(1), 2)
                    dict("調結表")(列數, 6) = dict("明細帳")(A.Value(1), 8)
                    dict("調結表")(列數, 7) = A.Value(0)
                    dict("調結表")(列數, 9) = dict("明細帳")(A.Value(1), 4)
                    dict("調結表")(列數, 10) = "明細帳"
                    dict("調結表")(列數, 11) = A.Value(1)
                    dict("明細帳")(A.Value(1), 10) = "⚠調結表"
                    dict("明細帳")(A.Value(1), 11) = 列數
                    A = A.Next
                    列數 = If((列數 Mod 34) = 28, 列數 + 13, 列數 + 1)
                End While
                A = dict("舊調結表(加：本基金已登帳而銀行未登帳之收入金額)").First
                While A IsNot Nothing
                    dict("調結表")(列數, 1) = "加：本基金已登帳而銀行未登帳之收入金額"
                    For Each i In {2, 3, 4, 5, 6, 9}
                        dict("調結表")(列數, i) = dict("舊調結表")(A.Value(1), i)
                    Next
                    dict("調結表")(列數, 7) = A.Value(0)
                    dict("調結表")(列數, 10) = "舊調結表"
                    dict("調結表")(列數, 11) = A.Value(1)
                    dict("舊調結表")(A.Value(1), 10) = "⚠調結表"
                    dict("舊調結表")(A.Value(1), 11) = 列數
                    A = A.Next
                    列數 = If((列數 Mod 34) = 28, 列數 + 13, 列數 + 1)
                End While
                A = dict("明細帳(收入)").First
                While A IsNot Nothing
                    dict("調結表")(列數, 1) = "加：本基金已登帳而銀行未登帳之收入金額"
                    dict("調結表")(列數, 2) = dict("明細帳")(A.Value(1), 1).SubString(0, 3)
                    dict("調結表")(列數, 3) = dict("明細帳")(A.Value(1), 1).SubString(3, 2)
                    dict("調結表")(列數, 4) = dict("明細帳")(A.Value(1), 1).SubString(5, 2)
                    dict("調結表")(列數, 5) = "#" & dict("明細帳")(A.Value(1), 2)
                    dict("調結表")(列數, 6) = dict("明細帳")(A.Value(1), 8)
                    dict("調結表")(列數, 7) = A.Value(0)
                    dict("調結表")(列數, 9) = dict("明細帳")(A.Value(1), 4)
                    dict("調結表")(列數, 10) = "明細帳"
                    dict("調結表")(列數, 11) = A.Value(1)
                    dict("明細帳")(A.Value(1), 10) = "⚠調結表"
                    dict("明細帳")(A.Value(1), 11) = 列數
                    A = A.Next
                    列數 = If((列數 Mod 34) = 28, 列數 + 13, 列數 + 1)
                End While
                A = dict("舊調結表(減：本基金未登帳而銀行已登帳之收入金額)").First
                While A IsNot Nothing
                    dict("調結表")(列數, 1) = "減：本基金未登帳而銀行已登帳之收入金額"
                    For Each i In {2, 3, 4, 5, 6, 9}
                        dict("調結表")(列數, i) = dict("舊調結表")(A.Value(1), i)
                    Next
                    dict("調結表")(列數, 7) = A.Value(0)
                    dict("調結表")(列數, 10) = "舊調結表"
                    dict("調結表")(列數, 11) = A.Value(1)
                    dict("舊調結表")(A.Value(1), 10) = "⚠調結表"
                    dict("舊調結表")(A.Value(1), 11) = 列數
                    A = A.Next
                    列數 = If((列數 Mod 34) = 28, 列數 + 13, 列數 + 1)
                End While
                A = dict("專案代收查詢(收入)").First
                B = dict("帳戶明細(收入)").First
                While A IsNot Nothing Or B IsNot Nothing
                    If If(A Is Nothing Or B Is Nothing, B Is Nothing, dict("專案代收查詢")(A.Value(1), 1) < dict("帳戶明細")(B.Value(1), 1))
                        dict("調結表")(列數, 1) = "減：本基金未登帳而銀行已登帳之收入金額"
                        dict("調結表")(列數, 2) = dict("專案代收查詢")(A.Value(1), 1).Year - 1911
                        dict("調結表")(列數, 3) = dict("專案代收查詢")(A.Value(1), 1).Month
                        dict("調結表")(列數, 4) = dict("專案代收查詢")(A.Value(1), 1).Day
                        dict("調結表")(列數, 7) = A.Value(0)
                        dict("調結表")(列數, 10) = "專案代收查詢"
                        dict("調結表")(列數, 11) = A.Value(1)
                        dict("專案代收查詢")(A.Value(1), 18) = "⚠調結表"
                        dict("專案代收查詢")(A.Value(1), 19) = 列數
                        A = A.Next
                    Else
                        dict("調結表")(列數, 1) = "減：本基金未登帳而銀行已登帳之收入金額"
                        dict("調結表")(列數, 2) = dict("帳戶明細")(B.Value(1), 1).Year - 1911
                        dict("調結表")(列數, 3) = dict("帳戶明細")(B.Value(1), 1).Month
                        dict("調結表")(列數, 4) = dict("帳戶明細")(B.Value(1), 1).Day
                        dict("調結表")(列數, 6) = dict("帳戶明細")(B.Value(1), 10)
                        dict("調結表")(列數, 7) = B.Value(0)
                        dict("調結表")(列數, 10) = "帳戶明細"
                        dict("調結表")(列數, 11) = B.Value(1)
                        dict("帳戶明細")(B.Value(1), 12) = "⚠調結表"
                        dict("帳戶明細")(B.Value(1), 13) = 列數
                        B = B.Next
                    End If
                    列數 = If((列數 Mod 34) = 28, 列數 + 13, 列數 + 1)
                End While
                A = dict("帳戶明細(支出)").First
                While A IsNot Nothing
                    dict("調結表")(列數, 1) = "加：本基金未登帳而銀行已登帳之支出金額"
                    dict("調結表")(列數, 2) = dict("帳戶明細")(A.Value(1), 1).Year - 1911
                    dict("調結表")(列數, 3) = dict("帳戶明細")(A.Value(1), 1).Month
                    dict("調結表")(列數, 4) = dict("帳戶明細")(A.Value(1), 1).Day
                    dict("調結表")(列數, 6) = dict("帳戶明細")(A.Value(1), 10)
                    dict("調結表")(列數, 7) = A.Value(0)
                    dict("調結表")(列數, 10) = "帳戶明細"
                    dict("調結表")(列數, 11) = A.Value(1)
                    dict("帳戶明細")(A.Value(1), 12) = "⚠調結表"
                    dict("帳戶明細")(A.Value(1), 13) = 列數
                    A = A.Next
                    列數 = If((列數 Mod 34) = 28, 列數 + 13, 列數 + 1)
                End While
                
                '調結表最後處理
                Dim 項目 As String = ""
                For i = 8 To 34 * 頁數 - 7
                    If dict("調結表")(i, 1) <> Nothing
                        If dict("調結表")(i, 1) = 項目
                            dict("調結表")(i, 1) = ""
                        Else
                            項目 = dict("調結表")(i, 1)
                        End If
                    End If
                    i = If((i Mod 34) = 28, i + 12, i)
                Next
                Dim 小合計位置 As Long = 0
                Dim 小合計 As String = ""
                For i = 8 To 34 * 頁數 - 7
                    If dict("調結表")(i, 7) <> Nothing
                        If dict("調結表")(i, 1) <> Nothing
                            小合計 = "=SUM(G" & i.ToString()
                        End If
                    End If
                    Dim j As Long = If((i Mod 34) = 28, i + 13, i + 1)
                    If dict("調結表")(j, 1) <> Nothing
                        小合計 = 小合計 & ":G" & i.ToString() & ")"
                        dict("調結表")(i, 8) = 小合計
                    Else If dict("調結表")(j, 7) = Nothing
                        小合計 = 小合計 & ":G" & i.ToString() & ")"
                        dict("調結表")(i, 8) = 小合計
                        Exit For
                    End If
                    小合計 = If((i Mod 34) = 28, 小合計 & ":G" & i.ToString() & ",G" & (i + 13).ToString(), 小合計)
                    i = If((i Mod 34) = 28, i + 12, i)
                Next
                Dim 大合計 As String = "=H7"
                For i = 8 To 34 * 頁數 - 7
                    If dict("調結表")(i, 1) <> Nothing
                        項目 = dict("調結表")(i, 1)
                    End If
                    If dict("調結表")(i, 8) <> Nothing
                        If 項目.SubString(0, 1) = "加"
                            大合計 = 大合計 & "+" & "H" & i.ToString()
                        Else If 項目.SubString(0, 1) = "減"
                            大合計 = 大合計 & "-" & "H" & i.ToString()
                        End If
                    End If
                    i = If((i Mod 34) = 28, i + 12, i)
                Next
                dict("調結表")(34 * 頁數 - 6, 8) = 大合計
                ''對應傳票資料
                'For i = 8 To 34 * 頁數 - 7
                '    If dict("調結表")(i, 7) <> Nothing
                '        If dict("調結表")(i, 1) <> Nothing
                '            Select Case dict("調結表")(i, 1)
                '                Case "減：本基金已登帳而銀行未登帳之支出金額"
                '                    mode = 1
                '                Case "加：本基金已登帳而銀行未登帳之收入金額"
                '                    mode = 2
                '                Case "減：本基金未登帳而銀行已登帳之收入金額"
                '                    mode = 3
                '                Case Else
                '                    mode = -1
                '            End Select
                '        End If
                '        If dict("調結表")(i, 7) <> Nothing
                '            Select Case mode
                '                Case 1
                '                    
                '                Case 2
                '                    
                '                Case 3
                '                    If dict("調結表")(i, 5) = Nothing
                '                        data.SelectCommand = _
                '                        "SELECT b.傳票號碼, b.登錄序號, b.摘要說明 " & _
                '                        "FROM (SELECT '' AS unused) a " & _
                '                        "LEFT JOIN " & _
                '                        "(SELECT TOP 1 傳票資料.傳票號碼, 傳票資料.登錄序號, 傳票資料.摘要說明 FROM 傳票資料 INNER JOIN 現金備查簿 " & _
                '                        "ON 傳票資料.年 = 現金備查簿.年 AND 傳票資料.傳票號碼 = 現金備查簿.傳票號碼 " & _
                '                        "WHERE (現金備查簿.結帳日期 > '" & taiwancalendarto(Me.結帳日期b.SelectedValue) & "' OR 現金備查簿.結帳日期 IS NULL) " & _
                '                        "AND (現金備查簿.收入金額409 > 0 OR 現金備查簿.支出金額409 > 0) " & _
                '                        "AND 傳票資料.收入金額 = '" & dict("調結表")(i, 7) & "' " & _
                '                        "ORDER BY 現金備查簿.結帳日期, 傳票資料.傳票號碼, 傳票資料.之) b " & _
                '                        "ON 1=1"
                '                        data_dv = data.Select(New DataSourceSelectArguments)
                '                        dict("調結表")(i, 5) = data_dv(0)(0).ToString()
                '                        dict("調結表")(i, 6) = data_dv(0)(1).ToString()
                '                        dict("調結表")(i, 9) = data_dv(0)(2).ToString()
                '                    End If
                '            End Select
                '        End If
                '    End If
                '    i = If((i Mod 34) = 28, i + 12, i)
                'Next
                'For i = 8 To 34 * 頁數 - 7
                '    If dict("調結表")(i, 7) <> Nothing
                '        If dict("調結表")(i, 1) <> Nothing
                '            Select Case dict("調結表")(i, 1)
                '                Case "減：本基金已登帳而銀行未登帳之支出金額"
                '                    mode = 1
                '                Case "加：本基金已登帳而銀行未登帳之收入金額"
                '                    mode = 2
                '                Case "減：本基金未登帳而銀行已登帳之收入金額"
                '                    mode = 3
                '                Case Else
                '                    mode = -1
                '            End Select
                '        End If
                '        If dict("調結表")(i, 7) <> Nothing
                '            Select Case mode
                '                Case 1
                '                    
                '                Case 2
                '                    
                '                Case 3
                '                    If dict("調結表")(i, 5) = Nothing
                '                        Dim sum As Long = CLng(dict("調結表")(i, 7))
                '                        Dim 備註 As String = dict("調結表")(i, 7).ToString()
                '                        For j = If((i + 1 Mod 34) = 28, i + 14, i + 1) To 34 * 頁數 - 7
                '                            If dict("調結表")(j, 7) <> Nothing
                '                                If dict("調結表")(j, 5) = Nothing
                '                                    sum = sum + CLng(dict("調結表")(j, 7))
                '                                    備註 = 備註 & "+" & dict("調結表")(j, 7).ToString()
                '                                    data.SelectCommand = _
                '                                    "SELECT b.傳票號碼, b.登錄序號, b.摘要說明 " & _
                '                                    "FROM " & _
                '                                    "(SELECT TOP 1 傳票資料.傳票號碼, 傳票資料.登錄序號, 傳票資料.摘要說明 FROM 傳票資料 INNER JOIN 現金備查簿 " & _
                '                                    "ON 傳票資料.年 = 現金備查簿.年 AND 傳票資料.傳票號碼 = 現金備查簿.傳票號碼 " & _
                '                                    "WHERE (現金備查簿.結帳日期 > '" & taiwancalendarto(Me.結帳日期b.SelectedValue) & "' OR 現金備查簿.結帳日期 IS NULL) " & _
                '                                    "AND (現金備查簿.收入金額409 > 0 OR 現金備查簿.支出金額409 > 0) " & _
                '                                    "AND 傳票資料.收入金額 = '" & sum.ToString() & "' " & _
                '                                    "ORDER BY 現金備查簿.結帳日期, 傳票資料.傳票號碼, 傳票資料.之) b"
                '                                    data_dv = data.Select(New DataSourceSelectArguments)
                '                                    If data_dv.Count > 0
                '                                        For k = i To j
                '                                            dict("調結表")(k, 5) = data_dv(0)(0).ToString()
                '                                            dict("調結表")(k, 6) = data_dv(0)(1).ToString()
                '                                            dict("調結表")(k, 9) = 備註 & "=" & sum.ToString()
                '                                        Next
                '                                        Exit For
                '                                    End If
                '                                End If
                '                            End If
                '                            j = If((j Mod 34) = 28, j + 12, j)
                '                        Next
                '                    End If
                '            End Select
                '        End If
                '    End If
                '    i = If((i Mod 34) = 28, i + 12, i)
                'Next
                
                'Finally
                xlWorksheet1 = xlWorkbook1.Worksheets("明細帳")
                xlWorksheet1.Range(xlWorksheet1.Cells(1, 1), xlWorksheet1.Cells(dict("明細帳").GetLength(0), dict("明細帳").GetLength(1))).Value = dict("明細帳")
                xlWorksheet1 = xlWorkbook1.Worksheets("舊調結表")
                xlWorksheet1.Range(xlWorksheet1.Cells(1, 1), xlWorksheet1.Cells(dict("舊調結表").GetLength(0), dict("舊調結表").GetLength(1))).Value = dict("舊調結表")
                xlWorksheet1 = xlWorkbook1.Worksheets("專案代收查詢")
                xlWorksheet1.Range(xlWorksheet1.Cells(1, 1), xlWorksheet1.Cells(dict("專案代收查詢").GetLength(0), dict("專案代收查詢").GetLength(1))).Value = dict("專案代收查詢")
                xlWorksheet1 = xlWorkbook1.Worksheets("帳戶明細")
                xlWorksheet1.Range(xlWorksheet1.Cells(1, 1), xlWorksheet1.Cells(dict("帳戶明細").GetLength(0), dict("帳戶明細").GetLength(1))).Value = dict("帳戶明細")
                xlWorksheet1 = xlWorkbook1.Worksheets("調結表")
                xlWorksheet1.Range(xlWorksheet1.Cells(1, 1), xlWorksheet1.Cells(dict("調結表").GetLength(0), dict("調結表").GetLength(1))).Value = dict("調結表")
                xlWorksheet1.Activate()
                xlWorkbook1.SaveAs(MyExcel1.Replace("xlsx", "xls"), FileFormat:=xlWorkbookNormal)
                xlWorkbook1.Close(SaveChanges := True)
                xlApp.Quit()
                ReleaseObject(xlWorksheet1)
                ReleaseObject(xlWorksheet2)
                ReleaseObject(xlWorkbook1)
                ReleaseObject(xlWorkbook2)
                ReleaseObject(xlApp)
                Response.Clear()
                Response.ClearHeaders()
                Response.Buffer = True
                Response.ContentType = "application/octet-stream"
                Dim downloadfilename = "中國信託調結表.xls"
                Response.AddHeader("Content-Disposition", "attachment;FileName=" + Uri.EscapeDataString(downloadfilename))
                Response.WriteFile(MyExcel1.Replace("xlsx", "xls"))
                System.Web.HttpContext.Current.ApplicationInstance.CompleteRequest()
                Response.Flush()
                System.IO.File.Delete(MyExcel1)
                System.IO.File.Delete(MyExcel1.Replace("xlsx", "xls"))
                Response.End()
        End Select
    End Sub
    Protected Sub GridView1_RowCommand(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles GridView1.RowCommand
        Select Case e.CommandName
            Case "CustomDelete"
                Dim i As Long = e.CommandSource.NamingContainer.RowIndex
                Dim MyExcel As String = MapPath(".\Excel\對帳\") & CType(Me.GridView1.Rows(i).FindControl("檔名"), HyperLink).Text
                System.IO.File.Delete(MyExcel)
                Me.GridView1.DeleteRow(i)
        End Select
    End Sub
End Class