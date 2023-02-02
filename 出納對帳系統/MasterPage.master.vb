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
Partial Class MasterPage
    Inherits System.Web.UI.MasterPage
    Dim con_14 As String = System.Web.Configuration.WebConfigurationManager.ConnectionStrings("ApplicationServices").ConnectionString
    Dim data As New SqlDataSource
    Dim data_dv As Data.DataView
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        data.ConnectionString = con_14
        If Not Page.IsPostBack Then
        End If
    End Sub
    Protected Sub Import(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not FileUpload1.HasFile
            Exit Sub
        End If
        
        Dim xlApp As New Excel.ApplicationClass()
        xlApp.DisplayAlerts = False
        xlApp.ScreenUpdating = false
        xlApp.EnableEvents = false
        Dim xlWorkbook As Excel.Workbook = Nothing
        Dim xlWorksheet As Excel.Worksheet = Nothing
        Dim Sheet1 As Object
        Dim Sheet2 As Object
        
        'ATOC
        For Each PostedFile As HttpPostedFile In FileUpload1.PostedFiles
            Dim MyGUID As String = Guid.NewGuid().ToString("N")
            Dim MyExcel As String = MapPath(".\Excel\Temp\") & MyGUID & ".xls"
            PostedFile.SaveAs(MyExcel)
            
            xlWorkbook = xlApp.Workbooks.Open(MyExcel, 0, False, 5, "", "", True, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", False, False, 0, True, 1, 0)
            xlWorksheet = xlWorkBook.Worksheets(1)'第一工作業
            
            If xlWorksheet.Cells(1, 1).Value = "spid"
                Sheet1 = xlWorksheet.Range(xlWorksheet.Cells(2, 1), xlWorksheet.Cells(xlWorksheet.Range("D65536").End(xlUp).Row, 10)).Value
                xlWorksheet =  xlWorkBook.Worksheets(2)'第二工作業
                Sheet2 = xlWorksheet.Range(xlWorksheet.Cells(2, 1), xlWorksheet.Cells(xlWorksheet.Range("D65536").End(xlUp).Row + 1, 30)).Value
                
                Dim 傳票送出納檔名 As String = PostedFile.FileName.Substring(0, 16)
                Dim 年 As String = PostedFile.FileName.Substring(4, 3)
                Dim 開票日期 As String = ""
                Dim 種類 As String = ""
                Dim 傳票號碼 As String = ""
                Dim 之 As String = ""
                Dim 會計科目及摘要 As String = ""
                Dim 支票編號 As String = ""
                Dim 收入金額 As String = ""
                Dim 支出金額 As String = ""
                Dim 收入金額405 As Long = 0
                Dim 支出金額405 As Long = 0
                Dim 餘額405 As String = ""
                Dim 收入金額409 As Long = 0
                Dim 支出金額409 As Long = 0
                Dim 餘額409 As String = ""
                Dim 廠商及備註 As String = ""
                Dim 名稱 As String = ""
                
                Dim 匯入銀行名稱 As String = ""
                Dim 匯入帳號 As String = ""
                Dim 收款人匯款戶名 As String = ""
                Dim 摘要說明 As String = ""
                Dim taketp As String = ""
                
                Dim j As Long = 1
                For i = 1 To Sheet2.GetLength(0) - 1
                    種類 = Sheet2(i, 1)
                    If 種類 = "1"
                        種類 = "收"
                    Else If 種類 = "2"
                        種類 = "支"
                    Else If 種類 = "3"
                        種類 = "現"
                    End If
                    傳票號碼 = Sheet2(i, 1).ToString() & CLng(Sheet2(i, 2)).ToString("000000")
                    會計科目及摘要 = Sheet1(j, 6).Replace("'", "''")
                    If Sheet2(i, 3) = 0 And Sheet2(i, 8) <> "t" 't為工程保留款，待工程竣工後才會支出，會另外有一筆真正支出用的款項，而t的只是作帳用，所以不計算
                        收入金額405 = 收入金額405 + Sheet2(i, 7)
                    Else If Sheet2(i, 3) = 1 And Sheet2(i, 8) <> "t"
                        支出金額405 = 支出金額405 + Sheet2(i, 7)
                    End If
                    If Sheet2(i, 4) = 1
                        廠商及備註 = Sheet2(i, 6)
                    End If
                    
                    If Sheet2(i + 1, 2) <> Sheet2(i, 2)
                        If IsNumeric(Sheet2(i, 9))'12/16 注意此行運作狀況
                            Response.Write( "<script language='javascript'>alert('格式有錯誤請檢查資料是否錯誤，I列')")
                            Response.Redirect(Request.Url.ToString())
                        Else
                            If Sheet2(i, 9) = "C002"
                                收入金額409 = 收入金額405
                                支出金額409 = 支出金額405
                                收入金額405 = 0
                                支出金額405 = 0
                            End If
                        End If
                        data.InsertCommand = _
                        "IF NOT EXISTS(SELECT * FROM 現金備查簿 WHERE 年 = '" & 年 & "' AND 傳票號碼 = '" & 傳票號碼 & "') " & _
                        "INSERT INTO 現金備查簿 " & _
                        "(傳票送出納檔名, 年, 種類, 傳票號碼, 會計科目及摘要, 支票編號, 收入金額405, 支出金額405, 餘額405, 收入金額409, 支出金額409, 餘額409, 廠商及備註) " & _
                        "VALUES " & _
                        "(NULLIF(N'" & 傳票送出納檔名 & "', ''), " & _
                        "NULLIF(N'" & 年 & "', ''), " & _
                        "NULLIF(N'" & 種類 & "', ''), " & _
                        "NULLIF(N'" & 傳票號碼 & "', ''), " & _
                        "NULLIF(N'" & 會計科目及摘要 & "', ''), " & _
                        "NULLIF(N'" & 支票編號 & "', ''), " & _
                        "NULLIF(N'" & 收入金額405.ToString() & "', '0'), " & _
                        "NULLIF(N'" & 支出金額405.ToString() & "', '0'), " & _
                        "NULLIF(N'" & 餘額405 & "', ''), " & _
                        "NULLIF(N'" & 收入金額409.ToString() & "', '0'), " & _
                        "NULLIF(N'" & 支出金額409.ToString() & "', '0'), " & _
                        "NULLIF(N'" & 餘額409 & "', ''), " & _
                        "NULLIF(N'" & 廠商及備註 & "', ''))"
                        data.Insert()
                        收入金額405 = 0
                        支出金額405 = 0
                        收入金額409 = 0
                        支出金額409 = 0
                        j = j + 1
                    End If
                Next
                
                '解決主計室重送的問題，要先刪除才可以再次匯入，避免把已經改過的資料覆蓋掉了
                Dim 已上傳 As Boolean = False
                j = 1
                For i = 1 To Sheet2.GetLength(0) - 1
                    開票日期 = ""
                    傳票號碼 = ""
                    之 = ""
                    名稱 = ""
                    收入金額 = ""
                    支出金額 = ""
                    匯入銀行名稱 = ""
                    匯入帳號 = ""
                    收款人匯款戶名 = ""
                    摘要說明 = ""
                    
                    開票日期 = Sheet1(j, 8)
                    傳票號碼 = Sheet2(i, 1).ToString() & CLng(Sheet2(i, 2)).ToString("000000")
                    之 = Sheet2(i, 4).ToString()
                    名稱 = Sheet2(i, 6).ToString()
                    匯入銀行名稱 = Sheet2(i, 17)
                    If Sheet2(i, 3) = 0
                        收入金額 = Sheet2(i, 7)
                    Else If Sheet2(i, 3) = 1
                        支出金額 = Sheet2(i, 7)
                    End If
                    匯入帳號 = Sheet2(i, 18)
                    收款人匯款戶名 = Sheet2(i, 19)
                    If 收款人匯款戶名 = ""
                        收款人匯款戶名 = ""
                    End If
                    收款人匯款戶名 = If(收款人匯款戶名 = "", "", 收款人匯款戶名)
                    taketp = Sheet2(i, 5)
                    '主計室會選錯
                    '1 = 匯款
                    '2 = 零用金
                    '3 = 沒用過
                    '4 = 零用金
                    '5 = 支票
                    '6 = 其他
                    
                    摘要說明 = Sheet1(j, 6).Replace("'", "''")
                    If 匯入帳號 <> ""
                        摘要說明 = "網路匯款"
                    Else If 收款人匯款戶名.Contains("電子支付")
                        摘要說明 = "電子支付"
                    Else If Regex.IsMatch(摘要說明, "[0-9]{3}年度零用金")
                        摘要說明 = Regex.Match(摘要說明, "[0-9]{3}年度零用金").ToString()
                    Else If taketp = "2" Or taketp = "4"
                        摘要說明 = "零用金"
                    Else If 摘要說明.Contains("ETC欠費")
                        摘要說明 = "ETC欠費"
                    Else If 摘要說明.Contains("營業稅")
                        摘要說明 = "營業稅"
                    Else If 摘要說明.Contains("公保費")
                        摘要說明 = "公保費"
                    Else If Regex.IsMatch(摘要說明, "[0-9]{3}年[0-9]{1,2}月份電話費")
                        摘要說明 = Regex.Match(摘要說明, "[0-9]{3}年[0-9]{1,2}月份電話費").ToString().Replace("年", "/").Replace("月份", "")
                    Else If Regex.IsMatch(摘要說明, "[0-9]{3}年[0-9]{1,2}月份地磅電費")
                        摘要說明 = Regex.Match(摘要說明, "[0-9]{3}年[0-9]{1,2}月份地磅電費").ToString().Replace("年", "/").Replace("月份", "").Replace("地磅電費", "地磅等電費")
                    Else If 摘要說明.Contains("退休撫卹基金")
                        摘要說明 = "退撫基金"
                    Else If 摘要說明.Contains("中分局電費")
                        摘要說明 = "中分局等電費"
                    Else If 摘要說明.Contains("勞工退休") And 摘要說明.Contains("舊制")
                        摘要說明 = "勞退舊制"
                    Else If 摘要說明.Contains("約聘僱") And 摘要說明.Contains("離職儲金")
                        摘要說明 = "約聘僱離職儲金"
                    Else If 名稱.Contains("建築師事務所")
                        摘要說明 = "執行業務所得等"
                    End If
                    
                    data.UpdateCommand = _
                        "IF N'" & 匯入帳號 & "' = '' " & _
                            "BEGIN " & _
                                "UPDATE 傳票資料 SET " & _
                                "傳票資料.開票日期 = NULLIF(N'" & 開票日期 & "',''), " & _
                                "傳票資料.傳票送出納檔名 = NULLIF(N'" & 傳票送出納檔名 & "',''), " & _
                                "傳票資料.摘要說明 = NULLIF(N'" & 摘要說明 & "','') " & _
                                "WHERE 傳票資料.年 = '" & 年 & "' AND 傳票資料.傳票號碼 = '" & 傳票號碼 & "' AND 傳票資料.之 = '" & 之 & "' " & _
                            "END " & _
                        "ELSE " & _
                            "BEGIN " & _
                                "UPDATE 傳票資料 SET " & _
                                "傳票資料.開票日期 = NULLIF(N'" & 開票日期 & "',''), " & _
                                "傳票資料.傳票送出納檔名 = NULLIF(N'" & 傳票送出納檔名 & "',''), " & _
                                "傳票資料.摘要說明 = NULLIF(N'" & 摘要說明 & "',''), " & _
                                "傳票資料.收款人代碼 = 收款人.收款人代碼, " & _
                                "傳票資料.收款人名稱 = 收款人.收款人名稱, " & _
                                "傳票資料.匯入銀行名稱 = NULLIF(N'" & 匯入銀行名稱 & "',''), " & _
                                "傳票資料.匯入銀行代碼 = 收款人.匯入銀行代碼, " & _
                                "傳票資料.匯入帳號 = 收款人.匯入帳號, " & _
                                "傳票資料.收款人匯款戶名 = 收款人.收款人匯款戶名, " & _
                                "傳票資料.收款人統編 = 收款人.收款人統編 " & _
                                "FROM 傳票資料 INNER JOIN (SELECT TOP 1 * FROM 收款人 WHERE 收款人.匯入帳號 LIKE N'%" & 匯入帳號 & "' ORDER BY 收款人.收款人代碼 DESC) AS 收款人 " & _
                                "ON 傳票資料.年 = '" & 年 & "' AND 傳票資料.傳票號碼 = '" & 傳票號碼 & "' AND 傳票資料.之 = '" & 之 & "' " & _
                            "END"
                    'data.Update()
                    
                    data.UpdateCommand = _
                    "UPDATE 傳票資料 SET " & _
                    "傳票資料.開票日期 = NULLIF(N'" & 開票日期 & "',''), " & _
                    "傳票資料.傳票送出納檔名 = NULLIF(N'" & 傳票送出納檔名 & "',''), " & _
                    "傳票資料.摘要說明 = NULLIF(N'" & 摘要說明 & "','') " & _
                    "WHERE 傳票資料.年 = '" & 年 & "' AND 傳票資料.傳票號碼 = '" & 傳票號碼 & "' AND 傳票資料.之 = '" & 之 & "'"
                    data.Update()
                    
                    data.UpdateCommand = "UPDATE 現金備查簿 SET " & _
                    "傳票送出納檔名 = NULLIF(N'" & 傳票送出納檔名 & "','') " & _
                    "WHERE 年 = '" & 年 & "' AND 傳票號碼 = '" & 傳票號碼 & "'"
                    data.Update()
                    
                    If i = 1
                        data.SelectCommand = "SELECT id FROM 傳票資料 WHERE 年 = '" & 年 & "' AND 傳票號碼 = '" & 傳票號碼 & "'"
                        data_dv = data.Select(New DataSourceSelectArguments)
                        已上傳 = If(data_dv.Count > 0, True, False)
                    End If
                    If 已上傳 = False AND Sheet2(i, 8) <> "t"
                        data.InsertCommand = "INSERT INTO 傳票資料 (傳票送出納檔名, 摘要說明, 年, 開票日期, 傳票號碼, 之, 名稱, 收入金額, 支出金額) VALUES (NULLIF(N'" & 傳票送出納檔名 & "',''), NULLIF(N'" & 摘要說明 & "',''), NULLIF(N'" & 年 & "',''), NULLIF(N'" & 開票日期 & "',''), NULLIF(N'" & 傳票號碼 & "',''), NULLIF(N'" & 之 & "',''), NULLIF(N'" & 名稱 & "',''), NULLIF(N'" & 收入金額 & "',''), NULLIF(N'" & 支出金額 & "',''))"
                        If 匯入帳號 <> ""
                            Dim 匯款上限 As String = "50000000"
                            While CLng(支出金額) > CLng(匯款上限)
                                data.InsertCommand = "INSERT INTO 傳票資料 (傳票送出納檔名, 摘要說明, 年, 開票日期, 傳票號碼, 之, 名稱, 收入金額, 支出金額, 匯入銀行名稱, 匯入帳號, 收款人匯款戶名) VALUES (NULLIF(N'" & 傳票送出納檔名 & "',''), NULLIF(N'" & 摘要說明 & "',''), NULLIF(N'" & 年 & "',''), NULLIF(N'" & 開票日期 & "',''), NULLIF(N'" & 傳票號碼 & "',''), NULLIF(N'" & 之 & "',''), NULLIF(N'" & 名稱 & "',''), NULLIF(N'" & "" & "',''), NULLIF(N'" & 匯款上限 & "',''), NULLIF(N'" & 匯入銀行名稱 & "',''), NULLIF(N'" & 匯入帳號 & "',''), NULLIF(N'" & 收款人匯款戶名 & "',''))"
                                data.Insert()
                                支出金額 = (CLng(支出金額) - CLng(匯款上限)).ToString()
                            End While
                            data.InsertCommand = "INSERT INTO 傳票資料 (傳票送出納檔名, 摘要說明, 年, 開票日期, 傳票號碼, 之, 名稱, 收入金額, 支出金額, 匯入銀行名稱, 匯入帳號, 收款人匯款戶名) VALUES (NULLIF(N'" & 傳票送出納檔名 & "',''), NULLIF(N'" & 摘要說明 & "',''), NULLIF(N'" & 年 & "',''), NULLIF(N'" & 開票日期 & "',''), NULLIF(N'" & 傳票號碼 & "',''), NULLIF(N'" & 之 & "',''), NULLIF(N'" & 名稱 & "',''), NULLIF(N'" & 收入金額 & "',''), NULLIF(N'" & 支出金額 & "',''), NULLIF(N'" & 匯入銀行名稱 & "',''), NULLIF(N'" & 匯入帳號 & "',''), NULLIF(N'" & 收款人匯款戶名 & "',''))"
                        End If
                        data.Insert()
                    End If
                    If Sheet2(i + 1, 2) <> Sheet2(i, 2)
                        j = j + 1
                    End If
                Next
            Else If xlWorksheet.Cells(1, 3).Value = "客戶基本資料清冊"
                Sheet1 = xlWorksheet.Range(xlWorksheet.Cells(6, 1), xlWorksheet.Cells(xlWorksheet.Range("A65536").End(xlUp).Row, 26)).Value
                
                Dim 匯入銀行代碼 As String = ""
                Dim 匯入帳號 As String = ""
                Dim 收款人匯款戶名 As String = ""
                Dim 收款人統編 As String = ""
                Dim 收款人EMAIL As String = ""
                
                For i As Long = 1 To Sheet1.GetLength(0)
                    匯入銀行代碼 = If(Sheet1(i, 11) = Nothing And Sheet1(i, 12) = Nothing, "", Sheet1(i, 11).ToString() & Sheet1(i, 12).ToString())
                    匯入帳號 = If(Sheet1(i, 13) = Nothing , "", Sheet1(i, 13).ToString())
                    收款人匯款戶名 = If(Sheet1(i, 14) = Nothing , "", Sheet1(i, 14).ToString())
                    收款人統編 = If(Sheet1(i, 7) = Nothing , "", Sheet1(i, 7).ToString())
                    收款人EMAIL = If(Sheet1(i, 26) = Nothing , "", Sheet1(i, 26).ToString())
                    
                    匯入銀行代碼 = 匯入銀行代碼.Replace(Chr(9), "")
                    匯入帳號 = 匯入帳號.Replace(Chr(9), "")
                    收款人匯款戶名 = 收款人匯款戶名.Replace(Chr(9), "")
                    收款人統編 = 收款人統編.Replace(Chr(9), "")
                    收款人EMAIL = 收款人EMAIL.Replace(Chr(9), "")
                    
                    data.InsertCommand = _
                    "IF NOT EXISTS( " & _
                        "SELECT * FROM 收款人 " & _
                        "WHERE 匯入帳號 = N'" & 匯入帳號 & "' AND 收款人匯款戶名 = N'" & 收款人匯款戶名 & "') " & _
                        "BEGIN " & _
                            "INSERT INTO 收款人 (序號, 收款人代碼, 匯入銀行代碼, 匯入帳號, 收款人匯款戶名, 收款人統編, 收款人EMAIL) " & _
                            "VALUES (" & _
                            "(SELECT (SELECT ISNULL(MAX(序號加收款人代碼) + 1, 1) FROM (VALUES (MAX(序號)), (MAX(收款人代碼))) AS VALUE(序號加收款人代碼)) FROM 收款人), " & _
                            "(SELECT (SELECT ISNULL(MAX(序號加收款人代碼) + 1, 1) FROM (VALUES (MAX(序號)), (MAX(收款人代碼))) AS VALUE(序號加收款人代碼)) FROM 收款人), " & _
                            "N'" & 匯入銀行代碼 & "', N'" & 匯入帳號 & "', N'" & 收款人匯款戶名 & "', N'" & 收款人統編 & "', N'" & 收款人EMAIL & "') " & _
                        "END " & _
                    "ELSE " & _
                        "BEGIN " & _
                            "UPDATE 收款人 SET " & _
                            "匯入銀行代碼 = CASE WHEN ISNULL(匯入銀行代碼, '') <> '' THEN 匯入銀行代碼 ELSE N'" & 匯入銀行代碼 & "' END, " & _
                            "收款人匯款戶名 = CASE WHEN ISNULL(收款人匯款戶名, '') <> '' THEN 收款人匯款戶名 ELSE N'" & 收款人匯款戶名 & "' END, " & _
                            "收款人統編 = CASE WHEN ISNULL(收款人統編, '') <> '' THEN 收款人統編 ELSE N'" & 收款人統編 & "' END, " & _
                            "收款人EMAIL = CASE WHEN ISNULL(收款人EMAIL, '') <> '' THEN 收款人EMAIL ELSE N'" & 收款人EMAIL & "' END " & _
                            "WHERE 匯入帳號 = N'" & 匯入帳號 & "' AND 收款人匯款戶名 = N'" & 收款人匯款戶名 & "' " & _
                        "END"
                    data.Insert()
                Next
                
                'data.DeleteCommand = _
                '"WITH CTE AS(SELECT RN = ROW_NUMBER() OVER (PARTITION BY 匯入帳號 ORDER BY 收款人代碼 DESC) FROM 收款人) " & _
                '"DELETE FROM CTE WHERE RN > 1"
                'data.Delete()
            Else If xlWorksheet.Cells(5, 1).Value = "傳票查詢表"
                Sheet1 = xlWorksheet.Range(xlWorksheet.Cells(1, 1), xlWorksheet.Cells(xlWorksheet.Range("J65536").End(xlUp).Row, 12)).Value
                Dim 開票日期 As String = ""
                Dim 種類 As String = "分"
                Dim 傳票號碼 As String = "4" & Sheet1(9, 3)
                Dim 會計科目及摘要 As String = ""
                Dim 金額 As Long = 0
                Dim 年 As String = Left(Sheet1(9, 1), 3)
                
                For i = 9 To Sheet1.GetLength(0)
                    If Sheet1(i, 2) = "收入" Or Sheet1(i, 2) = "支出" Or Sheet1(i, 2) = "現轉"
                        Me.Label1.Text = ""
                        Me.Label2.Text = "目前只能匯入只有分錄的傳票查詢表"
                        Exit Sub
                    End If
                    '翻頁
                    If (i Mod 34) = 32
                        i = i + 10
                    End If
                Next
                
                For i = 9 To Sheet1.GetLength(0)
                    If (Sheet1(i, 2) = "分錄" And ("4" & Sheet1(i, 3)) <> 傳票號碼) Or i = Sheet1.GetLength(0)
                        data.InsertCommand = _
                        "IF NOT EXISTS(" & _
                        "SELECT * FROM 分錄 " & _
                        "WHERE 年 = N'" & 年 & "' AND 傳票號碼 = N'" & 傳票號碼 & "') " & _
                            "BEGIN " & _
                                "INSERT INTO 分錄 (開票日期, 種類, 傳票號碼, 會計科目及摘要, 收入金額405, 年) " & _
                                "VALUES(" & _
                                "NULLIF(N'" & 開票日期 & "',''), " & _
                                "NULLIF(N'" & 種類 & "',''), " & _
                                "NULLIF(N'" & 傳票號碼 & "',''), " & _
                                "NULLIF(N'" & 會計科目及摘要 & "',''), " & _
                                "NULLIF(N'" & 金額.ToString() & "',''), " & _
                                "NULLIF(N'" & 年 & "','')) " & _
                            "END"
                        data.Insert()
                        'data.UpdateCommand = _
                        '"update 分錄 set 開票日期 = NULLIF(N'" & 開票日期 & "','') " & _
                        '"WHERE 年 = N'" & 年 & "' AND 傳票號碼 = N'" & 傳票號碼 & "'"
                        'data.Update()
                        
                        傳票號碼 = "4" & Sheet1(i, 3)
                        金額 = CLng("0" & Sheet1(i, 7))
                    Else
                        金額 = 金額 + CLng("0" & Sheet1(i, 7))
                    End If
                    '會計科目及摘要、開票日期
                    If Sheet1(i, 2) = "分錄" And ("4" & Sheet1(i, 3)) = 傳票號碼
                        會計科目及摘要 = Sheet1(i, 10)
                        開票日期 = Sheet1(i, 1).Replace("/", "")
                    Else
                        會計科目及摘要 = 會計科目及摘要 & Sheet1(i, 10)
                    End If
                    '翻頁
                    If (i Mod 34) = 32
                        i = i + 10
                    End If
                Next
            Else If xlWorksheet.Cells(6, 1).Value = "帳戶明細" Or xlWorksheet.Cells(6, 1).Value = "專案代收查詢"
                Try
                    File.Copy(MyExcel, MapPath(".\Excel\對帳\") & PostedFile.FileName, False)
                Catch
                End Try
                data.InsertCommand = _
                "IF NOT EXISTS(SELECT * FROM 對帳 WHERE 檔名 = N'" & PostedFile.FileName & "') " & _
                    "BEGIN " & _
                        "INSERT INTO 對帳 (帳戶, 上傳時間, 檔名, 備註, 起, 迄) " & _
                        "VALUES('409', GETDATE(), N'" & PostedFile.FileName & "', N'" & xlWorksheet.Cells(6, 1).Value & "', " & _
                        "NULLIF(N'" & If(xlWorksheet.Cells(6, 1).Value = "帳戶明細", xlWorksheet.Cells(16, 3).Value, xlWorksheet.Cells(17, 3).Value) & "', ''), " & _
                        "NULLIF(N'" & If(xlWorksheet.Cells(6, 1).Value = "帳戶明細", xlWorksheet.Cells(16, 7).Value, xlWorksheet.Cells(17, 7).Value) & "', '')) " & _
                    "END"
                data.Insert()
            Else If xlWorksheet.Cells(2, 1).Value = "  銀行名稱:土地銀行北台中分行" Or xlWorksheet.Cells(2, 1).Value = "  銀行名稱:中國信託銀行台中分行"
                Try
                    File.Copy(MyExcel, MapPath(".\Excel\對帳\") & PostedFile.FileName, False)
                Catch
                End Try
                data.InsertCommand = _
                "IF NOT EXISTS(SELECT * FROM 對帳 WHERE 檔名 = N'" & PostedFile.FileName & "') " & _
                    "BEGIN " & _
                        "INSERT INTO 對帳 (帳戶, 上傳時間, 檔名, 備註, 起, 迄) " & _
                        "VALUES('" & If(xlWorksheet.Cells(2, 1).Value = "  銀行名稱:土地銀行北台中分行", "405", "409") & "', " & _
                        "GETDATE(), N'" & PostedFile.FileName & "', N'調結表', NULL, " & _
                        "NULLIF(N'" & _
                        taiwancalendarto(If(Regex.Match(xlWorksheet.Cells(4, 2).Value.Replace(" ", ""), "[0-9]{3}[^0-9]+[0-9]{1,2}[^0-9]+[0-9]{1,2}"), "").ToString().Replace("年", "/").Replace("月", "/")) & 
                        "', '')) " & _
                    "END"
                Me.Label1.Text = data.InsertCommand
                data.Insert()
            End If
            
            xlWorkbook.Close(False)
            System.IO.File.Delete(MyExcel)
        Next
        xlApp.Quit()
        ReleaseObject(xlWorksheet)
        ReleaseObject(xlWorkbook)
        ReleaseObject(xlApp)
        Response.Redirect(Request.RawUrl)
    End Sub
End Class

