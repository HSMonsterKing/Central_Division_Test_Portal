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
Partial Class 上傳
    Inherits System.Web.UI.Page
    Dim con_14 As String = System.Web.Configuration.WebConfigurationManager.ConnectionStrings("ApplicationServices").ConnectionString
    Dim data As New SqlDataSource
    Dim data_dv As Data.DataView
    Dim data_dv1 As Data.DataView
    Dim data_dv2 As Data.DataView
    Dim data_dv3 As Data.DataView
    Dim i As Long = 0
    Dim j As Long = 0
    Dim k As Long = 0
    Public Function DBNullHandler(ByVal data_dv As Object) As String
        If IsDBNull(data_dv)
            Return ""
        Else
            Return data_dv
        End If
    End Function
    Public Function DBNullHandlerLong(ByVal data_dv As Object) As Long
        If IsDBNull(data_dv)
            Return 0
        Else
            Return data_dv
        End If
    End Function
    Public Function nothinghandler(ByVal a As Object) As Object
        If a IsNot Nothing
            return a
        Else
            return ""
        End If
    End Function
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not Page.IsPostBack Then
        Else
        End If
    End Sub
    Protected Sub Button3_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim _GUID As String = Guid.NewGuid().ToString("N")
        Dim MyExcel As String = MapPath(".\Excel\Temp\") & _GUID & ".xls"
        If FileUpload3.HasFile Then
            FileUpload3.SaveAs(MyExcel)
        Else
            Exit Sub
        End If
        
        data.ConnectionString = con_14
        Dim xlApp As New Excel.ApplicationClass()
        xlApp.DisplayAlerts = False
        xlApp.ScreenUpdating = false
        xlApp.EnableEvents = false
        Dim xlWorkBook As Excel.Workbook = xlApp.Workbooks.Open(MyExcel, 0, False, 5, "", "", True, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", False, False, 0, True, 1, 0)
        Dim xlWorkSheet As Excel.Worksheet = xlWorkBook.ActiveSheet
        
        Dim arr As Object
        arr = xlWorkSheet.Range(xlWorkSheet.Cells(1, 1), xlWorkSheet.Cells(xlWorkSheet.Range("D65536").End(-4162).Row, 13)).Value
        
        Dim _年月日2 As String = arr(4, 2)
        _年月日2 = taiwancalendarto(_年月日2)
        data.DeleteCommand = "delete from 現金備查簿 where year(年月日) = '" & _年月日2.Split("/")(0) & "'"
        data.Delete()
        
        For i As Long = 4 To xlWorkSheet.Range("D65536").End(-4162).Row
            If arr(i, 4) IsNot Nothing And Trim(nothinghandler(arr(i, 3))) <> ""
                Dim _序號 As String = nothinghandler(arr(i, 1))
                Dim _年月日 As String = nothinghandler(arr(i, 2))
                _年月日 = taiwancalendarto(_年月日)
                Dim _種類 As String = Trim(nothinghandler(arr(i, 3)))
                If _種類 = "收"
                    _種類 = "1"
                Else If _種類 = "支"
                    _種類 = "2"
                Else If _種類 = "現"
                    _種類 = "3"
                End If
                Dim _號數 As String = nothinghandler(arr(i, 4))
                Dim _會計科目及摘要 As String = nothinghandler(arr(i, 5)).ToString().Replace("'", "''")
                Dim _支票編號 As String = nothinghandler(arr(i, 6))
                Dim _收入金額405 As String = nothinghandler(arr(i, 7))
                Dim _支出金額405 As String = nothinghandler(arr(i, 8))
                Dim _餘額405 As String = nothinghandler(arr(i, 9))
                Dim _收入金額409 As String = nothinghandler(arr(i, 10))
                Dim _支出金額409 As String = nothinghandler(arr(i, 11))
                Dim _餘額409 As String = nothinghandler(arr(i, 12))
                Dim _廠商及備註 As String = nothinghandler(arr(i, 13))
                data.InsertCommand = "insert into 現金備查簿 (序號, 年月日, 種類, 號數, 會計科目及摘要, 支票編號, 收入金額405, 支出金額405, 餘額405, 收入金額409, 支出金額409, 餘額409, 廠商及備註) VALUES (NULLIF(N'" & _序號 & "',''),NULLIF(N'" & _年月日 & "',''),NULLIF(N'" & _種類 & "',''),NULLIF(N'" & _號數 & "',''),NULLIF(N'" & _會計科目及摘要 & "',''),NULLIF(N'" & _支票編號 & "',''),NULLIF(N'" & _收入金額405 & "',''),NULLIF(N'" & _支出金額405 & "',''),NULLIF(N'" & _餘額405 & "',''),NULLIF(N'" & _收入金額409 & "',''),NULLIF(N'" & _支出金額409 & "',''),NULLIF(N'" & _餘額409 & "',''),NULLIF(N'" & _廠商及備註 & "',''))"
                data.Insert()
            End If
        Next
        
        xlWorkBook.Save()
        xlWorkBook.Close()
        xlApp.Quit()
        ReleaseObject(xlWorkSheet)
        ReleaseObject(xlWorkBook)
        ReleaseObject(xlApp)
        System.IO.File.Delete(MyExcel)
    End Sub
    Protected Sub Button4_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button4.Click
        Dim _GUID As String = Guid.NewGuid().ToString("N")
        Dim MyExcel As String = MapPath(".\Excel\Temp\") & _GUID & ".xls"
        If FileUpload4.HasFile Then
            FileUpload4.SaveAs(MyExcel)
        Else
            Exit Sub
        End If
        
        data.ConnectionString = con_14
        Dim xlApp As New Excel.ApplicationClass()
        xlApp.DisplayAlerts = False
        xlApp.ScreenUpdating = false
        xlApp.EnableEvents = false
        Dim xlWorkBook As Excel.Workbook = xlApp.Workbooks.Open(MyExcel, 0, False, 5, "", "", True, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", False, False, 0, True, 1, 0)
        Dim wsSLIP As Excel.Worksheet = CType(xlWorkBook.Sheets("SLIP"), Excel.Worksheet)
        Dim wsPAYEE As Excel.Worksheet = CType(xlWorkBook.Sheets("PAYEE"), Excel.Worksheet)
        
        Dim SLIP As Object
        Dim PAYEE As Object
        SLIP = wsSLIP.Range(wsSLIP.Cells(2, 1), wsSLIP.Cells(wsSLIP.Range("D65536").End(-4162).Row, 7)).Value
        PAYEE = wsPAYEE.Range(wsPAYEE.Cells(2, 1), wsPAYEE.Cells(wsPAYEE.Range("D65536").End(-4162).Row + 1, 23)).Value 'Row + 1用於判斷檔案結尾
        
        Dim _序號 As String = ""
        Dim _年月日 As String = ""
        Dim _種類 As String = ""
        Dim _號數 As String = ""
        Dim _會計科目及摘要 As String = ""
        Dim _支票編號 As String = ""
        Dim _收入金額405 As Long = 0
        Dim _支出金額405 As Long = 0
        Dim _餘額405 As String = ""
        Dim _收入金額409 As String = ""
        Dim _支出金額409 As String = ""
        Dim _餘額409 As String = ""
        Dim _廠商及備註 As String = ""
        
        For i As Long = 1 To PAYEE.GetLength(0) - 1
            _號數 = PAYEE(i, 1).ToString() & CLng(PAYEE(i, 2)).ToString("000000")
            data.SelectCommand = "select id from 現金備查簿 where (year(年月日)=year(getdate()) or 年月日 is null) and 號數='" & _號數 & "'"
            data_dv = data.Select(New DataSourceSelectArguments)
            data.SelectCommand = "select id from 詳細金額 where 現金備查簿id=" & data_dv(0)(0) & " and seq=" & PAYEE(i, 4) & ""
            data_dv1 = data.Select(New DataSourceSelectArguments)
            If data_dv1.Count = 0
                If PAYEE(i, 3) = 0
                    data.InsertCommand = "insert into 詳細金額 (現金備查簿id, 收入金額, seq) VALUES (" & data_dv(0)(0) & "," & PAYEE(i, 7) & "," & PAYEE(i, 4) & ")"
                Else If PAYEE(i, 3) = 1
                    data.InsertCommand = "insert into 詳細金額 (現金備查簿id, 支出金額, seq) VALUES (" & data_dv(0)(0) & "," & PAYEE(i, 7) & "," & PAYEE(i, 4) & ")"
                End If
                data.Insert()
            End If
        Next
        
        xlWorkBook.Save()
        xlWorkBook.Close()
        xlApp.Quit()
        ReleaseObject(wsSLIP)
        ReleaseObject(wsPAYEE)
        ReleaseObject(xlWorkBook)
        ReleaseObject(xlApp)
        System.IO.File.Delete(MyExcel)
    End Sub
    Protected Sub Button5_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button5.Click
        Dim _GUID As String = Guid.NewGuid().ToString("N")
        Dim MyExcel As String = MapPath(".\Excel\Temp\") & _GUID & ".xls"
        If FileUpload5.HasFile Then
            FileUpload5.SaveAs(MyExcel)
        Else
            Exit Sub
        End If
        
        data.ConnectionString = con_14
        Dim xlApp As New Excel.ApplicationClass()
        xlApp.DisplayAlerts = False
        xlApp.ScreenUpdating = false
        xlApp.EnableEvents = false
        Dim xlWorkBook As Excel.Workbook = xlApp.Workbooks.Open(MyExcel, 0, False, 5, "", "", True, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", False, False, 0, True, 1, 0)
        Dim ws As Excel.Worksheet = CType(xlWorkBook.Sheets("客戶基本資料清冊"), Excel.Worksheet)
        
        Dim arr As Object
        arr = ws.Range(ws.Cells(6, 1), ws.Cells(ws.Range("A65536").End(-4162).Row, 26)).Value
        
        Dim _序號 As String = ""
        Dim _收款人代碼 As String = ""
        Dim _收款人名稱 As String = ""
        Dim _收款行代號 As String = ""
        Dim _收款人帳號 As String = ""
        Dim _收款人戶名 As String = ""
        Dim _收款人統編 As String = ""
        
        For i As Long = 1 To arr.GetLength(0)
            _序號 = arr(i, 2).ToString()
            _收款人代碼 = arr(i, 5).ToString()
            _收款人名稱 = arr(i, 6).ToString()
            _收款行代號 = arr(i, 11).ToString() & arr(i, 12).ToString()
            _收款人帳號 = arr(i, 13).ToString()
            _收款人戶名 = arr(i, 14).ToString()
            _收款人統編 = arr(i, 7).ToString()
            _序號 = _序號.Replace(Chr(9), "")
            _收款人代碼 = _收款人代碼.Replace(Chr(9), "")
            _收款人名稱 = _收款人名稱.Replace(Chr(9), "")
            _收款行代號 = _收款行代號.Replace(Chr(9), "")
            _收款人帳號 = _收款人帳號.Replace(Chr(9), "")
            _收款人戶名 = _收款人戶名.Replace(Chr(9), "")
            _收款人統編 = _收款人統編.Replace(Chr(9), "")
            data.SelectCommand = "select id from 收款人 where 序號='" & _序號 & "'"
            data_dv = data.Select(New DataSourceSelectArguments)
            If data_dv.Count = 0
                data.InsertCommand = "insert into 收款人 (序號, 收款人代碼, 收款人名稱, 匯入銀行代碼, 匯入帳號, 收款人匯款戶名, 收款人統編) VALUES (N'" & _序號 & "', N'" & _收款人代碼 & "', N'" & _收款人名稱 & "', N'" & _收款行代號 & "', N'" & _收款人帳號 & "', N'" & _收款人戶名 & "', N'" & _收款人統編 & "')"
                Try
                    data.Insert()
                Catch
                    Me.debug.Text = data.InsertCommand
                    Exit For
                End Try
            End If
        Next
        
        xlWorkBook.Save()
        xlWorkBook.Close()
        xlApp.Quit()
        ReleaseObject(ws)
        ReleaseObject(xlWorkBook)
        ReleaseObject(xlApp)
        System.IO.File.Delete(MyExcel)
    End Sub
    Protected Sub Button6_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button6.Click
        data.ConnectionString = con_14
        
        data.SelectCommand = "select id, 年月日 from 現金備查簿 where 年月日 is not null and 年月日!=''"
        data_dv = data.Select(New DataSourceSelectArguments)
        For i As Long = 0 To data_dv.Count - 1
            Dim id As String = data_dv(i)(0)
            Dim 新日期 As String = data_dv(i)(1)
            新日期 = (CLng(新日期.Split("-")(0)) - 1911).ToString() & "/" & 新日期.Split("-")(1) & "/" & 新日期.Split("-")(2)
            data.UpdateCommand = "update 現金備查簿 set 年月日='" & 新日期 & "' where id='" & id & "'"
            data.Update()
        Next
        
        data.SelectCommand = "select id, 年月日 from 現金備查簿 where 年月日 is null or 年月日=''"
        data_dv = data.Select(New DataSourceSelectArguments)
        For i As Long = 0 To data_dv.Count - 1
            Dim id As String = data_dv(i)(0)
            Dim 新日期 As String = "110"
            data.UpdateCommand = "update 現金備查簿 set 年月日='" & 新日期 & "' where id='" & id & "'"
            data.Update()
        Next
        
        data.UpdateCommand = "update 現金備查簿 set 種類='收' where 種類 Like '%1%'"
        data.Update()
        data.UpdateCommand = "update 現金備查簿 set 種類='支' where 種類 Like '%2%'"
        data.Update()
        data.UpdateCommand = "update 現金備查簿 set 種類='現' where 種類 Like '%3%'"
        data.Update()
    End Sub
    Protected Sub Button7_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button7.Click
        data.ConnectionString = con_14
        
        data.UpdateCommand = "update 現金備查簿 set 種類='收' where 種類 Like '%1%'"
        data.Update()
        data.UpdateCommand = "update 現金備查簿 set 種類='支' where 種類 Like '%2%'"
        data.Update()
        data.UpdateCommand = "update 現金備查簿 set 種類='現' where 種類 Like '%3%'"
        data.Update()
        
        data.UpdateCommand = "update 現金備查簿 set 年=year(年月日)-1911 where 年月日 is not null"
        data.Update()
        data.UpdateCommand = "update 現金備查簿 set 年=110 where 年月日 is null"
        data.Update()
    End Sub
    Protected Sub Button8_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button8.Click
        Dim _GUID As String = Guid.NewGuid().ToString("N")
        Dim MyExcel As String = MapPath(".\Excel\Temp\") & _GUID & ".xls"
        If FileUpload8.HasFile Then
            FileUpload8.SaveAs(MyExcel)
        Else
            Exit Sub
        End If
        
        data.ConnectionString = con_14
        Dim xlApp As New Excel.ApplicationClass()
        xlApp.DisplayAlerts = False
        xlApp.ScreenUpdating = false
        xlApp.EnableEvents = false
        Dim xlWorkBook As Excel.Workbook = xlApp.Workbooks.Open(MyExcel, 0, False, 5, "", "", True, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", False, False, 0, True, 1, 0)
        Dim xlWorkSheet As Excel.Worksheet = xlWorkBook.ActiveSheet
        
        Dim arr As Object
        arr = xlWorkSheet.Range(xlWorkSheet.Cells(1, 1), xlWorkSheet.Cells(xlWorkSheet.Range("B65536").End(-4162).Row, 18)).Value
        
        'data.DeleteCommand = "delete from 保管品 where 1=1"
        'data.Delete()
        
        '類種0是保證書
        '種類1是存單
        If arr(2, 1) = "原封戶保管品明細表"
            For i As Long = 7 To xlWorkSheet.Range("B65536").End(-4162).Row
                Dim _序號 As String = nothinghandler(arr(i, 1))
                Dim _日期 As String = nothinghandler(arr(i, 2))
                _日期 = _日期.Replace(".", "/")
                _日期 = taiwancalendarto(_日期)
                _日期 = If(IsDate(_日期), _日期, Today())
                Dim _保證書名稱或存單號碼 As String = Trim(nothinghandler(arr(i, 3)))
                Dim _收據編號 As String = nothinghandler(arr(i, 4))
                Dim _國保收據編號 As String = nothinghandler(arr(i, 5))
                Dim _戶名 As String = nothinghandler(arr(i, 6))
                Dim _品名 As String = nothinghandler(arr(i, 7))
                Dim _摘要 As String = nothinghandler(arr(i, 8))
                Dim _單位 As String = nothinghandler(arr(i, 9))
                Dim _數量 As String = nothinghandler(arr(i, 10))
                Dim _金額 As String = nothinghandler(arr(i, 11))
                Dim _保證書或存單保證期限 As String = nothinghandler(arr(i, 12))
                Dim _廠商保證責任期限 As String = nothinghandler(arr(i, 13))
                Dim _合約展延情形 As String = nothinghandler(arr(i, 14))
                Dim _保管處 As String = nothinghandler(arr(i, 15))
                Dim _承辦單位 As String = nothinghandler(arr(i, 16))
                Dim _承辦人 As String = nothinghandler(arr(i, 17))
                Dim _備考 As String = nothinghandler(arr(i, 18))
                Dim _種類 As String = "0"
                Dim _已收入 As String = "1"
                Dim _已支出 As String = "0"
                
                _序號 = "N'" & _序號 & "'"
                _日期 = "N'" & _日期 & "'"
                _保證書名稱或存單號碼 = "N'" & _保證書名稱或存單號碼 & "'"
                _收據編號 = "N'" & _收據編號 & "'"
                _國保收據編號 = "N'" & _國保收據編號 & "'"
                _戶名 = "N'" & _戶名 & "'"
                _品名 = "N'" & _品名 & "'"
                _摘要 = "N'" & _摘要 & "'"
                _單位 = "N'" & _單位 & "'"
                _數量 = "N'" & _數量 & "'"
                _金額 = "N'" & _金額 & "'"
                _保證書或存單保證期限 = "N'" & _保證書或存單保證期限 & "'"
                _廠商保證責任期限 = "N'" & _廠商保證責任期限 & "'"
                _合約展延情形 = "N'" & _合約展延情形 & "'"
                _保管處 = "N'" & _保管處 & "'"
                _承辦單位 = "N'" & _承辦單位 & "'"
                _承辦人 = "N'" & _承辦人 & "'"
                _備考 = "N'" & _備考 & "'"
                _種類 = "N'" & _種類 & "'"
                _已收入 = "N'" & _已收入 & "'"
                _已支出 = "N'" & _已支出 & "'"
                
                data.InsertCommand = "insert into 保管品明細表 " & _
                "(日期, 保證書名稱或存單號碼, 收據編號, 國保收據編號, 戶名, 品名, 摘要, 單位, 數量, 金額, 保證書或存單保證期限, 廠商保證責任期限, 合約展延情形, 保管處, 承辦單位, 承辦人, 備考, 種類, 已收入, 已支出) " & _
                "VALUES " & _
                "(" & _日期 & "," & _保證書名稱或存單號碼 & "," & _收據編號 & "," & _國保收據編號 & "," & _戶名 & "," & _品名 & "," & _摘要 & "," & _單位 & "," & _數量 & "," & _金額 & "," & _保證書或存單保證期限 & "," & _廠商保證責任期限 & "," & _合約展延情形 & "," & _保管處 & "," & _承辦單位 & "," & _承辦人 & "," & _備考 & "," & _種類 & "," & _已收入 & "," & _已支出 & ")"
                data.Insert()
            Next
        Else
            For i As Long = 7 To xlWorkSheet.Range("B65536").End(-4162).Row
                Dim _序號 As String = nothinghandler(arr(i, 1))
                Dim _日期 As String = nothinghandler(arr(i, 2))
                _日期 = _日期.Replace(".", "/")
                _日期 = taiwancalendarto(_日期)
                _日期 = If(IsDate(_日期), _日期, Today())
                Dim _收據編號 As String = nothinghandler(arr(i, 3))
                Dim _國保收據編號 As String = nothinghandler(arr(i, 4))
                Dim _保證書名稱或存單號碼 As String = Trim(nothinghandler(arr(i, 5)))
                Dim _戶名 As String = nothinghandler(arr(i, 6))
                Dim _品名 As String = nothinghandler(arr(i, 7))
                Dim _摘要 As String = nothinghandler(arr(i, 8))
                Dim _單位 As String = nothinghandler(arr(i, 9))
                Dim _數量 As String = nothinghandler(arr(i, 10))
                Dim _金額 As String = nothinghandler(arr(i, 11))
                Dim _保證書或存單保證期限 As String = nothinghandler(arr(i, 12))
                Dim _廠商保證責任期限 As String = nothinghandler(arr(i, 13))
                Dim _合約展延情形 As String = nothinghandler(arr(i, 14))
                Dim _保管處 As String = nothinghandler(arr(i, 15))
                Dim _承辦單位 As String = nothinghandler(arr(i, 16))
                Dim _承辦人 As String = nothinghandler(arr(i, 17))
                Dim _備考 As String = nothinghandler(arr(i, 18))
                Dim _種類 As String = "1"
                Dim _已收入 As String = "1"
                Dim _已支出 As String = "0"
                
                _序號 = "N'" & _序號 & "'"
                _日期 = "N'" & _日期 & "'"
                _保證書名稱或存單號碼 = "N'" & _保證書名稱或存單號碼 & "'"
                _收據編號 = "N'" & _收據編號 & "'"
                _國保收據編號 = "N'" & _國保收據編號 & "'"
                _戶名 = "N'" & _戶名 & "'"
                _品名 = "N'" & _品名 & "'"
                _摘要 = "N'" & _摘要 & "'"
                _單位 = "N'" & _單位 & "'"
                _數量 = "N'" & _數量 & "'"
                _金額 = "N'" & _金額 & "'"
                _保證書或存單保證期限 = "N'" & _保證書或存單保證期限 & "'"
                _廠商保證責任期限 = "N'" & _廠商保證責任期限 & "'"
                _合約展延情形 = "N'" & _合約展延情形 & "'"
                _保管處 = "N'" & _保管處 & "'"
                _承辦單位 = "N'" & _承辦單位 & "'"
                _承辦人 = "N'" & _承辦人 & "'"
                _備考 = "N'" & _備考 & "'"
                _種類 = "N'" & _種類 & "'"
                _已收入 = "N'" & _已收入 & "'"
                _已支出 = "N'" & _已支出 & "'"
                
                data.InsertCommand = "insert into 保管品明細表 " & _
                "(日期, 保證書名稱或存單號碼, 收據編號, 國保收據編號, 戶名, 品名, 摘要, 單位, 數量, 金額, 保證書或存單保證期限, 廠商保證責任期限, 合約展延情形, 保管處, 承辦單位, 承辦人, 備考, 種類, 已收入, 已支出) " & _
                "VALUES " & _
                "(" & _日期 & "," & _保證書名稱或存單號碼 & "," & _收據編號 & "," & _國保收據編號 & "," & _戶名 & "," & _品名 & "," & _摘要 & "," & _單位 & "," & _數量 & "," & _金額 & "," & _保證書或存單保證期限 & "," & _廠商保證責任期限 & "," & _合約展延情形 & "," & _保管處 & "," & _承辦單位 & "," & _承辦人 & "," & _備考 & "," & _種類 & "," & _已收入 & "," & _已支出 & ")"
                data.Insert()
            Next
        End If
        
        xlWorkBook.Save()
        xlWorkBook.Close()
        xlApp.Quit()
        ReleaseObject(xlWorkSheet)
        ReleaseObject(xlWorkBook)
        ReleaseObject(xlApp)
        System.IO.File.Delete(MyExcel)
    End Sub
End Class
