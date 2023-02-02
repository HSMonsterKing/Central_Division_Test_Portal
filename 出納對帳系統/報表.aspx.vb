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
Partial Class 報表
    Inherits System.Web.UI.Page
    Dim con_14 As String = System.Web.Configuration.WebConfigurationManager.ConnectionStrings("ApplicationServices").ConnectionString
    Dim data As New SqlDataSource
    Dim data_dv As Data.DataView
    Dim data_dv1 As Data.DataView
    Dim data_dv2 As Data.DataView
    Dim data_dv3 As Data.DataView
    Public Sub generatedropdownlist1()'新增日報表清單資訊
        data.SelectCommand = "select 結帳日期, id from 日報表 order by 結帳日期 desc"
        data_dv = data.Select(New DataSourceSelectArguments)
        Me.DropDownList1.Items.Clear()
        For i = 0 To data_dv.Count - 1
            Dim _結帳日期 As String = totaiwancalendar(data_dv(i)(0))
            _結帳日期 = "中華民國" & _結帳日期.Split("/")(0) & "年" & _結帳日期.Split("/")(1) & "月" & _結帳日期.Split("/")(2) & "日"
            Me.DropDownList1.Items.Add(_結帳日期)
            Me.DropDownList1.Items(i).Value = data_dv(i)(1)
        Next
    End Sub
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        data.ConnectionString = con_14
        If Not Page.IsPostBack Then
            generatedropdownlist1()
            DropDownList1_SelectedIndexChanged(sender, e)
        Else
        End If
    End Sub
    Protected Sub DropDownList1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DropDownList1.SelectedIndexChanged'日報表清單改變
        data.SelectCommand = "select * from 日報表 where id='" & Me.DropDownList1.SelectedValue & "'"
        data_dv = data.Select(New DataSourceSelectArguments)
        If data_dv.Count > 0
            '土地銀行北台中分行077056000014-中分局405專戶
            Me.LabelC6.Text = data_dv(0)(2).ToString()'上日結存
            Me.LabelD6.Text = data_dv(0)(3).ToString()'本日收入
            Me.LabelE6.Text = data_dv(0)(4).ToString()'本日支出
            Me.LabelC7.Text = data_dv(0)(5).ToString()
            Me.LabelD7.Text = data_dv(0)(6).ToString()
            Me.LabelE7.Text = data_dv(0)(7).ToString()
            Me.LabelC12.Text = data_dv(0)(8).ToString()
            Me.LabelE12.Text = data_dv(0)(9).ToString()
            Me.LabelC13.Text = data_dv(0)(10).ToString()
            Me.LabelE13.Text = data_dv(0)(11).ToString()
            Me.LabelC14.Text = data_dv(0)(12).ToString()
            Me.LabelE14.Text = data_dv(0)(13).ToString()
            Me.LabelC15.Text = data_dv(0)(14).ToString()
            Me.LabelE15.Text = data_dv(0)(15).ToString()
            Me.LabelF6.Text = Val(Me.LabelC6.Text) + Val(Me.LabelD6.Text) - Val(Me.LabelE6.Text)'本日結存
            Me.LabelF7.Text = Val(Me.LabelC7.Text) + Val(Me.LabelD7.Text) - Val(Me.LabelE7.Text)
            Me.LabelC11.Text = Val(Me.LabelC6.Text) + Val(Me.LabelC7.Text)'合計上日結存
            Me.LabelD11.Text = Val(Me.LabelD6.Text) + Val(Me.LabelD7.Text)'合計本日收入
            Me.LabelE11.Text = Val(Me.LabelE6.Text) + Val(Me.LabelE7.Text)'合計本日支出
            Me.LabelF11.Text = Val(Me.LabelC11.Text) + Val(Me.LabelD11.Text) - Val(Me.LabelE11.Text)'合計本日結存
            Me.LabelC6.Text = String.Format("{0:n0}", Val(Me.LabelC6.Text))'重新定義成貨幣
            Me.LabelD6.Text = String.Format("{0:n0}", Val(Me.LabelD6.Text))
            Me.LabelE6.Text = String.Format("{0:n0}", Val(Me.LabelE6.Text))
            Me.LabelC7.Text = String.Format("{0:n0}", Val(Me.LabelC7.Text))
            Me.LabelD7.Text = String.Format("{0:n0}", Val(Me.LabelD7.Text))
            Me.LabelE7.Text = String.Format("{0:n0}", Val(Me.LabelE7.Text))
            Me.LabelF6.Text = String.Format("{0:n0}", Val(Me.LabelF6.Text))
            Me.LabelF7.Text = String.Format("{0:n0}", Val(Me.LabelF7.Text))
            Me.LabelC11.Text = String.Format("{0:n0}", Val(Me.LabelC11.Text))
            Me.LabelD11.Text = String.Format("{0:n0}", Val(Me.LabelD11.Text))
            Me.LabelE11.Text = String.Format("{0:n0}", Val(Me.LabelE11.Text))
            Me.LabelF11.Text = String.Format("{0:n0}", Val(Me.LabelF11.Text))
            If Me.LabelE12.Text = ""
                Me.LabelE12.Text = "無開立傳票"
            End If
            If Me.LabelE13.Text = ""
                Me.LabelE13.Text = "無開立傳票"
            End If
            If Me.LabelE14.Text = ""
                Me.LabelE14.Text = "無開立傳票"
            End If
            If Me.LabelE15.Text = ""
                Me.LabelE15.Text = "無開立傳票"
            End If
        End If
    End Sub
    Protected Sub Button2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button2.Click'日報表下載
        Dim _GUID As String = Guid.NewGuid().ToString("N")
        Dim MyExcel As String = MapPath(".\Excel\Temp\") & _GUID & ".xls"
        System.IO.File.Copy(MapPath(".\Excel\現金備查簿.xls"), MyExcel)
        Dim xlApp As New Excel.ApplicationClass()
        xlApp.DisplayAlerts = False
        xlApp.ScreenUpdating = false
        xlApp.EnableEvents = false
        Dim xlWorkBook As Excel.Workbook = xlApp.Workbooks.Open(MyExcel, 0, False, 5, "", "", True, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", False, False, 0, True, 1, 0)
        Dim xlWorkSheet As Excel.Worksheet
        
        
        
        xlWorkSheet = CType(xlWorkBook.Sheets("備查簿"), Excel.Worksheet)
        data.SelectCommand = "select 結帳日期, year(結帳日期) - 1911 from 日報表 where id='" & Me.DropDownList1.SelectedValue & "'"
        data_dv = data.Select(New DataSourceSelectArguments)
        Dim _日報表結帳日期 As String = Me.DropDownList1.SelectedItem.Text
        Dim _年 As String = Me.DropDownList1.SelectedItem.Text
        If data_dv.Count > 0
            If Not IsDBNull(data_dv(0)(0))
                _日報表結帳日期 = data_dv(0)(0)
                _年 = data_dv(0)(1)
            Else
                _年 = _年.Substring(_年.IndexOf("國") + 1, 3)
                _日報表結帳日期 = _日報表結帳日期.Substring(_日報表結帳日期.IndexOf("國") + 1, 3) & "/" & _日報表結帳日期.Substring(_日報表結帳日期.IndexOf("年") + 1, 2) & "/" & _日報表結帳日期.Substring(_日報表結帳日期.IndexOf("月") + 1, 2)
            End If
        End If
        data.SelectCommand = "select * from 現金備查簿 where year(結帳日期)='" & _年 & "' + 1911 or (結帳日期 is null and YEAR(getdate())='" & _年 & "' + 1911) order by case when 序號 is null then 1 else 0 end, 序號, 傳票號碼"
        data_dv = data.Select(New DataSourceSelectArguments)
        Dim arr(data_dv.Count, 13) As Object
        For i = 0 To data_dv.Count - 1
            Dim _序號 As String = data_dv(i)(1).ToString()
            Dim _結帳日期 As String = data_dv(i)(2).ToString()
            _結帳日期 = totaiwancalendar(_結帳日期)
            Dim _種類 As String = data_dv(i)(3).ToString()
            If _種類 = "1"
                _種類 = "收"
            Else If _種類 = "2"
                _種類 = "支"
            Else If _種類 = "3"
                _種類 = "現"
            End If
            Dim _傳票號碼 As String = data_dv(i)(4).ToString()
            Dim _會計科目及摘要 As String = data_dv(i)(5).ToString()
            Dim _支票編號 As String = data_dv(i)(6).ToString()
            Dim _收入金額405 As String = data_dv(i)(7).ToString()
            Dim _支出金額405 As String = data_dv(i)(8).ToString()
            Dim _餘額405 As String = data_dv(i)(9).ToString()
            Dim _收入金額409 As String = data_dv(i)(10).ToString()
            Dim _支出金額409 As String = data_dv(i)(11).ToString()
            Dim _餘額409 As String = data_dv(i)(12).ToString()
            Dim _廠商及備註 As String = data_dv(i)(13).ToString()
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
            'xlWorkSheet.Cells(4 + i, 1) = _序號
            'xlWorkSheet.Cells(4 + i, 2) = _結帳日期
            'xlWorkSheet.Cells(4 + i, 3) = _種類
            'xlWorkSheet.Cells(4 + i, 4) = _傳票號碼
            'xlWorkSheet.Cells(4 + i, 5) = _會計科目及摘要
            'xlWorkSheet.Cells(4 + i, 6) = _支票編號
            'xlWorkSheet.Cells(4 + i, 7) = _收入金額405
            'xlWorkSheet.Cells(4 + i, 8) = _支出金額405
            'xlWorkSheet.Cells(4 + i, 9) = _餘額405
            'xlWorkSheet.Cells(4 + i, 10) = _收入金額409
            'xlWorkSheet.Cells(4 + i, 11) = _支出金額409
            'xlWorkSheet.Cells(4 + i, 12) = _餘額409
            'xlWorkSheet.Cells(4 + i, 13) = _廠商及備註
        Next
        xlWorkSheet.Range(xlWorkSheet.Cells(4, 1), xlWorkSheet.Cells(4 + data_dv.Count - 1, 13)).Value = arr
        If _年 <> "109"
            data.SelectCommand = "select top 1 餘額405, 餘額409 from 現金備查簿 where year(結帳日期)='" & _年 & "' + 1910 order by 序號 desc"
            data_dv = data.Select(New DataSourceSelectArguments)
            If data_dv.Count > 0
                xlWorkSheet.Cells(3, 9) = data_dv(0)(0)
                xlWorkSheet.Cells(3, 12) = data_dv(0)(1)
            End If
        End If
        xlWorkSheet.Range("E:E").WrapText = False
        xlWorkSheet.Range("M:M").WrapText = False
        'xlWorkSheet.Range("E:E").ShrinkToFit = True
        
        
        
        xlWorkSheet = CType(xlWorkBook.Sheets("每日數"), Excel.Worksheet)
        data.SelectCommand = "select * from 現金備查簿 where 結帳日期='" & _日報表結帳日期 & "' order by 傳票號碼"
        data_dv = data.Select(New DataSourceSelectArguments)
        Dim arr2(data_dv.Count + 2, 9) As Object
        For i = 0 To data_dv.Count - 1
            arr2(i, 0) = data_dv(i)(2).ToString()
            arr2(i, 0) = totaiwancalendar(arr2(i, 0))
            arr2(i, 1) = data_dv(i)(3).ToString()
            If arr2(i, 1) = "1"
                arr2(i, 1) = "收"
            Else If arr2(i, 1) = "2"
                arr2(i, 1) = "支"
            Else If arr2(i, 1) = "3"
                arr2(i, 1) = "現"
            End If
            arr2(i, 2) = data_dv(i)(4).ToString()
            arr2(i, 3) = data_dv(i)(5).ToString()
            arr2(i, 4) = data_dv(i)(6).ToString().Replace(Chr(13), "").Trim(Chr(10))
            arr2(i, 5) = data_dv(i)(7).ToString()
            arr2(i, 6) = data_dv(i)(8).ToString()
            arr2(i, 7) = data_dv(i)(10).ToString()
            arr2(i, 8) = data_dv(i)(11).ToString()
        Next
        
        arr2(data_dv.Count + 1, 3) = "合                         計"
        arr2(data_dv.Count + 1, 5) = "=SUM(G3:G" & (data_dv.Count + 3).ToString() & ")"
        arr2(data_dv.Count + 1, 6) = "=SUM(H3:H" & (data_dv.Count + 3).ToString() & ")"
        arr2(data_dv.Count + 1, 7) = "=SUM(I3:I" & (data_dv.Count + 3).ToString() & ")"
        arr2(data_dv.Count + 1, 8) = "=SUM(J3:J" & (data_dv.Count + 3).ToString() & ")"
        
        xlWorkSheet.Range(xlWorkSheet.Cells(3, 2), xlWorkSheet.Cells(3 + data_dv.Count + 1, 10)).Borders.LineStyle = 1
        xlWorkSheet.Range(xlWorkSheet.Cells(3, 2), xlWorkSheet.Cells(3 + data_dv.Count + 1, 10)).Value = arr2
        xlWorkSheet.Rows(data_dv.Count + 4).PageBreak = -4142
        xlWorkSheet.Range("E:E").WrapText = False
        
        
        
        xlWorkSheet = CType(xlWorkBook.Sheets("日報表"), Excel.Worksheet)
        xlWorkSheet.Activate()
        If Me.DropDownList1.SelectedItem IsNot Nothing
            xlWorkSheet.Cells(4, 2) = Me.DropDownList1.SelectedItem.Text
        End If
        Me.LabelC6.Text = Trim(Me.LabelC6.Text)
        Me.LabelD6.Text = Trim(Me.LabelD6.Text)
        Me.LabelE6.Text = Trim(Me.LabelE6.Text)
        Me.LabelC7.Text = Trim(Me.LabelC7.Text)
        Me.LabelD7.Text = Trim(Me.LabelD7.Text)
        Me.LabelE7.Text = Trim(Me.LabelE7.Text)
        Me.LabelC12.Text = Trim(Me.LabelC12.Text)
        Me.LabelE12.Text = Trim(Me.LabelE12.Text)
        Me.LabelC13.Text = Trim(Me.LabelC13.Text)
        Me.LabelE13.Text = Trim(Me.LabelE13.Text)
        Me.LabelC14.Text = Trim(Me.LabelC14.Text)
        Me.LabelE14.Text = Trim(Me.LabelE14.Text)
        Me.LabelC15.Text = Trim(Me.LabelC15.Text)
        Me.LabelE15.Text = Trim(Me.LabelE15.Text)
        xlWorkSheet.Cells(6, 3) = Me.LabelC6.Text
        xlWorkSheet.Cells(6, 4) = Me.LabelD6.Text
        xlWorkSheet.Cells(6, 5) = Me.LabelE6.Text
        xlWorkSheet.Cells(7, 3) = Me.LabelC7.Text
        xlWorkSheet.Cells(7, 4) = Me.LabelD7.Text
        xlWorkSheet.Cells(7, 5) = Me.LabelE7.Text
        xlWorkSheet.Cells(12, 3) = Me.LabelC12.Text
        xlWorkSheet.Cells(12, 5) = Me.LabelE12.Text
        xlWorkSheet.Cells(13, 3) = Me.LabelC13.Text
        xlWorkSheet.Cells(13, 5) = Me.LabelE13.Text
        xlWorkSheet.Cells(14, 3) = Me.LabelC14.Text
        xlWorkSheet.Cells(14, 5) = Me.LabelE14.Text
        xlWorkSheet.Cells(15, 3) = Me.LabelC15.Text
        xlWorkSheet.Cells(15, 5) = Me.LabelE15.Text
        If Me.LabelE12.Text = ""
            xlWorkSheet.Cells(12, 5) = "無開立傳票"
        End If
        If Me.LabelE13.Text = ""
            xlWorkSheet.Cells(13, 5) = "無開立傳票"
        End If
        If Me.LabelE14.Text = ""
            xlWorkSheet.Cells(14, 5) = "無開立傳票"
        End If
        If Me.LabelE15.Text = ""
            xlWorkSheet.Cells(15, 5) = "無開立傳票"
        End If
        
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
        data.SelectCommand = "select 結帳日期 from 日報表 where id='" & Me.DropDownList1.SelectedValue & "'"
        data_dv = data.Select(New DataSourceSelectArguments)
        Dim datadate As Date
        Dim downloadfilename
        If data_dv.Count > 0
            datadate = data_dv(0)(0)
            downloadfilename = "日報表 " & datadate.ToString("MMdd") & ".xls"
        Else
            downloadfilename = "日報表.xls"
        End If
        Response.AddHeader("Content-Disposition", "attachment;FileName=" + Uri.EscapeDataString(downloadfilename))
        Response.WriteFile(MyExcel)
        System.Web.HttpContext.Current.ApplicationInstance.CompleteRequest()
        Response.Flush()
        System.IO.File.Delete(MyExcel)
        Response.End()
    End Sub
    Protected Sub Button3_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button3.Click'日報表刪除
        Dim id As String = Me.DropDownList1.SelectedValue
        data.UpdateCommand = _
            "UPDATE " & _
                "現金備查簿 " & _
            "SET " & _
                "序號 = NULL, " & _
                "結帳日期 = NULL " & _
            "FROM 現金備查簿 INNER JOIN 日報表 " & _
                "ON (傳票號碼 BETWEEN C12 AND E12 " & _
                    "OR 傳票號碼 BETWEEN C13 AND E13 " & _
                    "OR 傳票號碼 BETWEEN C14 AND E14 " & _
                ") AND 現金備查簿.結帳日期 = 日報表.結帳日期 " & _
            "WHERE 日報表.id = " & id & ""
        data.Update()
        data.UpdateCommand = _
            "UPDATE " & _
                "分錄 " & _
            "SET " & _
                "序號 = NULL, " & _
                "結帳日期 = NULL " & _
            "FROM 分錄 INNER JOIN 日報表 " & _
                "ON (傳票號碼 BETWEEN C15 AND E15 " & _
                ") AND 分錄.結帳日期 = 日報表.結帳日期 " & _
            "WHERE 日報表.id = " & id & ""
        data.Update()
        
        data.DeleteCommand = "delete from 日報表 where id = " & id & ""
        data.Delete()
        generatedropdownlist1()
        DropDownList1_SelectedIndexChanged(sender, e)
    End Sub
    Protected Sub Download2(ByVal sender As Object, ByVal e As System.EventArgs)'月報表下載
        Dim _GUID As String = Guid.NewGuid().ToString("N")
        Dim MyExcel As String = MapPath(".\Excel\Temp\") & _GUID & ".xls"
        System.IO.File.Copy(MapPath(".\Excel\現金備查簿.xls"), MyExcel)
        Dim xlApp As New Excel.ApplicationClass()
        xlApp.DisplayAlerts = False
        xlApp.ScreenUpdating = false
        xlApp.EnableEvents = false
        Dim xlWorkBook As Excel.Workbook = xlApp.Workbooks.Open(MyExcel, 0, False, 5, "", "", True, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", False, False, 0, True, 1, 0)
        Dim xlWorkSheet As Excel.Worksheet
        
        xlWorkSheet = CType(xlWorkBook.Sheets("備查簿"), Excel.Worksheet)
        xlWorkSheet.Activate()
        Dim _年 As String = Me.結帳日期a.SelectedValue.Split("/")(0)
        xlWorkSheet.PageSetup.CenterHeader = Regex.Replace(xlWorkSheet.PageSetup.CenterHeader, "[0-9]{3}年度", _年 & "年度")
        Dim 結帳日期a As String = (CLng(Me.結帳日期a.SelectedValue.Split("/")(0)) + 1911).ToString() & "/" & Me.結帳日期a.SelectedValue.Split("/")(1) & "/" & Me.結帳日期a.SelectedValue.Split("/")(2)
        Dim 結帳日期b As String = (CLng(Me.結帳日期b.SelectedValue.Split("/")(0)) + 1911).ToString() & "/" & Me.結帳日期b.SelectedValue.Split("/")(1) & "/" & Me.結帳日期b.SelectedValue.Split("/")(2)
        data.SelectCommand = "select * from 現金備查簿 where 結帳日期 between '" & 結帳日期a & "' and '" & 結帳日期b & "' order by 傳票號碼"
        data_dv = data.Select(New DataSourceSelectArguments)
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
        
        xlWorkSheet.Rows(3).Delete()
        xlWorkSheet.Range("E:E").WrapText = False
        xlWorkSheet.Range("M:M").WrapText = False
        
        
        
        xlWorkSheet = xlWorkbook.Worksheets("每日數")
        Dim arr2(data_dv.Count + 2, 9) As Object
        For i = 0 To data_dv.Count - 1
            arr2(i, 0) = data_dv(i)(2).ToString()
            arr2(i, 0) = totaiwancalendar(arr2(i, 0))
            arr2(i, 1) = data_dv(i)(3).ToString()
            If arr2(i, 1) = "1"
                arr2(i, 1) = "收"
            Else If arr2(i, 1) = "2"
                arr2(i, 1) = "支"
            Else If arr2(i, 1) = "3"
                arr2(i, 1) = "現"
            End If
            arr2(i, 2) = data_dv(i)(4).ToString()
            arr2(i, 3) = data_dv(i)(5).ToString()
            arr2(i, 4) = data_dv(i)(6).ToString().Replace(Chr(13), "").Trim(Chr(10))
            arr2(i, 5) = data_dv(i)(7).ToString()
            arr2(i, 6) = data_dv(i)(8).ToString()
            arr2(i, 7) = data_dv(i)(10).ToString()
            arr2(i, 8) = data_dv(i)(11).ToString()
        Next
        
        arr2(data_dv.Count + 1, 3) = "合                         計"
        arr2(data_dv.Count + 1, 5) = "=SUM(G3:G" & (data_dv.Count + 3).ToString() & ")"
        arr2(data_dv.Count + 1, 6) = "=SUM(H3:H" & (data_dv.Count + 3).ToString() & ")"
        arr2(data_dv.Count + 1, 7) = "=SUM(I3:I" & (data_dv.Count + 3).ToString() & ")"
        arr2(data_dv.Count + 1, 8) = "=SUM(J3:J" & (data_dv.Count + 3).ToString() & ")"
        
        xlWorkSheet.Range(xlWorkSheet.Cells(3, 2), xlWorkSheet.Cells(3 + data_dv.Count + 1, 10)).Borders.LineStyle = 1
        xlWorkSheet.Range(xlWorkSheet.Cells(3, 2), xlWorkSheet.Cells(3 + data_dv.Count + 1, 10)).Value = arr2
        xlWorkSheet.Rows(data_dv.Count + 4).PageBreak = -4142
        xlWorkSheet.Range("E:E").WrapText = False
        
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
        Dim downloadfilename As String = "月報表.xls"
        Response.AddHeader("Content-Disposition", "attachment;FileName=" + Uri.EscapeDataString(downloadfilename))
        Response.WriteFile(MyExcel)
        System.Web.HttpContext.Current.ApplicationInstance.CompleteRequest()
        Response.Flush()
        System.IO.File.Delete(MyExcel)
        Response.End()
    End Sub
End Class
