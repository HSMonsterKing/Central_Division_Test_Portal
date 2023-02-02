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
Partial Class 保管品
    Inherits System.Web.UI.Page
    Dim con_14 As String = System.Web.Configuration.WebConfigurationManager.ConnectionStrings("ApplicationServices").ConnectionString
    Dim data As New SqlDataSource
    Dim data_dv As Data.DataView
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        data.ConnectionString = con_14
        Me.GridView1.PageSize = Me.Input1.Text
        Me.GridView2.PageSize = Me.Input1.Text
        
        If Not Page.IsPostBack Then
            DropDownList1_SelectedIndexChanged(sender, e)
            
            Me.GridView1.PageIndex = Int32.MaxValue
            Me.GridView2.PageIndex = Int32.MaxValue
        Else
        End If
    End Sub
    Protected Sub DropDownList1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Select Case Me.DropDownList1.SelectedValue
            Case "保證書明細表"
                Me.GridView1.columns(3).HeaderText = "保證書<br>名稱"
                Me.GridView1.columns(12).HeaderText = "保證書<br>保證期限"
                Me.Panel1_1.Visible = False
                Me.Panel1_2.Visible = False
                Me.Button4.Visible = True
                Me.GridView1.Visible = True
                Me.GridView2.Visible = False
            Case "定存單明細表"
                Me.GridView1.columns(3).HeaderText = "支票/定期<br>存單號碼"
                Me.GridView1.columns(12).HeaderText = "存單<br>保證期限"
                Me.Panel1_1.Visible = False
                Me.Panel1_2.Visible = False
                Me.Button4.Visible = True
                Me.GridView1.Visible = True
                Me.GridView2.Visible = False
            Case "保證書紀錄簿"
                Me.Panel1_1.Visible = True
                Me.Panel1_2.Visible = True
                Me.Button4.Visible = False
                Me.GridView1.Visible = False
                Me.GridView2.Visible = True
            Case "定存單紀錄簿"
                Me.Panel1_1.Visible = True
                Me.Panel1_2.Visible = True
                Me.Button4.Visible = False
                Me.GridView1.Visible = False
                Me.GridView2.Visible = True
        End Select
        
        Me.GridView1.PageIndex = Int32.MaxValue
    End Sub
    Protected Sub Insert(ByVal sender As Object, ByVal e As System.EventArgs)
        Select Case Me.DropDownList1.SelectedValue
            Case "保證書明細表", "定存單明細表"
                Me.SqlDataSource1.Insert()
            Case "保證書紀錄簿", "定存單紀錄簿"
                Me.SqlDataSource2.Insert()
        End Select
        
        Me.GridView1.PageIndex = Int32.MaxValue
    End Sub
    Protected Sub Update(ByVal sender As Object, ByVal e As System.EventArgs)
        Select Case Me.DropDownList1.SelectedValue
            Case "保證書明細表", "定存單明細表"
                For i = 0 To Me.GridView1.Rows.Count - 1
                    Dim Label0 As String = CType(Me.GridView1.Rows(i).FindControl("Label0"), Label).Text
                    Dim Label1 As String = CType(Me.GridView1.Rows(i).FindControl("Label1"), Label).Text
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
                    Dim TextBox13 As String = CType(Me.GridView1.Rows(i).FindControl("TextBox13"), TextBox).Text
                    Dim TextBox14 As String = CType(Me.GridView1.Rows(i).FindControl("TextBox14"), TextBox).Text
                    Dim TextBox15 As String = CType(Me.GridView1.Rows(i).FindControl("TextBox15"), TextBox).Text
                    Dim TextBox16 As String = CType(Me.GridView1.Rows(i).FindControl("TextBox16"), TextBox).Text
                    Dim TextBox17 As String = CType(Me.GridView1.Rows(i).FindControl("TextBox17"), TextBox).Text
                    Dim TextBox18 As String = CType(Me.GridView1.Rows(i).FindControl("TextBox18"), TextBox).Text
                    
                    'nvarchar
                    Label0 = Label0.Replace("'", "")
                    Label1 = Label1.Replace("'", "")
                    TextBox3 = TextBox3.Replace("'", "")
                    TextBox4 = TextBox4.Replace("'", "")
                    TextBox5 = TextBox5.Replace("'", "")
                    TextBox6 = TextBox6.Replace("'", "")
                    TextBox7 = TextBox7.Replace("'", "")
                    TextBox8 = TextBox8.Replace("'", "")
                    TextBox9 = TextBox9.Replace("'", "")
                    TextBox12 = TextBox12.Replace("'", "")
                    TextBox13 = TextBox13.Replace("'", "")
                    TextBox14 = TextBox14.Replace("'", "")
                    TextBox15 = TextBox15.Replace("'", "")
                    TextBox16 = TextBox16.Replace("'", "")
                    TextBox17 = TextBox17.Replace("'", "")
                    TextBox18 = TextBox18.Replace("'", "")
                    Label0 = "N'" & Label0 & "'"
                    Label1 = "N'" & Label1 & "'"
                    TextBox3 = "N'" & TextBox3 & "'"
                    TextBox4 = "N'" & TextBox4 & "'"
                    TextBox5 = "N'" & TextBox5 & "'"
                    TextBox6 = "N'" & TextBox6 & "'"
                    TextBox7 = "N'" & TextBox7 & "'"
                    TextBox8 = "N'" & TextBox8 & "'"
                    TextBox9 = "N'" & TextBox9 & "'"
                    TextBox12 = "N'" & TextBox12 & "'"
                    TextBox13 = "N'" & TextBox13 & "'"
                    TextBox14 = "N'" & TextBox14 & "'"
                    TextBox15 = "N'" & TextBox15 & "'"
                    TextBox16 = "N'" & TextBox16 & "'"
                    TextBox17 = "N'" & TextBox17 & "'"
                    TextBox18 = "N'" & TextBox18 & "'"
                    
                    'date
                    TextBox2 = TextBox2.Replace(".", "/")
                    TextBox2 = Regex.Replace(TextBox2, "[^0-9/]", "")
                    TextBox2 = Regex.Replace(TextBox2, "/{2,}", "/")
                    TextBox2 = taiwancalendarto(TextBox2)
                    TextBox2 = "N'" & TextBox2 & "'"
                    TextBox2 = "NULLIF(" & TextBox2 & ", '')"
                    
                    'bigint
                    TextBox10 = Regex.Replace(TextBox10, "[^0-9]", "")
                    TextBox10 = If(TextBox10 = "", "0", TextBox10)
                    TextBox11 = Regex.Replace(TextBox11, "[^0-9]", "")
                    TextBox11 = If(TextBox11 = "", "0", TextBox11)
                    
                    data.UpdateCommand = "UPDATE 保管品明細表 SET " & _
                    "日期 = " & TextBox2 & ", " & _
                    "保證書名稱或存單號碼 = " & TextBox3 & ", " & _
                    "收據編號 = " & TextBox4 & ", " & _
                    "國保收據編號 = " & TextBox5 & ", " & _
                    "戶名 = " & TextBox6 & ", " & _
                    "品名 = " & TextBox7 & ", " & _
                    "摘要 = " & TextBox8 & ", " & _
                    "單位 = " & TextBox9 & ", " & _
                    "數量 = " & TextBox10 & ", " & _
                    "金額 = " & TextBox11 & ", " & _
                    "保證書或存單保證期限 = " & TextBox12 & ", " & _
                    "廠商保證責任期限 = " & TextBox13 & ", " & _
                    "合約展延情形 = " & TextBox14 & ", " & _
                    "保管處 = " & TextBox15 & ", " & _
                    "承辦單位 = " & TextBox16 & ", " & _
                    "承辦人 = " & TextBox17 & ", " & _
                    "備考 = " & TextBox18 & " " & _
                    "WHERE id = " & Label0 & " "
                    data.Update()
                Next
                
                Me.GridView1.DataBind()
            Case "保證書紀錄簿", "定存單紀錄簿"
                For i = 0 To Me.GridView2.Rows.Count - 1
                    Dim Label0 As String = CType(Me.GridView2.Rows(i).FindControl("Label0"), Label).Text
                    Dim TextBox1 As String = CType(Me.GridView2.Rows(i).FindControl("TextBox1"), TextBox).Text
                    Dim TextBox2 As String = CType(Me.GridView2.Rows(i).FindControl("TextBox2"), TextBox).Text
                    Dim TextBox3 As String = CType(Me.GridView2.Rows(i).FindControl("TextBox3"), TextBox).Text
                    Dim TextBox4 As String = CType(Me.GridView2.Rows(i).FindControl("TextBox4"), TextBox).Text
                    Dim TextBox5 As String = CType(Me.GridView2.Rows(i).FindControl("TextBox5"), TextBox).Text
                    Dim TextBox6 As String = CType(Me.GridView2.Rows(i).FindControl("TextBox6"), TextBox).Text
                    Dim TextBox7 As String = CType(Me.GridView2.Rows(i).FindControl("TextBox7"), TextBox).Text
                    Dim TextBox8 As String = CType(Me.GridView2.Rows(i).FindControl("TextBox8"), TextBox).Text
                    Dim TextBox9 As String = CType(Me.GridView2.Rows(i).FindControl("TextBox9"), TextBox).Text
                    Dim TextBox10 As String = CType(Me.GridView2.Rows(i).FindControl("TextBox10"), TextBox).Text
                    
                    'nvarchar
                    Label0 = Label0.Replace("'", "")
                    TextBox1 = TextBox1.Replace("'", "")
                    TextBox4 = TextBox4.Replace("'", "")
                    TextBox5 = TextBox5.Replace("'", "")
                    TextBox9 = TextBox9.Replace("'", "")
                    TextBox10 = TextBox10.Replace("'", "")
                    Label0 = "N'" & Label0 & "'"
                    TextBox1 = "N'" & TextBox1 & "'"
                    TextBox4 = "N'" & TextBox4 & "'"
                    TextBox5 = "N'" & TextBox5 & "'"
                    TextBox9 = "N'" & TextBox9 & "'"
                    TextBox10 = "N'" & TextBox10 & "'"
                    
                    'date
                    TextBox2 = TextBox2.Replace(".", "/")
                    TextBox2 = Regex.Replace(TextBox2, "[^0-9/]", "")
                    TextBox2 = Regex.Replace(TextBox2, "/{2,}", "/")
                    TextBox2 = taiwancalendarto(TextBox2)
                    TextBox2 = "N'" & TextBox2 & "'"
                    TextBox2 = "NULLIF(" & TextBox2 & ", '')"
                    
                    'bit
                    TextBox3 = Regex.Replace(TextBox3, "[^收支]{1,}", "")
                    Select Case TextBox3
                        Case "收"
                            TextBox3 = "0"
                        Case "支"
                            TextBox3 = "1"
                        Case Else
                            TextBox3 = "收支"
                    End Select
                    
                    'bigint
                    TextBox6 = Regex.Replace(TextBox6, "[^0-9]", "")
                    TextBox6 = If(TextBox6 = "", "0", TextBox6)
                    TextBox7 = Regex.Replace(TextBox7, "[^0-9]", "")
                    TextBox7 = If(TextBox7 = "", "0", TextBox7)
                    TextBox8 = Regex.Replace(TextBox8, "[^0-9]", "")
                    TextBox8 = If(TextBox8 = "", "0", TextBox8)
                    
                    data.UpdateCommand = "UPDATE 保管品紀錄簿 SET " & _
                    "序號 = " & TextBox1 & ", " & _
                    "日期 = " & TextBox2 & ", " & _
                    "收支 = " & TextBox3 & ", " & _
                    "國保收據編號 = " & TextBox4 & ", " & _
                    "摘要 = " & TextBox5 & ", " & _
                    "收入金額 = " & TextBox6 & ", " & _
                    "支出金額 = " & TextBox7 & ", " & _
                    "餘額 = " & TextBox8 & ", " & _
                    "戶名 = " & TextBox9 & ", " & _
                    "品名 = " & TextBox10 & " " & _
                    "WHERE id = " & Label0 & " "
                    data.Update()
                Next
                
                '重算餘額
                data.UpdateCommand = _
                "WITH CTE AS " & _
                "(SELECT *, " & _
                    "(SELECT TOP 1 (CASE WHEN 收入金額 = 0 AND 支出金額 = 0 THEN 餘額 ELSE 0 END) FROM 保管品紀錄簿 WHERE 種類 = 0 ORDER BY 序號) " & _
                    "+ " & _
                    "(SUM(收入金額 - 支出金額) OVER (ORDER BY 序號 " & _
                    "ROWS BETWEEN UNBOUNDED PRECEDING AND CURRENT ROW)) " & _
                    "AS RunningTotal " & _
                "FROM 保管品紀錄簿 WHERE 種類 = 0) " & _
                "UPDATE CTE SET 餘額 = RunningTotal"
                data.Update()
                
                data.UpdateCommand = _
                "WITH CTE AS " & _
                "(SELECT *, " & _
                    "(SELECT TOP 1 (CASE WHEN 收入金額 = 0 AND 支出金額 = 0 THEN 餘額 ELSE 0 END) FROM 保管品紀錄簿 WHERE 種類 = 1 ORDER BY 序號) " & _
                    "+ " & _
                    "(SUM(收入金額 - 支出金額) OVER (ORDER BY 序號 " & _
                    "ROWS BETWEEN UNBOUNDED PRECEDING AND CURRENT ROW)) " & _
                    "AS RunningTotal " & _
                "FROM 保管品紀錄簿 WHERE 種類 = 1) " & _
                "UPDATE CTE SET 餘額 = RunningTotal"
                data.Update()
                
                Me.GridView2.DataBind()
        End Select
    End Sub
    Protected Sub Download(ByVal sender As Object, ByVal e As System.EventArgs)
        Select Case Me.DropDownList1.SelectedValue
            Case "保證書明細表"
                Dim _GUID As String = Guid.NewGuid().ToString("N")
                Dim MyExcel As String = MapPath(".\Excel\Temp\") & _GUID & ".xls"
                System.IO.File.Copy(MapPath(".\Excel\保證書明細表.xls"), MyExcel)
                Dim xlApp As New Excel.ApplicationClass()
                xlApp.DisplayAlerts = False
                xlApp.ScreenUpdating = false
                xlApp.EnableEvents = false
                Dim xlWorkBook As Excel.Workbook = xlApp.Workbooks.Open(MyExcel, 0, False, 5, "", "", True, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", False, False, 0, True, 1, 0)
                Dim xlWorkSheet As Excel.Worksheet
                
                xlWorkSheet = CType(xlWorkBook.Sheets("保證書"), Excel.Worksheet)
                data.SelectCommand = _
                    "SELECT * " & _
                    "FROM 保管品明細表 " & _
                    "WHERE (種類=0 AND 已支出=0) " & _
                    "ORDER BY " & _
                    "CASE WHEN 日期 IS NULL THEN 1 ELSE 0 END, 日期, " & _
                    "CASE WHEN 收據編號 = '' THEN 1 ELSE 0 END, 收據編號, " & _
                    "CASE WHEN 國保收據編號 = '' THEN 1 ELSE 0 END, 國保收據編號"
                data_dv = data.Select(New DataSourceSelectArguments)
                Dim arr(data_dv.Count, 18) As Object
                For i = 0 To data_dv.Count - 1
                    Dim _序號 As String = "=ROW()-6"
                    Dim _日期 As String = data_dv(i)("日期").ToString()
                    _日期 = totaiwancalendar(_日期)
                    _日期 = _日期.Replace("/", ".")
                    Dim _保證書名稱 As String = data_dv(i)("保證書名稱或存單號碼").ToString()
                    Dim _收據編號 As String = data_dv(i)("收據編號").ToString()
                    Dim _國保收據編號 As String = data_dv(i)("國保收據編號").ToString()
                    Dim _戶名 As String = data_dv(i)("戶名").ToString()
                    Dim _品名 As String = data_dv(i)("品名").ToString()
                    Dim _摘要 As String = data_dv(i)("摘要").ToString()
                    Dim _單位 As String = data_dv(i)("單位").ToString()
                    Dim _數量 As String = data_dv(i)("數量").ToString()
                    Dim _金額 As String = data_dv(i)("金額").ToString()
                    Dim _保證書保證期限 As String = data_dv(i)("保證書或存單保證期限").ToString()
                    Dim _廠商保證責任期限 As String = data_dv(i)("廠商保證責任期限").ToString()
                    Dim _合約展延情形 As String = data_dv(i)("合約展延情形").ToString()
                    Dim _保管處 As String = data_dv(i)("保管處").ToString()
                    Dim _承辦單位 As String = data_dv(i)("承辦單位").ToString()
                    Dim _承辦人 As String = data_dv(i)("承辦人").ToString()
                    Dim _備考 As String = data_dv(i)("備考").ToString()
                    arr(i, 0) = _序號'.Replace(Chr(13), "").Trim(Chr(10))
                    arr(i, 1) = _日期
                    arr(i, 2) = _保證書名稱
                    arr(i, 3) = _收據編號
                    arr(i, 4) = _國保收據編號
                    arr(i, 5) = _戶名
                    arr(i, 6) = _品名
                    arr(i, 7) = _摘要
                    arr(i, 8) = _單位
                    arr(i, 9) = _數量
                    arr(i, 10) = _金額
                    arr(i, 11) = _保證書保證期限
                    arr(i, 12) = _廠商保證責任期限
                    arr(i, 13) = _合約展延情形
                    arr(i, 14) = _保管處
                    arr(i, 15) = _承辦單位
                    arr(i, 16) = _承辦人
                    arr(i, 17) = _備考
                Next
                If data_dv.Count > 1
                    xlWorkSheet.Range(xlWorkSheet.Cells(8, 1), xlWorkSheet.Cells(7 + data_dv.Count - 1, 18)).EntireRow.Insert()
                End If
                xlWorkSheet.Range(xlWorkSheet.Cells(7, 1), xlWorkSheet.Cells(7 + data_dv.Count - 1, 18)).Value = arr
                
                xlWorkSheet.Cells(3, 1).Value = "中華民國 " & (Today().Year - 1911).ToString() & " 年 " & Today().Month.ToString() &" 月 " & Date.DaysInMonth(Today().Year, Today().Month) & " 日"
                
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
                downloadfilename = "保證書明細表.xls"
                Response.AddHeader("Content-Disposition", "attachment;FileName=" + Uri.EscapeDataString(downloadfilename))
                Response.WriteFile(MyExcel)
                System.Web.HttpContext.Current.ApplicationInstance.CompleteRequest()
                Response.Flush()
                System.IO.File.Delete(MyExcel)
                Response.End()
            Case "定存單明細表"
                Dim _GUID As String = Guid.NewGuid().ToString("N")
                Dim MyExcel As String = MapPath(".\Excel\Temp\") & _GUID & ".xls"
                System.IO.File.Copy(MapPath(".\Excel\定存單明細表.xls"), MyExcel)
                Dim xlApp As New Excel.ApplicationClass()
                xlApp.DisplayAlerts = False
                xlApp.ScreenUpdating = false
                xlApp.EnableEvents = false
                Dim xlWorkBook As Excel.Workbook = xlApp.Workbooks.Open(MyExcel, 0, False, 5, "", "", True, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", False, False, 0, True, 1, 0)
                Dim xlWorkSheet As Excel.Worksheet
                
                xlWorkSheet = CType(xlWorkBook.Sheets("定存單"), Excel.Worksheet)
                data.SelectCommand = _
                    "SELECT * " & _
                    "FROM 保管品明細表 " & _
                    "WHERE (種類=1 AND 已支出=0) " & _
                    "ORDER BY " & _
                    "CASE WHEN 日期 IS NULL THEN 1 ELSE 0 END, 日期, " & _
                    "CASE WHEN 收據編號 = '' THEN 1 ELSE 0 END, 收據編號, " & _
                    "CASE WHEN 國保收據編號 = '' THEN 1 ELSE 0 END, 國保收據編號"
                data_dv = data.Select(New DataSourceSelectArguments)
                Dim arr(data_dv.Count, 18) As Object
                For i = 0 To data_dv.Count - 1
                    Dim _序號 As String = "=ROW()-6"
                    Dim _日期 As String = data_dv(i)("日期").ToString()
                    _日期 = totaiwancalendar(_日期)
                    _日期 = _日期.Replace("/", ".")
                    Dim _收據編號 As String = data_dv(i)("收據編號").ToString()
                    Dim _國保收據編號 As String = data_dv(i)("國保收據編號").ToString()
                    Dim _存單號碼 As String = data_dv(i)("保證書名稱或存單號碼").ToString()
                    Dim _戶名 As String = data_dv(i)("戶名").ToString()
                    Dim _品名 As String = data_dv(i)("品名").ToString()
                    Dim _摘要 As String = data_dv(i)("摘要").ToString()
                    Dim _單位 As String = data_dv(i)("單位").ToString()
                    Dim _數量 As String = data_dv(i)("數量").ToString()
                    Dim _金額 As String = data_dv(i)("金額").ToString()
                    Dim _存單保證期限 As String = data_dv(i)("保證書或存單保證期限").ToString()
                    Dim _廠商保證責任期限 As String = data_dv(i)("廠商保證責任期限").ToString()
                    Dim _合約展延情形 As String = data_dv(i)("合約展延情形").ToString()
                    Dim _保管處 As String = data_dv(i)("保管處").ToString()
                    Dim _承辦單位 As String = data_dv(i)("承辦單位").ToString()
                    Dim _承辦人 As String = data_dv(i)("承辦人").ToString()
                    Dim _備考 As String = data_dv(i)("備考").ToString()
                    arr(i, 0) = _序號'.Replace(Chr(13), "").Trim(Chr(10))
                    arr(i, 1) = _日期
                    arr(i, 2) = _收據編號
                    arr(i, 3) = _國保收據編號
                    arr(i, 4) = _存單號碼
                    arr(i, 5) = _戶名
                    arr(i, 6) = _品名
                    arr(i, 7) = _摘要
                    arr(i, 8) = _單位
                    arr(i, 9) = _數量
                    arr(i, 10) = _金額
                    arr(i, 11) = _存單保證期限
                    arr(i, 12) = _廠商保證責任期限
                    arr(i, 13) = _合約展延情形
                    arr(i, 14) = _保管處
                    arr(i, 15) = _承辦單位
                    arr(i, 16) = _承辦人
                    arr(i, 17) = _備考
                Next
                If data_dv.Count > 1
                    xlWorkSheet.Range(xlWorkSheet.Cells(8, 1), xlWorkSheet.Cells(7 + data_dv.Count - 1, 18)).EntireRow.Insert()
                End If
                xlWorkSheet.Range(xlWorkSheet.Cells(7, 1), xlWorkSheet.Cells(7 + data_dv.Count - 1, 18)).Value = arr
                
                xlWorkSheet.Cells(3, 1).Value = "中華民國 " & (Today().Year - 1911).ToString() & " 年 " & Today().Month.ToString() &" 月 " & Date.DaysInMonth(Today().Year, Today().Month) & " 日"
                
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
                downloadfilename = "定存單明細表.xls"
                Response.AddHeader("Content-Disposition", "attachment;FileName=" + Uri.EscapeDataString(downloadfilename))
                Response.WriteFile(MyExcel)
                System.Web.HttpContext.Current.ApplicationInstance.CompleteRequest()
                Response.Flush()
                System.IO.File.Delete(MyExcel)
                Response.End()
            Case "保證書紀錄簿", "定存單紀錄簿"
        End Select
    End Sub
    Public Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        'GridView1變更時，更新GridView3
        '例如刪除一行資料時，重新計算總數量和總金額
        Me.GridView3.DataBind()
    End Sub
    Protected Sub GridView1_RowCommand(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles GridView1.RowCommand
        Select Case Me.DropDownList1.SelectedValue
            Case "保證書明細表", "定存單明細表"
                If e.CommandName = "收入"
                    Update(sender, e)
                    
                    Dim CurrentRow = Me.GridView1.Rows(e.CommandSource.NamingContainer.RowIndex)
                    Dim Label0 As String = CType(CurrentRow.FindControl("Label0"), Label).Text
                    
                    Label0 = Label0.Replace("'", "")
                    Label0 = "N'" & Label0 & "'"
                    
                    data.UpdateCommand = _
                    "SET XACT_ABORT ON " & _
                    "BEGIN TRANSACTION " & _
                    "IF (SELECT 已收入 FROM 保管品明細表 WHERE id = " & Label0 & ") = 0 " & _
                        "BEGIN " & _
                            "UPDATE 保管品明細表 SET 已收入 = 1 WHERE id = " & Label0 & " " & _
                            "INSERT INTO 保管品紀錄簿 " & _
                            "(保管品明細表id, 種類, 序號, 日期, 收支, 國保收據編號, 摘要, 收入金額, 支出金額, 餘額, 戶名, 品名) " & _
                            "(SELECT id, 種類, (SELECT ISNULL(MAX(序號), 0)+1 FROM 保管品紀錄簿 WHERE 種類 = 保管品明細表.種類), 日期, 0, 國保收據編號, 摘要, 金額, 0, (SELECT ISNULL((SELECT 餘額 FROM 保管品紀錄簿 WHERE 序號 = (SELECT ISNULL(MAX(序號), 0) FROM 保管品紀錄簿 WHERE 種類 = 保管品明細表.種類) AND 種類 = 保管品明細表.種類), 0)) + 金額, 戶名, 品名 FROM 保管品明細表 WHERE id = " & Label0 & ") " & _
                        "END " & _
                    "COMMIT TRANSACTION " & _
                    "SET XACT_ABORT OFF "
                    data.Update()
                    
                    Me.GridView1.DataBind()
                Else If e.CommandName = "支出"
                    Update(sender, e)
                    
                    Dim CurrentRow = Me.GridView1.Rows(e.CommandSource.NamingContainer.RowIndex)
                    Dim Label0 As String = CType(CurrentRow.FindControl("Label0"), Label).Text
                    
                    Label0 = Label0.Replace("'", "")
                    Label0 = "N'" & Label0 & "'"
                    
                    data.UpdateCommand = _
                    "SET XACT_ABORT ON " & _
                    "BEGIN TRANSACTION " & _
                    "IF (SELECT 已支出 FROM 保管品明細表 WHERE id = " & Label0 & ") = 0 " & _
                        "BEGIN " & _
                            "UPDATE 保管品明細表 SET 已支出 = 1 WHERE id = " & Label0 & " " & _
                            "INSERT INTO 保管品紀錄簿 " & _
                            "(保管品明細表id, 種類, 序號, 日期, 收支, 國保收據編號, 摘要, 收入金額, 支出金額, 餘額, 戶名, 品名) " & _
                            "(SELECT id, 種類, (SELECT ISNULL(MAX(序號), 0)+1 FROM 保管品紀錄簿 WHERE 種類 = 保管品明細表.種類), NULL, 1, 國保收據編號, (SELECT ISNULL((SELECT MAX(摘要) FROM 保管品紀錄簿 WHERE 保管品明細表id = " & Label0 & "), 摘要)), 0, 金額, (SELECT ISNULL((SELECT 餘額 FROM 保管品紀錄簿 WHERE 序號 = (SELECT ISNULL(MAX(序號), 0) FROM 保管品紀錄簿 WHERE 種類 = 保管品明細表.種類) AND 種類 = 保管品明細表.種類), 0)) - 金額, (SELECT ISNULL((SELECT MAX(戶名) FROM 保管品紀錄簿 WHERE 保管品明細表id = " & Label0 & "), 戶名)), (SELECT ISNULL((SELECT MAX(品名) FROM 保管品紀錄簿 WHERE 保管品明細表id = " & Label0 & "), 品名)) FROM 保管品明細表 WHERE id = " & Label0 & ") " & _
                        "END " & _
                    "COMMIT TRANSACTION " & _
                    "SET XACT_ABORT OFF "
                    data.Update()
                    
                    Me.GridView1.DataBind()
                End If
            Case "保證書紀錄簿", "定存單紀錄簿"
        End Select
    End Sub
    Protected Sub GridView2_RowCommand(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles GridView2.RowCommand
        Select Case Me.DropDownList1.SelectedValue
            Case "保證書明細表", "定存單明細表"
            Case "保證書紀錄簿"
                If e.CommandName = "轉明細表"
                    Update(sender, e)
                    
                    Dim CurrentRow = Me.GridView2.Rows(e.CommandSource.NamingContainer.RowIndex)
                    Dim Label0 As String = CType(CurrentRow.FindControl("Label0"), Label).Text
                    
                    data.UpdateCommand = _
                    "SET XACT_ABORT ON " & _
                    "BEGIN TRANSACTION " & _
                    "DECLARE @TempTable TABLE (id int) " & _
                    "IF (SELECT COUNT(*) FROM 保管品明細表 WHERE id = (SELECT 保管品明細表id FROM 保管品紀錄簿 WHERE id = " & Label0 & ")) = 1 " & _
                        "BEGIN " & _
                            "UPDATE 保管品明細表 SET 已支出 = 0 WHERE id = (SELECT 保管品明細表id FROM 保管品紀錄簿 WHERE id = " & Label0 & ") " & _
                        "END " & _
                    "ELSE " & _
                        "BEGIN " & _
                            "INSERT INTO 保管品明細表 " & _
                            "(種類, 已收入, 已支出, 日期, 國保收據編號, 摘要, 單位, 數量, 金額, 合約展延情形, 戶名, 品名) " & _
                            "OUTPUT INSERTED.id INTO @TempTable (id) " & _
                            "(SELECT 0, " & _
                                "1, " & _
                                "0, " & _
                                "日期, " & _
                                "國保收據編號, " & _
                                "摘要, " & _
                                "N'包', " & _
                                "1, " & _
                                "CASE WHEN 收支 = 0 THEN 收入金額 ELSE 支出金額 END, " & _
                                "N'無', " & _
                                "戶名, " & _
                                "品名 " & _
                            "FROM 保管品紀錄簿 WHERE id = N'" & Label0 & "') " & _
                            "UPDATE 保管品紀錄簿 SET 保管品明細表id = (SELECT MAX(id) FROM @TempTable) WHERE id = N'" & Label0 & "' " & _
                        "END " & _
                    "COMMIT TRANSACTION " & _
                    "SET XACT_ABORT OFF "
                    data.Update()
                Else If e.CommandName = "支出"
                    Update(sender, e)
                    
                    Dim CurrentRow = Me.GridView2.Rows(e.CommandSource.NamingContainer.RowIndex)
                    Dim Label0 As String = CType(CurrentRow.FindControl("Label0"), Label).Text
                    
                    data.UpdateCommand = _
                        "INSERT INTO 保管品紀錄簿 " & _
                        "(保管品明細表id, 種類, 序號, 日期, 收支, 國保收據編號, 摘要, 收入金額, 支出金額, 餘額, 戶名, 品名) " & _
                        "(SELECT 保管品明細表id, " & _
                            "0, " & _
                            "(SELECT ISNULL(MAX(序號), 0)+1 FROM 保管品紀錄簿 WHERE 種類 = 0), " & _
                            "NULL, " & _
                            "1, " & _
                            "國保收據編號, " & _
                            "摘要, " & _
                            "0, " & _
                            "收入金額, " & _
                            "(SELECT ISNULL((SELECT 餘額 FROM 保管品紀錄簿 WHERE 序號 = (SELECT ISNULL(MAX(序號), 0) FROM 保管品紀錄簿 WHERE 種類 = 0) AND 種類 = 0), 0)) - 收入金額, " & _
                            "戶名, " & _
                            "品名 " & _
                        "FROM 保管品紀錄簿 WHERE id = N'" & Label0 & "')"
                    data.Update()
                    
                    Me.GridView2.DataBind()
                Else If e.CommandName = "下載"
                    Update(sender, e)
                    
                    Dim CurrentRow = Me.GridView2.Rows(e.CommandSource.NamingContainer.RowIndex)
                    Dim TextBox1 As String = CType(CurrentRow.FindControl("TextBox1"), TextBox).Text
                    
                    Dim _GUID As String = Guid.NewGuid().ToString("N")
                    Dim MyExcel As String = MapPath(".\Excel\Temp\") & _GUID & ".xls"
                    System.IO.File.Copy(MapPath(".\Excel\保證書紀錄簿.xls"), MyExcel)
                    Dim xlApp As New Excel.ApplicationClass()
                    xlApp.DisplayAlerts = False
                    xlApp.ScreenUpdating = false
                    xlApp.EnableEvents = false
                    Dim xlWorkBook As Excel.Workbook = xlApp.Workbooks.Open(MyExcel, 0, False, 5, "", "", True, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", False, False, 0, True, 1, 0)
                    Dim xlWorkSheet As Excel.Worksheet
                    
                    xlWorkSheet = CType(xlWorkBook.Sheets("保證書"), Excel.Worksheet)
                    data.SelectCommand = _
                        "SELECT TOP 38" & _
                        "YEAR(日期) - 1911 AS 年, " & _
                        "MONTH(日期) AS 月, " & _
                        "DAY(日期) AS 日, " & _
                        "CASE WHEN 收支 = 0 THEN '收' ELSE '支' END AS 收支, " & _
                        "國保收據編號, " & _
                        "摘要, " & _
                        "SUBSTRING(REVERSE(CAST(收入金額 AS VARCHAR(20))),9,12) AS 收入金額_億, " & _
                        "SUBSTRING(REVERSE(CAST(收入金額 AS VARCHAR(20))),8,1) AS 收入金額_千萬, " & _
                        "SUBSTRING(REVERSE(CAST(收入金額 AS VARCHAR(20))),7,1) AS 收入金額_百萬, " & _
                        "SUBSTRING(REVERSE(CAST(收入金額 AS VARCHAR(20))),6,1) AS 收入金額_十萬, " & _
                        "SUBSTRING(REVERSE(CAST(收入金額 AS VARCHAR(20))),5,1) AS 收入金額_萬, " & _
                        "SUBSTRING(REVERSE(CAST(收入金額 AS VARCHAR(20))),4,1) AS 收入金額_千, " & _
                        "SUBSTRING(REVERSE(CAST(收入金額 AS VARCHAR(20))),3,1) AS 收入金額_百, " & _
                        "SUBSTRING(REVERSE(CAST(收入金額 AS VARCHAR(20))),2,1) AS 收入金額_十, " & _
                        "SUBSTRING(REVERSE(CAST(收入金額 AS VARCHAR(20))),1,1) AS 收入金額_個, " & _
                        "SUBSTRING(REVERSE(CAST(支出金額 AS VARCHAR(20))),9,12) AS 支出金額_億, " & _
                        "SUBSTRING(REVERSE(CAST(支出金額 AS VARCHAR(20))),8,1) AS 支出金額_千萬, " & _
                        "SUBSTRING(REVERSE(CAST(支出金額 AS VARCHAR(20))),7,1) AS 支出金額_百萬, " & _
                        "SUBSTRING(REVERSE(CAST(支出金額 AS VARCHAR(20))),6,1) AS 支出金額_十萬, " & _
                        "SUBSTRING(REVERSE(CAST(支出金額 AS VARCHAR(20))),5,1) AS 支出金額_萬, " & _
                        "SUBSTRING(REVERSE(CAST(支出金額 AS VARCHAR(20))),4,1) AS 支出金額_千, " & _
                        "SUBSTRING(REVERSE(CAST(支出金額 AS VARCHAR(20))),3,1) AS 支出金額_百, " & _
                        "SUBSTRING(REVERSE(CAST(支出金額 AS VARCHAR(20))),2,1) AS 支出金額_十, " & _
                        "SUBSTRING(REVERSE(CAST(支出金額 AS VARCHAR(20))),1,1) AS 支出金額_個, " & _
                        "SUBSTRING(REVERSE(CAST(餘額 AS VARCHAR(20))),9,12) AS 餘額_億, " & _
                        "SUBSTRING(REVERSE(CAST(餘額 AS VARCHAR(20))),8,1) AS 餘額_千萬, " & _
                        "SUBSTRING(REVERSE(CAST(餘額 AS VARCHAR(20))),7,1) AS 餘額_百萬, " & _
                        "SUBSTRING(REVERSE(CAST(餘額 AS VARCHAR(20))),6,1) AS 餘額_十萬, " & _
                        "SUBSTRING(REVERSE(CAST(餘額 AS VARCHAR(20))),5,1) AS 餘額_萬, " & _
                        "SUBSTRING(REVERSE(CAST(餘額 AS VARCHAR(20))),4,1) AS 餘額_千, " & _
                        "SUBSTRING(REVERSE(CAST(餘額 AS VARCHAR(20))),3,1) AS 餘額_百, " & _
                        "SUBSTRING(REVERSE(CAST(餘額 AS VARCHAR(20))),2,1) AS 餘額_十, " & _
                        "SUBSTRING(REVERSE(CAST(餘額 AS VARCHAR(20))),1,1) AS 餘額_個, " & _
                        "戶名, " & _
                        "品名 " & _
                        "FROM 保管品紀錄簿 " & _
                        "WHERE (種類 = 0 AND 序號 >= " & If(TextBox1 = "", "1", TextBox1) & ") " & _
                        "ORDER BY 序號"
                    data_dv = data.Select(New DataSourceSelectArguments)
                    Dim arr(19, 51) As Object
                    For i = 0 To Math.Min(18, data_dv.Count - 1)
                        Dim _年 As String = data_dv(i)("年").ToString()
                        Dim _月 As String = data_dv(i)("月").ToString()
                        Dim _日 As String = data_dv(i)("日").ToString()
                        Dim _收支 As String = data_dv(i)("收支").ToString()
                        Dim _國保收據編號 As String = data_dv(i)("國保收據編號").ToString()
                        Dim _摘要 As String = data_dv(i)("摘要").ToString()
                        Dim _收入金額_億 As String = data_dv(i)("收入金額_億").ToString()
                        Dim _收入金額_千萬 As String = data_dv(i)("收入金額_千萬").ToString()
                        Dim _收入金額_百萬 As String = data_dv(i)("收入金額_百萬").ToString()
                        Dim _收入金額_十萬 As String = data_dv(i)("收入金額_十萬").ToString()
                        Dim _收入金額_萬 As String = data_dv(i)("收入金額_萬").ToString()
                        Dim _收入金額_千 As String = data_dv(i)("收入金額_千").ToString()
                        Dim _收入金額_百 As String = data_dv(i)("收入金額_百").ToString()
                        Dim _收入金額_十 As String = data_dv(i)("收入金額_十").ToString()
                        Dim _收入金額_個 As String = data_dv(i)("收入金額_個").ToString()
                        Dim _支出金額_億 As String = data_dv(i)("支出金額_億").ToString()
                        Dim _支出金額_千萬 As String = data_dv(i)("支出金額_千萬").ToString()
                        Dim _支出金額_百萬 As String = data_dv(i)("支出金額_百萬").ToString()
                        Dim _支出金額_十萬 As String = data_dv(i)("支出金額_十萬").ToString()
                        Dim _支出金額_萬 As String = data_dv(i)("支出金額_萬").ToString()
                        Dim _支出金額_千 As String = data_dv(i)("支出金額_千").ToString()
                        Dim _支出金額_百 As String = data_dv(i)("支出金額_百").ToString()
                        Dim _支出金額_十 As String = data_dv(i)("支出金額_十").ToString()
                        Dim _支出金額_個 As String = data_dv(i)("支出金額_個").ToString()
                        Dim _餘額_億 As String = data_dv(i)("餘額_億").ToString()
                        Dim _餘額_千萬 As String = data_dv(i)("餘額_千萬").ToString()
                        Dim _餘額_百萬 As String = data_dv(i)("餘額_百萬").ToString()
                        Dim _餘額_十萬 As String = data_dv(i)("餘額_十萬").ToString()
                        Dim _餘額_萬 As String = data_dv(i)("餘額_萬").ToString()
                        Dim _餘額_千 As String = data_dv(i)("餘額_千").ToString()
                        Dim _餘額_百 As String = data_dv(i)("餘額_百").ToString()
                        Dim _餘額_十 As String = data_dv(i)("餘額_十").ToString()
                        Dim _餘額_個 As String = data_dv(i)("餘額_個").ToString()
                        Dim _戶名 As String = data_dv(i)("戶名").ToString()
                        Dim _品名 As String = data_dv(i)("品名").ToString()
                        arr(i, 0) = _年
                        arr(i, 1) = _月
                        arr(i, 3) = _日
                        arr(i, 5) = _收支
                        arr(i, 7) = _國保收據編號
                        arr(i, 9) = _摘要
                        arr(i, 18) = _收入金額_億
                        arr(i, 19) = _收入金額_千萬
                        arr(i, 20) = _收入金額_百萬
                        arr(i, 21) = _收入金額_十萬
                        arr(i, 22) = _收入金額_萬
                        arr(i, 23) = _收入金額_千
                        arr(i, 24) = _收入金額_百
                        arr(i, 25) = _收入金額_十
                        arr(i, 26) = _收入金額_個
                        arr(i, 27) = _支出金額_億
                        arr(i, 28) = _支出金額_千萬
                        arr(i, 29) = _支出金額_百萬
                        arr(i, 30) = _支出金額_十萬
                        arr(i, 31) = _支出金額_萬
                        arr(i, 32) = _支出金額_千
                        arr(i, 33) = _支出金額_百
                        arr(i, 34) = _支出金額_十
                        arr(i, 35) = _支出金額_個
                        arr(i, 36) = _餘額_億
                        arr(i, 37) = _餘額_千萬
                        arr(i, 38) = _餘額_百萬
                        arr(i, 39) = _餘額_十萬
                        arr(i, 40) = _餘額_萬
                        arr(i, 41) = _餘額_千
                        arr(i, 42) = _餘額_百
                        arr(i, 43) = _餘額_十
                        arr(i, 44) = _餘額_個
                        arr(i, 47) = _戶名
                        arr(i, 49) = _品名
                    Next
                    xlWorkSheet.Range(xlWorkSheet.Cells(11, 1), xlWorkSheet.Cells(29, 51)).Value = arr
                    Array.Clear(arr, 0, arr.Length)
                    For i = 0 To data_dv.Count - 1 - 19
                        Dim _年 As String = data_dv(i + 19)("年").ToString()
                        Dim _月 As String = data_dv(i + 19)("月").ToString()
                        Dim _日 As String = data_dv(i + 19)("日").ToString()
                        Dim _收支 As String = data_dv(i + 19)("收支").ToString()
                        Dim _國保收據編號 As String = data_dv(i + 19)("國保收據編號").ToString()
                        Dim _摘要 As String = data_dv(i + 19)("摘要").ToString()
                        Dim _收入金額_億 As String = data_dv(i + 19)("收入金額_億").ToString()
                        Dim _收入金額_千萬 As String = data_dv(i + 19)("收入金額_千萬").ToString()
                        Dim _收入金額_百萬 As String = data_dv(i + 19)("收入金額_百萬").ToString()
                        Dim _收入金額_十萬 As String = data_dv(i + 19)("收入金額_十萬").ToString()
                        Dim _收入金額_萬 As String = data_dv(i + 19)("收入金額_萬").ToString()
                        Dim _收入金額_千 As String = data_dv(i + 19)("收入金額_千").ToString()
                        Dim _收入金額_百 As String = data_dv(i + 19)("收入金額_百").ToString()
                        Dim _收入金額_十 As String = data_dv(i + 19)("收入金額_十").ToString()
                        Dim _收入金額_個 As String = data_dv(i + 19)("收入金額_個").ToString()
                        Dim _支出金額_億 As String = data_dv(i + 19)("支出金額_億").ToString()
                        Dim _支出金額_千萬 As String = data_dv(i + 19)("支出金額_千萬").ToString()
                        Dim _支出金額_百萬 As String = data_dv(i + 19)("支出金額_百萬").ToString()
                        Dim _支出金額_十萬 As String = data_dv(i + 19)("支出金額_十萬").ToString()
                        Dim _支出金額_萬 As String = data_dv(i + 19)("支出金額_萬").ToString()
                        Dim _支出金額_千 As String = data_dv(i + 19)("支出金額_千").ToString()
                        Dim _支出金額_百 As String = data_dv(i + 19)("支出金額_百").ToString()
                        Dim _支出金額_十 As String = data_dv(i + 19)("支出金額_十").ToString()
                        Dim _支出金額_個 As String = data_dv(i + 19)("支出金額_個").ToString()
                        Dim _餘額_億 As String = data_dv(i + 19)("餘額_億").ToString()
                        Dim _餘額_千萬 As String = data_dv(i + 19)("餘額_千萬").ToString()
                        Dim _餘額_百萬 As String = data_dv(i + 19)("餘額_百萬").ToString()
                        Dim _餘額_十萬 As String = data_dv(i + 19)("餘額_十萬").ToString()
                        Dim _餘額_萬 As String = data_dv(i + 19)("餘額_萬").ToString()
                        Dim _餘額_千 As String = data_dv(i + 19)("餘額_千").ToString()
                        Dim _餘額_百 As String = data_dv(i + 19)("餘額_百").ToString()
                        Dim _餘額_十 As String = data_dv(i + 19)("餘額_十").ToString()
                        Dim _餘額_個 As String = data_dv(i + 19)("餘額_個").ToString()
                        Dim _戶名 As String = data_dv(i + 19)("戶名").ToString()
                        Dim _品名 As String = data_dv(i + 19)("品名").ToString()
                        arr(i, 0) = _年
                        arr(i, 1) = _月
                        arr(i, 3) = _日
                        arr(i, 5) = _收支
                        arr(i, 7) = _國保收據編號
                        arr(i, 9) = _摘要
                        arr(i, 18) = _收入金額_億
                        arr(i, 19) = _收入金額_千萬
                        arr(i, 20) = _收入金額_百萬
                        arr(i, 21) = _收入金額_十萬
                        arr(i, 22) = _收入金額_萬
                        arr(i, 23) = _收入金額_千
                        arr(i, 24) = _收入金額_百
                        arr(i, 25) = _收入金額_十
                        arr(i, 26) = _收入金額_個
                        arr(i, 27) = _支出金額_億
                        arr(i, 28) = _支出金額_千萬
                        arr(i, 29) = _支出金額_百萬
                        arr(i, 30) = _支出金額_十萬
                        arr(i, 31) = _支出金額_萬
                        arr(i, 32) = _支出金額_千
                        arr(i, 33) = _支出金額_百
                        arr(i, 34) = _支出金額_十
                        arr(i, 35) = _支出金額_個
                        arr(i, 36) = _餘額_億
                        arr(i, 37) = _餘額_千萬
                        arr(i, 38) = _餘額_百萬
                        arr(i, 39) = _餘額_十萬
                        arr(i, 40) = _餘額_萬
                        arr(i, 41) = _餘額_千
                        arr(i, 42) = _餘額_百
                        arr(i, 43) = _餘額_十
                        arr(i, 44) = _餘額_個
                        arr(i, 47) = _戶名
                        arr(i, 49) = _品名
                    Next
                    xlWorkSheet.Range(xlWorkSheet.Cells(41, 1), xlWorkSheet.Cells(59, 51)).Value = arr
                    
                    '承上頁的餘額
                    data.SelectCommand = _
                        "SELECT TOP 1" & _
                        "SUBSTRING(REVERSE(CAST(餘額 AS VARCHAR(20))),9,12) AS 餘額_億, " & _
                        "SUBSTRING(REVERSE(CAST(餘額 AS VARCHAR(20))),8,1) AS 餘額_千萬, " & _
                        "SUBSTRING(REVERSE(CAST(餘額 AS VARCHAR(20))),7,1) AS 餘額_百萬, " & _
                        "SUBSTRING(REVERSE(CAST(餘額 AS VARCHAR(20))),6,1) AS 餘額_十萬, " & _
                        "SUBSTRING(REVERSE(CAST(餘額 AS VARCHAR(20))),5,1) AS 餘額_萬, " & _
                        "SUBSTRING(REVERSE(CAST(餘額 AS VARCHAR(20))),4,1) AS 餘額_千, " & _
                        "SUBSTRING(REVERSE(CAST(餘額 AS VARCHAR(20))),3,1) AS 餘額_百, " & _
                        "SUBSTRING(REVERSE(CAST(餘額 AS VARCHAR(20))),2,1) AS 餘額_十, " & _
                        "SUBSTRING(REVERSE(CAST(餘額 AS VARCHAR(20))),1,1) AS 餘額_個 " & _
                        "FROM 保管品紀錄簿 " & _
                        "WHERE (種類 = 0 AND 序號 < " & If(TextBox1 = "", "1", TextBox1) & ") " & _
                        "ORDER BY 序號 DESC"
                    data_dv = data.Select(New DataSourceSelectArguments)
                    ReDim arr(1, 9)
                    For i = 0 To data_dv.Count - 1
                        Dim _餘額_億 As String = data_dv(i)("餘額_億").ToString()
                        Dim _餘額_千萬 As String = data_dv(i)("餘額_千萬").ToString()
                        Dim _餘額_百萬 As String = data_dv(i)("餘額_百萬").ToString()
                        Dim _餘額_十萬 As String = data_dv(i)("餘額_十萬").ToString()
                        Dim _餘額_萬 As String = data_dv(i)("餘額_萬").ToString()
                        Dim _餘額_千 As String = data_dv(i)("餘額_千").ToString()
                        Dim _餘額_百 As String = data_dv(i)("餘額_百").ToString()
                        Dim _餘額_十 As String = data_dv(i)("餘額_十").ToString()
                        Dim _餘額_個 As String = data_dv(i)("餘額_個").ToString()
                        arr(i, 0) = _餘額_億
                        arr(i, 1) = _餘額_千萬
                        arr(i, 2) = _餘額_百萬
                        arr(i, 3) = _餘額_十萬
                        arr(i, 4) = _餘額_萬
                        arr(i, 5) = _餘額_千
                        arr(i, 6) = _餘額_百
                        arr(i, 7) = _餘額_十
                        arr(i, 8) = _餘額_個
                    Next
                    xlWorkSheet.Range(xlWorkSheet.Cells(10, 37), xlWorkSheet.Cells(10, 46)).Value = arr
                    xlWorkSheet.Range(xlWorkSheet.Cells(40, 37), xlWorkSheet.Cells(40, 46)).Value = xlWorkSheet.Range(xlWorkSheet.Cells(29, 37), xlWorkSheet.Cells(29, 46)).Value
                    
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
                    downloadfilename = "保證書紀錄簿.xls"
                    Response.AddHeader("Content-Disposition", "attachment;FileName=" + Uri.EscapeDataString(downloadfilename))
                    Response.WriteFile(MyExcel)
                    System.Web.HttpContext.Current.ApplicationInstance.CompleteRequest()
                    Response.Flush()
                    System.IO.File.Delete(MyExcel)
                    Response.End()
                End If
            Case "定存單紀錄簿"
                If e.CommandName = "轉明細表"
                    Update(sender, e)
                    
                    Dim CurrentRow = Me.GridView2.Rows(e.CommandSource.NamingContainer.RowIndex)
                    Dim Label0 As String = CType(CurrentRow.FindControl("Label0"), Label).Text
                    
                    data.UpdateCommand = _
                    "SET XACT_ABORT ON " & _
                    "BEGIN TRANSACTION " & _
                    "DECLARE @TempTable TABLE (id int) " & _
                    "IF (SELECT COUNT(*) FROM 保管品明細表 WHERE id = (SELECT 保管品明細表id FROM 保管品紀錄簿 WHERE id = " & Label0 & ")) = 1 " & _
                        "BEGIN " & _
                            "UPDATE 保管品明細表 SET 已支出 = 0 WHERE id = (SELECT 保管品明細表id FROM 保管品紀錄簿 WHERE id = " & Label0 & ") " & _
                        "END " & _
                    "ELSE " & _
                        "BEGIN " & _
                            "INSERT INTO 保管品明細表 " & _
                            "(種類, 已收入, 已支出, 日期, 國保收據編號, 摘要, 單位, 數量, 金額, 合約展延情形, 戶名, 品名) " & _
                            "OUTPUT INSERTED.id INTO @TempTable (id) " & _
                            "(SELECT 1, " & _
                                "1, " & _
                                "0, " & _
                                "日期, " & _
                                "國保收據編號, " & _
                                "摘要, " & _
                                "N'包', " & _
                                "1, " & _
                                "CASE WHEN 收支 = 0 THEN 收入金額 ELSE 支出金額 END, " & _
                                "N'無', " & _
                                "戶名, " & _
                                "品名 " & _
                            "FROM 保管品紀錄簿 WHERE id = N'" & Label0 & "') " & _
                            "UPDATE 保管品紀錄簿 SET 保管品明細表id = (SELECT MAX(id) FROM @TempTable) WHERE id = N'" & Label0 & "' " & _
                        "END " & _
                    "COMMIT TRANSACTION " & _
                    "SET XACT_ABORT OFF "
                    data.Update()
                Else If e.CommandName = "支出"
                    Update(sender, e)
                    
                    Dim CurrentRow = Me.GridView2.Rows(e.CommandSource.NamingContainer.RowIndex)
                    Dim Label0 As String = CType(CurrentRow.FindControl("Label0"), Label).Text
                    
                    data.UpdateCommand = _
                        "INSERT INTO 保管品紀錄簿 " & _
                        "(保管品明細表id, 種類, 序號, 日期, 收支, 國保收據編號, 摘要, 收入金額, 支出金額, 餘額, 戶名, 品名) " & _
                        "(SELECT 保管品明細表id, " & _
                            "1, " & _
                            "(SELECT ISNULL(MAX(序號), 0)+1 FROM 保管品紀錄簿 WHERE 種類 = 1), " & _
                            "NULL, " & _
                            "1, " & _
                            "國保收據編號, " & _
                            "摘要, " & _
                            "0, " & _
                            "收入金額, " & _
                            "(SELECT ISNULL((SELECT 餘額 FROM 保管品紀錄簿 WHERE 序號 = (SELECT ISNULL(MAX(序號), 0) FROM 保管品紀錄簿 WHERE 種類 = 1) AND 種類 = 1), 0)) - 收入金額, " & _
                            "戶名, " & _
                            "品名 " & _
                        "FROM 保管品紀錄簿 WHERE id = N'" & Label0 & "')"
                    data.Update()
                    
                    Me.GridView2.DataBind()
                Else If e.CommandName = "下載"
                    Update(sender, e)
                    
                    Dim CurrentRow = Me.GridView2.Rows(e.CommandSource.NamingContainer.RowIndex)
                    Dim TextBox1 As String = CType(CurrentRow.FindControl("TextBox1"), TextBox).Text
                    
                    Dim _GUID As String = Guid.NewGuid().ToString("N")
                    Dim MyExcel As String = MapPath(".\Excel\Temp\") & _GUID & ".xls"
                    System.IO.File.Copy(MapPath(".\Excel\定存單紀錄簿.xls"), MyExcel)
                    Dim xlApp As New Excel.ApplicationClass()
                    xlApp.DisplayAlerts = False
                    xlApp.ScreenUpdating = false
                    xlApp.EnableEvents = false
                    Dim xlWorkBook As Excel.Workbook = xlApp.Workbooks.Open(MyExcel, 0, False, 5, "", "", True, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", False, False, 0, True, 1, 0)
                    Dim xlWorkSheet As Excel.Worksheet
                    
                    xlWorkSheet = CType(xlWorkBook.Sheets("定存單"), Excel.Worksheet)
                    data.SelectCommand = _
                        "SELECT TOP 38" & _
                        "YEAR(日期) - 1911 AS 年, " & _
                        "MONTH(日期) AS 月, " & _
                        "DAY(日期) AS 日, " & _
                        "CASE WHEN 收支 = 0 THEN '收' ELSE '支' END AS 收支, " & _
                        "國保收據編號, " & _
                        "摘要, " & _
                        "SUBSTRING(REVERSE(CAST(收入金額 AS VARCHAR(20))),9,12) AS 收入金額_億, " & _
                        "SUBSTRING(REVERSE(CAST(收入金額 AS VARCHAR(20))),8,1) AS 收入金額_千萬, " & _
                        "SUBSTRING(REVERSE(CAST(收入金額 AS VARCHAR(20))),7,1) AS 收入金額_百萬, " & _
                        "SUBSTRING(REVERSE(CAST(收入金額 AS VARCHAR(20))),6,1) AS 收入金額_十萬, " & _
                        "SUBSTRING(REVERSE(CAST(收入金額 AS VARCHAR(20))),5,1) AS 收入金額_萬, " & _
                        "SUBSTRING(REVERSE(CAST(收入金額 AS VARCHAR(20))),4,1) AS 收入金額_千, " & _
                        "SUBSTRING(REVERSE(CAST(收入金額 AS VARCHAR(20))),3,1) AS 收入金額_百, " & _
                        "SUBSTRING(REVERSE(CAST(收入金額 AS VARCHAR(20))),2,1) AS 收入金額_十, " & _
                        "SUBSTRING(REVERSE(CAST(收入金額 AS VARCHAR(20))),1,1) AS 收入金額_個, " & _
                        "SUBSTRING(REVERSE(CAST(支出金額 AS VARCHAR(20))),9,12) AS 支出金額_億, " & _
                        "SUBSTRING(REVERSE(CAST(支出金額 AS VARCHAR(20))),8,1) AS 支出金額_千萬, " & _
                        "SUBSTRING(REVERSE(CAST(支出金額 AS VARCHAR(20))),7,1) AS 支出金額_百萬, " & _
                        "SUBSTRING(REVERSE(CAST(支出金額 AS VARCHAR(20))),6,1) AS 支出金額_十萬, " & _
                        "SUBSTRING(REVERSE(CAST(支出金額 AS VARCHAR(20))),5,1) AS 支出金額_萬, " & _
                        "SUBSTRING(REVERSE(CAST(支出金額 AS VARCHAR(20))),4,1) AS 支出金額_千, " & _
                        "SUBSTRING(REVERSE(CAST(支出金額 AS VARCHAR(20))),3,1) AS 支出金額_百, " & _
                        "SUBSTRING(REVERSE(CAST(支出金額 AS VARCHAR(20))),2,1) AS 支出金額_十, " & _
                        "SUBSTRING(REVERSE(CAST(支出金額 AS VARCHAR(20))),1,1) AS 支出金額_個, " & _
                        "SUBSTRING(REVERSE(CAST(餘額 AS VARCHAR(20))),9,12) AS 餘額_億, " & _
                        "SUBSTRING(REVERSE(CAST(餘額 AS VARCHAR(20))),8,1) AS 餘額_千萬, " & _
                        "SUBSTRING(REVERSE(CAST(餘額 AS VARCHAR(20))),7,1) AS 餘額_百萬, " & _
                        "SUBSTRING(REVERSE(CAST(餘額 AS VARCHAR(20))),6,1) AS 餘額_十萬, " & _
                        "SUBSTRING(REVERSE(CAST(餘額 AS VARCHAR(20))),5,1) AS 餘額_萬, " & _
                        "SUBSTRING(REVERSE(CAST(餘額 AS VARCHAR(20))),4,1) AS 餘額_千, " & _
                        "SUBSTRING(REVERSE(CAST(餘額 AS VARCHAR(20))),3,1) AS 餘額_百, " & _
                        "SUBSTRING(REVERSE(CAST(餘額 AS VARCHAR(20))),2,1) AS 餘額_十, " & _
                        "SUBSTRING(REVERSE(CAST(餘額 AS VARCHAR(20))),1,1) AS 餘額_個, " & _
                        "戶名, " & _
                        "品名 " & _
                        "FROM 保管品紀錄簿 " & _
                        "WHERE (種類 = 1 AND 序號 >= " & If(TextBox1 = "", "1", TextBox1) & ") " & _
                        "ORDER BY 序號"
                    data_dv = data.Select(New DataSourceSelectArguments)
                    Dim arr(19, 51) As Object
                    For i = 0 To Math.Min(18, data_dv.Count - 1)
                        Dim _年 As String = data_dv(i)("年").ToString()
                        Dim _月 As String = data_dv(i)("月").ToString()
                        Dim _日 As String = data_dv(i)("日").ToString()
                        Dim _收支 As String = data_dv(i)("收支").ToString()
                        Dim _國保收據編號 As String = data_dv(i)("國保收據編號").ToString()
                        Dim _摘要 As String = data_dv(i)("摘要").ToString()
                        Dim _收入金額_億 As String = data_dv(i)("收入金額_億").ToString()
                        Dim _收入金額_千萬 As String = data_dv(i)("收入金額_千萬").ToString()
                        Dim _收入金額_百萬 As String = data_dv(i)("收入金額_百萬").ToString()
                        Dim _收入金額_十萬 As String = data_dv(i)("收入金額_十萬").ToString()
                        Dim _收入金額_萬 As String = data_dv(i)("收入金額_萬").ToString()
                        Dim _收入金額_千 As String = data_dv(i)("收入金額_千").ToString()
                        Dim _收入金額_百 As String = data_dv(i)("收入金額_百").ToString()
                        Dim _收入金額_十 As String = data_dv(i)("收入金額_十").ToString()
                        Dim _收入金額_個 As String = data_dv(i)("收入金額_個").ToString()
                        Dim _支出金額_億 As String = data_dv(i)("支出金額_億").ToString()
                        Dim _支出金額_千萬 As String = data_dv(i)("支出金額_千萬").ToString()
                        Dim _支出金額_百萬 As String = data_dv(i)("支出金額_百萬").ToString()
                        Dim _支出金額_十萬 As String = data_dv(i)("支出金額_十萬").ToString()
                        Dim _支出金額_萬 As String = data_dv(i)("支出金額_萬").ToString()
                        Dim _支出金額_千 As String = data_dv(i)("支出金額_千").ToString()
                        Dim _支出金額_百 As String = data_dv(i)("支出金額_百").ToString()
                        Dim _支出金額_十 As String = data_dv(i)("支出金額_十").ToString()
                        Dim _支出金額_個 As String = data_dv(i)("支出金額_個").ToString()
                        Dim _餘額_億 As String = data_dv(i)("餘額_億").ToString()
                        Dim _餘額_千萬 As String = data_dv(i)("餘額_千萬").ToString()
                        Dim _餘額_百萬 As String = data_dv(i)("餘額_百萬").ToString()
                        Dim _餘額_十萬 As String = data_dv(i)("餘額_十萬").ToString()
                        Dim _餘額_萬 As String = data_dv(i)("餘額_萬").ToString()
                        Dim _餘額_千 As String = data_dv(i)("餘額_千").ToString()
                        Dim _餘額_百 As String = data_dv(i)("餘額_百").ToString()
                        Dim _餘額_十 As String = data_dv(i)("餘額_十").ToString()
                        Dim _餘額_個 As String = data_dv(i)("餘額_個").ToString()
                        Dim _戶名 As String = data_dv(i)("戶名").ToString()
                        Dim _品名 As String = data_dv(i)("品名").ToString()
                        arr(i, 0) = _年
                        arr(i, 1) = _月
                        arr(i, 3) = _日
                        arr(i, 5) = _收支
                        arr(i, 7) = _國保收據編號
                        arr(i, 9) = _摘要
                        arr(i, 18) = _收入金額_億
                        arr(i, 19) = _收入金額_千萬
                        arr(i, 20) = _收入金額_百萬
                        arr(i, 21) = _收入金額_十萬
                        arr(i, 22) = _收入金額_萬
                        arr(i, 23) = _收入金額_千
                        arr(i, 24) = _收入金額_百
                        arr(i, 25) = _收入金額_十
                        arr(i, 26) = _收入金額_個
                        arr(i, 27) = _支出金額_億
                        arr(i, 28) = _支出金額_千萬
                        arr(i, 29) = _支出金額_百萬
                        arr(i, 30) = _支出金額_十萬
                        arr(i, 31) = _支出金額_萬
                        arr(i, 32) = _支出金額_千
                        arr(i, 33) = _支出金額_百
                        arr(i, 34) = _支出金額_十
                        arr(i, 35) = _支出金額_個
                        arr(i, 36) = _餘額_億
                        arr(i, 37) = _餘額_千萬
                        arr(i, 38) = _餘額_百萬
                        arr(i, 39) = _餘額_十萬
                        arr(i, 40) = _餘額_萬
                        arr(i, 41) = _餘額_千
                        arr(i, 42) = _餘額_百
                        arr(i, 43) = _餘額_十
                        arr(i, 44) = _餘額_個
                        arr(i, 47) = _戶名
                        arr(i, 49) = _品名
                    Next
                    xlWorkSheet.Range(xlWorkSheet.Cells(11, 1), xlWorkSheet.Cells(29, 51)).Value = arr
                    Array.Clear(arr, 0, arr.Length)
                    For i = 0 To data_dv.Count - 1 - 19
                        Dim _年 As String = data_dv(i + 19)("年").ToString()
                        Dim _月 As String = data_dv(i + 19)("月").ToString()
                        Dim _日 As String = data_dv(i + 19)("日").ToString()
                        Dim _收支 As String = data_dv(i + 19)("收支").ToString()
                        Dim _國保收據編號 As String = data_dv(i + 19)("國保收據編號").ToString()
                        Dim _摘要 As String = data_dv(i + 19)("摘要").ToString()
                        Dim _收入金額_億 As String = data_dv(i + 19)("收入金額_億").ToString()
                        Dim _收入金額_千萬 As String = data_dv(i + 19)("收入金額_千萬").ToString()
                        Dim _收入金額_百萬 As String = data_dv(i + 19)("收入金額_百萬").ToString()
                        Dim _收入金額_十萬 As String = data_dv(i + 19)("收入金額_十萬").ToString()
                        Dim _收入金額_萬 As String = data_dv(i + 19)("收入金額_萬").ToString()
                        Dim _收入金額_千 As String = data_dv(i + 19)("收入金額_千").ToString()
                        Dim _收入金額_百 As String = data_dv(i + 19)("收入金額_百").ToString()
                        Dim _收入金額_十 As String = data_dv(i + 19)("收入金額_十").ToString()
                        Dim _收入金額_個 As String = data_dv(i + 19)("收入金額_個").ToString()
                        Dim _支出金額_億 As String = data_dv(i + 19)("支出金額_億").ToString()
                        Dim _支出金額_千萬 As String = data_dv(i + 19)("支出金額_千萬").ToString()
                        Dim _支出金額_百萬 As String = data_dv(i + 19)("支出金額_百萬").ToString()
                        Dim _支出金額_十萬 As String = data_dv(i + 19)("支出金額_十萬").ToString()
                        Dim _支出金額_萬 As String = data_dv(i + 19)("支出金額_萬").ToString()
                        Dim _支出金額_千 As String = data_dv(i + 19)("支出金額_千").ToString()
                        Dim _支出金額_百 As String = data_dv(i + 19)("支出金額_百").ToString()
                        Dim _支出金額_十 As String = data_dv(i + 19)("支出金額_十").ToString()
                        Dim _支出金額_個 As String = data_dv(i + 19)("支出金額_個").ToString()
                        Dim _餘額_億 As String = data_dv(i + 19)("餘額_億").ToString()
                        Dim _餘額_千萬 As String = data_dv(i + 19)("餘額_千萬").ToString()
                        Dim _餘額_百萬 As String = data_dv(i + 19)("餘額_百萬").ToString()
                        Dim _餘額_十萬 As String = data_dv(i + 19)("餘額_十萬").ToString()
                        Dim _餘額_萬 As String = data_dv(i + 19)("餘額_萬").ToString()
                        Dim _餘額_千 As String = data_dv(i + 19)("餘額_千").ToString()
                        Dim _餘額_百 As String = data_dv(i + 19)("餘額_百").ToString()
                        Dim _餘額_十 As String = data_dv(i + 19)("餘額_十").ToString()
                        Dim _餘額_個 As String = data_dv(i + 19)("餘額_個").ToString()
                        Dim _戶名 As String = data_dv(i + 19)("戶名").ToString()
                        Dim _品名 As String = data_dv(i + 19)("品名").ToString()
                        arr(i, 0) = _年
                        arr(i, 1) = _月
                        arr(i, 3) = _日
                        arr(i, 5) = _收支
                        arr(i, 7) = _國保收據編號
                        arr(i, 9) = _摘要
                        arr(i, 18) = _收入金額_億
                        arr(i, 19) = _收入金額_千萬
                        arr(i, 20) = _收入金額_百萬
                        arr(i, 21) = _收入金額_十萬
                        arr(i, 22) = _收入金額_萬
                        arr(i, 23) = _收入金額_千
                        arr(i, 24) = _收入金額_百
                        arr(i, 25) = _收入金額_十
                        arr(i, 26) = _收入金額_個
                        arr(i, 27) = _支出金額_億
                        arr(i, 28) = _支出金額_千萬
                        arr(i, 29) = _支出金額_百萬
                        arr(i, 30) = _支出金額_十萬
                        arr(i, 31) = _支出金額_萬
                        arr(i, 32) = _支出金額_千
                        arr(i, 33) = _支出金額_百
                        arr(i, 34) = _支出金額_十
                        arr(i, 35) = _支出金額_個
                        arr(i, 36) = _餘額_億
                        arr(i, 37) = _餘額_千萬
                        arr(i, 38) = _餘額_百萬
                        arr(i, 39) = _餘額_十萬
                        arr(i, 40) = _餘額_萬
                        arr(i, 41) = _餘額_千
                        arr(i, 42) = _餘額_百
                        arr(i, 43) = _餘額_十
                        arr(i, 44) = _餘額_個
                        arr(i, 47) = _戶名
                        arr(i, 49) = _品名
                    Next
                    xlWorkSheet.Range(xlWorkSheet.Cells(41, 1), xlWorkSheet.Cells(59, 51)).Value = arr
                    
                    '承上頁的餘額
                    data.SelectCommand = _
                        "SELECT TOP 1" & _
                        "SUBSTRING(REVERSE(CAST(餘額 AS VARCHAR(20))),9,12) AS 餘額_億, " & _
                        "SUBSTRING(REVERSE(CAST(餘額 AS VARCHAR(20))),8,1) AS 餘額_千萬, " & _
                        "SUBSTRING(REVERSE(CAST(餘額 AS VARCHAR(20))),7,1) AS 餘額_百萬, " & _
                        "SUBSTRING(REVERSE(CAST(餘額 AS VARCHAR(20))),6,1) AS 餘額_十萬, " & _
                        "SUBSTRING(REVERSE(CAST(餘額 AS VARCHAR(20))),5,1) AS 餘額_萬, " & _
                        "SUBSTRING(REVERSE(CAST(餘額 AS VARCHAR(20))),4,1) AS 餘額_千, " & _
                        "SUBSTRING(REVERSE(CAST(餘額 AS VARCHAR(20))),3,1) AS 餘額_百, " & _
                        "SUBSTRING(REVERSE(CAST(餘額 AS VARCHAR(20))),2,1) AS 餘額_十, " & _
                        "SUBSTRING(REVERSE(CAST(餘額 AS VARCHAR(20))),1,1) AS 餘額_個 " & _
                        "FROM 保管品紀錄簿 " & _
                        "WHERE (種類 = 1 AND 序號 < " & If(TextBox1 = "", "1", TextBox1) & ") " & _
                        "ORDER BY 序號 DESC"
                    data_dv = data.Select(New DataSourceSelectArguments)
                    ReDim arr(1, 9)
                    For i = 0 To data_dv.Count - 1
                        Dim _餘額_億 As String = data_dv(i)("餘額_億").ToString()
                        Dim _餘額_千萬 As String = data_dv(i)("餘額_千萬").ToString()
                        Dim _餘額_百萬 As String = data_dv(i)("餘額_百萬").ToString()
                        Dim _餘額_十萬 As String = data_dv(i)("餘額_十萬").ToString()
                        Dim _餘額_萬 As String = data_dv(i)("餘額_萬").ToString()
                        Dim _餘額_千 As String = data_dv(i)("餘額_千").ToString()
                        Dim _餘額_百 As String = data_dv(i)("餘額_百").ToString()
                        Dim _餘額_十 As String = data_dv(i)("餘額_十").ToString()
                        Dim _餘額_個 As String = data_dv(i)("餘額_個").ToString()
                        arr(i, 0) = _餘額_億
                        arr(i, 1) = _餘額_千萬
                        arr(i, 2) = _餘額_百萬
                        arr(i, 3) = _餘額_十萬
                        arr(i, 4) = _餘額_萬
                        arr(i, 5) = _餘額_千
                        arr(i, 6) = _餘額_百
                        arr(i, 7) = _餘額_十
                        arr(i, 8) = _餘額_個
                    Next
                    xlWorkSheet.Range(xlWorkSheet.Cells(10, 37), xlWorkSheet.Cells(10, 46)).Value = arr
                    xlWorkSheet.Range(xlWorkSheet.Cells(40, 37), xlWorkSheet.Cells(40, 46)).Value = xlWorkSheet.Range(xlWorkSheet.Cells(29, 37), xlWorkSheet.Cells(29, 46)).Value
                    
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
                    downloadfilename = "定存單紀錄簿.xls"
                    Response.AddHeader("Content-Disposition", "attachment;FileName=" + Uri.EscapeDataString(downloadfilename))
                    Response.WriteFile(MyExcel)
                    System.Web.HttpContext.Current.ApplicationInstance.CompleteRequest()
                    Response.Flush()
                    System.IO.File.Delete(MyExcel)
                    Response.End()
                End If
        End Select
    End Sub
End Class
