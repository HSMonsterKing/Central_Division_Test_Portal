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
Partial Class 結算
    Inherits System.Web.UI.Page
    Dim con_14 As String = System.Web.Configuration.WebConfigurationManager.ConnectionStrings("ApplicationServices").ConnectionString
    Dim data As New SqlDataSource
    Dim data_dv As Data.DataView
    Dim data_dv1 As Data.DataView
    Dim data_dv2 As Data.DataView
    Dim data_dv3 As Data.DataView
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        data.ConnectionString = con_14
        If Not Page.IsPostBack Then
            For i = 109 To Now.Year() - 1910
                Me.DropDownList1.Items.Add(i)
                Me.DropDownList1.Items(i - 109).Value = i
            Next
            Me.DropDownList1.SelectedIndex = Me.DropDownList1.Items.IndexOf(Me.DropDownList1.Items.FindByValue(Now.Year - 1911))
            For i = 1 To 12
                Me.DropDownList2.Items.Add((i).ToString("00"))
                Me.DropDownList2.Items(i - 1).Value = (i).ToString("00")
            Next
            Me.DropDownList2.SelectedIndex = Me.DropDownList2.Items.IndexOf(Me.DropDownList2.Items.FindByValue(Now.Month.ToString("00")))
            DropDownList1_SelectedIndexChanged(sender, e)
            Me.DropDownList3.SelectedIndex = Me.DropDownList3.Items.IndexOf(Me.DropDownList3.Items.FindByValue(Now.Day.ToString("00")))
            
            data.SelectCommand = "select top 1 DATEADD(day, 5, 結帳日期) from 日報表 order by 結帳日期 desc"
            data_dv = data.Select(New DataSourceSelectArguments)
            If data_dv.Count > 0
                If Not IsDBNull(data_dv(0)(0))
                    Dim _結帳日期 As String
                    _結帳日期 = data_dv(0)(0)
                    If DateAndTime.Weekday(_結帳日期, 2) = 6
                        _結帳日期 = Convert.ToDateTime(_結帳日期).AddDays(-1)
                    Else If DateAndTime.Weekday(_結帳日期, 2) = 7
                        _結帳日期 = Convert.ToDateTime(_結帳日期).AddDays(1)
                    End If
                    Me.DropDownList1.SelectedIndex = Me.DropDownList1.Items.IndexOf(Me.DropDownList1.Items.FindByValue(Year(_結帳日期) - 1911))
                    DropDownList1_SelectedIndexChanged(sender, e)
                    Me.DropDownList2.SelectedIndex = Me.DropDownList2.Items.IndexOf(Me.DropDownList2.Items.FindByValue(Month(_結帳日期).ToString("00")))
                    DropDownList1_SelectedIndexChanged(sender, e)
                    Me.DropDownList3.SelectedIndex = Me.DropDownList3.Items.IndexOf(Me.DropDownList3.Items.FindByValue(Day(_結帳日期).ToString("00")))
                End If
            End If
            settextbox()
        Else
        End If
    End Sub
    Protected Sub DropDownList1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DropDownList1.SelectedIndexChanged
        Dim currentdate = Me.DropDownList3.SelectedValue
        Me.DropDownList3.Items.Clear()
        For i = 0 To DateTime.DaysInMonth((CLng(Me.DropDownList1.SelectedValue) + 1911), CLng(Me.DropDownList2.SelectedValue)) - 1
            Me.DropDownList3.Items.Add((i + 1).ToString("00"))
            Me.DropDownList3.Items(i).Value = (i + 1).ToString("00")
        Next
        If Me.DropDownList3.Items.IndexOf(Me.DropDownList3.Items.FindByValue(currentdate)) = -1
            Me.DropDownList3.SelectedIndex = Me.DropDownList3.Items.Count - 1
        Else
            Me.DropDownList3.SelectedIndex = Me.DropDownList3.Items.IndexOf(Me.DropDownList3.Items.FindByValue(currentdate))
        End If
        settextbox()
    End Sub
    Protected Sub DropDownList2_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DropDownList2.SelectedIndexChanged
        DropDownList1_SelectedIndexChanged(sender, e)
        Me.DropDownList3.SelectedIndex = Me.DropDownList3.Items.IndexOf(Me.DropDownList3.Items.FindByValue("01"))
        settextbox()
    End Sub
    Protected Sub DropDownList3_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DropDownList3.SelectedIndexChanged
        settextbox()
    End Sub
    Public Sub settextbox()
        Dim _結帳日期 As String = Me.DropDownList1.SelectedValue & "/" & Me.DropDownList2.SelectedValue & "/" & Me.DropDownList3.SelectedValue
        _結帳日期 = taiwancalendarto(_結帳日期)
        data.SelectCommand = "select top 1 * from 日報表 where year(結帳日期)=" & Me.DropDownList1.SelectedValue & " + 1911 and 結帳日期<'" & _結帳日期 & "' order by 結帳日期 desc"
        data_dv = data.Select(New DataSourceSelectArguments)
        If data_dv.Count > 0
            If Not IsDBNull(data_dv(0)(9))
                Me.TextBox1.Text = data_dv(0)(9) + 1
            End If
            If Not IsDBNull(data_dv(0)(11))
                Me.TextBox3.Text = data_dv(0)(11) + 1
            End If
            If Not IsDBNull(data_dv(0)(13))
                Me.TextBox5.Text = data_dv(0)(13) + 1
            End If
            If Not IsDBNull(data_dv(0)(15))
                Me.TextBox7.Text = data_dv(0)(15) + 1
            End If
        Else
            Me.TextBox1.Text = "1000001"
            Me.TextBox3.Text = "2000001"
            Me.TextBox5.Text = "3000001"
            Me.TextBox7.Text = "4000001"
        End If
    End Sub
    Protected Sub reorder(ByVal _年 As String)
        data.UpdateCommand = "WITH CTE AS (SELECT id, 序號, ROW_NUMBER() OVER (ORDER BY 結帳日期, 傳票號碼) AS RN FROM 現金備查簿 WHERE YEAR(結帳日期)=" & _年 & " + 1911) UPDATE CTE SET 序號 = RN"
        data.Update()
        data.UpdateCommand = "WITH CTE AS (SELECT id, 序號, ROW_NUMBER() OVER (ORDER BY 結帳日期, 傳票號碼) AS RN FROM 分錄 WHERE YEAR(結帳日期)=" & _年 & " + 1911) UPDATE CTE SET 序號 = RN"
        data.Update()
    End Sub
    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.TextBox1.Text = Trim(Me.TextBox1.Text)
        Me.TextBox2.Text = Trim(Me.TextBox2.Text)
        Me.TextBox3.Text = Trim(Me.TextBox3.Text)
        Me.TextBox4.Text = Trim(Me.TextBox4.Text)
        Me.TextBox5.Text = Trim(Me.TextBox5.Text)
        Me.TextBox6.Text = Trim(Me.TextBox6.Text)
        Me.TextBox7.Text = Trim(Me.TextBox7.Text)
        Me.TextBox8.Text = Trim(Me.TextBox8.Text)
        If Me.TextBox1.Text <> ""
            Me.TextBox1.Text = CLng(Me.TextBox1.Text).ToString("000000")
            Me.TextBox1.Text = "1" & Strings.Right(Me.TextBox1.Text, 6)
        Else
            Me.TextBox2.Text = ""
        End If
        If Me.TextBox2.Text <> ""
            Me.TextBox2.Text = CLng(Me.TextBox2.Text).ToString("000000")
            Me.TextBox2.Text = "1" & Strings.Right(Me.TextBox2.Text, 6)
        Else
            Me.TextBox1.Text = ""
        End If
        If Me.TextBox3.Text <> ""
            Me.TextBox3.Text = CLng(Me.TextBox3.Text).ToString("000000")
            Me.TextBox3.Text = "2" & Strings.Right(Me.TextBox3.Text, 6)
        Else
            Me.TextBox4.Text = ""
        End If
        If Me.TextBox4.Text <> ""
            Me.TextBox4.Text = CLng(Me.TextBox4.Text).ToString("000000")
            Me.TextBox4.Text = "2" & Strings.Right(Me.TextBox4.Text, 6)
        Else
            Me.TextBox3.Text = ""
        End If
        If Me.TextBox5.Text <> ""
            Me.TextBox5.Text = CLng(Me.TextBox5.Text).ToString("000000")
            Me.TextBox5.Text = "3" & Strings.Right(Me.TextBox5.Text, 6)
        Else
            Me.TextBox6.Text = ""
        End If
        If Me.TextBox6.Text <> ""
            Me.TextBox6.Text = CLng(Me.TextBox6.Text).ToString("000000")
            Me.TextBox6.Text = "3" & Strings.Right(Me.TextBox6.Text, 6)
        Else
            Me.TextBox5.Text = ""
        End If
        If Me.TextBox7.Text <> ""
            Me.TextBox7.Text = CLng(Me.TextBox7.Text).ToString("000000")
            Me.TextBox7.Text = "4" & Strings.Right(Me.TextBox7.Text, 6)
        Else
            Me.TextBox8.Text = ""
        End If
        If Me.TextBox8.Text <> ""
            Me.TextBox8.Text = CLng(Me.TextBox8.Text).ToString("000000")
            Me.TextBox8.Text = "4" & Strings.Right(Me.TextBox8.Text, 6)
        Else
            Me.TextBox7.Text = ""
        End If
        
        Dim _C6 As Long = 0
        Dim _D6 As Long = 0
        Dim _E6 As Long = 0
        Dim _C7 As Long = 0
        Dim _D7 As Long = 0
        Dim _E7 As Long = 0
        Dim _C12 As String = ""
        Dim _E12 As String = ""
        Dim _C13 As String = ""
        Dim _E13 As String = ""
        Dim _C14 As String = ""
        Dim _E14 As String = ""
        Dim _C15 As String = ""
        Dim _E15 As String = ""
        
        Dim _結帳日期 As String = Me.DropDownList1.SelectedValue & "/" & Me.DropDownList2.SelectedValue & "/" & Me.DropDownList3.SelectedValue
        _結帳日期 = taiwancalendarto(_結帳日期)
        
        '考慮重新結算的情況
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
            "WHERE 日報表.結帳日期 >= '" & _結帳日期 & "'"
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
            "WHERE 日報表.結帳日期 >= '" & _結帳日期 & "'"
        data.Update()
        data.DeleteCommand = "delete from 日報表 where 結帳日期>='" & _結帳日期 & "'"
        data.Delete()
        
        data.SelectCommand = "select top 1 * from 日報表 where 結帳日期<'" & _結帳日期 & "' order by 結帳日期 desc"
        data_dv = data.Select(New DataSourceSelectArguments)
        _C6 = data_dv(0)(2) + data_dv(0)(3) - data_dv(0)(4)
        _C7 = data_dv(0)(5) + data_dv(0)(6) - data_dv(0)(7)
        _C12 = Me.TextBox1.Text
        _C13 = Me.TextBox3.Text
        _C14 = Me.TextBox5.Text
        _C15 = Me.TextBox7.Text
        _E12 = Me.TextBox2.Text
        _E13 = Me.TextBox4.Text
        _E14 = Me.TextBox6.Text
        _E15 = Me.TextBox8.Text
        
        Dim _餘額405 As Long = 0
        Dim _餘額409 As Long = 0
        data.SelectCommand = "select top 1 C6, D6, E6, C7, D7, E7 from 日報表 where 結帳日期<'" & _結帳日期 & "' order by 結帳日期 desc"
        data_dv = data.Select(New DataSourceSelectArguments)
        If data_dv.Count > 0
            _餘額405 = data_dv(0)(0) + data_dv(0)(1) - data_dv(0)(2)
            _餘額409 = data_dv(0)(3) + data_dv(0)(4) - data_dv(0)(5)
        End If
        
        For k = 1 To 3
            Dim a As String = ""
            Dim b As String = ""
            If k = 1
                a = Me.TextBox1.Text
                b = Me.TextBox2.Text
            Else If k = 2
                a = Me.TextBox3.Text
                b = Me.TextBox4.Text
            Else If k = 3
                a = Me.TextBox5.Text
                b = Me.TextBox6.Text
            End If
            data.SelectCommand = "select * from 現金備查簿 where 傳票號碼>='" & a & "' and 傳票號碼<='" & b & "' and (年='" & Me.DropDownList1.SelectedValue & "') order by 傳票號碼"
            data_dv = data.Select(New DataSourceSelectArguments)
            For i = 0 To data_dv.Count - 1
                Dim _收入金額405 As Long = CLng("0" & data_dv(i)(7).ToString())
                Dim _支出金額405 As Long = CLng("0" & data_dv(i)(8).ToString())
                Dim _收入金額409 As Long = CLng("0" & data_dv(i)(10).ToString())
                Dim _支出金額409 As Long = CLng("0" & data_dv(i)(11).ToString())
                _餘額405 = _餘額405 + _收入金額405 - _支出金額405
                _餘額409 = _餘額409 + _收入金額409 - _支出金額409
                data.UpdateCommand = "update 現金備查簿 set 結帳日期='" & _結帳日期 & "', 餘額405='" & _餘額405 & "', 餘額409='" & _餘額409 & "'  where id='" & data_dv(i)(0) & "'"
                data.Update()
                _D6 = _D6 + CLng("0" & data_dv(i)(7).ToString())
                _D7 = _D7 + CLng("0" & data_dv(i)(10).ToString())
                _E6 = _E6 + CLng("0" & data_dv(i)(8).ToString())
                _E7 = _E7 + CLng("0" & data_dv(i)(11).ToString())
            Next
            data.UpdateCommand = "update 分錄 set 結帳日期='" & _結帳日期 & "' where 傳票號碼>='" & Me.TextBox7.Text & "' and 傳票號碼<='" & Me.TextBox8.Text & "' and (年='" & Me.DropDownList1.SelectedValue & "')"
            data.Update()
        Next
        
        data.InsertCommand = "insert into 日報表 (結帳日期, C6, C7, C12, C13, C14, C15, E6, E7, E12, E13, E14, E15, D6, D7) VALUES (N'" & _結帳日期 & "', NULLIF(N'" & _C6 & "',''), NULLIF(N'" & _C7 & "',''), NULLIF(N'" & _C12 & "',''), NULLIF(N'" & _C13 & "',''), NULLIF(N'" & _C14 & "',''), NULLIF(N'" & _C15 & "',''), NULLIF(N'" & _E6 & "',''), NULLIF(N'" & _E7 & "',''), NULLIF(N'" & _E12 & "',''), NULLIF(N'" & _E13 & "',''), NULLIF(N'" & _E14 & "',''), NULLIF(N'" & _E15 & "',''), NULLIF(N'" & _D6 & "',''), NULLIF(N'" & _D7 & "',''))"
        data.Insert()
        
        reorder(Me.DropDownList1.SelectedValue)
        
        Me.Debug.Text = "結算完成！"
    End Sub
End Class
