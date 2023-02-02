Imports System.Windows.Forms
Imports Microsoft.Office.Interop
Partial Class 統計
    Inherits System.Web.UI.Page
    Dim con_14 As String = "Data Source=10.52.0.178;Initial Catalog=大宗郵件;User ID=qaz;Password=1qaz@WSX"
    Dim data As New SqlDataSource
    Dim data_dv As Data.DataView
    Dim data_dv1 As Data.DataView
    Dim data_dv2 As Data.DataView
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not Page.IsPostBack Then
            Dim r1, r2, r3 As Long
            Dim _登入帳號 As String
            Dim _承辦人帳號 As String = ""
            _登入帳號 = Request.ServerVariables("REMOTE_HOST")
            If Len(Trim(_登入帳號)) <= 0 Then
                _登入帳號 = System.Security.Principal.WindowsIdentity.GetCurrent().Name
            End If
            r1 = Len(_登入帳號)
            For i As Long = 1 To r1
                If Mid(_登入帳號, i, 1) = "\" Then
                    r2 = i
                End If
            Next
            r3 = r1 - r2
            _登入帳號 = Right(_登入帳號, r3)
            Me.TextBox2.Text = _登入帳號
            Dim _年, _月, _日 As String
            '_年 = (Now.Year - 1911).ToString
            '_月 = Now.Month.ToString("00")
            '_日 = Now.Day.ToString("00")
            data.ConnectionString = con_14
            data.SelectCommand = "SELECT 年, 月, 日 FROM 大宗郵件執據_操作者 WHERE 帳號='" & Me.TextBox2.Text & "'"
            data.DataBind()
            data_dv = data.Select(New DataSourceSelectArguments)
            Try
                _年 = data_dv(0)(0)
                _月 = data_dv(0)(1)
                _日 = data_dv(0)(2)
            Catch
                _年 = (Now.Year - 1911).ToString
                _月 = Now.Month.ToString("00")
                _日 = Now.Day.ToString("00")
                data.ConnectionString = con_14
                data.UpdateCommand = "update 大宗郵件執據_操作者 set 年 = '" & Me.DropDownList1.SelectedValue & "', 月 = '" & Me.DropDownList2.SelectedValue & "', 日 = '" & Me.DropDownList3.SelectedValue & "' WHERE 帳號='" & Me.TextBox2.Text & "'"
                data.Update()
                data.DataBind()
            End Try
            
            Dim j As Long
            For i As Long = 0 To Val(_年) - 108
                j = i + 109
                Me.DropDownList1.Items.Add(j.ToString("000"))
                Me.DropDownList1.Items(i).Value = j.ToString("000")
                Me.DropDownList4.Items.Add(j.ToString("000"))
                Me.DropDownList4.Items(i).Value = j.ToString("000")
            Next
            Me.DropDownList1.DataBind()
            Me.DropDownList1.SelectedValue = _年
            Me.DropDownList4.DataBind()
            Me.DropDownList4.SelectedValue = _年
            For i As Long = 0 To 11
                Me.DropDownList2.Items.Add((i + 1).ToString("00"))
                Me.DropDownList2.Items(i).Value = (i + 1).ToString("00")
                Me.DropDownList5.Items.Add((i + 1).ToString("00"))
                Me.DropDownList5.Items(i).Value = (i + 1).ToString("00")
            Next
            Me.DropDownList2.DataBind()
            Me.DropDownList2.SelectedValue = _月
            Me.DropDownList5.DataBind()
            Me.DropDownList5.SelectedValue = _月
            For i As Long = 0 To DateTime.DaysInMonth(Now.Year, Now.Month) - 1
                Me.DropDownList3.Items.Add((i + 1).ToString("00"))
                Me.DropDownList3.Items(i).Value = (i + 1).ToString("00")
                Me.DropDownList6.Items.Add((i + 1).ToString("00"))
                Me.DropDownList6.Items(i).Value = (i + 1).ToString("00")
            Next
            Me.DropDownList3.DataBind()
            Me.DropDownList3.SelectedValue = _日
            Me.DropDownList6.DataBind()
            Me.DropDownList6.SelectedValue = _日
            Me.RadioButtonList1.SelectedValue = 1
            '  Button4.OnClientClick = "window.open('日曆.aspx','','menubar=no,status=no,scrollbars=yes,top=100,left=200,toolbar=no,width=450,height=300');"
            GenTreeNode2()
        End If
    End Sub
    Protected Sub DropDownList1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DropDownList1.SelectedIndexChanged
        GenTreeNode()
    End Sub
    Protected Sub DropDownList2_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DropDownList2.SelectedIndexChanged
        GenTreeNode()
    End Sub
    Protected Sub DropDownList4_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DropDownList4.SelectedIndexChanged
        GenTreeNode1()
    End Sub
    Protected Sub DropDownList5_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DropDownList5.SelectedIndexChanged
        GenTreeNode1()
    End Sub
    Protected Sub Button3_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button3.Click
        GenTreeNode2()
    End Sub
    Protected Sub GenTreeNode()
        Dim _日 As String = Me.DropDownList3.SelectedValue
        Me.DropDownList3.Items.Clear()
        For i As Long = 0 To DateTime.DaysInMonth((Val(Me.DropDownList1.SelectedValue) + 1911), Val(Me.DropDownList2.SelectedValue)) - 1
            Me.DropDownList3.Items.Add((i + 1).ToString("00"))
            Me.DropDownList3.Items(i).Value = (i + 1).ToString("00")
        Next
        Me.DropDownList3.DataBind()
        Try
            Me.DropDownList3.SelectedValue = _日
        Catch ex As Exception
        End Try
    End Sub
    Protected Sub GenTreeNode1()
        Dim _日 As String = Me.DropDownList6.SelectedValue
        Me.DropDownList6.Items.Clear()
        For i As Long = 0 To DateTime.DaysInMonth((Val(Me.DropDownList4.SelectedValue) + 1911), Val(Me.DropDownList5.SelectedValue)) - 1
            Me.DropDownList6.Items.Add((i + 1).ToString("00"))
            Me.DropDownList6.Items(i).Value = (i + 1).ToString("00")
        Next
        Me.DropDownList6.DataBind()
        Try
            Me.DropDownList6.SelectedValue = _日
        Catch ex As Exception
        End Try
    End Sub
    'Protected Sub Button19_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button19.Click
    '    GenTreeNode2()
    'End Sub
    Protected Sub GenTreeNode2()
        Me.Label1.Text = ""
        Me.Label2.Text = ""
        data.ConnectionString = con_14
        Dim _件數合計 As Decimal = 0
        Dim _郵資合計 As Decimal = 0
        Dim _id As Long = 0
        Dim _件數 As Long = 0
        Dim _郵資 As Long = 0
        Dim _序號 As Long = 0
        Dim _日期1 As String = Me.DropDownList1.SelectedValue + Me.DropDownList2.SelectedValue + Me.DropDownList3.SelectedValue
        Dim _日期2 As String = Me.DropDownList4.SelectedValue + Me.DropDownList5.SelectedValue + Me.DropDownList6.SelectedValue
        data.SelectCommand = "SELECT id,序號 FROM  大宗郵件執據_郵寄種類 ORDER BY 排序"
        data.DataBind()
        data_dv = data.Select(New DataSourceSelectArguments)
        For i As Long = 0 To data_dv.Count - 1
            _id = data_dv(i)(0).ToString()
            _序號 = data_dv(i)(1).ToString()
            _件數 = 0
            _郵資 = 0
            If Me.RadioButtonList1.SelectedValue = 1 Then
                data.SelectCommand = "SELECT SUM(件數) AS a FROM 大宗郵件執據 where (郵寄種類='" & _序號 & "') and (年+月+日>='" & _日期1 & "') and (年+月+日<='" & _日期2 & "')"
            End If
            If Me.RadioButtonList1.SelectedValue = 2 Then
                data.SelectCommand = "SELECT SUM(件數) AS a FROM 大宗郵件執據 where 收費小組=0 and (郵寄種類='" & _序號 & "') and (年+月+日>='" & _日期1 & "') and (年+月+日<='" & _日期2 & "')"
            End If
            If Me.RadioButtonList1.SelectedValue = 3 Then
                data.SelectCommand = "SELECT SUM(件數) AS a FROM 大宗郵件執據 where 收費小組=1 and (郵寄種類='" & _序號 & "') and (年+月+日>='" & _日期1 & "') and (年+月+日<='" & _日期2 & "')"
            End If
            data.DataBind()
            data_dv1 = data.Select(New DataSourceSelectArguments)
            Try
                _件數 = data_dv1(0)(0).ToString()
            Catch ex As Exception
            End Try
            _件數合計 = _件數合計 + _件數
            data.UpdateCommand = "update 大宗郵件執據_郵寄種類 set 件數='" & _件數 & "' where id='" & _id & "'"
            data.Update()
            data.DataBind()
            If Me.RadioButtonList1.SelectedValue = 1 Then
                data.SelectCommand = "SELECT SUM(郵資) AS a FROM 大宗郵件執據 where (郵寄種類='" & _序號 & "') and (年+月+日>='" & _日期1 & "') and (年+月+日<='" & _日期2 & "')"
            End If
            If Me.RadioButtonList1.SelectedValue = 2 Then
                data.SelectCommand = "SELECT SUM(郵資) AS a FROM 大宗郵件執據 where 收費小組=0 and (郵寄種類='" & _序號 & "') and (年+月+日>='" & _日期1 & "') and (年+月+日<='" & _日期2 & "')"
            End If
            If Me.RadioButtonList1.SelectedValue = 3 Then
                data.SelectCommand = "SELECT SUM(郵資) AS a FROM 大宗郵件執據 where 收費小組=1 and (郵寄種類='" & _序號 & "') and (年+月+日>='" & _日期1 & "') and (年+月+日<='" & _日期2 & "')"
            End If
            data.DataBind()
            data_dv1 = data.Select(New DataSourceSelectArguments)
            Try
                _郵資 = data_dv1(0)(0).ToString()
            Catch ex As Exception
            End Try
            _郵資合計 = _郵資合計 + _郵資
            data.UpdateCommand = "update 大宗郵件執據_郵寄種類 set 郵資='" & _郵資 & "' where id='" & _id & "'"
            data.Update()
            data.DataBind()
        Next
        Me.SqlDataSource1.DataBind()
        Me.GridView1.DataBind()
        Me.Label1.Text = "件數合計:" + Format(CType(_件數合計.ToString(), Decimal), "#,###,###,###,###")
        Me.Label2.Text = "郵資合計:" + Format(CType(_郵資合計.ToString(), Decimal), "#,###,###,###,###")
        If _件數合計 = 0 Then
            Me.Label1.Text = "件數合計:0"
        End If
        If _郵資合計 = 0 Then
            Me.Label2.Text = "郵資合計:0"
        End If
    End Sub
    Protected Sub RadioButtonList1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles RadioButtonList1.SelectedIndexChanged 
        GenTreeNode2()
    End Sub
    Protected Sub Button4_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button4.Click
        data.ConnectionString = con_14
        Dim _GUID As String = Guid.NewGuid().ToString("N")
        Dim MyExcel As String = "C:\大宗郵件\Excel\Temp\" & _GUID & ".xls"
        System.IO.File.Copy("C:\大宗郵件\Excel\郵務種類日報表.xls", MyExcel)
        'Dim MyExcel As String = "C:\大宗測試\Excel\Temp\" & _GUID & ".xls"
        'System.IO.File.Copy("C:\大宗測試\Excel\郵務種類日報表.xls", MyExcel)
        Dim xlApp As Excel.ApplicationClass
        xlApp = New Excel.ApplicationClass()
        xlApp.DisplayAlerts = False
        xlApp.ScreenUpdating = false
        xlApp.EnableEvents = false
        Dim xlWorkBook As Excel.Workbook = xlApp.Workbooks.Open(MyExcel, 0, False, 5, "", "", True, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", False, False, 0, True, 1, 0)
        Dim xlWorkSheet As Excel.Worksheet = xlWorkBook.ActiveSheet

        data.SelectCommand = "SELECT distinct 日 FROM  大宗郵件執據 where 年='" & Me.DropDownList1.SelectedValue & "' and 月='" & Me.DropDownList2.SelectedValue & "'ORDER BY 日"
        If Me.RadioButtonList1.SelectedValue = 2 Then
            data.SelectCommand = "SELECT distinct 日 FROM  大宗郵件執據 where 年='" & Me.DropDownList1.SelectedValue & "' and 月='" & Me.DropDownList2.SelectedValue & "' and 收費小組 = 0 ORDER BY 日"
        End If
        If Me.RadioButtonList1.SelectedValue = 3 Then
            data.SelectCommand = "SELECT distinct 日 FROM  大宗郵件執據 where 年='" & Me.DropDownList1.SelectedValue & "' and 月='" & Me.DropDownList2.SelectedValue & "' and 收費小組 = 1 ORDER BY 日"
        End If
        
        data.DataBind()
        data_dv = data.Select(New DataSourceSelectArguments)
        Dim _日 As String = ""
        Dim dat As String = ""
        Dim _件數 As Long = 0
        Dim _郵資 As Long = 0
        Dim _序號 As Long = 0
        If Me.RadioButtonList1.SelectedValue = 1 Then
            xlWorkSheet.Cells(1, 1) = "交通部高速公路局中區養護工程分局"
        End If
        If Me.RadioButtonList1.SelectedValue = 2 Then
            xlWorkSheet.Cells(1, 1) = "交通部高速公路局中區養護工程分局(本分局)"
        End If
        If Me.RadioButtonList1.SelectedValue = 3 Then
            xlWorkSheet.Cells(1, 1) = "交通部高速公路局中區養護工程分局(業務科)"
        End If
        xlWorkSheet.Cells(3, 1) = Me.DropDownList1.SelectedValue + "年"
        xlWorkSheet.Cells(5, 1) = Me.DropDownList2.SelectedValue
        'data.SelectCommand = "SELECT TOP 9 序號, 郵寄種類 FROM 大宗郵件執據_郵寄種類 ORDER BY 排序"
        data.SelectCommand = "SELECT TOP 9 序號, 郵寄種類 FROM 大宗郵件執據_郵寄種類 where 件數>0 ORDER BY 排序"
        'data.SelectCommand = "SELECT TOP 12 序號, 郵寄種類 FROM 大宗郵件執據_郵寄種類 ORDER BY 排序"
        data.DataBind()
        data_dv1 = data.Select(New DataSourceSelectArguments)
        For j As Long = 0 To data_dv1.Count - 1
            '標題
            xlWorkSheet.Cells(3, 3 + (j<<1)) = Trim(data_dv1(j)(1).ToString())
        Next
        Dim tmp As String
        For i As Long = 0 To data_dv.Count - 1
            _日 = data_dv(i)(0).ToString()
            xlWorkSheet.Cells(i + 5, 2) = _日
            For j As Long = 0 To data_dv1.Count - 1
                _序號 = data_dv1(j)(0).ToString()
                _件數 = 0
                If Me.RadioButtonList1.SelectedValue = 1 Then
                    data.SelectCommand = "SELECT SUM(件數) AS a FROM 大宗郵件執據 where  (郵寄種類='" & _序號 & "') and 年='" & Me.DropDownList1.SelectedValue & "' and 月='" & Me.DropDownList2.SelectedValue & "' and 日='" & _日 & "'"
                End If
                If Me.RadioButtonList1.SelectedValue = 2 Then
                    data.SelectCommand = "SELECT SUM(件數) AS a FROM 大宗郵件執據 where 收費小組=0 and  (郵寄種類='" & _序號 & "') and 年='" & Me.DropDownList1.SelectedValue & "' and 月='" & Me.DropDownList2.SelectedValue & "' and 日='" & _日 & "'"
                End If
                If Me.RadioButtonList1.SelectedValue = 3 Then
                    data.SelectCommand = "SELECT SUM(件數) AS a FROM 大宗郵件執據 where  (收費小組=1 or 收費小組=-1) and  (郵寄種類='" & _序號 & "') and 年='" & Me.DropDownList1.SelectedValue & "' and 月='" & Me.DropDownList2.SelectedValue & "' and 日='" & _日 & "'"
                End If
                data.DataBind()
                data_dv2 = data.Select(New DataSourceSelectArguments)
                Try
                    _件數 = data_dv2(0)(0).ToString()
                Catch ex As Exception
                End Try
                If _件數 > 0 Then
                    xlWorkSheet.Cells(i + 5, (j<<1) + 3) = _件數
                End If
            Next
            tmp = (i + 5).ToString
            xlWorkSheet.Cells(i + 5, 21) = "=sum(c"+tmp+",e"+tmp+",g"+tmp+",i"+tmp+",k"+tmp+",m"+tmp+",o"+tmp+",q"+tmp+",s"+tmp+")"
            dat = "a" + (i + 5).ToString + ":l" + (i + 5).ToString
            xlWorkSheet.Range(dat).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle = 1
            xlWorkSheet.Range(dat).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = 1
            For j As Long = 1 To 12
                xlWorkSheet.Cells(i + 5, j).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = 1
                xlWorkSheet.Cells(i + 5, j).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle = 1
            Next
        Next
        For i As Long = 0 To data_dv.Count - 1
            _日 = data_dv(i)(0).ToString()
            xlWorkSheet.Cells(i + 5, 2) = _日
            For j As Long = 0 To data_dv1.Count - 1
                _序號 = data_dv1(j)(0).ToString()
                _郵資 = 0
                If Me.RadioButtonList1.SelectedValue = 1 Then
                    data.SelectCommand = "SELECT SUM(郵資) AS a FROM 大宗郵件執據 where  (郵寄種類='" & _序號 & "') and 年='" & Me.DropDownList1.SelectedValue & "' and 月='" & Me.DropDownList2.SelectedValue & "' and 日='" & _日 & "'"
                End If
                If Me.RadioButtonList1.SelectedValue = 2 Then
                    data.SelectCommand = "SELECT SUM(郵資) AS a FROM 大宗郵件執據 where 收費小組=0 and  (郵寄種類='" & _序號 & "') and 年='" & Me.DropDownList1.SelectedValue & "' and 月='" & Me.DropDownList2.SelectedValue & "' and 日='" & _日 & "'"
                End If
                If Me.RadioButtonList1.SelectedValue = 3 Then
                    data.SelectCommand = "SELECT SUM(郵資) AS a FROM 大宗郵件執據 where  (收費小組=1 or 收費小組=-1) and  (郵寄種類='" & _序號 & "') and 年='" & Me.DropDownList1.SelectedValue & "' and 月='" & Me.DropDownList2.SelectedValue & "' and 日='" & _日 & "'"
                End If
                data.DataBind()
                data_dv2 = data.Select(New DataSourceSelectArguments)
                Try
                    _郵資 = data_dv2(0)(0).ToString()
                Catch ex As Exception
                End Try
                If _郵資 > 0 Then
                    xlWorkSheet.Cells(i + 5, (j<<1) + 4) = _郵資
                End If
            Next
            tmp = (i + 5).ToString
            xlWorkSheet.Cells(i + 5, 22) = "=sum(d"+tmp+",f"+tmp+",h"+tmp+",j"+tmp+",l"+tmp+",n"+tmp+",p"+tmp+",r"+tmp+",t"+tmp+")"
            dat = "a" + (i + 5).ToString + ":l" + (i + 5).ToString
            xlWorkSheet.Range(dat).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle = 1
            xlWorkSheet.Range(dat).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = 1
            For j As Long = 1 To 12
                xlWorkSheet.Cells(i + 5, j).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = 1
                xlWorkSheet.Cells(i + 5, j).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle = 1
            Next
        Next
        dat = "a" + (data_dv.Count + 5).ToString + ":b" + (data_dv.Count + 5).ToString
        xlWorkSheet.Range(dat).MergeCells = True '合併
        xlWorkSheet.Range(dat).HorizontalAlignment = -4108 '置中
        xlWorkSheet.Cells(data_dv.Count + 5, 1) = "合計"
        xlWorkSheet.Cells(data_dv.Count + 5, 3) = "=sum(c5:c" + (data_dv.Count + 4).ToString + ")"
        xlWorkSheet.Cells(data_dv.Count + 5, 4) = "=sum(d5:d" + (data_dv.Count + 4).ToString + ")"
        xlWorkSheet.Cells(data_dv.Count + 5, 5) = "=sum(e5:e" + (data_dv.Count + 4).ToString + ")"
        xlWorkSheet.Cells(data_dv.Count + 5, 6) = "=sum(f5:f" + (data_dv.Count + 4).ToString + ")"
        xlWorkSheet.Cells(data_dv.Count + 5, 7) = "=sum(g5:g" + (data_dv.Count + 4).ToString + ")"
        xlWorkSheet.Cells(data_dv.Count + 5, 8) = "=sum(h5:h" + (data_dv.Count + 4).ToString + ")"
        xlWorkSheet.Cells(data_dv.Count + 5, 9) = "=sum(i5:i" + (data_dv.Count + 4).ToString + ")"
        xlWorkSheet.Cells(data_dv.Count + 5, 10) = "=sum(j5:j" + (data_dv.Count + 4).ToString + ")"
        xlWorkSheet.Cells(data_dv.Count + 5, 11) = "=sum(k5:k" + (data_dv.Count + 4).ToString + ")"
        xlWorkSheet.Cells(data_dv.Count + 5, 12) = "=sum(l5:l" + (data_dv.Count + 4).ToString + ")"
        xlWorkSheet.Cells(data_dv.Count + 5, 13) = "=sum(m5:m" + (data_dv.Count + 4).ToString + ")"
        xlWorkSheet.Cells(data_dv.Count + 5, 14) = "=sum(n5:n" + (data_dv.Count + 4).ToString + ")"
        xlWorkSheet.Cells(data_dv.Count + 5, 15) = "=sum(o5:o" + (data_dv.Count + 4).ToString + ")"
        xlWorkSheet.Cells(data_dv.Count + 5, 16) = "=sum(p5:p" + (data_dv.Count + 4).ToString + ")"
        xlWorkSheet.Cells(data_dv.Count + 5, 17) = "=sum(q5:q" + (data_dv.Count + 4).ToString + ")"
        xlWorkSheet.Cells(data_dv.Count + 5, 18) = "=sum(r5:r" + (data_dv.Count + 4).ToString + ")"
        xlWorkSheet.Cells(data_dv.Count + 5, 19) = "=sum(s5:s" + (data_dv.Count + 4).ToString + ")"
        xlWorkSheet.Cells(data_dv.Count + 5, 20) = "=sum(t5:t" + (data_dv.Count + 4).ToString + ")"
        xlWorkSheet.Cells(data_dv.Count + 5, 21) = "=sum(u5:u" + (data_dv.Count + 4).ToString + ")"
        xlWorkSheet.Cells(data_dv.Count + 5, 22) = "=sum(v5:v" + (data_dv.Count + 4).ToString + ")"
        dat = "a" + (data_dv.Count + 6).ToString + ":V" + (data_dv.Count + 6).ToString
        xlWorkSheet.Range(dat).MergeCells = True '合併
        xlWorkSheet.Cells(data_dv.Count + 6, 1) = Trim(Me.TextBox1.Text)
        dat = "a" + (data_dv.Count + 7).ToString + ":V" + (data_dv.Count + 7).ToString
        xlWorkSheet.Range(dat).MergeCells = True '合併
        xlWorkSheet.Cells(data_dv.Count + 7, 1) = Trim(Me.TextBox3.Text)
        dat = "a" + (data_dv.Count + 8).ToString + ":V" + (data_dv.Count + 8).ToString
        xlWorkSheet.Range(dat).MergeCells = True '合併
        xlWorkSheet.Cells(data_dv.Count + 8, 1) = Trim(Me.TextBox4.Text)
        dat = "a" + (data_dv.Count + 9).ToString + ":V" + (data_dv.Count + 9).ToString
        xlWorkSheet.Range(dat).MergeCells = True '合併
        xlWorkSheet.Cells(data_dv.Count + 9, 1) = Trim(Me.TextBox5.Text)

        xlWorkSheet.Range("c5:c" & (data_dv.Count + 5).ToString()).NumberFormat = "#件;#件;0件;@"
        xlWorkSheet.Range("d5:d" & (data_dv.Count + 5).ToString()).NumberFormat = "#元;#元;0元;@"
        xlWorkSheet.Range("e5:e" & (data_dv.Count + 5).ToString()).NumberFormat = "#件;#件;0件;@"
        xlWorkSheet.Range("f5:f" & (data_dv.Count + 5).ToString()).NumberFormat = "#元;#元;0元;@"
        xlWorkSheet.Range("g5:g" & (data_dv.Count + 5).ToString()).NumberFormat = "#件;#件;0件;@"
        xlWorkSheet.Range("h5:h" & (data_dv.Count + 5).ToString()).NumberFormat = "#元;#元;0元;@"
        xlWorkSheet.Range("i5:i" & (data_dv.Count + 5).ToString()).NumberFormat = "#件;#件;0件;@"
        xlWorkSheet.Range("j5:j" & (data_dv.Count + 5).ToString()).NumberFormat = "#元;#元;0元;@"
        xlWorkSheet.Range("k5:k" & (data_dv.Count + 5).ToString()).NumberFormat = "#件;#件;0件;@"
        xlWorkSheet.Range("l5:l" & (data_dv.Count + 5).ToString()).NumberFormat = "#元;#元;0元;@"
        xlWorkSheet.Range("m5:m" & (data_dv.Count + 5).ToString()).NumberFormat = "#件;#件;0件;@"
        xlWorkSheet.Range("n5:n" & (data_dv.Count + 5).ToString()).NumberFormat = "#元;#元;0元;@"
        xlWorkSheet.Range("o5:o" & (data_dv.Count + 5).ToString()).NumberFormat = "#件;#件;0件;@"
        xlWorkSheet.Range("p5:p" & (data_dv.Count + 5).ToString()).NumberFormat = "#元;#元;0元;@"
        xlWorkSheet.Range("q5:q" & (data_dv.Count + 5).ToString()).NumberFormat = "#件;#件;0件;@"
        xlWorkSheet.Range("r5:r" & (data_dv.Count + 5).ToString()).NumberFormat = "#元;#元;0元;@"
        xlWorkSheet.Range("s5:s" & (data_dv.Count + 5).ToString()).NumberFormat = "#件;#件;0件;@"
        xlWorkSheet.Range("t5:t" & (data_dv.Count + 5).ToString()).NumberFormat = "#元;#元;0元;@"
        xlWorkSheet.Range("u5:u" & (data_dv.Count + 5).ToString()).NumberFormat = "#件;#件;0件;@"
        xlWorkSheet.Range("v5:v" & (data_dv.Count + 5).ToString()).NumberFormat = "#元;#元;0元;@"

        xlWorkSheet.Range("c5:c" & (data_dv.Count + 4).ToString()).NumberFormat = "#件;#件;;@"
        xlWorkSheet.Range("d5:d" & (data_dv.Count + 4).ToString()).NumberFormat = "#元;#元;;@"
        xlWorkSheet.Range("e5:e" & (data_dv.Count + 4).ToString()).NumberFormat = "#件;#件;;@"
        xlWorkSheet.Range("f5:f" & (data_dv.Count + 4).ToString()).NumberFormat = "#元;#元;;@"
        xlWorkSheet.Range("g5:g" & (data_dv.Count + 4).ToString()).NumberFormat = "#件;#件;;@"
        xlWorkSheet.Range("h5:h" & (data_dv.Count + 4).ToString()).NumberFormat = "#元;#元;;@"
        xlWorkSheet.Range("i5:i" & (data_dv.Count + 4).ToString()).NumberFormat = "#件;#件;;@"
        xlWorkSheet.Range("j5:j" & (data_dv.Count + 4).ToString()).NumberFormat = "#元;#元;;@"
        xlWorkSheet.Range("k5:k" & (data_dv.Count + 4).ToString()).NumberFormat = "#件;#件;;@"
        xlWorkSheet.Range("l5:l" & (data_dv.Count + 4).ToString()).NumberFormat = "#元;#元;;@"
        xlWorkSheet.Range("m5:m" & (data_dv.Count + 4).ToString()).NumberFormat = "#件;#件;;@"
        xlWorkSheet.Range("n5:n" & (data_dv.Count + 4).ToString()).NumberFormat = "#元;#元;;@"
        xlWorkSheet.Range("o5:o" & (data_dv.Count + 4).ToString()).NumberFormat = "#件;#件;;@"
        xlWorkSheet.Range("p5:p" & (data_dv.Count + 4).ToString()).NumberFormat = "#元;#元;;@"
        xlWorkSheet.Range("q5:q" & (data_dv.Count + 4).ToString()).NumberFormat = "#件;#件;;@"
        xlWorkSheet.Range("r5:r" & (data_dv.Count + 4).ToString()).NumberFormat = "#元;#元;;@"
        xlWorkSheet.Range("s5:s" & (data_dv.Count + 4).ToString()).NumberFormat = "#件;#件;;@"
        xlWorkSheet.Range("t5:t" & (data_dv.Count + 4).ToString()).NumberFormat = "#元;#元;;@"
        'xlWorkSheet.Range("u5:u" & (data_dv.Count + 4).ToString()).NumberFormat = "#件;#件;;@"
        'xlWorkSheet.Range("v5:v" & (data_dv.Count + 4).ToString()).NumberFormat = "#元;#元;;@"

        xlWorkSheet.Range("A3:V" & (data_dv.Count + 5).ToString()).Borders.LineStyle = 1
        xlWorkSheet.Range("A" & (data_dv.Count + 6).ToString() & ":V" & (data_dv.Count + 6).ToString()).Borders(8).LineStyle = 1
        xlWorkSheet.Range("A" & (data_dv.Count + 6).ToString() & ":V" & (data_dv.Count + 9).ToString()).Borders(7).LineStyle = 1
        xlWorkSheet.Range("A" & (data_dv.Count + 6).ToString() & ":V" & (data_dv.Count + 9).ToString()).Borders(10).LineStyle = 1
        xlWorkSheet.Range("A" & (data_dv.Count + 9).ToString() & ":V" & (data_dv.Count + 9).ToString()).Borders(9).LineStyle = 1
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
        Dim downloadfilename = "郵務種類日報表.xls"
        Response.AddHeader("Content-Disposition", "attachment;FileName=" + Uri.EscapeDataString(downloadfilename))
        Response.WriteFile(MyExcel)
        System.Web.HttpContext.Current.ApplicationInstance.CompleteRequest()
        Response.Flush()
        System.IO.File.Delete(MyExcel)
        Response.End()
    End Sub
    Protected Sub Button5_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button5.Click
        data.ConnectionString = con_14
        Dim _GUID As String = Guid.NewGuid().ToString("N")
        Dim MyExcel As String = "C:\大宗郵件\Excel\Temp\" & _GUID & ".xls"
        System.IO.File.Copy("C:\大宗郵件\Excel\每日郵資統計表.xls", MyExcel)
        Dim xlApp As Excel.ApplicationClass
        xlApp = New Excel.ApplicationClass()
        xlApp.DisplayAlerts = False
        xlApp.ScreenUpdating = false
        xlApp.EnableEvents = false
        Dim xlWorkBook As Excel.Workbook = xlApp.Workbooks.Open(MyExcel, 0, False, 5, "", "", True, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", False, False, 0, True, 1, 0)
        Dim xlWorkSheet As Excel.Worksheet = xlWorkBook.ActiveSheet
        
        data.SelectCommand = "SELECT distinct 日 FROM  大宗郵件執據 where 年='" & Me.DropDownList1.SelectedValue & "' and 月='" & Me.DropDownList2.SelectedValue & "'ORDER BY 日"
        data.DataBind()
        data_dv = data.Select(New DataSourceSelectArguments)
        Dim _日 As String = ""
        Dim _月 As String = ""
        Dim _郵資 As Long = 0
        If Me.RadioButtonList1.SelectedValue = 1 Then
            xlWorkSheet.Cells(1, 1) = Me.DropDownList1.SelectedValue + "年每日郵資統計"
        End If
        If Me.RadioButtonList1.SelectedValue = 2 Then
            xlWorkSheet.Cells(1, 1) = Me.DropDownList1.SelectedValue + "年每日郵資統計(中分局本部)"
        End If
        If Me.RadioButtonList1.SelectedValue = 3 Then
            xlWorkSheet.Cells(1, 1) = Me.DropDownList1.SelectedValue + "年每日郵資統計(業務科)"
        End If
        For i As Long = 1 To 31
            _日 = i.ToString("00")
            For j As Long = 1 To 12
                _月 = j.ToString("00")
                _郵資 = 0
                If Me.RadioButtonList1.SelectedValue = 1 Then
                    data.SelectCommand = "SELECT SUM(郵資) AS a FROM 大宗郵件執據 where 年='" & Me.DropDownList1.SelectedValue & "' and 月='" & _月 & "' and 日='" & _日 & "'"
                End If
                If Me.RadioButtonList1.SelectedValue = 2 Then
                    data.SelectCommand = "SELECT SUM(郵資) AS a FROM 大宗郵件執據 where 收費小組=0 and 年='" & Me.DropDownList1.SelectedValue & "' and 月='" & _月 & "' and 日='" & _日 & "'"
                End If
                If Me.RadioButtonList1.SelectedValue = 3 Then
                    data.SelectCommand = "SELECT SUM(郵資) AS a FROM 大宗郵件執據 where  (收費小組=1 or 收費小組=-1) and 年='" & Me.DropDownList1.SelectedValue & "' and 月='" & _月 & "' and 日='" & _日 & "'"
                End If
                data.DataBind()
                data_dv = data.Select(New DataSourceSelectArguments)
                Try
                    _郵資 = data_dv(0)(0).ToString()
                Catch ex As Exception
                End Try
                If _郵資 > 0 Then
                    xlWorkSheet.Cells(i + 2, j + 1) = _郵資
                End If
            Next
        Next
        xlWorkSheet.Cells(35, 1) = Trim(Me.TextBox1.Text)
        xlWorkSheet.Cells(36, 1) = Trim(Me.TextBox3.Text)
        xlWorkSheet.Cells(37, 1) = Trim(Me.TextBox4.Text)
        xlWorkSheet.Cells(38, 1) = Trim(Me.TextBox5.Text)
        
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
        Dim downloadfilename = "每日郵資統計表.xls"
        Response.AddHeader("Content-Disposition", "attachment;FileName=" + Uri.EscapeDataString(downloadfilename))
        Response.WriteFile(MyExcel)
        System.Web.HttpContext.Current.ApplicationInstance.CompleteRequest()
        Response.Flush()
        System.IO.File.Delete(MyExcel)
        Response.End()
    End Sub
    Protected Sub Button6_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button6.Click
        data.ConnectionString = con_14
        Dim _GUID As String = Guid.NewGuid().ToString("N")
        Dim MyExcel As String = "C:\大宗郵件\Excel\Temp\" & _GUID & ".xls"
        System.IO.File.Copy("C:\大宗郵件\Excel\每日件數統計表.xls", MyExcel)
        Dim xlApp As Excel.ApplicationClass
        xlApp = New Excel.ApplicationClass()
        xlApp.DisplayAlerts = False
        xlApp.ScreenUpdating = false
        xlApp.EnableEvents = false
        Dim xlWorkBook As Excel.Workbook = xlApp.Workbooks.Open(MyExcel, 0, False, 5, "", "", True, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", False, False, 0, True, 1, 0)
        Dim xlWorkSheet As Excel.Worksheet = xlWorkBook.ActiveSheet
        
        data.SelectCommand = "SELECT distinct 日 FROM  大宗郵件執據 where 年='" & Me.DropDownList1.SelectedValue & "' and 月='" & Me.DropDownList2.SelectedValue & "'ORDER BY 日"
        data.DataBind()
        data_dv = data.Select(New DataSourceSelectArguments)
        Dim _日 As String = ""
        Dim _月 As String = ""
        Dim _件數 As Long = 0
        If Me.RadioButtonList1.SelectedValue = 1 Then
            xlWorkSheet.Cells(1, 1) = Me.DropDownList1.SelectedValue + "年每日件數統計"
        End If
        If Me.RadioButtonList1.SelectedValue = 2 Then
            xlWorkSheet.Cells(1, 1) = Me.DropDownList1.SelectedValue + "年每日件數統計(中分局本部)"
        End If
        If Me.RadioButtonList1.SelectedValue = 3 Then
            xlWorkSheet.Cells(1, 1) = Me.DropDownList1.SelectedValue + "年每日件數統計(業務科)"
        End If
        For i As Long = 1 To 31
            _日 = i.ToString("00")
            For j As Long = 1 To 12
                _月 = j.ToString("00")
                _件數 = 0
                If Me.RadioButtonList1.SelectedValue = 1 Then
                    data.SelectCommand = "SELECT SUM(件數) AS a FROM 大宗郵件執據 where 年='" & Me.DropDownList1.SelectedValue & "' and 月='" & _月 & "' and 日='" & _日 & "'"
                End If
                If Me.RadioButtonList1.SelectedValue = 2 Then
                    data.SelectCommand = "SELECT SUM(件數) AS a FROM 大宗郵件執據 where 收費小組=0 and 年='" & Me.DropDownList1.SelectedValue & "' and 月='" & _月 & "' and 日='" & _日 & "'"
                End If
                If Me.RadioButtonList1.SelectedValue = 3 Then
                    data.SelectCommand = "SELECT SUM(件數) AS a FROM 大宗郵件執據 where  (收費小組=1 or 收費小組=-1) and 年='" & Me.DropDownList1.SelectedValue & "' and 月='" & _月 & "' and 日='" & _日 & "'"
                End If
                data.DataBind()
                data_dv = data.Select(New DataSourceSelectArguments)
                Try
                    _件數 = data_dv(0)(0).ToString()
                Catch ex As Exception
                End Try
                If _件數 > 0 Then
                    xlWorkSheet.Cells(i + 2, j + 1) = _件數
                End If
            Next
        Next
        xlWorkSheet.Cells(35, 1) = Trim(Me.TextBox1.Text)
        xlWorkSheet.Cells(36, 1) = Trim(Me.TextBox3.Text)
        xlWorkSheet.Cells(37, 1) = Trim(Me.TextBox4.Text)
        xlWorkSheet.Cells(38, 1) = Trim(Me.TextBox5.Text)
        
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
        Dim downloadfilename = "每日件數統計表.xls"
        Response.AddHeader("Content-Disposition", "attachment;FileName=" + Uri.EscapeDataString(downloadfilename))
        Response.WriteFile(MyExcel)
        System.Web.HttpContext.Current.ApplicationInstance.CompleteRequest()
        Response.Flush()
        System.IO.File.Delete(MyExcel)
        Response.End()
    End Sub
    Protected Sub Button7_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button7.Click
        data.ConnectionString = con_14
        Dim _GUID As String = Guid.NewGuid().ToString("N")
        Dim MyExcel As String = "C:\大宗郵件\Excel\Temp\" & _GUID & ".xls"
        System.IO.File.Copy("C:\大宗郵件\Excel\每日郵資暨件數統計表.xls", MyExcel)
        Dim xlApp As Excel.ApplicationClass
        xlApp = New Excel.ApplicationClass()
        xlApp.DisplayAlerts = False
        xlApp.ScreenUpdating = false
        xlApp.EnableEvents = false
        Dim xlWorkBook As Excel.Workbook = xlApp.Workbooks.Open(MyExcel, 0, False, 5, "", "", True, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", False, False, 0, True, 1, 0)
        Dim xlWorkSheet As Excel.Worksheet = xlWorkBook.ActiveSheet
        
        data.SelectCommand = "SELECT distinct 日 FROM  大宗郵件執據 where 年='" & Me.DropDownList1.SelectedValue & "' and 月='" & Me.DropDownList2.SelectedValue & "'ORDER BY 日"
        data.DataBind()
        data_dv = data.Select(New DataSourceSelectArguments)
        Dim _日 As String = ""
        Dim _月 As String = ""
        Dim _郵資 As Long = 0
        If Me.RadioButtonList1.SelectedValue = 1 Then
            xlWorkSheet.Cells(1, 1) = Me.DropDownList1.SelectedValue + "年每日郵資暨件數統計"
        End If
        If Me.RadioButtonList1.SelectedValue = 2 Then
            xlWorkSheet.Cells(1, 1) = Me.DropDownList1.SelectedValue + "年每日郵資暨件數統計(中分局本部)"
        End If
        If Me.RadioButtonList1.SelectedValue = 3 Then
            xlWorkSheet.Cells(1, 1) = Me.DropDownList1.SelectedValue + "年每日郵資暨件數統計(業務科)"
        End If
        For i As Long = 1 To 31
            _日 = i.ToString("00")
            For j As Long = 1 To 12
                _月 = j.ToString("00")
                _郵資 = 0
                If Me.RadioButtonList1.SelectedValue = 1 Then
                    data.SelectCommand = "SELECT SUM(郵資) AS a FROM 大宗郵件執據 where 年='" & Me.DropDownList1.SelectedValue & "' and 月='" & _月 & "' and 日='" & _日 & "'"
                End If
                If Me.RadioButtonList1.SelectedValue = 2 Then
                    data.SelectCommand = "SELECT SUM(郵資) AS a FROM 大宗郵件執據 where 收費小組=0 and 年='" & Me.DropDownList1.SelectedValue & "' and 月='" & _月 & "' and 日='" & _日 & "'"
                End If
                If Me.RadioButtonList1.SelectedValue = 3 Then
                    data.SelectCommand = "SELECT SUM(郵資) AS a FROM 大宗郵件執據 where  (收費小組=1 or 收費小組=-1) and 年='" & Me.DropDownList1.SelectedValue & "' and 月='" & _月 & "' and 日='" & _日 & "'"
                End If
                data.DataBind()
                data_dv = data.Select(New DataSourceSelectArguments)
                Try
                    _郵資 = data_dv(0)(0).ToString()
                Catch ex As Exception
                End Try
                If _郵資 > 0 Then
                    xlWorkSheet.Cells(i + 2, (j<<1) + 1) = _郵資
                End If
            Next
        Next
        Dim _件數 As Long = 0
        If Me.RadioButtonList1.SelectedValue = 1 Then
            xlWorkSheet.Cells(1, 1) = Me.DropDownList1.SelectedValue + "年每日件數統計"
        End If
        If Me.RadioButtonList1.SelectedValue = 2 Then
            xlWorkSheet.Cells(1, 1) = Me.DropDownList1.SelectedValue + "年每日件數統計(中分局本部)"
        End If
        If Me.RadioButtonList1.SelectedValue = 3 Then
            xlWorkSheet.Cells(1, 1) = Me.DropDownList1.SelectedValue + "年每日件數統計(業務科)"
        End If
        For i As Long = 1 To 31
            _日 = i.ToString("00")
            For j As Long = 1 To 12
                _月 = j.ToString("00")
                _件數 = 0
                If Me.RadioButtonList1.SelectedValue = 1 Then
                    data.SelectCommand = "SELECT SUM(件數) AS a FROM 大宗郵件執據 where 年='" & Me.DropDownList1.SelectedValue & "' and 月='" & _月 & "' and 日='" & _日 & "'"
                End If
                If Me.RadioButtonList1.SelectedValue = 2 Then
                    data.SelectCommand = "SELECT SUM(件數) AS a FROM 大宗郵件執據 where 收費小組=0 and 年='" & Me.DropDownList1.SelectedValue & "' and 月='" & _月 & "' and 日='" & _日 & "'"
                End If
                If Me.RadioButtonList1.SelectedValue = 3 Then
                    data.SelectCommand = "SELECT SUM(件數) AS a FROM 大宗郵件執據 where  (收費小組=1 or 收費小組=-1) and 年='" & Me.DropDownList1.SelectedValue & "' and 月='" & _月 & "' and 日='" & _日 & "'"
                End If
                data.DataBind()
                data_dv = data.Select(New DataSourceSelectArguments)
                Try
                    _件數 = data_dv(0)(0).ToString()
                Catch ex As Exception
                End Try
                If _件數 > 0 Then
                    xlWorkSheet.Cells(i + 2, (j<<1) + 0) = _件數
                End If
            Next
        Next
        xlWorkSheet.Cells(35, 1) = Trim(Me.TextBox1.Text)
        xlWorkSheet.Cells(36, 1) = Trim(Me.TextBox3.Text)
        xlWorkSheet.Cells(37, 1) = Trim(Me.TextBox4.Text)
        xlWorkSheet.Cells(38, 1) = Trim(Me.TextBox5.Text)
        
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
        Dim downloadfilename = "每日郵資暨件數統計表.xls"
        Response.AddHeader("Content-Disposition", "attachment;FileName=" + Uri.EscapeDataString(downloadfilename))
        Response.WriteFile(MyExcel)
        System.Web.HttpContext.Current.ApplicationInstance.CompleteRequest()
        Response.Flush()
        System.IO.File.Delete(MyExcel)
        Response.End()
    End Sub
End Class


