Imports Microsoft.Office.Interop
Imports System.Diagnostics
Partial Class 大宗郵件執據_查詢
    Inherits System.Web.UI.Page
    'Dim con_wf2 As String = "Data Source=edocsql.freeway.gov.tw\SQL2012;Initial Catalog=CFW_wf2;User ID=CFWedocdb01;Password=99$edocdb$CFW"
    Dim con_wf2 As String = "Data Source=edocsqlplus.freeway.gov.tw\SQL2019,54399;Initial Catalog=CFW_wf2;User ID=CFWedocdb01;Password=99$edocdb$CFW"
    Dim con_14 As String = "Data Source=10.52.0.178;Initial Catalog=大宗郵件;User ID=qaz;Password=1qaz@WSX"
    Dim data As New SqlDataSource
    Dim data_dv As Data.DataView
    Dim data_dv1 As Data.DataView
    Dim data_dv2 As Data.DataView
    '開始 金額轉換大寫
    Public Shared Function UpperMoney(ByVal Money As String) As String
    Money = Money.Replace("-", "").Replace(".", "")
    Dim Number As String = "零壹貳叄肆伍陸柒捌玖"
    Dim Unit As String = "元拾佰仟萬拾佰仟億拾佰仟萬"
    Dim str As String = ""
    For i As Long = 0 To Money.Length - 1
    Dim c As String = Money.Chars(i)
    Dim Index As Long = Money.Length - 1 - i
    str &= Number(c) & "　"　& Unit(Index) & "　"
    Next
    str = str & "整"
    Return str
    End Function
    '結尾 金額轉換大寫
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not Page.IsPostBack Then
            Dim r1, r2, r3 As Long
            Dim card_id As String
            card_id = Request.ServerVariables("REMOTE_HOST")
            If Len(Trim(card_id)) <= 0 Then
                card_id = System.Security.Principal.WindowsIdentity.GetCurrent().Name
            End If
            r1 = Len(card_id)
            For i As Long = 1 To r1
                If Mid(card_id, i, 1) = "\" Then
                    r2 = i
                End If
            Next
            r3 = r1 - r2
            card_id = Right(card_id, r3)
            Me.TextBox2.Text = card_id
            
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
            Next
            Me.DropDownList1.DataBind()
            Me.DropDownList1.SelectedValue = _年
            For i As Long = 0 To 11
                Me.DropDownList2.Items.Add((i + 1).ToString("00"))
                Me.DropDownList2.Items(i).Value = (i + 1).ToString("00")
            Next
            Me.DropDownList2.DataBind()
            Me.DropDownList2.SelectedValue = _月
            For i As Long = 0 To DateTime.DaysInMonth(Now.Year, Now.Month) - 1
                Me.DropDownList3.Items.Add((i + 1).ToString("00"))
                Me.DropDownList3.Items(i).Value = (i + 1).ToString("00")
            Next
            Me.DropDownList3.DataBind()
            Me.DropDownList3.SelectedValue = _日
            
            refresh_dropdownlist5()
        End If
    End Sub
    Protected Sub refresh_dropdownlist5()
        Dim temp = Me.DropDownList5.SelectedValue
        Me.DropDownList5.Items.Clear()
        data.ConnectionString = con_14
        data.SelectCommand = "SELECT * FROM (SELECT DISTINCT b.郵寄種類, b.序號, b.排序 FROM 大宗郵件執據 a INNER JOIN 大宗郵件執據_郵寄種類 b ON a.郵寄種類 = b.序號 WHERE a.年='" & Me.DropDownList1.SelectedValue & "' AND a.月='" & Me.DropDownList2.SelectedValue & "' AND a.日='" & Me.DropDownList3.SelectedValue & "') AS c ORDER BY c.排序"
        data.DataBind()
        data_dv = data.Select(New DataSourceSelectArguments)
        For i As Long = 0 To data_dv.Count - 1
            Me.DropDownList5.Items.Add(Trim(data_dv(i)(0).ToString()))
            Me.DropDownList5.Items(i).Value = Trim(data_dv(i)(1).ToString())
        Next
        Me.DropDownList5.Items.Add("全部")
        Me.DropDownList5.Items(data_dv.Count).Value = 0
        Me.Label3.Text = ""
        DropDownList5.SelectedIndex = DropDownList5.Items.IndexOf(DropDownList5.Items.FindByValue(temp))
    End Sub
    Protected Sub DropDownList1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DropDownList1.SelectedIndexChanged
        Me.DropDownList3.Items.Clear()
        For i As Long = 0 To DateTime.DaysInMonth((Val(Me.DropDownList1.SelectedValue) + 1911), Val(Me.DropDownList2.SelectedValue)) - 1
            Me.DropDownList3.Items.Add((i + 1).ToString("00"))
            Me.DropDownList3.Items(i).Value = (i + 1).ToString("00")
        Next
        data.ConnectionString = con_14
        data.UpdateCommand = "update 大宗郵件執據_操作者 set 年 = '" & Me.DropDownList1.SelectedValue & "', 月 = '" & Me.DropDownList2.SelectedValue & "', 日 = '" & Me.DropDownList3.SelectedValue & "' WHERE 帳號='" & Me.TextBox2.Text & "'"
        data.Update()
        data.DataBind()
        Me.DropDownList3.DataBind()
        Me.GridView2.DataBind()
        refresh_dropdownlist5()
    End Sub
    Protected Sub DropDownList2_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DropDownList2.SelectedIndexChanged
        Me.DropDownList3.Items.Clear()
        For i As Long = 0 To DateTime.DaysInMonth((Val(Me.DropDownList1.SelectedValue) + 1911), Val(Me.DropDownList2.SelectedValue)) - 1
            Me.DropDownList3.Items.Add((i + 1).ToString("00"))
            Me.DropDownList3.Items(i).Value = (i + 1).ToString("00")
        Next
        data.ConnectionString = con_14
        data.UpdateCommand = "update 大宗郵件執據_操作者 set 年 = '" & Me.DropDownList1.SelectedValue & "', 月 = '" & Me.DropDownList2.SelectedValue & "', 日 = '" & Me.DropDownList3.SelectedValue & "' WHERE 帳號='" & Me.TextBox2.Text & "'"
        data.Update()
        data.DataBind()
        Me.DropDownList3.DataBind()
        Me.GridView2.DataBind()
        refresh_dropdownlist5()
    End Sub
    Protected Sub DropDownList3_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DropDownList3.SelectedIndexChanged
        data.ConnectionString = con_14
        data.UpdateCommand = "update 大宗郵件執據_操作者 set 年 = '" & Me.DropDownList1.SelectedValue & "', 月 = '" & Me.DropDownList2.SelectedValue & "', 日 = '" & Me.DropDownList3.SelectedValue & "' WHERE 帳號='" & Me.TextBox2.Text & "'"
        data.Update()
        data.DataBind()
        Me.DropDownList3.DataBind()
        refresh_dropdownlist5()
    End Sub

    Protected Sub GridView2_RowDeleted(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewDeletedEventArgs) Handles GridView2.RowDeleted
        data.ConnectionString = con_14
        Dim _序號 As Long = 0
        Dim _id As Long = 0
        Dim _郵寄種類 As Long = 0
        data.SelectCommand = "SELECT 序號 FROM  大宗郵件執據_郵寄種類  ORDER BY 序號"
        data.DataBind()
        data_dv = data.Select(New DataSourceSelectArguments)
        For i As Long = 0 To data_dv.Count - 1
            _郵寄種類 = data_dv(i)(0).ToString()
            data.SelectCommand = "SELECT id FROM  大宗郵件執據 where 年='" & Me.DropDownList1.SelectedValue & "' and 月='" & Me.DropDownList2.SelectedValue & "' and 日='" & Me.DropDownList3.SelectedValue & "' and 郵寄種類='" & _郵寄種類 & "' ORDER BY 序號"
            data.DataBind()
            data_dv1 = data.Select(New DataSourceSelectArguments)
            For j As Long = 0 To data_dv1.Count - 1
                _序號 = j + 1
                _id = data_dv1(j)(0).ToString()
                data.UpdateCommand = "update 大宗郵件執據 set 序號='" & _序號 & "' where id='" & _id & "'"
                data.Update()
                data.DataBind()
            Next
        Next
        Me.SqlDataSource3.DataBind()
        Me.GridView2.DataBind()
    End Sub
    Protected Sub GridView2_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView2.SelectedIndexChanged
        '存檔
        data.ConnectionString = con_14
        For i = 0 To Me.GridView2.Rows.Count - 1
            Dim _收費小組 As Long = 0
            Dim _id As Long = CType(Me.GridView2.Rows(i).FindControl("Label1"), Label).Text()
            Dim _序號 As String = CType(Me.GridView2.Rows(i).FindControl("TextBox5"), TextBox).Text()
            Dim _掛號號碼 As String = CType(Me.GridView2.Rows(i).FindControl("TextBox15"), TextBox).Text()
            Dim _備註 As String = Trim(CType(Me.GridView2.Rows(i).FindControl("TextBox9"), TextBox).Text())
            Dim _重量 As String = CType(Me.GridView2.Rows(i).FindControl("TextBox7"), TextBox).Text()
            Dim _郵資 As String = CType(Me.GridView2.Rows(i).FindControl("TextBox8"), TextBox).Text()
            Dim _收件人 As String = CType(Me.GridView2.Rows(i).FindControl("TextBox12"), TextBox).Text()
            Dim _文號 As String = CType(Me.GridView2.Rows(i).FindControl("TextBox14"), TextBox).Text()
            Dim _件數 As Long = CType(Me.GridView2.Rows(i).FindControl("TextBox18"), TextBox).Text()
            Dim _郵遞區號 As String = CType(Me.GridView2.Rows(i).FindControl("TextBox56"), TextBox).Text()
            Dim _地址 As String = CType(Me.GridView2.Rows(i).FindControl("TextBox13"), TextBox).Text()
            Dim _郵寄種類 As Long = CType(Me.GridView2.Rows(i).FindControl("DropDownList4"), DropDownList).SelectedValue
            If CType(Me.GridView2.Rows(i).FindControl("CheckBox1"), CheckBox).Checked Then
                _收費小組 = 1
            Else
                _收費小組 = 0
            End If
            Dim _掛號類別 As String = ""
            data.SelectCommand = "SELECT 掛號類別 FROM  大宗郵件執據_郵寄種類 where 序號 ='" & _郵寄種類 & " '"
            data.DataBind()
            data_dv = data.Select(New DataSourceSelectArguments)
            _掛號類別 = Trim(data_dv(0)(0).ToString())
            If Trim(_掛號號碼) <> ""
                _掛號號碼 = Clng(Trim(_掛號號碼)).ToString("000000")
            End If
            data.UpdateCommand = "update 大宗郵件執據 set 序號='" & _序號 & "' , 掛號號碼=NULLIF('" & _掛號號碼 & "', '') , 備註=N'" & _備註 & "' , 重量='" & _重量 & "' , 郵資='" & _郵資 & "' , 郵寄種類='" & _郵寄種類 & "' , 收件人=N'" & _收件人 & "' , 郵遞區號=NULLIF('" & _郵遞區號 & "', ''), 地址=N'" & _地址 & "' , 文號='" & _文號 & "' , 掛號類別=NULLIF('" & _掛號類別 & "', '') where id='" & _id & "'"
            data.Update()
            data.DataBind()
            data.UpdateCommand = "update 大宗郵件執據 set 收費小組='" & _收費小組 & "' , 件數='" & _件數 & "' where id='" & _id & "'"
            data.Update()
            data.DataBind()
        Next
        Me.SqlDataSource3.DataBind()
        Me.GridView2.DataBind()
        refresh_dropdownlist5()
    End Sub
    Protected Sub Button5_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button5.Click
        data.ConnectionString = con_14
        For i = 0 To Me.GridView2.Rows.Count - 1
            Dim _收費小組 As Long = 0
            Dim _id As Long = CType(Me.GridView2.Rows(i).FindControl("Label1"), Label).Text()
            Dim _序號 As String = CType(Me.GridView2.Rows(i).FindControl("TextBox5"), TextBox).Text()
            Dim _掛號號碼 As String = CType(Me.GridView2.Rows(i).FindControl("TextBox15"), TextBox).Text()
            Dim _備註 As String = Trim(CType(Me.GridView2.Rows(i).FindControl("TextBox9"), TextBox).Text())
            Dim _重量 As String = CType(Me.GridView2.Rows(i).FindControl("TextBox7"), TextBox).Text()
            Dim _郵資 As String = CType(Me.GridView2.Rows(i).FindControl("TextBox8"), TextBox).Text()
            Dim _收件人 As String = CType(Me.GridView2.Rows(i).FindControl("TextBox12"), TextBox).Text()
            Dim _文號 As String = CType(Me.GridView2.Rows(i).FindControl("TextBox14"), TextBox).Text()
            Dim _件數 As Long = CType(Me.GridView2.Rows(i).FindControl("TextBox18"), TextBox).Text()
            Dim _郵遞區號 As String = CType(Me.GridView2.Rows(i).FindControl("TextBox56"), TextBox).Text()
            Dim _地址 As String = CType(Me.GridView2.Rows(i).FindControl("TextBox13"), TextBox).Text()
            Dim _郵寄種類 As Long = CType(Me.GridView2.Rows(i).FindControl("DropDownList4"), DropDownList).SelectedValue
            If CType(Me.GridView2.Rows(i).FindControl("CheckBox1"), CheckBox).Checked Then
                _收費小組 = 1
            Else
                _收費小組 = 0
            End If
            Dim _掛號類別 As String = ""
            data.SelectCommand = "SELECT 掛號類別 FROM  大宗郵件執據_郵寄種類 where 序號 ='" & _郵寄種類 & " '"
            data.DataBind()
            data_dv = data.Select(New DataSourceSelectArguments)
            _掛號類別 = Trim(data_dv(0)(0).ToString())
            If Trim(_掛號號碼) <> ""
                _掛號號碼 = Clng(Trim(_掛號號碼)).ToString("000000")
            End If
            data.UpdateCommand = "update 大宗郵件執據 set 序號='" & _序號 & "' , 掛號號碼=NULLIF('" & _掛號號碼 & "', '') , 備註=N'" & _備註 & "' , 重量='" & _重量 & "' , 郵資='" & _郵資 & "' , 郵寄種類='" & _郵寄種類 & "' , 收件人=N'" & _收件人 & "' , 郵遞區號=NULLIF('" & _郵遞區號 & "', ''), 地址=N'" & _地址 & "' , 文號='" & _文號 & "' , 掛號類別=NULLIF('" & _掛號類別 & "', '') where id='" & _id & "'"
            data.Update()
            data.DataBind()
            data.UpdateCommand = "update 大宗郵件執據 set 收費小組='" & _收費小組 & "' , 件數='" & _件數 & "' where id='" & _id & "'"
            data.Update()
            data.DataBind()
        Next
        Me.SqlDataSource3.DataBind()
        Me.GridView2.DataBind()
    End Sub
    Protected Sub Button2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button2.Click
        data.ConnectionString = con_14
        data.SelectCommand = "SELECT 序號,id FROM  大宗郵件執據 where 年='" & Me.DropDownList1.SelectedValue & "' and 月='" & Me.DropDownList2.SelectedValue & "' and 日='" & Me.DropDownList3.SelectedValue & "'  and 郵寄種類 = " & Me.DropDownList5.SelectedValue & " and 收件人 LIKE '%" & Trim(Me.TextBox19.Text) & "%' and 文號 LIKE '%" & Trim(Me.TextBox20.Text) & "%' ORDER BY 序號"
        data.DataBind()
        data_dv = data.Select(New DataSourceSelectArguments)
        Dim _id As Long = 0
        For i As Long = 0 To data_dv.Count - 1
            _id = data_dv(i)(1).ToString()
            data.UpdateCommand = "update 大宗郵件執據 set 序號='" & i + 1 & "' where id='" & _id & "'"
            data.Update()
            data.DataBind()
        Next
        Me.DropDownList3.DataBind()
        Me.GridView2.DataBind()
    End Sub
    Protected Sub DropDownList5_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DropDownList5.SelectedIndexChanged

    End Sub
    Protected Sub Button3_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button3.Click
        If Len(Trim(Me.TextBox1.Text)) > 0 Then
            data.ConnectionString = con_14
            data.SelectCommand = "SELECT a.掛號號碼,a.id FROM 大宗郵件執據 a INNER JOIN 大宗郵件執據_郵寄種類 b ON a.郵寄種類 = b.序號 WHERE a.年='" & Me.DropDownList1.SelectedValue & "' AND a.月='" & Me.DropDownList2.SelectedValue & "' AND a.日='" & Me.DropDownList3.SelectedValue & "' AND (0='" & Me.DropDownList5.SelectedValue & "' OR a.郵寄種類='" & Me.DropDownList5.SelectedValue & "') AND ((a.收件人 IS NULL AND ''='" & Trim(Me.TextBox19.Text) & "') OR a.收件人 LIKE N'%" & Trim(Me.TextBox19.Text) & "%') AND ((a.文號 IS NULL AND ''='" & Trim(Me.TextBox20.Text) & "') OR a.文號 LIKE N'%" & Trim(Me.TextBox20.Text) & "%') ORDER BY b.排序, a.序號"
            data.DataBind()
            data_dv = data.Select(New DataSourceSelectArguments)
            Dim _掛號號碼 As Long = Val(Me.TextBox1.Text)
            Dim _id As Long = 0
            Dim _begin As Long = 1
            Dim _end As Long = data_dv.Count
            Try
                If Trim(Me.TextBox30.Text) <> "" And CLng(Trim(Me.TextBox30.Text)) > _begin
                    _begin = CLng(Trim(Me.TextBox30.Text))
                End If
            Catch
            End Try
            Try
                If Trim(Me.TextBox31.Text) <> "" And CLng(Trim(Me.TextBox31.Text)) < _end
                    _end = CLng(Trim(Me.TextBox31.Text))
                End If
            Catch
            End Try
            For i As Long = _begin To _end
                _id = data_dv(i - 1)(1).ToString()
                data.UpdateCommand = "update 大宗郵件執據 set 掛號號碼='" & _掛號號碼.ToString("000000") & "' where id='" & _id & "'"
                data.Update()
                data.DataBind()
                _掛號號碼 = _掛號號碼 + 1
            Next
            Me.TextBox30.Text = ""
            If Trim(Me.TextBox31.Text) <> ""
                Me.TextBox30.Text = CLng(Trim(Me.TextBox31.Text)) + 1
            End IF
            Me.TextBox31.Text = ""
            Me.TextBox1.Text = ""
            Me.DropDownList3.DataBind()
            Me.GridView2.DataBind()
        End If
    End Sub
    Protected Sub Button8_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button8.Click
        data.ConnectionString = con_14
        data.SelectCommand = "SELECT id FROM  大宗郵件執據 where 年='" & Me.DropDownList1.SelectedValue & "' and 月='" & Me.DropDownList2.SelectedValue & "' and 日='" & Me.DropDownList3.SelectedValue & "'  and 郵寄種類 = " & Me.DropDownList5.SelectedValue & " and 收件人 LIKE '%" & Trim(Me.TextBox19.Text) & "%' and 文號 LIKE '%" & Trim(Me.TextBox20.Text) & "%' ORDER BY 序號"
        data.DataBind()
        data_dv = data.Select(New DataSourceSelectArguments)
        Dim _重量 As Long = Me.TextBox10.Text
        Dim _id As Long = 0
        For i As Long = 0 To data_dv.Count - 1
            _id = data_dv(i)(0).ToString()
            data.UpdateCommand = "update 大宗郵件執據 set 重量='" & _重量 & "' where 重量=0 and  id='" & _id & "'"
            data.Update()
            data.DataBind()
        Next
        Me.DropDownList3.DataBind()
        Me.GridView2.DataBind()

    End Sub
    Protected Sub Button9_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button9.Click
        data.ConnectionString = con_14
        Dim _郵資 As Long = Me.TextBox11.Text
        Dim _id As Long = 0
        Dim _重量 As Long = 0
        If Trim(Me.DropDownList5.SelectedValue) <> "" Then
            data.SelectCommand = "SELECT id FROM  大宗郵件執據 where  郵資=0 and 郵寄種類='" & Me.DropDownList5.SelectedValue & "' and 年='" & Me.DropDownList1.SelectedValue & "' and 月='" & Me.DropDownList2.SelectedValue & "' and 日='" & Me.DropDownList3.SelectedValue & "'"
            data.DataBind()
            data_dv = data.Select(New DataSourceSelectArguments)
            For i As Long = 0 To data_dv.Count - 1
                _id = data_dv(i)(0).ToString()
                data.UpdateCommand = "update 大宗郵件執據 set 郵資='" & _郵資 & "' where id='" & _id & "'"
                data.Update()
                data.DataBind()
            Next
            Me.DropDownList3.DataBind()
            Me.GridView2.DataBind()
        End If
    End Sub
    
    Protected Sub Button10_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button10.Click
        data.ConnectionString = con_14
        '防呆
        'data.SelectCommand = "SELECT 郵局 FROM 大宗郵件執據_郵寄種類 WHERE 序號='" & Me.DropDownList5.SelectedValue & "' AND 郵局=0"
        'data_dv = data.Select(New DataSourceSelectArguments)
        'If data_dv.Count > 0
        '    ScriptManager.RegisterStartupScript(Me, Page.GetType, "Script5", "setTimeout(function() { alert('所選非郵局郵寄種類。'); }, 100);", True)
        '    Exit Sub
        'End If
        
        Dim _GUID As String = Guid.NewGuid().ToString("N")
        Dim MyExcel As String = "C:\大宗郵件\Excel\Temp\" & _GUID & ".xls"
        System.IO.File.Copy("C:\大宗郵件\Excel\交寄大宗函件執據.xls", MyExcel)
        '0元會計入會顯示
        Dim xlApp As Excel.ApplicationClass
        xlApp = New Excel.ApplicationClass()
        xlApp.DisplayAlerts = False
        xlApp.ScreenUpdating = false
        xlApp.EnableEvents = false
        Dim xlWorkBook As Excel.Workbook = xlApp.Workbooks.Open(MyExcel, 0, False, 5, "", "", True, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", False, False, 0, True, 1, 0)
        Dim xlWorkSheet As Excel.Worksheet = xlWorkBook.ActiveSheet
        
        Dim _年 As String = Me.DropDownList1.SelectedValue
        Dim _月 As String = Me.DropDownList2.SelectedValue
        Dim _日 As String = Me.DropDownList3.SelectedValue
        data.SelectCommand = "SELECT a.序號,a.掛號號碼,a.收件人,a.地址,a.文號,a.郵資 FROM 大宗郵件執據 a INNER JOIN 大宗郵件執據_郵寄種類 b ON a.郵寄種類 = b.序號 WHERE a.年='" & Me.DropDownList1.SelectedValue & "' AND a.月='" & Me.DropDownList2.SelectedValue & "' AND a.日='" & Me.DropDownList3.SelectedValue & "' AND (0='" & Me.DropDownList5.SelectedValue & "' OR a.郵寄種類='" & Me.DropDownList5.SelectedValue & "') AND ((a.收件人 IS NULL AND ''='" & Trim(Me.TextBox19.Text) & "') OR a.收件人 LIKE N'%" & Trim(Me.TextBox19.Text) & "%') AND ((a.文號 IS NULL AND ''='" & Trim(Me.TextBox20.Text) & "') OR a.文號 LIKE N'%" & Trim(Me.TextBox20.Text) & "%') ORDER BY b.排序, a.序號"
        data_dv = data.Select(New DataSourceSelectArguments)
        
        xlWorkSheet.Cells(2, 1).Value = "交寄大宗    " & Me.DropDownList5.SelectedItem.Text & "    函件執據"
        xlWorkSheet.Cells(4, 5).Value = "民國" & _年 & "年" & _月 & "月" & _日 & "日"
        
        Dim arr((data_dv.Count * 2) + 100, 8) As Object
        
        Dim i As Long = 6
        Dim j As Long = 0
        For j = 0 To data_dv.Count - 1
            If i Mod 21 = 0 Or i Mod 21 = 1 Or i Mod 21 = 2
                xlWorkSheet.Cells(i, 5).RowHeight = 21
                j = j - 1
            Else
                xlWorkSheet.Range(xlWorkSheet.Cells(i, 5), xlWorkSheet.Cells(i, 6)).MergeCells = True
                arr(i - 6, 0) = j + 1
                arr(i - 6, 1) = data_dv(j)(1)
                arr(i - 6, 2) = data_dv(j)(2)
                arr(i - 6, 3) = data_dv(j)(3)
                arr(i - 6, 4) = data_dv(j)(4)
                arr(i - 6, 7) = data_dv(j)(5)
                data.SelectCommand = "SELECT 重量 FROM  大宗郵件執據_資費表 where 序號='" & Me.DropDownList5.SelectedValue & "' and 郵資 = '" & data_dv(j)(5).ToString() & "'"
                data_dv1 = data.Select(New DataSourceSelectArguments)
                If data_dv1.Count > 0
                    arr(i - 6, 6) = data_dv1(0)(0).toString
                End If
            End If
            i = i + 1
        Next
        For i = i To (((i \ 21) * 21) + 23)
            If i Mod 21 = 0 Or i Mod 21 = 1 Or i Mod 21 = 2
                xlWorkSheet.Cells(i, 5).RowHeight = 21
            Else
                xlWorkSheet.Range(xlWorkSheet.Cells(i, 5), xlWorkSheet.Cells(i, 6)).MergeCells = True
            End If
        Next
        '統計
        i = 6
        Dim _郵資 As Long = 0
        For j = 0 To data_dv.Count - 1
            Select (i Mod 21)
                Case 0
                    arr(i - 6, 4) = "累計函件共"
                    arr(i - 6, 5) = j.ToString() + "件照收無誤"
                    j = j - 1
                Case 1
                    arr(i - 6, 4) = "郵資共計"
                    arr(i - 6, 5) = _郵資.ToString() + "元整"
                    j = j - 1
                Case 2
                    arr(i - 6, 4) = "經辦員簽章"
                    j = j - 1
                Case Else
                    _郵資 = _郵資 + data_dv(j)(5)
            End Select
            i = i + 1
        Next
        For i = i To (((i \ 21) * 21) + 23)
            Select (i Mod 21)
                Case 0
                    arr(i - 6, 4) = "累計函件共"
                    arr(i - 6, 5) = j.ToString() + "件照收無誤"
                Case 1
                    arr(i - 6, 4) = "郵資共計"
                    arr(i - 6, 5) = _郵資.ToString() + "元整"
                Case 2
                    arr(i - 6, 4) = "經辦員簽章"
                Case Else
            End Select
        Next
        
        xlWorkSheet.Range(xlWorkSheet.Cells(6, 1), xlWorkSheet.Cells((data_dv.Count * 2) + 100, 8)).Value = arr
        
        '框線
        i = 6
        For j = 0 To data_dv.Count - 1
            Select (i Mod 21)
                Case 0
                    j = j - 1
                Case 1
                    j = j - 1
                Case 2
                    j = j - 1
                Case 20
                    If i = 20
                        xlWorkSheet.Range(xlWorkSheet.Cells(i - 14, 1), xlWorkSheet.Cells(i, 8)).Borders.LineStyle = 1
                    Else
                        xlWorkSheet.Range(xlWorkSheet.Cells(i - 17, 1), xlWorkSheet.Cells(i, 8)).Borders.LineStyle = 1
                    End If
                Case Else
            End Select
            i = i + 1
        Next
        If i <> ((i \ 21) * 21)
            For i = i To (((i \ 21) * 21) + 23)
                Select (i Mod 21)
                    Case 0
                    Case 1
                    Case 2
                    Case 20
                        If i = 20
                            xlWorkSheet.Range(xlWorkSheet.Cells(i - 14, 1), xlWorkSheet.Cells(i, 8)).Borders.LineStyle = 1
                        Else
                            xlWorkSheet.Range(xlWorkSheet.Cells(i - 17, 1), xlWorkSheet.Cells(i, 8)).Borders.LineStyle = 1
                        End If
                    Case Else
                End Select
            Next
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
        Dim downloadfilename = "交寄大宗函件執據 " & Me.DropDownList2.SelectedValue & Me.DropDownList3.SelectedValue & ".xls"
        Response.AddHeader("Content-Disposition", "attachment;FileName=" + Uri.EscapeDataString(downloadfilename))
        Response.WriteFile(MyExcel)
        System.Web.HttpContext.Current.ApplicationInstance.CompleteRequest()
        Response.Flush()
        System.IO.File.Delete(MyExcel)
        Response.End()
    End Sub
    
    Protected Sub Button50_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button50.Click    
        data.ConnectionString = con_14
        '防呆
        'data.SelectCommand = "SELECT 郵局 FROM 大宗郵件執據_郵寄種類 WHERE 序號='" & Me.DropDownList5.SelectedValue & "' AND 郵局=0"
        'data_dv = data.Select(New DataSourceSelectArguments)
        'If data_dv.Count > 0
        '    ScriptManager.RegisterStartupScript(Me, Page.GetType, "Script5", "setTimeout(function() { alert('所選非郵局郵寄種類。'); }, 100);", True)
        '    Exit Sub
        'End If
        
        Dim _GUID As String = Guid.NewGuid().ToString("N")
        Dim MyExcel As String = MapPath(".\Excel\Temp\") & _GUID & ".xls"
        System.IO.File.Copy(MapPath(".\Excel\特約郵件郵費單.xls"), MyExcel)
        '0元不計入不顯示
        Dim xlApp As Excel.ApplicationClass
        xlApp = New Excel.ApplicationClass()
        xlApp.DisplayAlerts = False
        xlApp.ScreenUpdating = false
        xlApp.EnableEvents = false
        Dim xlWorkBook As Excel.Workbook = xlApp.Workbooks.Open(MyExcel, 0, False, 5, "", "", True, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", False, False, 0, True, 1, 0)
        Dim xlWorkSheet As Excel.Worksheet = xlWorkBook.ActiveSheet
        
        Dim _序號 As Long = 0
        Dim _郵寄種類 As String = ""
        Dim _件數 As Long = 0
        Dim _重量 As String = ""
        Dim _郵資 As Long = 0
        Dim r As Long = 2
        Dim r1 As Long = 0
        Dim tot1 As Long = 0
        Dim tot2 As Long = 0
        Dim tot3 As Long = 0
        Dim tot4 As Long = 0
        _郵寄種類 = DropDownList5.SelectedValue
        If _郵寄種類 = DropDownList5.SelectedValue Then
            data.SelectCommand = "SELECT DISTINCT 郵資 FROM  大宗郵件執據 where 年='" & Me.DropDownList1.SelectedValue & "' and 月='" & Me.DropDownList2.SelectedValue & "' and 日='" & Me.DropDownList3.SelectedValue & "' and (0='" & _郵寄種類 & "' or 郵寄種類='" & _郵寄種類 & "') and ((收件人 is null and ''='" & Trim(Me.TextBox19.Text) & "') or 收件人 Like N'%" & Trim(Me.TextBox19.Text) & "%') and ((文號 is null and ''='" & Trim(Me.TextBox20.Text) & "') or 文號 Like N'%" & Trim(Me.TextBox20.Text) & "%') ORDER BY 郵資"
            data.DataBind()
            data_dv1 = data.Select(New DataSourceSelectArguments)
            For j As Long = 0 To data_dv1.Count - 1
                _郵資 = data_dv1(j)(0).ToString()
                data.SelectCommand = "SELECT id FROM  大宗郵件執據 where 年='" & Me.DropDownList1.SelectedValue & "' and 月='" & Me.DropDownList2.SelectedValue & "' and 日='" & Me.DropDownList3.SelectedValue & "' and 郵資='" & _郵資 & "' and (0='" & _郵寄種類 & "' or 郵寄種類='" & _郵寄種類 & "') and ((收件人 is null and ''='" & Trim(Me.TextBox19.Text) & "') or 收件人 Like N'%" & Trim(Me.TextBox19.Text) & "%') and ((文號 is null and ''='" & Trim(Me.TextBox20.Text) & "') or 文號 Like N'%" & Trim(Me.TextBox20.Text) & "%')"
                data.DataBind()
                data_dv2 = data.Select(New DataSourceSelectArguments)
                If data_dv2.Count > 0 Then
                    _件數 = data_dv2.Count
                    r = r + 1
                    xlWorkSheet.Cells(r+5, 8) = _件數
                    data.SelectCommand = "SELECT 重量 FROM  大宗郵件執據_資費表 where 序號='" & _郵寄種類 & "' and 郵資 = '" & _郵資 & "'"
                    data.DataBind()
                    data_dv = data.Select(New DataSourceSelectArguments)
                    If data_dv.Count > 0
                        _重量 = data_dv(0)(0).toString
                    Else
                        _重量 = ""
                    End If
                    xlWorkSheet.Cells(r+5, 9) = _重量
                    xlWorkSheet.Cells(r+5, 10) = _郵資
                    xlWorkSheet.Cells(r+5, 12) = String.Format("{0:0,0}", _件數 * _郵資)
                    tot1 = tot1 + _件數
                    tot2 = tot2 + (_件數 * _郵資)
                    tot3 = tot3 + _件數
                    tot4 = tot4 + (_件數 * _郵資)
                End If
            Next
        End If
        xlWorkSheet.Cells(24, 8) = tot1'合計件數
        xlWorkSheet.Cells(24, 12) = tot2.toString("N0")'合計郵資

        xlWorkSheet.Cells(29, 2) = tot2.toString("N0") + "元"
        xlWorkSheet.Cells(29, 9) = tot2.toString("N0") + "元"

        '零元修正
        If  xlWorkSheet.Cells(8, 10).value = "0"
            xlWorkSheet.Cells(24, 8) = xlWorkSheet.Cells(24, 8).value - xlWorkSheet.Cells(8, 8).value
            For i As Long = 8 To 22
                For j As Long = 8 To 12
                    xlWorkSheet.Cells(i, j) = xlWorkSheet.Cells(i + 1, j)
                Next
            Next
        End If
        
        Dim _金額大寫 As String = UpperMoney(tot2.toString)
        xlWorkSheet.Cells(31, 3) = _金額大寫
        xlWorkSheet.Cells(32, 3) = _金額大寫
        xlWorkSheet.Cells(33, 3) = _金額大寫
        xlWorkSheet.Cells(34, 3) = _金額大寫
        Dim _郵寄種類分割 As String = Trim(Me.DropDownList5.SelectedItem.ToString)
        For i As Long = 0 To _郵寄種類分割.Length - 1
            xlWorkSheet.Cells(i + 8, 5) = _郵寄種類分割.Substring(i, 1)
        Next
        xlWorkSheet.Cells(5, 12) = "交寄日期：" + Me.DropDownList1.SelectedValue + "年" + Me.DropDownList2.SelectedValue + "月" + Me.DropDownList3.SelectedValue + "日"
        
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
        Dim downloadfilename = "特約郵件郵費單 " & Me.DropDownList2.SelectedValue & Me.DropDownList3.SelectedValue & ".xls"
        Response.AddHeader("Content-Disposition", "attachment;FileName=" + Uri.EscapeDataString(downloadfilename))
        Response.WriteFile(MyExcel)
        System.Web.HttpContext.Current.ApplicationInstance.CompleteRequest()
        Response.Flush()
        System.IO.File.Delete(MyExcel)
        Response.End()
    End Sub
    Protected Sub Button51_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button51.Click    
        data.ConnectionString = con_14
        Dim _GUID As String = Guid.NewGuid().ToString("N")
        Dim MyExcel As String = "C:\大宗郵件\Excel\Temp\" & _GUID & ".xls"
        System.IO.File.Copy("C:\大宗郵件\Excel\地址標籤.xls", MyExcel)
        Dim xlApp As Excel.ApplicationClass
        xlApp = New Excel.ApplicationClass()
        xlApp.DisplayAlerts = False
        xlApp.ScreenUpdating = false
        xlApp.EnableEvents = false
        Dim xlWorkBook As Excel.Workbook = xlApp.Workbooks.Open(MyExcel, 0, False, 5, "", "", True, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", False, False, 0, True, 1, 0)
        Dim xlWorkSheet As Excel.Worksheet = xlWorkBook.ActiveSheet
        
        Dim _序號 As Long = 0
        Dim _郵寄種類 As String = ""
        Dim _件數 As Long = 0
        Dim _重量 As String = ""
        Dim _郵資 As Long = 0
        _郵寄種類 = DropDownList5.SelectedValue
        data.SelectCommand = "SELECT a.郵遞區號, a.地址, a.文號, a.收件人 FROM 大宗郵件執據 a INNER JOIN 大宗郵件執據_郵寄種類 b ON a.郵寄種類 = b.序號 WHERE a.年='" & Me.DropDownList1.SelectedValue & "' AND a.月='" & Me.DropDownList2.SelectedValue & "' AND a.日='" & Me.DropDownList3.SelectedValue & "' AND (0='" & Me.DropDownList5.SelectedValue & "' OR a.郵寄種類='" & Me.DropDownList5.SelectedValue & "') AND ((a.收件人 IS NULL AND ''='" & Trim(Me.TextBox19.Text) & "') OR a.收件人 LIKE N'%" & Trim(Me.TextBox19.Text) & "%') AND ((a.文號 IS NULL AND ''='" & Trim(Me.TextBox20.Text) & "') OR a.文號 LIKE N'%" & Trim(Me.TextBox20.Text) & "%') ORDER BY b.排序, a.序號"
        data.DataBind()
        data_dv = data.Select(New DataSourceSelectArguments)
        
        Dim _空 As Long = 0
        Me.TextBox52.Text = Trim(Me.TextBox52.Text)
        If Me.TextBox52.Text <> ""
            _空 = CLng(Me.TextBox52.Text)
        End If
        For i As Long = 0 To data_dv.Count - 1 + _空
            xlWorkSheet.Cells(((i Or 1)>>1)*5 + 1, ((i And 1)<<2) + 1).RowHeight = i And 15
            xlWorkSheet.Cells(((i Or 1)>>1)*5 + 2, ((i And 1)<<2) + 1).RowHeight = 13.75
            xlWorkSheet.Range("A" & (((i Or 1)>>1)*5 + 2).ToString & ":A" & (((i Or 1)>>1)*5 + 3).ToString).MergeCells = True
            xlWorkSheet.Range("E" & (((i Or 1)>>1)*5 + 2).ToString & ":E" & (((i Or 1)>>1)*5 + 3).ToString).MergeCells = True
            If i >= _空
                xlWorkSheet.Cells(((i Or 1)>>1)*5 + 3, ((i And 1)<<2) + 2) = data_dv(i - _空)(1)
                xlWorkSheet.Cells(((i Or 1)>>1)*5 + 4, ((i And 1)<<2) + 1) = data_dv(i - _空)(2)
                xlWorkSheet.Cells(((i Or 1)>>1)*5 + 5, ((i And 1)<<2) + 1) = data_dv(i - _空)(3)
                xlWorkSheet.Cells(((i Or 1)>>1)*5 + 2, ((i And 1)<<2) + 1) = data_dv(i - _空)(0)
            End If
            xlWorkSheet.Cells(((i Or 1)>>1)*5 + 3, ((i And 1)<<2) + 2).VerticalAlignment = -4160
            'xlWorkSheet.Cells(((i Or 1)>>1)*5 + 3, ((i And 1)<<2) + 2).RowHeight = 33
            xlWorkSheet.Cells(((i Or 1)>>1)*5 + 4, ((i And 1)<<2) + 1).VerticalAlignment = -4160
            xlWorkSheet.Cells(((i Or 1)>>1)*5 + 5, ((i And 1)<<2) + 1).VerticalAlignment = -4160
            xlWorkSheet.Cells(((i Or 1)>>1)*5 + 4, ((i And 1)<<2) + 1).HorizontalAlignment = -4152
            xlWorkSheet.Range("A" & (((i Or 1)>>1)*5 + 4).ToString & ":B" & (((i Or 1)>>1)*5 + 4).ToString).MergeCells = True
            xlWorkSheet.Range("E" & (((i Or 1)>>1)*5 + 4).ToString & ":F" & (((i Or 1)>>1)*5 + 4).ToString).MergeCells = True
            xlWorkSheet.Range("A" & (((i Or 1)>>1)*5 + 5).ToString & ":B" & (((i Or 1)>>1)*5 + 5).ToString).MergeCells = True
            xlWorkSheet.Range("E" & (((i Or 1)>>1)*5 + 5).ToString & ":F" & (((i Or 1)>>1)*5 + 5).ToString).MergeCells = True
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
        Dim downloadfilename = "地址標籤 " & Me.DropDownList2.SelectedValue & Me.DropDownList3.SelectedValue & ".xls"
        Response.AddHeader("Content-Disposition", "attachment;FileName=" + Uri.EscapeDataString(downloadfilename))
        Response.WriteFile(MyExcel)
        System.Web.HttpContext.Current.ApplicationInstance.CompleteRequest()
        Response.Flush()
        System.IO.File.Delete(MyExcel)
        Response.End()
    End Sub
    Protected Sub Button52_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button52.Click    
        data.ConnectionString = con_14
        Dim _GUID As String = Guid.NewGuid().ToString("N")
        'Dim MyExcel As String = "C:\大宗郵件\Excel\Temp\" & _GUID & ".xls"
        Dim MyExcel As String = "C:\大宗郵件\Excel\Temp\" & _GUID & ""
        System.IO.File.Copy("C:\大宗郵件\Excel\新地址標籤.xls", MyExcel)
        Dim xlApp As Excel.ApplicationClass
        xlApp = New Excel.ApplicationClass()
        xlApp.DisplayAlerts = False
        xlApp.ScreenUpdating = false
        xlApp.EnableEvents = false
        Dim xlWorkBook As Excel.Workbook = xlApp.Workbooks.Open(MyExcel, 0, False, 5, "", "", True, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", False, False, 0, True, 1, 0)
        Dim xlWorkSheet As Excel.Worksheet = xlWorkBook.ActiveSheet
        
        Dim _序號 As Long = 0
        Dim _郵寄種類 As String = ""
        Dim _件數 As Long = 0
        Dim _重量 As String = ""
        Dim _郵資 As Long = 0
        _郵寄種類 = DropDownList5.SelectedValue
        data.SelectCommand = "SELECT a.郵遞區號, a.地址, a.收件人, a.文號 FROM 大宗郵件執據 a INNER JOIN 大宗郵件執據_郵寄種類 b ON a.郵寄種類 = b.序號 WHERE a.年='" & Me.DropDownList1.SelectedValue & "' AND a.月='" & Me.DropDownList2.SelectedValue & "' AND a.日='" & Me.DropDownList3.SelectedValue & "' AND (0='" & Me.DropDownList5.SelectedValue & "' OR a.郵寄種類='" & Me.DropDownList5.SelectedValue & "') AND ((a.收件人 IS NULL AND ''='" & Trim(Me.TextBox19.Text) & "') OR a.收件人 LIKE N'%" & Trim(Me.TextBox19.Text) & "%') AND ((a.文號 IS NULL AND ''='" & Trim(Me.TextBox20.Text) & "') OR a.文號 LIKE N'%" & Trim(Me.TextBox20.Text) & "%') ORDER BY b.排序, a.序號"
        data.DataBind()
        data_dv = data.Select(New DataSourceSelectArguments)
        Dim arr(data_dv.Count * 5 + 2, 8) As Object
        Dim a As Long = 0
        Dim b As Long = 0
        arr(2, 2) = "<頁面空白>"
        If data_dv.Count > 0
            arr(2, 2) = ""
        End If
        
        Dim _空 As Long = 0
        Me.TextBox52.Text = Trim(Me.TextBox52.Text)
        If Me.TextBox52.Text <> ""
            _空 = CLng(Me.TextBox52.Text)
        End If
        For i As Long = 0 To data_dv.Count - 1 + _空
            If i And 1
                If i >= _空
                    arr(1 + b, 5) = data_dv(i - _空)(0)
                    arr(1 + b, 6) = data_dv(i - _空)(1)
                    arr(2 + b, 5) = data_dv(i - _空)(2)
                    arr(3 + b, 5) = data_dv(i - _空)(3)
                End If
                xlWorkSheet.Range("F" & (3 + b).ToString() & ":" & "G" & (3 + b).ToString()).MergeCells = True
                xlWorkSheet.Range("F" & (4 + b).ToString() & ":" & "G" & (4 + b).ToString()).MergeCells = True
                xlWorkSheet.Cells(4 + b, 6).HorizontalAlignment = -4152
                b = b + 5
            Else
                If i >= _空
                    arr(1 + a, 1) = data_dv(i - _空)(0)
                    arr(1 + a, 2) = data_dv(i - _空)(1)
                    arr(2 + a, 1) = data_dv(i - _空)(2)
                    arr(3 + a, 1) = data_dv(i - _空)(3)
                End If
                xlWorkSheet.Range("B" & (3 + a).ToString() & ":" & "C" & (3 + a).ToString()).MergeCells = True
                xlWorkSheet.Range("B" & (4 + a).ToString() & ":" & "C" & (4 + a).ToString()).MergeCells = True
                xlWorkSheet.Cells(4 + a, 2).HorizontalAlignment = -4152
                xlWorkSheet.Cells(1 + a, 1).RowHeight = 14.25
                xlWorkSheet.Cells(5 + a, 1).RowHeight = 14.25
                a = a + 5
            End If
        Next
        'xlWorkSheet.Columns(9).PageBreak = -4135
        Try
            xlWorkSheet.Range(xlWorkSheet.Cells(1, 1), xlWorkSheet.Cells(data_dv.Count * 5, 8)).Value = arr
        Catch
        End Try
        
        xlWorkBook.Save()
        Dim MyPdf As String = "C:\大宗郵件\Excel\Temp\" & _GUID & ".pdf"
        xlWorkBook.PrintOut(Preview:=False, ActivePrinter:="Microsoft Print To PDF", PrintToFile:=True, PrToFileName:=MyPdf)
        xlWorkBook.Close()
        xlApp.Quit()
        ReleaseObject(xlWorkSheet)
        ReleaseObject(xlWorkBook)
        ReleaseObject(xlApp)
        
        Response.Clear()
        Response.ClearHeaders()
        Response.Buffer = True
        Response.ContentType = "application/octet-stream"
        Dim downloadfilename = "地址標籤 " & Me.DropDownList2.SelectedValue & Me.DropDownList3.SelectedValue & ".pdf"
        Response.AddHeader("Content-Disposition", "attachment;FileName=" + Uri.EscapeDataString(downloadfilename))
        Response.WriteFile(MyPdf)
        System.Web.HttpContext.Current.ApplicationInstance.CompleteRequest()
        Response.Flush()
        System.IO.File.Delete(MyExcel)
        System.IO.File.Delete(MyPdf)
        Response.End()
    End Sub
    Protected Sub Button11_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button11.Click
        data.ConnectionString = con_14
        Dim _GUID As String = Guid.NewGuid().ToString("N")
        Dim MyExcel As String = "C:\大宗郵件\Excel\Temp\" & _GUID & ".xls"
        System.IO.File.Copy("C:\大宗郵件\Excel\合計.xls", MyExcel)
        Dim xlApp As Excel.ApplicationClass
        xlApp = New Excel.ApplicationClass()
        xlApp.DisplayAlerts = False
        xlApp.ScreenUpdating = false
        xlApp.EnableEvents = false
        Dim xlWorkBook As Excel.Workbook = xlApp.Workbooks.Open(MyExcel, 0, False, 5, "", "", True, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", False, False, 0, True, 1, 0)
        Dim xlWorkSheet As Excel.Worksheet = xlWorkBook.ActiveSheet

        Dim _序號 As Long = 0
        Dim _郵寄種類 As Long = 0
        Dim _種類 As String = ""
        Dim _件數 As Long = 0
        Dim _郵資 As Long = 0
        Dim r As Long = 2
        Dim r1 As Long = 0
        Dim tot1 As Long = 0
        Dim tot2 As Long = 0
        Dim tot3 As Long = 0
        Dim tot4 As Long = 0
        Dim _範圍 As String = ""
        data.SelectCommand = "SELECT 序號,郵寄種類 FROM  大宗郵件執據_郵寄種類  ORDER BY 序號"
        data.DataBind()
        data_dv = data.Select(New DataSourceSelectArguments)
        xlWorkSheet.Cells(1, 2) = Me.DropDownList1.SelectedValue + "年" + Me.DropDownList2.SelectedValue + "月" + Me.DropDownList3.SelectedValue + "日"
        For i As Long = 0 To data_dv.Count - 1
            _件數 = 0
            _郵寄種類 = data_dv(i)(0).ToString()
            If tot1 <> 0 Then
                r = r + 1
                xlWorkSheet.Cells(r, 2) = "合計"
                xlWorkSheet.Cells(r, 3) = tot1
                xlWorkSheet.Cells(r, 5) = tot2
                _範圍 = "b" + (r).ToString + ":e" + (r).ToString
                xlWorkSheet.Range(_範圍).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle = 1
                xlWorkSheet.Range(_範圍).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = 1
                For k As Long = 2 To 5
                    xlWorkSheet.Cells(r, k).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = 1
                    xlWorkSheet.Cells(r, k).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle = 1
                Next
                tot1 = 0
                tot2 = 0
            End If
            r1 = 0
            If 1=1 Then
                _種類 = data_dv(i)(1).ToString()
                data.SelectCommand = "SELECT DISTINCT 郵資 FROM  大宗郵件執據 where 年='" & Me.DropDownList1.SelectedValue & "' and 月='" & Me.DropDownList2.SelectedValue & "' and 日='" & Me.DropDownList3.SelectedValue & "' and 郵寄種類='" & _郵寄種類 & "' ORDER BY 郵資"
                data.DataBind()
                data_dv1 = data.Select(New DataSourceSelectArguments)
                For j As Long = 0 To data_dv1.Count - 1
                    _郵資 = data_dv1(j)(0).ToString()
                    data.SelectCommand = "SELECT id FROM  大宗郵件執據 where 年='" & Me.DropDownList1.SelectedValue & "' and 月='" & Me.DropDownList2.SelectedValue & "' and 日='" & Me.DropDownList3.SelectedValue & "' and 郵資='" & _郵資 & "' and 郵寄種類='" & _郵寄種類 & "'"
                    data.DataBind()
                    data_dv2 = data.Select(New DataSourceSelectArguments)
                    If data_dv2.Count > 0 Then
                        _件數 = data_dv2.Count
                        r = r + 1
                        If r1 = 0 Then
                            r1 = 1
                            xlWorkSheet.Cells(r, 2) = _種類
                        End If
                        xlWorkSheet.Cells(r, 3) = _件數
                        xlWorkSheet.Cells(r, 4) = _郵資
                        xlWorkSheet.Cells(r, 5) = _件數 * _郵資
                        tot1 = tot1 + _件數
                        tot2 = tot2 + (_件數 * _郵資)
                        tot3 = tot3 + _件數
                        tot4 = tot4 + (_件數 * _郵資)
                        _範圍 = "b" + (r).ToString + ":e" + (r).ToString
                        xlWorkSheet.Range(_範圍).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle = 1
                        xlWorkSheet.Range(_範圍).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = 1
                        For k As Long = 2 To 5
                            xlWorkSheet.Cells(r, k).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = 1
                            xlWorkSheet.Cells(r, k).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle = 1
                        Next
                    End If
                Next
            End If
        Next
        r = r + 1
        xlWorkSheet.Cells(r, 2) = "總計"
        xlWorkSheet.Cells(r, 3) = tot3
        xlWorkSheet.Cells(r, 5) = tot4
        _範圍 = "b" + (r).ToString + ":e" + (r).ToString
        xlWorkSheet.Range(_範圍).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle = 1
        xlWorkSheet.Range(_範圍).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = 1
        For k As Long = 2 To 5
            xlWorkSheet.Cells(r, k).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = 1
            xlWorkSheet.Cells(r, k).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle = 1
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
        Dim downloadfilename = "合計 " & Me.DropDownList2.SelectedValue & Me.DropDownList3.SelectedValue & ".xls"
        Response.AddHeader("Content-Disposition", "attachment;FileName=" + Uri.EscapeDataString(downloadfilename))
        Response.WriteFile(MyExcel)
        System.Web.HttpContext.Current.ApplicationInstance.CompleteRequest()
        Response.Flush()
        System.IO.File.Delete(MyExcel)
        Response.End()
    End Sub
    Protected Sub Button12_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button12.Click
        data.ConnectionString = con_14
        data.SelectCommand = "SELECT 掛號號碼,id FROM  大宗郵件執據 where 年='" & Me.DropDownList1.SelectedValue & "' and 月='" & Me.DropDownList2.SelectedValue & "' and 日='" & Me.DropDownList3.SelectedValue & "'  and 序號>='" & Me.TextBox16.Text & "' and 郵寄種類 = " & Me.DropDownList5.SelectedValue & " and 收件人 LIKE '%" & Trim(Me.TextBox19.Text) & "%' and 文號 LIKE '%" & Trim(Me.TextBox20.Text) & "%' ORDER BY 序號"
        data.DataBind()
        data_dv = data.Select(New DataSourceSelectArguments)
        Dim _id As Long = 0
        For i As Long = 0 To data_dv.Count - 1
            _id = data_dv(i)(1).ToString()
            data.UpdateCommand = "update 大宗郵件執據 set 掛號類別='" & Me.TextBox17.Text & "' where id='" & _id & "'"
            data.Update()
            data.DataBind()
        Next
        Me.DropDownList3.DataBind()
        Me.GridView2.DataBind()
    End Sub
    Protected Sub Button13_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button13.Click
        data.ConnectionString = con_14
        data.SelectCommand = "SELECT 序號,掛號號碼,收件人,地址,文號,備註,重量,郵資,掛號類別 FROM  大宗郵件執據 where 郵資<=0 and 年='" & Me.DropDownList1.SelectedValue & "' and 月='" & Me.DropDownList2.SelectedValue & "' and 日='" & Me.DropDownList3.SelectedValue & "' and 郵寄種類 = " & Me.DropDownList5.SelectedValue & " and 收件人 LIKE '%" & Trim(Me.TextBox19.Text) & "%' and 文號 LIKE '%" & Trim(Me.TextBox20.Text) & "%' ORDER BY 序號"
        data.DataBind()
        data_dv = data.Select(New DataSourceSelectArguments)
        If data_dv.Count > 0 Then
            Me.Label3.Text = "有" + data_dv.Count.ToString + "件沒有輸入郵資"
        Else
            Me.Label3.Text = ""
        End If
    End Sub
    Protected Sub Button14_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button14.Click
        data.ConnectionString = con_14
        Dim _GUID As String = Guid.NewGuid().ToString("N")
        Dim MyExcel As String = "C:\大宗郵件\Excel\Temp\" & _GUID & ".xls"
        System.IO.File.Copy("C:\大宗郵件\Excel\合計.xls", MyExcel)
        Dim xlApp As Excel.ApplicationClass
        xlApp = New Excel.ApplicationClass()
        xlApp.DisplayAlerts = False
        xlApp.ScreenUpdating = false
        xlApp.EnableEvents = false
        Dim xlWorkBook As Excel.Workbook = xlApp.Workbooks.Open(MyExcel, 0, False, 5, "", "", True, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", False, False, 0, True, 1, 0)
        Dim xlWorkSheet As Excel.Worksheet = xlWorkBook.ActiveSheet

        Dim _序號 As Long = 0
        Dim _郵寄種類 As Long = 0
        Dim _種類 As String = ""
        Dim _件數 As Long = 0
        Dim _郵資 As Long = 0
        Dim r As Long = 2
        Dim r1 As Long = 0
        Dim tot1 As Long = 0
        Dim tot2 As Long = 0
        Dim tot3 As Long = 0
        Dim tot4 As Long = 0
        Dim _範圍 As String = ""
        data.SelectCommand = "SELECT 序號,郵寄種類 FROM  大宗郵件執據_郵寄種類  ORDER BY 序號"
        data.DataBind()
        data_dv = data.Select(New DataSourceSelectArguments)
        xlWorkSheet.Cells(1, 2) = Me.DropDownList1.SelectedValue + "年" + Me.DropDownList2.SelectedValue + "月" + Me.DropDownList3.SelectedValue + "日"
        For i As Long = 0 To data_dv.Count - 1
            _郵寄種類 = data_dv(i)(0).ToString()
            If tot1 <> 0 Then
                r = r + 1
                xlWorkSheet.Cells(r, 2) = "合計"
                xlWorkSheet.Cells(r, 3) = tot1
                xlWorkSheet.Cells(r, 5) = tot2
                _範圍 = "b" + (r).ToString + ":e" + (r).ToString
                xlWorkSheet.Range(_範圍).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle = 1
                xlWorkSheet.Range(_範圍).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = 1
                For k As Long = 2 To 5
                    xlWorkSheet.Cells(r, k).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = 1
                    xlWorkSheet.Cells(r, k).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle = 1
                Next
                tot1 = 0
                tot2 = 0
            End If
            r1 = 0
            If _郵寄種類 >= 18 And _郵寄種類 <= 28 Then
                _種類 = data_dv(i)(1).ToString()
                data.SelectCommand = "SELECT DISTINCT 郵資 FROM  大宗郵件執據 where 年='" & Me.DropDownList1.SelectedValue & "' and 月='" & Me.DropDownList2.SelectedValue & "' and 日='" & Me.DropDownList3.SelectedValue & "' and 郵寄種類='" & _郵寄種類 & "' ORDER BY 郵資"
                data.DataBind()
                data_dv1 = data.Select(New DataSourceSelectArguments)
                _件數 = 0
                For j As Long = 0 To data_dv1.Count - 1
                    _郵資 = data_dv1(j)(0).ToString()
                    data.SelectCommand = "SELECT COUNT(id) FROM  大宗郵件執據 where 年='" & Me.DropDownList1.SelectedValue & "' and 月='" & Me.DropDownList2.SelectedValue & "' and 日='" & Me.DropDownList3.SelectedValue & "' and 郵資='" & _郵資 & "' and 郵寄種類='" & _郵寄種類 & "'"
                    data.DataBind()
                    data_dv2 = data.Select(New DataSourceSelectArguments)
                    If data_dv2.Count > 0 Then
                        _件數 = data_dv2(0)(0).ToString()
                        r = r + 1
                        If r1 = 0 Then
                            r1 = 1
                            xlWorkSheet.Cells(r, 2) = _種類
                        End If
                        xlWorkSheet.Cells(r, 3) = _件數
                        xlWorkSheet.Cells(r, 4) = _郵資
                        xlWorkSheet.Cells(r, 5) = _件數 * _郵資
                        tot1 = tot1 + _件數
                        tot2 = tot2 + (_件數 * _郵資)
                        tot3 = tot3 + _件數
                        tot4 = tot4 + (_件數 * _郵資)
                        _範圍 = "b" + (r).ToString + ":e" + (r).ToString
                        xlWorkSheet.Range(_範圍).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle = 1
                        xlWorkSheet.Range(_範圍).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = 1
                        For k As Long = 2 To 5
                            xlWorkSheet.Cells(r, k).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = 1
                            xlWorkSheet.Cells(r, k).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle = 1
                        Next
                    End If
                Next
            End If
        Next
        r = r + 1
        xlWorkSheet.Cells(r, 2) = "總計"
        xlWorkSheet.Cells(r, 3) = tot3
        xlWorkSheet.Cells(r, 5) = tot4
        _範圍 = "b" + (r).ToString + ":e" + (r).ToString
        xlWorkSheet.Range(_範圍).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle = 1
        xlWorkSheet.Range(_範圍).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = 1
        For k As Long = 2 To 5
            xlWorkSheet.Cells(r, k).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = 1
            xlWorkSheet.Cells(r, k).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle = 1
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
        Dim downloadfilename = "*合計 " & Me.DropDownList2.SelectedValue & Me.DropDownList3.SelectedValue & ".xls"
        Response.AddHeader("Content-Disposition", "attachment;FileName=" + Uri.EscapeDataString(downloadfilename))
        Response.WriteFile(MyExcel)
        System.Web.HttpContext.Current.ApplicationInstance.CompleteRequest()
        Response.Flush()
        System.IO.File.Delete(MyExcel)
        Response.End()
    End Sub
    Private Sub GridView2_PreRender(sender As Object, e As EventArgs) Handles GridView2.PreRender
        data.ConnectionString = con_14
        data.SelectCommand = "SELECT 郵資 FROM  大宗郵件執據 where 年='" & Me.DropDownList1.SelectedValue & "' and 月='" & Me.DropDownList2.SelectedValue & "' and 日='" & Me.DropDownList3.SelectedValue & "' and (0='" & Me.DropDownList5.SelectedValue & "' or 郵寄種類='" & Me.DropDownList5.SelectedValue & "') and ((收件人 is null and ''='" & Trim(Me.TextBox19.Text) & "') or 收件人 Like N'%" & Trim(Me.TextBox19.Text) & "%') and ((文號 is null and ''='" & Trim(Me.TextBox20.Text) & "') or 文號 Like N'%" & Trim(Me.TextBox20.Text) & "%')"
        data.DataBind()
        data_dv = data.Select(New DataSourceSelectArguments)
        Me.Label2.Text = ""
        Dim _郵資 As Long = 0
        For i As Long = 0 To data_dv.Count - 1
            _郵資 = _郵資 + data_dv(i)(0).ToString()
        Next
        Me.Label2.Text = "共" + data_dv.Count.ToString + "件郵資合計" + _郵資.ToString + "元"
    End Sub
End Class


