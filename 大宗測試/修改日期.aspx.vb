Partial Class 修改日期
    Inherits System.Web.UI.Page
    Dim con_14 As String = "Data Source=10.52.0.178;Initial Catalog=大宗郵件;User ID=qaz;Password=1qaz@WSX"
    Dim data As New SqlDataSource
    Dim data_dv As Data.DataView
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
            Me.TextBox5.Text = card_id
            
            Dim _年, _月, _日 As String
            '_年 = (Now.Year - 1911).ToString
            '_月 = Now.Month.ToString("00")
            '_日 = Now.Day.ToString("00")
            data.ConnectionString = con_14
            data.SelectCommand = "SELECT 年, 月, 日 FROM 大宗郵件執據_操作者 WHERE 帳號='" & Me.TextBox5.Text & "'"
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
                data.UpdateCommand = "update 大宗郵件執據_操作者 set 年 = '" & Me.DropDownList1.SelectedValue & "', 月 = '" & Me.DropDownList2.SelectedValue & "', 日 = '" & Me.DropDownList3.SelectedValue & "' WHERE 帳號='" & Me.TextBox5.Text & "'"
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
        End If
    End Sub
    Protected Sub DropDownList1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DropDownList1.SelectedIndexChanged
        Me.DropDownList3.Items.Clear()
        For i As Long = 0 To DateTime.DaysInMonth((Val(Me.DropDownList1.SelectedValue) + 1911), Val(Me.DropDownList2.SelectedValue)) - 1
            Me.DropDownList3.Items.Add((i + 1).ToString("00"))
            Me.DropDownList3.Items(i).Value = (i + 1).ToString("00")
        Next
        Me.DropDownList3.DataBind()
    End Sub
    Protected Sub DropDownList2_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DropDownList2.SelectedIndexChanged
        Me.DropDownList3.Items.Clear()
        For i As Long = 0 To DateTime.DaysInMonth((Val(Me.DropDownList1.SelectedValue) + 1911), Val(Me.DropDownList2.SelectedValue)) - 1
            Me.DropDownList3.Items.Add((i + 1).ToString("00"))
            Me.DropDownList3.Items(i).Value = (i + 1).ToString("00")
        Next
        Me.DropDownList3.DataBind()
    End Sub
    Protected Sub DropDownList3_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DropDownList3.SelectedIndexChanged
    End Sub
    Protected Sub DropDownList4_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DropDownList4.SelectedIndexChanged
        Me.DropDownList6.Items.Clear()
        For i As Long = 0 To DateTime.DaysInMonth((Val(Me.DropDownList4.SelectedValue) + 1911), Val(Me.DropDownList5.SelectedValue)) - 1
            Me.DropDownList6.Items.Add((i + 1).ToString("00"))
            Me.DropDownList6.Items(i).Value = (i + 1).ToString("00")
        Next
        Me.DropDownList6.DataBind()
    End Sub
    Protected Sub DropDownList5_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DropDownList5.SelectedIndexChanged
        Me.DropDownList6.Items.Clear()
        For i As Long = 0 To DateTime.DaysInMonth((Val(Me.DropDownList4.SelectedValue) + 1911), Val(Me.DropDownList5.SelectedValue)) - 1
            Me.DropDownList6.Items.Add((i + 1).ToString("00"))
            Me.DropDownList6.Items(i).Value = (i + 1).ToString("00")
        Next
        Me.DropDownList6.DataBind()
    End Sub
    'Protected Sub Button3_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button3.Click
        'data.ConnectionString = con_14
        'data.UpdateCommand = "update 大宗郵件執據 set 年='" & Me.DropDownList4.SelectedValue & "' , 月='" & Me.DropDownList5.SelectedValue & "' , 日='" & Me.'DropDownList6.SelectedValue & "'  where 年='" & Me.DropDownList1.SelectedValue & "' and 月='" & Me.DropDownList2.SelectedValue & "' and 日='" & Me.DropDownList3.SelectedValue & "'"
    '    data.Update()
    '    data.DataBind()
    'End Sub
    Protected Sub Button3_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button3.Click
        Me.TextBox30.Text = Trim(Me.TextBox30.Text)
        Me.TextBox31.Text = Trim(Me.TextBox31.Text)
        data.ConnectionString = con_14
        If Me.TextBox30.Text = "" And Me.TextBox31.Text = ""
            data.UpdateCommand = "update 大宗郵件執據 set 年='" & Me.DropDownList4.SelectedValue & "' , 月='" & Me.DropDownList5.SelectedValue & "' , 日='" & Me.DropDownList6.SelectedValue & "'  where 年='" & Me.DropDownList1.SelectedValue & "' and 月='" & Me.DropDownList2.SelectedValue & "' and 日='" & Me.DropDownList3.SelectedValue & "'"
        ElseIf Me.TextBox30.Text <> "" And Me.TextBox31.Text <> ""
            data.UpdateCommand = "update 大宗郵件執據 set 年='" & Me.DropDownList4.SelectedValue & "' , 月='" & Me.DropDownList5.SelectedValue & "' , 日='" & Me.DropDownList6.SelectedValue & "'  where 年='" & Me.DropDownList1.SelectedValue & "' and 月='" & Me.DropDownList2.SelectedValue & "' and 日='" & Me.DropDownList3.SelectedValue & "' and 序號 >= " & CLng(Me.TextBox30.Text) & " and 序號 <= " & CLng(Me.TextBox31.Text) & " "
        ElseIf Me.TextBox30.Text <> ""
            data.UpdateCommand = "update 大宗郵件執據 set 年='" & Me.DropDownList4.SelectedValue & "' , 月='" & Me.DropDownList5.SelectedValue & "' , 日='" & Me.DropDownList6.SelectedValue & "'  where 年='" & Me.DropDownList1.SelectedValue & "' and 月='" & Me.DropDownList2.SelectedValue & "' and 日='" & Me.DropDownList3.SelectedValue & "' and 序號 >= " & CLng(Me.TextBox30.Text) & ""
        ElseIf Me.TextBox31.Text <> ""
            data.UpdateCommand = "update 大宗郵件執據 set 年='" & Me.DropDownList4.SelectedValue & "' , 月='" & Me.DropDownList5.SelectedValue & "' , 日='" & Me.DropDownList6.SelectedValue & "'  where 年='" & Me.DropDownList1.SelectedValue & "' and 月='" & Me.DropDownList2.SelectedValue & "' and 日='" & Me.DropDownList3.SelectedValue & "' and 序號 <= " & CLng(Me.TextBox31.Text) & " "
        Else
            Exit Sub
        End If
        data.Update()
        data.DataBind()
    End Sub
End Class

