Partial Class 搜尋
    Inherits System.Web.UI.Page
    'Dim con_wf2 As String = "Data Source=edocsql.freeway.gov.tw\SQL2012;Initial Catalog=CFW_wf2;User ID=CFWedocdb01;Password=99$edocdb$CFW"
    Dim con_wf2 As String = "Data Source=edocsqlplus.freeway.gov.tw\SQL2019,54399;Initial Catalog=CFW_wf2;User ID=CFWedocdb01;Password=99$edocdb$CFW"
    Dim con_14 As String = "Data Source=10.52.0.178;Initial Catalog=大宗郵件;User ID=qaz;Password=1qaz@WSX"
    Dim data As New SqlDataSource
    Dim data_dv As Data.DataView
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
            Me.TextBox6.Text = _登入帳號

            Dim _年, _月, _日 As String
            '_年 = (Now.Year - 1911).ToString
            '_月 = Now.Month.ToString("00")
            '_日 = Now.Day.ToString("00")
            data.ConnectionString = con_14
            data.SelectCommand = "SELECT 年, 月, 日 FROM 大宗郵件執據_操作者 WHERE 帳號='" & Me.TextBox6.Text & "'"
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
                data.UpdateCommand = "update 大宗郵件執據_操作者 set 年 = '" & Me.DropDownList1.SelectedValue & "', 月 = '" & Me.DropDownList2.SelectedValue & "', 日 = '" & Me.DropDownList3.SelectedValue & "' WHERE 帳號='" & Me.TextBox6.Text & "'"
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
    Protected Sub Button3_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button3.Click
        Me.TextBox2.Text = Me.DropDownList1.SelectedValue + Me.DropDownList2.SelectedValue + Me.DropDownList3.SelectedValue
        Me.TextBox3.Text = Me.DropDownList4.SelectedValue + Me.DropDownList5.SelectedValue + Me.DropDownList6.SelectedValue
        Me.TextBox4.Text = Me.TextBox1.Text
        Me.TextBox5.Text = ""
    End Sub
    Protected Sub Button4_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button4.Click
        Me.TextBox2.Text = "00000000"
        Me.TextBox3.Text = "9999999"
        Me.TextBox4.Text = ""
        Me.TextBox5.Text = Me.TextBox1.Text
    End Sub
End Class
