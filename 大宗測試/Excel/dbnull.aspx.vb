If Not IsDBNull(data_dv(0)(0))
    _序號 = data_dv(0)(0)
End If

dd4.SelectedIndex = dd4.Items.IndexOf(dd4.Items.FindByValue(_序號))

Me.DropDownList2.SelectedIndex = Me.DropDownList2.Items.IndexOf(Me.DropDownList2.Items.FindByValue("請選擇"))

NULLIF('" & _掛號號碼 & "', '')

N'" & _收件人 & "'