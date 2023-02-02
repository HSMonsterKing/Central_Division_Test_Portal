Partial Class 日曆
    Inherits System.Web.UI.Page
    Protected Sub Calendar1_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Calendar1.SelectionChanged
        Me.TextBox1.Text = Me.Calendar1.SelectedDate
    End Sub
End Class
