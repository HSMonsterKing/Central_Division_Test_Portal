Partial Class 資費表
    Inherits System.Web.UI.Page
    Dim con_14 As String = "Data Source=10.52.0.178;Initial Catalog=大宗郵件;User ID=qaz;Password=1qaz@WSX"
    Dim data As New SqlDataSource
    Dim data_dv As Data.DataView
    Protected Sub GridView1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.SelectedIndexChanged
        data.ConnectionString = con_14
        Dim _id As Long = CType(Me.GridView1.SelectedRow.FindControl("Label1"), Label).Text()
        Dim _a As String = CType(Me.GridView1.SelectedRow.FindControl("TextBox1"), TextBox).Text()
        data.UpdateCommand = "update 大宗郵件執據_資費表 set 郵資='" & _a & "' where id='" & _id & "'"
        data.Update()
        data.DataBind()
    End Sub
End Class
