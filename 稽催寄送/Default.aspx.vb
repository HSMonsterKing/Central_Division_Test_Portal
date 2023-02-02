
Partial Class index
    Inherits System.Web.UI.Page
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not Page.IsPostBack Then
            Response.Redirect("./稽催寄送.aspx")
        End If 
    End Sub
End Class
