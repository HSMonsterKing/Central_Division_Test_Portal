
Partial Class index
    Inherits System.Web.UI.Page

    Dim con As String = "Data Source=10.52.0.178;Initial Catalog=大宗郵件;User ID=qaz;Password=1qaz@WSX"
    Dim data As New SqlDataSource
    Dim data_dv As Data.DataView
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not Page.IsPostBack Then
            Dim r1, r2, r3 As Integer
            Dim _登入帳號 As String
            Dim _承辦人帳號 As String = ""
            _登入帳號 = Request.ServerVariables("REMOTE_HOST")
            If Len(Trim(_登入帳號)) <= 0 Then
                _登入帳號 = System.Security.Principal.WindowsIdentity.GetCurrent().Name
            End If
            r1 = Len(_登入帳號)
            For i As Integer = 1 To r1
                If Mid(_登入帳號, i, 1) = "\" Then
                    r2 = i
                End If
            Next
            r3 = r1 - r2
            _登入帳號 = Right(_登入帳號, r3)
            data.ConnectionString = con
            data.SelectCommand = "SELECT id FROM 大宗郵件執據_操作者 where 帳號='" & _登入帳號 & " ' "
            data.DataBind()
            data_dv = data.Select(New DataSourceSelectArguments)
            If data_dv.Count > 0 Then
                Response.Redirect("./大宗郵件執據.aspx")
            Else
                Response.Redirect("./搜尋.aspx")
            End If


            'If _登入帳號 = "lien1" Or _登入帳號 = "rita" Or _登入帳號 = "montego" Or _登入帳號 = "silk" Or _登入帳號 = "wei2712" Or _登入帳號 = "ingjen" Then
            '    Response.Redirect("./大宗函件/大宗函件執據.aspx")
            'Else
            '    Response.Redirect("./大宗函件/搜尋.aspx")
            'End If
        End If
    End Sub
End Class

