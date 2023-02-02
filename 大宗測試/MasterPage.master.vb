
Partial Class MasterPage
    Inherits System.Web.UI.MasterPage
    Dim con_Freeway As String = "Data Source=10.52.0.178;Initial Catalog=Freeway;User ID=qaz;Password=1qaz@WSX"
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
            If _登入帳號 = "10.52.3.155" Then
                Me.Button23.Visible = True
            Else
                Me.Button23.Visible = False
            End If
            data.ConnectionString = con_Freeway
            data.SelectCommand = "SELECT USER_NM FROM 姓名 where USER_ID='" & _登入帳號 & "'"
            data.DataBind()
            data_dv = data.Select(New DataSourceSelectArguments)
            Dim _姓名 As String = Trim(data_dv(0)(0).ToString())
            Me.Title1.Text = "~~歡迎" + _姓名 + "使用本系統，本系統於109年11月啟用~~"




            data.ConnectionString = con
            data.SelectCommand = "SELECT id FROM  大宗郵件執據_操作者 where 帳號='" & _登入帳號 & " ' "
            data.DataBind()
            data_dv = data.Select(New DataSourceSelectArguments)
            If data_dv.Count <= 0 Then
                Me.Button6.Visible = False
                Me.Button22.Visible = False
                Me.Button17.Visible = False
                Me.Button16.Visible = False
                Me.Button18.Visible = False
                Me.Button19.Visible = False
                Me.Button20.Visible = False
                Me.Button21.Visible = False
                Me.Button24.Visible = False
            End If
        End If
    End Sub
End Class

