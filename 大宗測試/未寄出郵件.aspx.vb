Partial Class 未寄出郵件
    Inherits System.Web.UI.Page
    'Dim con_wf2 As String = "Data Source=edocsql.freeway.gov.tw\SQL2012;Initial Catalog=CFW_wf2;User ID=CFWedocdb01;Password=99$edocdb$CFW"
    Dim con_wf2 As String = "Data Source=edocsqlplus.freeway.gov.tw\SQL2019,54399;Initial Catalog=CFW_wf2;User ID=CFWedocdb01;Password=99$edocdb$CFW"
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
            r2 = r1
            For i As Long = 1 To r1
                If Mid(_登入帳號, i, 1) = "\" Then
                    r2 = i
                End If
            Next
            If r1 = r2 Then
                r3 = r1
            Else
                r3 = r1 - r2
            End If
            _登入帳號 = Right(_登入帳號, r3)
            Calendar1.SelectedDate = DateTime.Today
            GenTreeNode()
        End If
    End Sub
    Protected Sub Calendar1_SelectionChanged(sender As Object, e As EventArgs) Handles Calendar1.SelectionChanged
        GenTreeNode()
    End Sub
    Protected Sub GenTreeNode()
        Me.ListBox1.Items.Clear()
        Me.ListBox2.Items.Clear()
        data.ConnectionString = con_wf2
        Dim da1 As Date = DateSerial(Me.Calendar1.SelectedDate.Year, Me.Calendar1.SelectedDate.Month, Me.Calendar1.SelectedDate.Day).AddHours(0).AddMinutes(0).AddSeconds(0)
        Dim da2 As Date = DateSerial(Me.Calendar1.SelectedDate.Year, Me.Calendar1.SelectedDate.Month, Me.Calendar1.SelectedDate.Day + 1).AddHours(0).AddMinutes(0).AddSeconds(0)
        Dim da3 As Date = DateSerial(Me.Calendar1.SelectedDate.Year, Me.Calendar1.SelectedDate.Month, Me.Calendar1.SelectedDate.Day).AddHours(0).AddMinutes(0).AddSeconds(0)

        Dim _FLow_id As Long = 0
        Dim _id As Long = 0
        Dim _create_doc_no As String = ""
        Dim _ACC_NAME As String = ""
        Dim _ACC_NAME1 As String = ""
        Dim _文號 As String = ""
        Dim _收件人 As String = ""
        data.SelectCommand = "SELECT distinct FLow_id,REG_DATE FROM  cur_flow_data where   ROLE_ID=26 and  REG_DATE>'" & da1 & " ' and REG_DATE<'" & da2 & "' ORDER BY REG_DATE"
        data.DataBind()
        data_dv = data.Select(New DataSourceSelectArguments)
        If data_dv.Count > 0 Then
            For i As Long = 0 To data_dv.Count - 1
                _FLow_id = data_dv(i)(0).ToString()
                da3 = data_dv(i)(1).ToString()
                data.SelectCommand = "SELECT create_doc_no FROM   cur_flow where id='" & _FLow_id & "'"
                data.DataBind()
                data_dv1 = data.Select(New DataSourceSelectArguments)
                If data_dv1.Count > 0 Then
                    _create_doc_no = data_dv1(0)(0).ToString()
                    data.SelectCommand = "SELECT id FROM  CREATE_DOC where FULL_NO ='" & _create_doc_no & " '"
                    data.DataBind()
                    data_dv1 = data.Select(New DataSourceSelectArguments)
                    If data_dv1.Count > 0 Then
                        _id = data_dv1(0)(0).ToString()
                        data.SelectCommand = "SELECT ACC_NAME FROM  ACCEPTER where  (deli_way=2 or deli_way=4 ) and  DOC_ID ='" & _id & " 'ORDER BY ID"
                        data.DataBind()
                        data_dv1 = data.Select(New DataSourceSelectArguments)
                        For j As Long = 0 To data_dv1.Count - 1
                            _ACC_NAME = data_dv1(j)(0).ToString()
                            _ACC_NAME1 = _ACC_NAME + "."
                            data.ConnectionString = con_14
                            data.SelectCommand = "SELECT id,文號,收件人  FROM  大宗郵件執據 where (收件人='" & _ACC_NAME1 & "' or 收件人 ='" & _ACC_NAME & "') and 文號 LIKE '%" & _create_doc_no & "%'"
                            data.DataBind()
                            data_dv2 = data.Select(New DataSourceSelectArguments)
                            If data_dv2.Count <= 0 Then
                                Me.ListBox1.Items.Add(da3 + "  " + _create_doc_no + " " + _ACC_NAME)
                            End If
                            data.ConnectionString = con_wf2
                        Next
                    End If
                End If
            Next
        End If
        Dim A As Long = 0
        data.SelectCommand = "SELECT distinct FLow_id,REG_DATE,id FROM   history_flow_data where   ROLE_ID=26 and  REG_DATE>'" & da1 & " ' and REG_DATE<'" & da2 & "' ORDER BY REG_DATE"
        data.DataBind()
        data_dv = data.Select(New DataSourceSelectArguments)
        If data_dv.Count > 0 Then
            For i As Long = 0 To data_dv.Count - 1
                _FLow_id = data_dv(i)(0).ToString()
                da3 = data_dv(i)(1).ToString()
                data.SelectCommand = "SELECT create_doc_no FROM    history_flow where id='" & _FLow_id & "'"
                data.DataBind()
                data_dv1 = data.Select(New DataSourceSelectArguments)
                If data_dv1.Count > 0 Then
                    _create_doc_no = data_dv1(0)(0).ToString()
                    data.SelectCommand = "SELECT id FROM  CREATE_DOC where FULL_NO ='" & _create_doc_no & " '"
                    data.DataBind()
                    data_dv1 = data.Select(New DataSourceSelectArguments)
                    If data_dv1.Count > 0 Then
                        _id = data_dv1(0)(0).ToString()
                        data.SelectCommand = "SELECT ACC_NAME FROM  ACCEPTER where (deli_way=2 or deli_way=4 ) and  DOC_ID ='" & _id & " 'ORDER BY ID"
                        data.DataBind()
                        data_dv1 = data.Select(New DataSourceSelectArguments)
                        For j As Long = 0 To data_dv1.Count - 1
                            _ACC_NAME = data_dv1(j)(0).ToString()
                            _ACC_NAME1 = _ACC_NAME + "."
                            data.ConnectionString = con_14
                            data.SelectCommand = "SELECT id,文號,收件人  FROM  大宗郵件執據 where (收件人='" & _ACC_NAME1 & "' or 收件人 ='" & _ACC_NAME & "') and 文號 LIKE '%" & _create_doc_no & "%'"
                            data.DataBind()
                            data_dv2 = data.Select(New DataSourceSelectArguments)
                            If data_dv2.Count <= 0 Then
                                Me.ListBox2.Items.Add(da3 + "  " + _create_doc_no + " " + _ACC_NAME)
                            End If
                            data.ConnectionString = con_wf2
                        Next
                    End If
                End If
            Next
        End If
    End Sub
End Class
