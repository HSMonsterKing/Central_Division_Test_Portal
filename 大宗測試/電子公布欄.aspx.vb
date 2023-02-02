
Partial Class 電子公布欄
    Inherits System.Web.UI.Page
    Dim str As String = ""
    'Dim con_wf2 As String = "Data Source=edocsql.freeway.gov.tw\SQL2012;Initial Catalog=CFW_wf2;User ID=CFWedocdb01;Password=99$edocdb$CFW"
    Dim con_wf2 As String = "Data Source=edocsqlplus.freeway.gov.tw\SQL2019,54399;Initial Catalog=CFW_wf2;User ID=CFWedocdb01;Password=99$edocdb$CFW"
    'Dim con_ConnList As String = "Data Source=edocsql.freeway.gov.tw\SQL2012;Initial Catalog=CFW_bbslist;User ID=CFWedocdb01;Password=99$edocdb$CFW"
    Dim con_ConnList As String = "Data Source=edocsqlplus.freeway.gov.tw\SQL2019,54399;Initial Catalog=CFW_bbslist;User ID=CFWedocdb01;Password=99$edocdb$CFW"
    Dim data As New SqlDataSource
    Dim data_dv As Data.DataView
    Protected Sub Button24_Click(sender As Object, e As EventArgs) Handles Button24.Click
        'data.ConnectionString = con_ConnList
        'str = "SELECT * FROM bbs_viewState"
        'data.SelectCommand = str
        'data.DataBind()
        'data_dv = data.Select(New DataSourceSelectArguments)

        'Me.GridView1.DataSource = data_dv
        'Me.GridView1.DataBind()
        Me.Label2.Text = "montego"
        Me.Label3.Text = "吳丁贊"
        data.ConnectionString = con_wf2
        str = "SELECT * FROM COM_DATA where STATUS=0 and STR_DATA2='" & Me.Label2.Text & "'"
        data.SelectCommand = str
        data.DataBind()
        data_dv = data.Select(New DataSourceSelectArguments)
        Dim _起始時間 As Date = DateSerial(Now.Year, Now.Month, Now.Day)
        Dim _終止時間 As Date = DateSerial(Now.Year, Now.Month, Now.Day)
        data.ConnectionString = con_ConnList
        str = "SELECT id,physicalDocID,createTime,subDocType3 FROM doc where createTime >='" & _起始時間 & "'and createTime<='" & _終止時間 & "'"

        str = "SELECT id,physicalDocID,createTime,subDocType3 FROM doc where id<1000"

        'str = "SELECT * FROM Users where commonName='" & Me.Label3.Text & "'"
        data.SelectCommand = str
        data.DataBind()
        data_dv = data.Select(New DataSourceSelectArguments)
        Me.GridView1.DataSource = data_dv
        Me.GridView1.DataBind()
    End Sub
End Class
