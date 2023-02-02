Imports Microsoft.Office.Interop
Imports System.Data.SqlClient
Imports System.Data.Sql
Imports System.Data.SqlTypes
Imports System.Configuration
Imports System.Text.RegularExpressions
Partial Class 電子交換失敗
    Inherits System.Web.UI.Page
    'Dim con_wf2 As String = "Data Source=edocsql.freeway.gov.tw\SQL2012;Initial Catalog=CFW_wf2;User ID=CFWedocdb01;Password=99$edocdb$CFW"
    Dim con_wf2 As String = "Data Source=edocsqlplus.freeway.gov.tw\SQL2019,54399;Initial Catalog=CFW_wf2;User ID=CFWedocdb01;Password=99$edocdb$CFW"
    Dim con_14 As String = "Data Source=10.52.0.178;Initial Catalog=大宗郵件;User ID=qaz;Password=1qaz@WSX"
    Dim data As New SqlDataSource
    Dim data_dv As Data.DataView
    Dim data_dv1 As Data.DataView
    Dim data_dv2 As Data.DataView
    Dim data_dv3 As Data.DataView
    Private misValue As Object
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
            Me.TextBox2.Text = card_id

            Dim _年, _月, _日 As String
            _年 = (Now.Year - 1911).ToString
            _月 = Now.Month.ToString("00")
            _日 = Now.Day.ToString("00")
            
            Dim j As Long
            For i As Long = 0 To Val(_年) - 108
                j = i + 109
                Me.DropDownList1.Items.Add(j.ToString("000"))
                Me.DropDownList1.Items(i).Value = j.ToString("000")
            Next
            Me.DropDownList1.DataBind()
            Me.DropDownList1.SelectedValue = _年
            For i As Long = 0 To 11
                Me.DropDownList2.Items.Add((i + 1).ToString("00"))
                Me.DropDownList2.Items(i).Value = (i + 1).ToString("00")
            Next
            Me.DropDownList2.DataBind()
            Me.DropDownList2.SelectedValue = _月
            For i As Long = 0 To DateTime.DaysInMonth(Now.Year, Now.Month) - 1
                Me.DropDownList3.Items.Add((i + 1).ToString("00"))
                Me.DropDownList3.Items(i).Value = (i + 1).ToString("00")
            Next
            Me.DropDownList3.DataBind()
            Me.DropDownList3.SelectedValue = _日
            For i As Long = 0 To 1
                data.ConnectionString = con_14
                data.SelectCommand = "SELECT 批號 FROM 大宗郵件執據 WHERE 批號='"&(i + 1).ToString("0")&"' and 年='" & Me.DropDownList1.SelectedValue & "' and 月='" & Me.DropDownList2.SelectedValue & "' and 日='" & Me.DropDownList3.SelectedValue & "'"
                data.DataBind()
                data_dv = data.Select(New DataSourceSelectArguments)
                Me.DropDownList7.Items.Add((i + 1).ToString("0") + "(共" + data_dv.Count.ToString("0") + "件)")
                Me.DropDownList7.Items(i).Value = (i + 1).ToString("0")
            Next
            data.ConnectionString = con_14
            data.DeleteCommand = "DELETE FROM 大宗郵件執據_bak WHERE 帳號='" & Me.TextBox2.Text & "'"
            data.Delete()
            data.DataBind()
            Me.SqlDataSource1.DataBind()
            Me.GridView1.DataBind()
            Me.Label2.Text = ""
            Me.Button10.Visible = False
            Me.Button11.Visible = False
        End If
    End Sub
    Protected Sub DropDownList1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DropDownList1.SelectedIndexChanged
        Me.DropDownList3.Items.Clear()
        For i As Long = 0 To DateTime.DaysInMonth((Val(Me.DropDownList1.SelectedValue) + 1911), Val(Me.DropDownList2.SelectedValue)) - 1
            Me.DropDownList3.Items.Add((i + 1).ToString("00"))
            Me.DropDownList3.Items(i).Value = (i + 1).ToString("00")
        Next
        Me.DropDownList3.DataBind()
        Me.DropDownList7.Items.Clear()
        For i As Long = 0 To 1
            data.ConnectionString = con_14
            data.SelectCommand = "SELECT 批號 FROM 大宗郵件執據 WHERE 批號='"&(i + 1).ToString("0")&"' and 年='" & Me.DropDownList1.SelectedValue & "' and 月='" & Me.DropDownList2.SelectedValue & "' and 日='" & Me.DropDownList3.SelectedValue & "'"
            data.DataBind()
            data_dv = data.Select(New DataSourceSelectArguments)
            Me.DropDownList7.Items.Add((i + 1).ToString("0") + "(共" + data_dv.Count.ToString("0") + "件)")
            Me.DropDownList7.Items(i).Value = (i + 1).ToString("0")
        Next
        Me.DropDownList7.DataBind()
    End Sub
    Protected Sub DropDownList2_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DropDownList2.SelectedIndexChanged
        Me.DropDownList3.Items.Clear()
        For i As Long = 0 To DateTime.DaysInMonth((Val(Me.DropDownList1.SelectedValue) + 1911), Val(Me.DropDownList2.SelectedValue)) - 1
            Me.DropDownList3.Items.Add((i + 1).ToString("00"))
            Me.DropDownList3.Items(i).Value = (i + 1).ToString("00")
        Next
        Me.DropDownList3.DataBind()
        Me.DropDownList7.Items.Clear()
        For i As Long = 0 To 1
            data.ConnectionString = con_14
            data.SelectCommand = "SELECT 批號 FROM 大宗郵件執據 WHERE 批號='"&(i + 1).ToString("0")&"' and 年='" & Me.DropDownList1.SelectedValue & "' and 月='" & Me.DropDownList2.SelectedValue & "' and 日='" & Me.DropDownList3.SelectedValue & "'"
            data.DataBind()
            data_dv = data.Select(New DataSourceSelectArguments)
            Me.DropDownList7.Items.Add((i + 1).ToString("0") + "(共" + data_dv.Count.ToString("0") + "件)")
            Me.DropDownList7.Items(i).Value = (i + 1).ToString("0")
        Next
        Me.DropDownList7.DataBind()
    End Sub
    Protected Sub DropDownList3_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DropDownList3.SelectedIndexChanged
        Me.DropDownList7.Items.Clear()
        For i As Long = 0 To 1
            data.ConnectionString = con_14
            data.SelectCommand = "SELECT 批號 FROM 大宗郵件執據 WHERE 批號='"&(i + 1).ToString("0")&"' and 年='" & Me.DropDownList1.SelectedValue & "' and 月='" & Me.DropDownList2.SelectedValue & "' and 日='" & Me.DropDownList3.SelectedValue & "'"
            data.DataBind()
            data_dv = data.Select(New DataSourceSelectArguments)
            Me.DropDownList7.Items.Add((i + 1).ToString("0") + "(共" + data_dv.Count.ToString("0") + "件)")
            Me.DropDownList7.Items(i).Value = (i + 1).ToString("0")
        Next
        Me.DropDownList7.DataBind()
    End Sub
    Protected Sub Button2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button2.Click
        data.ConnectionString = con_14
        Me.Label2.Text = ""
        Me.TextBox3.Text = ""
        Me.TextBox3.Visible = False
        data.DeleteCommand = "DELETE FROM 大宗郵件執據_bak WHERE 帳號='" & Me.TextBox2.Text & "'"
        data.Delete()
        data.DataBind()
        Dim _掛號類別 As String = ""
        data.SelectCommand = "SELECT 掛號類別 FROM  大宗郵件執據_郵寄種類 where 序號 ='1'"
        data.DataBind()
        data_dv = data.Select(New DataSourceSelectArguments)
        If data_dv.Count > 0
            _掛號類別 = Trim(data_dv(0)(0).ToString())
        End If
        data.ConnectionString = con_wf2
        Dim _收費小組 As Long = 0
        data.SelectCommand = "SELECT int_data2 FROM CREATE_DOC INNER JOIN  COM_DATA ON CREATE_DOC.CREATE_MAN = COM_DATA.ID INNER JOIN  USERS ON COM_DATA.STR_DATA2 = USERS.EMP_NO INNER JOIN  DEPT ON USERS.DEPT_ID = DEPT.ID  where (int_data2=5788 or int_data2=11286) and FULL_NO='" & Me.TextBox1.Text & "'"
        data.DataBind()
        data_dv = data.Select(New DataSourceSelectArguments)
        If data_dv.Count > 0 Then
            _收費小組 = 1
        Else
            _收費小組 = 0
        End If
        data.SelectCommand = "SELECT id,SUBJECT FROM  CREATE_DOC where FULL_NO ='" & Me.TextBox1.Text & " '"
        data.DataBind()
        data_dv = data.Select(New DataSourceSelectArguments)
        Dim _id As Long = 0
        Dim _收件人 As String = ""
        Dim _地址 As String = ""
        Dim _郵遞區號 As String = ""
        Dim _yn As Long = 1
        ''增加部份開始
        Dim _收件 As String = ""
        Dim _TEMP_DOC_ID As Long = 1
        Dim _長度 As Long = 0
        Dim r, r1 As Long
        ''增加部份 結束
        If data_dv.Count > 0 Then
            Me.TextBox3.Visible = True
            Me.TextBox3.Text = "主旨：" + Trim(data_dv(0)(1).ToString())
            _id = data_dv(0)(0).ToString()
            data.SelectCommand = "SELECT ACC_NAME,ADDR,ADDR_CODE,TEMP_DOC_ID,DELI_WAY FROM  ACCEPTER where  DOC_ID ='" & _id & " 'ORDER BY ID"
            data.DataBind()
            data_dv = data.Select(New DataSourceSelectArguments)
            ''增加部份開始
            _TEMP_DOC_ID = 0
            If data_dv.Count > 0 Then
                _TEMP_DOC_ID = data_dv(0)(3).ToString()
            End If
            data.SelectCommand = "SELECT EXCUSER FROM  TEMP_DOC_INFO where ID  ='" & _TEMP_DOC_ID & " '"
            data.DataBind()
            data_dv2 = data.Select(New DataSourceSelectArguments)
            _收件 = ""
            If data_dv.Count > 0 Then
                _收件 = Trim(data_dv(0)(0).ToString())
            End If
            _長度 = Len(Trim(_收件))
            r = 0
            r1 = 0
            For k As Long = 1 To _長度
                If Mid(_收件, k, 1) = "," And r = 0 Then
                    _收件人 = Left(_收件, k - 1)
                    r1 = r1 + 1
                    _收件 = Right(_收件, _長度 - k)
                    _長度 = Len(Trim(_收件))
                    r = 1
                End If
            Next
            Dim _郵寄種類 As Long = 1
            'data.ConnectionString = con_14
            'data.SelectCommand = "SELECT 郵資 FROM 大宗郵件執據_資費表 WHERE 序號='"& _郵寄種類 &"' ORDER BY 郵資"
            'data.DataBind()
            'data_dv = data.Select(New DataSourceSelectArguments)
            
            Dim _郵資 As Long = 28
            data.ConnectionString = con_wf2
            data.SelectCommand = "SELECT ACC_NAME,ADDR,ADDR_CODE,TEMP_DOC_ID,DELI_WAY FROM  ACCEPTER where  DOC_ID ='" & _id & " 'ORDER BY ID"
            data.DataBind()
            data_dv = data.Select(New DataSourceSelectArguments)
            ''增加部份 結束
            For i As Long = 0 To data_dv.Count - 1
                _收件人 = _收件
                ''增加部份開始
                If data_dv.Count <> r1 Then
                    r = 0
                    For k As Long = 1 To _長度
                        If Mid(_收件, k, 1) = "," And r = 0 Then
                            _收件人 = Left(_收件, k - 1)
                            _收件 = Right(_收件, _長度 - k)
                            _長度 = Len(Trim(_收件))
                            r = 1
                        End If
                    Next
                    _收件人 = data_dv(i)(0).ToString()
                End If
                ''增加部份 結束
                If data_dv(i)(4).ToString() = 2 Or data_dv(i)(4).ToString() = 4 Then
                    _地址 = data_dv(i)(1).ToString()
                    _郵遞區號 = data_dv(i)(2).ToString()
                    data.ConnectionString = con_14
                    data.InsertCommand = "INSERT INTO 大宗郵件執據_bak(帳號,收件人,地址,文號,郵資,郵寄種類,郵遞區號,yn,掛號類別,收費小組,附件) VALUES ('" & Me.TextBox2.Text & "',NULLIF(N'" & _收件人 & "', ''),NULLIF(N'" & _地址 & "', ''),'" & Trim(Me.TextBox1.Text) & "','" & _郵資 & "','" & _郵寄種類 & "',NULLIF('" & _郵遞區號 & "', ''),'" & _yn & "',NULLIF('" & _掛號類別 & "', ''),'" & _收費小組 & "','0')"
                    data.Insert()
                    data.DataBind()
                End If
            Next
        End If
        Me.SqlDataSource1.DataBind()
        Me.GridView1.DataBind()
        Me.Button10.Visible = True
        Me.Button11.Visible = True
    End Sub
    '存檔
    Protected Sub GridView1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.SelectedIndexChanged
        data.ConnectionString = con_14
        For i = 0 To Me.GridView1.Rows.Count - 1
            Dim _收費小組 As Long = 0
            Dim _id As Long = CType(Me.GridView1.Rows(i).FindControl("Label1"), Label).Text()
            Dim _備註 As String = Trim(CType(Me.GridView1.Rows(i).FindControl("TextBox9"), TextBox).Text())
            Dim _重量 As String = CType(Me.GridView1.Rows(i).FindControl("TextBox7"), TextBox).Text()
            Dim _郵資 As String = CType(Me.GridView1.Rows(i).FindControl("DropDownList8"), DropDownList).SelectedValue
            Dim _郵寄種類 As Long = CType(Me.GridView1.Rows(i).FindControl("DropDownList4"), DropDownList).SelectedValue
            Dim _收件人 As String = CType(Me.GridView1.Rows(i).FindControl("TextBox18"), TextBox).Text()
            Dim _郵遞區號 As String = CType(Me.GridView1.Rows(i).FindControl("TextBox56"), TextBox).Text()
            Dim _地址 As String = CType(Me.GridView1.Rows(i).FindControl("TextBox19"), TextBox).Text()
            Dim _文號 As String = CType(Me.GridView1.Rows(i).FindControl("TextBox22"), TextBox).Text()
            Dim _附件 As Long = 0
            If CType(Me.GridView1.Rows(i).FindControl("CheckBox40"), CheckBox).Checked Then
                _附件 = 1
            Else
                _附件 = 0
            End If
            If CType(Me.GridView1.Rows(i).FindControl("CheckBox2"), CheckBox).Checked Then
                _收費小組 = 1
            Else
                _收費小組 = 0
            End If
            Dim _掛號類別 As String = ""
            data.SelectCommand = "SELECT 掛號類別 FROM  大宗郵件執據_郵寄種類 where 序號 ='" & _郵寄種類 & " '"
            data.DataBind()
            data_dv = data.Select(New DataSourceSelectArguments)
            _掛號類別 = Trim(data_dv(0)(0).ToString())
            data.UpdateCommand = "update 大宗郵件執據_bak set 備註=N'" & _備註 & "' , 重量='" & _重量 & "' , 郵資='" & _郵資 & "' , 郵寄種類='" & _郵寄種類 & "' , 收件人=N'" & _收件人 & "' , 郵遞區號=NULLIF('" & _郵遞區號 & "', ''), 地址=N'" & _地址 & "' , 文號='" & _文號 & "', 附件='" & _附件 & "', 掛號類別='" & _掛號類別 & "' , 收費小組='" & _收費小組 & "' where id='" & _id & "'"
            data.Update()
            data.DataBind()
        Next
        Me.SqlDataSource1.DataBind()
        Me.GridView1.DataBind()
    End Sub

    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        If e.Row.RowIndex > -1 Then
            Dim dd4 As New DropDownList
            Dim _id As Long = CType(e.Row.FindControl("Label1"), Label).Text()
            data.ConnectionString = con_14
            data.SelectCommand = "SELECT 郵寄種類, 郵資 FROM 大宗郵件執據_bak WHERE id='"& _id &"'"
            data.DataBind()
            data_dv = data.Select(New DataSourceSelectArguments)
            Dim _序號 As Long = 1
            If Not IsDBNull(data_dv(0)(0))
                _序號 = data_dv(0)(0)
            End If
            Dim _郵資 As Long = 28
            If Not IsDBNull(data_dv(0)(1))
                _郵資 = data_dv(0)(1)
            End If
            data.SelectCommand = "SELECT 郵寄種類, 序號 FROM 大宗郵件執據_郵寄種類 ORDER BY 排序"
            data.DataBind()
            data_dv = data.Select(New DataSourceSelectArguments)
            dd4 = e.Row.FindControl("DropDownList4")
            dd4.Items.Clear()
            For i As Long = 0 To data_dv.Count - 1
                dd4.Items.Add(data_dv(i)(0))
                dd4.Items(i).Value = data_dv(i)(1)
            Next
            dd4.SelectedIndex = dd4.Items.IndexOf(dd4.Items.FindByValue(_序號))
            dd4.DataBind()
            Dim dd8 As New DropDownList
            data.SelectCommand = "SELECT 郵資 FROM 大宗郵件執據_資費表 WHERE 序號='"& _序號 &"' ORDER BY 郵資"
            data.DataBind()
            data_dv = data.Select(New DataSourceSelectArguments)
            dd8 = e.Row.FindControl("DropDownList8")
            dd8.Items.Clear()
            For i As Long = 0 To data_dv.Count - 1
                dd8.Items.Add(data_dv(i)(0))
                dd8.Items(i).Value = data_dv(i)(0)
            Next
            dd8.SelectedIndex = dd8.Items.IndexOf(dd8.Items.FindByValue(_郵資))
            dd8.DataBind()

            Dim cb40 As New CheckBox
            cb40 = e.Row.FindControl("CheckBox40")
            data.SelectCommand = "SELECT 附件 FROM 大宗郵件執據_bak WHERE id='"& _id &"'"
            data.DataBind()
            data_dv = data.Select(New DataSourceSelectArguments)
            If Not IsDBNull(data_dv(0)(0))
                cb40.checked = data_dv(0)(0)
            End If
            cb40.DataBind()
        End If
    End Sub

    Protected Sub GridView1_CheckBox40_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim _id As Long = CType(Me.GridView1.Rows(sender.NamingContainer.RowIndex).FindControl("Label1"), Label).Text()
        Dim _文號 As String = CType(Me.GridView1.Rows(sender.NamingContainer.RowIndex).FindControl("TextBox22"), TextBox).Text()
        If sender.checked
            _文號 = _文號.Replace(" ", "")
            _文號 = _文號.Trim()
            _文號 = Long.Parse(Regex.Replace(_文號, "[^\d]", ""))
            _文號 = _文號 + "附件"
        Else
            _文號 = Long.Parse(Regex.Replace(_文號, "[^\d]", ""))
        End If
        'data.ConnectionString = con_14
        'data.UpdateCommand = "update 大宗郵件執據_bak set 附件='" & sender.checked & "' where id='" & _id & "'"
        'data.Update()
        'data.DataBind()
        CType(Me.GridView1.Rows(sender.NamingContainer.RowIndex).FindControl("TextBox22"), TextBox).Text() = _文號
    End Sub

    Protected Sub GridView2_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView2.RowDataBound
        If e.Row.RowIndex > -1 Then
            Dim dd4 As New DropDownList
            Dim _id As Long = CType(e.Row.FindControl("Label1"), Label).Text()
            data.ConnectionString = con_14
            data.SelectCommand = "SELECT 郵寄種類, 郵資 FROM 大宗郵件執據 WHERE id='"& _id &"'"
            data.DataBind()
            data_dv = data.Select(New DataSourceSelectArguments)
            Dim _序號 As Long = 1
            If Not IsDBNull(data_dv(0)(0))
                _序號 = data_dv(0)(0)
            End If
            Dim _郵資 As Long = 28
            If Not IsDBNull(data_dv(0)(1))
                _郵資 = data_dv(0)(1)
            End If
            data.SelectCommand = "SELECT 郵寄種類, 序號 FROM 大宗郵件執據_郵寄種類 ORDER BY 排序"
            data.DataBind()
            data_dv = data.Select(New DataSourceSelectArguments)
            dd4 = e.Row.FindControl("DropDownList4")
            dd4.Items.Clear()
            For i As Long = 0 To data_dv.Count - 1
                dd4.Items.Add(data_dv(i)(0))
                dd4.Items(i).Value = data_dv(i)(1)
            Next
            dd4.SelectedIndex = dd4.Items.IndexOf(dd4.Items.FindByValue(_序號))
            dd4.DataBind()
            Dim dd8 As New DropDownList
            data.SelectCommand = "SELECT 郵資 FROM 大宗郵件執據_資費表 WHERE 序號='"& _序號 &"' ORDER BY 郵資"
            data.DataBind()
            data_dv = data.Select(New DataSourceSelectArguments)
            dd8 = e.Row.FindControl("DropDownList8")
            dd8.Items.Clear()
            For i As Long = 0 To data_dv.Count - 1
                dd8.Items.Add(data_dv(i)(0))
                dd8.Items(i).Value = data_dv(i)(0)
            Next
            If dd8.Items.IndexOf(dd8.Items.FindByValue(_郵資)) = -1
                dd8.Items.Add(_郵資)
                dd8.Items(data_dv.Count).Value = _郵資
            End If
            dd8.SelectedIndex = dd8.Items.IndexOf(dd8.Items.FindByValue(_郵資))
            dd8.DataBind()
        End If
    End Sub
    Protected Sub GridView1_DropDownList4_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim _id As Long = CType(Me.GridView1.Rows(sender.NamingContainer.RowIndex).FindControl("Label1"), Label).Text()
        Dim dd8 As DropDownList = CType(Me.GridView1.Rows(sender.NamingContainer.RowIndex).FindControl("DropDownList8"), DropDownList)
        Dim _序號 As Long = sender.SelectedValue
        data.ConnectionString = con_14
        data.SelectCommand = "SELECT 郵資 FROM 大宗郵件執據_資費表 WHERE 序號='" & _序號 &"' ORDER BY 郵資"
        data.DataBind()
        data_dv = data.Select(New DataSourceSelectArguments)
        dd8.Items.Clear()
        For i As Long = 0 To data_dv.Count - 1
            dd8.Items.Add(data_dv(i)(0))
            dd8.Items(i).Value = data_dv(i)(0)
        Next
        Dim _郵資 As String = ""
        If data_dv.Count > 1
            If Trim(data_dv(0)(0).toString) = "0"
                _郵資 = data_dv(1)(0).toString
            Else
                _郵資 = data_dv(0)(0).toString
            End If
        ElseIf data_dv.Count = 1
            _郵資 = data_dv(0)(0).toString
        End If
        If dd8.Items.IndexOf(dd8.Items.FindByValue(_郵資)) = -1
            dd8.Items.Add(_郵資)
            dd8.Items(data_dv.Count).Value = _郵資
        End If
        dd8.SelectedIndex = dd8.Items.IndexOf(dd8.Items.FindByValue(_郵資))
        dd8.DataBind()
        'data.UpdateCommand = "update 大宗郵件執據_bak set 郵資='" & _郵資 & "' , 郵寄種類='" & _序號 & "' where id='" & _id & "'"
        'data.Update()
        'data.DataBind()
        'Me.SqlDataSource1.DataBind()
        'Me.GridView1.DataBind()
    End Sub
    Protected Sub GridView2_DropDownList4_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim _id As Long = CType(Me.GridView2.Rows(sender.NamingContainer.RowIndex).FindControl("Label1"), Label).Text()
        Dim dd8 As DropDownList = CType(Me.GridView2.Rows(sender.NamingContainer.RowIndex).FindControl("DropDownList8"), DropDownList)
        Dim _序號 As Long = sender.SelectedValue
        data.ConnectionString = con_14
        data.SelectCommand = "SELECT 郵資 FROM 大宗郵件執據_資費表 WHERE 序號='" & _序號 &"' ORDER BY 郵資"
        data.DataBind()
        data_dv = data.Select(New DataSourceSelectArguments)
        dd8.Items.Clear()
        For i As Long = 0 To data_dv.Count - 1
            dd8.Items.Add(data_dv(i)(0))
            dd8.Items(i).Value = data_dv(i)(0)
        Next
        Dim _郵資 As String = ""
        If data_dv.Count > 1
            If Trim(data_dv(0)(0).toString) = "0"
                _郵資 = data_dv(1)(0).toString
            Else
                _郵資 = data_dv(0)(0).toString
            End If
        ElseIf data_dv.Count = 1
            _郵資 = data_dv(0)(0).toString
        End If
        If dd8.Items.IndexOf(dd8.Items.FindByValue(_郵資)) = -1
            dd8.Items.Add(_郵資)
            dd8.Items(data_dv.Count).Value = _郵資
        End If
        dd8.SelectedIndex = dd8.Items.IndexOf(dd8.Items.FindByValue(_郵資))
        dd8.DataBind()
        'data.UpdateCommand = "update 大宗郵件執據 set 郵資='" & _郵資 & "', 郵寄種類='" & _序號 & "' where id='" & _id & "'"
        'data.Update()
        'data.DataBind()
        'Me.SqlDataSource3.DataBind()
        'Me.GridView2.DataBind()
    End Sub

    Protected Sub Button3_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button3.Click
    '先存檔再說
        GridView1_SelectedIndexChanged(Nothing, Nothing)
        GridView2_SelectedIndexChanged(Nothing, Nothing)
        If  Me.GridView1.Rows.Count = 0
            return
        End If

        data.ConnectionString = con_14
        Me.Label2.Text = ""
        Dim _id As Long = 0
        Dim _yn As Long = 0
        Dim _收費小組 As Long = 0
        For i As Long = 0 To Me.GridView1.Rows.Count - 1
            _id = CType(Me.GridView1.Rows(i).FindControl("Label1"), Label).Text
            _yn = CType(Me.GridView1.Rows(i).FindControl("CheckBox1"), CheckBox).Checked
            If CType(Me.GridView1.Rows(i).FindControl("CheckBox2"), CheckBox).Checked Then
                _收費小組 = 1
            Else
                _收費小組 = 0
            End If
            data.UpdateCommand = "update 大宗郵件執據_bak set yn='" & _yn & "' , 收費小組='" & _收費小組 & "' where id='" & _id & "'"
            data.Update()
            data.DataBind()
        Next
        Dim _序號 As String = ""
        Dim _重量 As String = ""
        Dim _郵資 As String = ""
        Dim _郵寄種類 As Long = 0
        Dim _收件人 As String = ""
        Dim _地址 As String = ""
        Dim _文號 As String = ""
        Dim _備註 As String = ""
        Dim _郵遞區號 As String = ""
        Dim _掛號類別 As String = ""
        Dim _件數 As Long = 1
        data.SelectCommand = "SELECT 序號 FROM  大宗郵件執據 where 年='" & Me.DropDownList1.SelectedValue & " ' and 月='" & Me.DropDownList2.SelectedValue & " ' and 日='" & Me.DropDownList3.SelectedValue & "' and 批號='" & Me.DropDownList7.SelectedValue & "' ORDER BY 序號 DESC"
        data.DataBind()
        data_dv = data.Select(New DataSourceSelectArguments)
        If data_dv.Count > 0 Then
            _序號 = data_dv(0)(0).ToString() + 1
        Else
            _序號 = 1
        End If
        data.SelectCommand = "SELECT 序號,掛號號碼,收件人,地址,文號,備註,重量,郵資,郵寄種類,郵遞區號,yn,掛號類別,收費小組 FROM  大宗郵件執據_bak where 帳號='" & Me.TextBox2.Text & " ' order by 序號"
        data.DataBind()
        data_dv = data.Select(New DataSourceSelectArguments)
        Dim a As String = ""
        Dim b As String = ""
        For i As Long = 0 To data_dv.Count - 1
            If data_dv(i)(10).ToString() <> 0 Then
                '_掛號號碼 = data_dv(i)(1).ToString()
                _收件人 = Trim(data_dv(i)(2).ToString())
                _地址 = Trim(data_dv(i)(3).ToString())
                _文號 = Trim(data_dv(i)(4).ToString())
                _備註 = Trim(data_dv(i)(5).ToString())
                _重量 = data_dv(i)(6).ToString()
                _郵資 = data_dv(i)(7).ToString()
                _郵寄種類 = data_dv(i)(8).ToString()
                _郵遞區號 = data_dv(i)(9).ToString()
                _掛號類別 = data_dv(i)(11).ToString()
                _收費小組 = data_dv(i)(12).ToString()
                data.SelectCommand = "SELECT 文號,id,收件人 FROM  大宗郵件執據 where '、' + 收件人 + '、' Like N'%、" & _收件人 & "、%' and 年='" & Me.DropDownList1.SelectedValue & "' and 月='" & Me.DropDownList2.SelectedValue & "' and 日='" & Me.DropDownList3.SelectedValue & "' and 批號='" & Me.DropDownList7.SelectedValue & "'"
                data.DataBind()
                data_dv1 = data.Select(New DataSourceSelectArguments)
                If data_dv1.Count <= 0 Then
                    data.SelectCommand = "SELECT 文號,id,收件人 FROM  大宗郵件執據 where 地址 = N'" & _地址 & " ' and 年='" & Me.DropDownList1.SelectedValue & "' and 月='" & Me.DropDownList2.SelectedValue & "' and 日='" & Me.DropDownList3.SelectedValue & "' and 批號='" & Me.DropDownList7.SelectedValue & "'"
                    data.DataBind()
                    data_dv2 = data.Select(New DataSourceSelectArguments)
                    If data_dv2.Count <= 0 Then
                        data.InsertCommand = "INSERT INTO 大宗郵件執據(年,月,日,批號,序號,收件人,地址,文號,備註,重量,郵資,郵寄種類,郵遞區號,掛號類別,件數,收費小組) VALUES ('" & Me.DropDownList1.SelectedValue & "','" & Me.DropDownList2.SelectedValue & "','" & Me.DropDownList3.SelectedValue & "','" & Me.DropDownList7.SelectedValue & "',NULLIF('" & _序號 & "',''),NULLIF(N'" & _收件人 & "',''),NULLIF(N'" & _地址 & "',''),'" & _文號 & "','" & _備註 & "','" & _重量 & "','" & _郵資 & "','" & _郵寄種類 & "','" & _郵遞區號 & "',NULLIF('" & _掛號類別 & "',''),'" & _件數 & "','" & _收費小組 & "')"
                        data.Insert()
                        data.DataBind()
                        _序號 = _序號 + 1
                    Else
                        a = Trim(data_dv2(0)(0).ToString())
                        b = Trim(data_dv2(0)(2).ToString())
                        a = a + "、" + _文號
                        b = b + "、" + _收件人
                        _id = data_dv2(0)(1).ToString()
                        data.UpdateCommand = "update 大宗郵件執據 set 收件人 = '" & b & "', 文號='" & a & "' where id='" & _id & "'"
                        data.Update()
                        data.DataBind()
                        data.SelectCommand = "SELECT 序號 FROM  大宗郵件執據 where id='" & _id & "'"
                        data.DataBind()
                        data_dv3 = data.Select(New DataSourceSelectArguments)
                        Try
                            Me.Label2.Text = _收件人 + " 因地址相同 併入序號" + data_dv3(0)(0).toString
                        Catch
                        End Try
                    End If
                Else
                    a = Trim(data_dv1(0)(0).ToString())
                    a = a + "、" + _文號
                    _id = data_dv1(0)(1).ToString()
                    data.UpdateCommand = "update 大宗郵件執據 set 文號='" & a & "' where id='" & _id & "'"
                    data.Update()
                    data.DataBind()
                    data.SelectCommand = "SELECT 序號 FROM  大宗郵件執據 where id='" & _id & "'"
                    data.DataBind()
                    data_dv3 = data.Select(New DataSourceSelectArguments)
                    Try
                        Me.Label2.Text = _收件人 + " 因收件人相同 併入序號" + data_dv3(0)(0).toString
                    Catch
                    End Try
                End If
            End If
        Next
        Me.SqlDataSource3.DataBind()
        Me.GridView2.DataBind()
        data.DeleteCommand = "DELETE FROM 大宗郵件執據_bak WHERE 帳號='" & Me.TextBox2.Text & "'"
        data.Delete()
        data.DataBind()
        Me.SqlDataSource1.DataBind()
        Me.GridView1.DataBind()
        Me.TextBox3.Text = ""
        Me.Button10.Visible = False
        Me.Button11.Visible = False
        Dim d7index as Long = Me.DropDownList7.SelectedValue
        Me.DropDownList7.Items.Clear()
        For i As Long = 0 To 1
            data.ConnectionString = con_14
            data.SelectCommand = "SELECT 批號 FROM 大宗郵件執據 WHERE 批號='"&(i + 1).ToString("0")&"' and 年='" & Me.DropDownList1.SelectedValue & "' and 月='" & Me.DropDownList2.SelectedValue & "' and 日='" & Me.DropDownList3.SelectedValue & "'"
            data.DataBind()
            data_dv = data.Select(New DataSourceSelectArguments)
            Me.DropDownList7.Items.Add((i + 1).ToString("0") + "(共" + data_dv.Count.ToString("0") + "件)")
            Me.DropDownList7.Items(i).Value = (i + 1).ToString("0")
        Next
        Me.DropDownList7.SelectedValue = d7index
        Me.DropDownList7.DataBind()
        data.ConnectionString = con_14

        '重算序號
        data.SelectCommand = "SELECT id FROM  大宗郵件執據 where 年='" & Me.DropDownList1.SelectedValue & "' and 月='" & Me.DropDownList2.SelectedValue & "' and 日='" & Me.DropDownList3.SelectedValue & "' and 批號='" & Me.DropDownList7.SelectedValue & "' ORDER BY 序號"
        data.DataBind()
        data_dv = data.Select(New DataSourceSelectArguments)
        For i As Long = 0 to data_dv.Count - 1
            data.UpdateCommand = "update 大宗郵件執據 set 序號='" & (i+1) & "' where id='" & data_dv(i)(0) & "'"
            data.Update()
            data.DataBind()
        Next

        Me.SqlDataSource3.DataBind()
        Me.GridView2.DataBind()
    End Sub
    '存檔
    Protected Sub GridView2_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView2.SelectedIndexChanged
        data.ConnectionString = con_14
        For i = 0 To Me.GridView2.Rows.Count - 1
            Dim _收費小組 As Long = 0
            Dim _id As Long = CType(Me.GridView2.Rows(i).FindControl("Label1"), Label).Text()
            Dim _序號 As String = CType(Me.GridView2.Rows(i).FindControl("TextBox5"), TextBox).Text()
            Dim _掛號號碼 As String = CType(Me.GridView2.Rows(i).FindControl("TextBox25"), TextBox).Text()
            Dim _備註 As String = Trim(CType(Me.GridView2.Rows(i).FindControl("TextBox9"), TextBox).Text())
            Dim _重量 As String = CType(Me.GridView2.Rows(i).FindControl("TextBox7"), TextBox).Text()
            Dim _郵資 As String = CType(Me.GridView2.Rows(i).FindControl("DropDownList8"), DropDownList).SelectedValue
            Dim _收件人 As String = CType(Me.GridView2.Rows(i).FindControl("TextBox20"), TextBox).Text()
            Dim _文號 As String = CType(Me.GridView2.Rows(i).FindControl("TextBox22"), TextBox).Text()
            Dim _郵遞區號 As String = CType(Me.GridView2.Rows(i).FindControl("TextBox56"), TextBox).Text()
            Dim _地址 As String = CType(Me.GridView2.Rows(i).FindControl("TextBox21"), TextBox).Text()
            Dim _件數 As Long = CType(Me.GridView2.Rows(i).FindControl("TextBox26"), TextBox).Text()
            Dim _郵寄種類 As Long = CType(Me.GridView2.Rows(i).FindControl("DropDownList4"), DropDownList).SelectedValue
            Dim _掛號類別 As String = ""
            If CType(Me.GridView2.Rows(i).FindControl("CheckBox3"), CheckBox).Checked Then
                _收費小組 = 1
            Else
                _收費小組 = 0
            End If
            data.SelectCommand = "SELECT 掛號類別 FROM  大宗郵件執據_郵寄種類 where 序號 ='" & _郵寄種類 & " '"
            data.DataBind()
            data_dv = data.Select(New DataSourceSelectArguments)
            _掛號類別 = Trim(data_dv(0)(0).ToString())
            If Trim(_掛號號碼) <> ""
                _掛號號碼 = Clng(Trim(_掛號號碼)).ToString("000000")
            End If
            data.UpdateCommand = "update 大宗郵件執據 set 序號='" & _序號 & "' , 掛號號碼=NULLIF('" & _掛號號碼 & "', '') , 備註=N'" & _備註 & "' , 重量='" & _重量 & "' , 郵資='" & _郵資 & "' , 郵寄種類='" & _郵寄種類 & "' , 收件人=N'" & _收件人 & "', 郵遞區號=NULLIF('" & _郵遞區號 & "', ''), 地址=N'" & _地址 & "' , 文號='" & _文號 & "' , 掛號類別='" & _掛號類別 & "' , 收費小組='" & _收費小組 & "', 件數='" & _件數 & "' where id='" & _id & "'"
            data.Update()
            data.DataBind()
        Next
        Me.SqlDataSource3.DataBind()
        Me.GridView2.DataBind()
    End Sub

    Protected Sub Button5_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button5.Click
        data.ConnectionString = con_14
        Dim _id As Long = 0
        Dim _序號 As String = ""
        Dim _掛號號碼 As String = ""
        Dim _掛號類別 As String = ""
        Dim _備註 As String = ""
        Dim _重量 As String = ""
        Dim _文號 As String = ""
        Dim _收件人 As String = ""
        Dim _地址 As String = ""
        Dim _郵資 As String = ""
        Dim _郵寄種類 As Long = 0
        Dim _件數 As Long = 0
        Dim _收費小組 As Long = 0
        For i As Long = 0 To Me.GridView2.Rows.Count - 1
            _id = CType(Me.GridView2.Rows(i).FindControl("Label1"), Label).Text()
            _序號 = CType(Me.GridView2.Rows(i).FindControl("TextBox5"), TextBox).Text()
            _掛號號碼 = CType(Me.GridView2.Rows(i).FindControl("TextBox25"), TextBox).Text()
            _掛號類別 = CType(Me.GridView2.Rows(i).FindControl("TextBox6"), TextBox).Text()
            _備註 = Trim(CType(Me.GridView2.Rows(i).FindControl("TextBox9"), TextBox).Text())
            _收件人 = CType(Me.GridView2.Rows(i).FindControl("TextBox20"), TextBox).Text()
            _地址 = Trim(CType(Me.GridView2.Rows(i).FindControl("TextBox21"), TextBox).Text())
            _文號 = Trim(CType(Me.GridView2.Rows(i).FindControl("TextBox22"), TextBox).Text())
            _重量 = CType(Me.GridView2.Rows(i).FindControl("TextBox7"), TextBox).Text()
            _郵資 = CType(Me.GridView2.Rows(i).FindControl("DropDownList8"), DropDownList).SelectedValue
            _郵寄種類 = CType(Me.GridView2.Rows(i).FindControl("DropDownList4"), DropDownList).SelectedValue
            _件數 = CType(Me.GridView2.Rows(i).FindControl("TextBox26"), TextBox).Text()
            If CType(Me.GridView2.Rows(i).FindControl("CheckBox3"), CheckBox).Checked Then
                _收費小組 = 1
            Else
                _收費小組 = 0
            End If
            data.SelectCommand = "SELECT 重量 FROM  大宗郵件執據_資費表 where 序號 ='" & _郵寄種類 & " ' and 郵資='" & _郵資 & "'"
            data.DataBind()
            data_dv = data.Select(New DataSourceSelectArguments)
            If data_dv.Count > 0 Then
                _重量 = data_dv(0)(0).ToString()
            End If
            data.ConnectionString = con_14
            If Trim(_掛號號碼) <> ""
                _掛號號碼 = Clng(Trim(_掛號號碼)).ToString("000000")
            End If
            data.UpdateCommand = "update 大宗郵件執據 set 序號='" & _序號 & "' , 掛號號碼=NULLIF('" & _掛號號碼 & "', '') , 備註=N'" & _備註 & "' , 重量='" & _重量 & "' , 郵資='" & _郵資 & "' , 郵寄種類='" & _郵寄種類 & "' , 收件人=N'" & _收件人 & "' , 地址=N'" & _地址 & "' , 文號='" & _文號 & "' , 掛號類別='" & _掛號類別 & "' , 收費小組='" & _收費小組 & "' where id='" & _id & "'"
            data.Update()
            data.DataBind()
            data.UpdateCommand = "update 大宗郵件執據 set 件數='" & _件數 & "' where id='" & _id & "'"
            data.Update()
            data.DataBind()
        Next
        Me.SqlDataSource3.DataBind()
        Me.GridView2.DataBind()
    End Sub

    Protected Sub GridView2_RowCommand(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles GridView2.RowCommand
        'Me.TextBox1.Text = e.CommandName
        'if e.CommandName = "delete"
        if e.CommandName = "delete2"
            '重算序號
            data.ConnectionString = con_14
            data.SelectCommand = "SELECT id FROM  大宗郵件執據 where 年='" & Me.DropDownList1.SelectedValue & "' and 月='" & Me.DropDownList2.SelectedValue & "' and 日='" & Me.DropDownList3.SelectedValue & "' and 批號='" & Me.DropDownList7.SelectedValue & "' ORDER BY 序號"
            data.DataBind()
            data_dv = data.Select(New DataSourceSelectArguments)
            For i As Long = 0 to data_dv.Count - 1
                data.UpdateCommand = "update 大宗郵件執據 set 序號='" & (i+1) & "' where id='" & data_dv(i)(0) & "'"
                data.Update()
                data.DataBind()
            Next

            Me.SqlDataSource3.DataBind()
            Me.GridView2.DataBind()

            Dim d7index as Long = Me.DropDownList7.SelectedValue
            Me.DropDownList7.Items.Clear()
            For i As Long = 0 To 1
                data.ConnectionString = con_14
                data.SelectCommand = "SELECT 批號 FROM 大宗郵件執據 WHERE 批號='"&(i + 1).ToString("0")&"' and 年='" & Me.DropDownList1.SelectedValue & "' and 月='" & Me.DropDownList2.SelectedValue & "' and 日='" & Me.DropDownList3.SelectedValue & "'"
                data.DataBind()
                data_dv = data.Select(New DataSourceSelectArguments)
                if i <> d7index - 1
                    Me.DropDownList7.Items.Add((i + 1).ToString("0") + "(共" + data_dv.Count.ToString("0") + "件)")
                Else 
                    Me.DropDownList7.Items.Add((i + 1).ToString("0") + "(共" + (data_dv.Count - 1).ToString("0") + "件)")
                End if
                Me.DropDownList7.Items(i).Value = (i + 1).ToString("0")
            Next
            Me.DropDownList7.SelectedValue = d7index
            Me.DropDownList7.DataBind()
        End If
    End Sub

    Protected Sub Button9_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button9.Click
        data.ConnectionString = con_14
        Me.Label2.Text = ""
        Me.TextBox3.Text = ""
        Me.TextBox3.Visible = False
        data.DeleteCommand = "DELETE FROM 大宗郵件執據_bak WHERE 帳號='" & Me.TextBox2.Text & "'"
        data.Delete()
        data.DataBind()
        Dim _掛號類別 As String = ""
        data.SelectCommand = "SELECT 掛號類別 FROM  大宗郵件執據_郵寄種類 where 序號 ='1'"
        data.DataBind()
        data_dv = data.Select(New DataSourceSelectArguments)
        _掛號類別 = Trim(data_dv(0)(0).ToString())
        data.ConnectionString = con_wf2
        Dim _收費小組 As Long = 0
        data.SelectCommand = "SELECT int_data2 FROM CREATE_DOC INNER JOIN  COM_DATA ON CREATE_DOC.CREATE_MAN = COM_DATA.ID INNER JOIN  USERS ON COM_DATA.STR_DATA2 = USERS.EMP_NO INNER JOIN  DEPT ON USERS.DEPT_ID = DEPT.ID  where (int_data2=5788 or int_data2=11286) and FULL_NO='" & Me.TextBox1.Text & "'"
        data.DataBind()
        data_dv = data.Select(New DataSourceSelectArguments)
        If data_dv.Count > 0 Then
            _收費小組 = 1
        End If
        data.ConnectionString = con_14
        data.SelectCommand = "SELECT 郵資 from 大宗郵件執據_資費表 where 序號 ='1'"
        data.DataBind()
        data_dv = data.Select(New DataSourceSelectArguments)
        Dim _郵資 As Long = data_dv(0)(0)
        data.ConnectionString = con_wf2
        data.SelectCommand = "SELECT id,SUBJECT FROM  CREATE_DOC where FULL_NO ='" & Me.TextBox1.Text & " '"
        data.DataBind()
        data_dv = data.Select(New DataSourceSelectArguments)
        Dim _id As Long = 0
        Dim _收件人 As String = ""
        Dim _地址 As String = ""
        Dim _郵遞區號 As String = ""
        Dim _yn As Long = 1
        If data_dv.Count > 0 Then
            Me.TextBox3.Visible = True
            Me.TextBox3.Text = "主旨：" + Trim(data_dv(0)(1).ToString())
            _id = data_dv(0)(0).ToString()
            data.SelectCommand = "SELECT ACC_NAME,ADDR,ADDR_CODE FROM  ACCEPTER where DELI_WAY=1  and DOC_ID ='" & _id & " 'ORDER BY ID"
            data.DataBind()
            data_dv = data.Select(New DataSourceSelectArguments)
            data.ConnectionString = con_14
            Dim _郵寄種類 As Long = 1
            For i As Long = 0 To data_dv.Count - 1
                _收件人 = data_dv(i)(0).ToString()
                _地址 = data_dv(i)(1).ToString()
                _郵遞區號 = data_dv(i)(2).ToString()
                data.InsertCommand = "INSERT INTO 大宗郵件執據_bak(帳號,收件人,地址,文號,郵寄種類,郵遞區號,yn,掛號類別,收費小組,郵資,附件) VALUES ('" & Me.TextBox2.Text & "',NULLIF(N'" & _收件人 & "',''),NULLIF(N'" & _地址 & "',''),NULLIF('" & Trim(Me.TextBox1.Text) & "',''),'" & _郵寄種類 & "',NULLIF('" & _郵遞區號 & "',''),'" & _yn & "',NULLIF('" & _掛號類別 & "',''),'" & _收費小組 & "','" & _郵資 & "','0')"
                data.Insert()
                data.DataBind()
            Next
        End If
        Me.SqlDataSource1.DataBind()
        Me.GridView1.DataBind()
        Me.Button10.Visible = True
        Me.Button11.Visible = True
    End Sub
    Protected Sub Button10_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button10.Click
        'Select Case ID, 序號, 掛號號碼, 收件人, 地址, 文號, 備註, 重量, 郵資 ,yn,郵寄種類,收費小組
        'From 大宗郵件執據_bak
        'Where 帳號 =@_帳號

        data.ConnectionString = con_14
        data.UpdateCommand = "update 大宗郵件執據_bak set yn=1 where 帳號 ='" & Me.TextBox2.Text & " '"
        data.Update()
        data.DataBind()
        Me.SqlDataSource1.DataBind()
        Me.GridView1.DataBind()





        'For i As Long = 0 To Me.GridView1.Rows.Count - 1
        '    CType(Me.GridView1.Rows(i).FindControl("CheckBox1"), CheckBox).Checked = True
        'Next
    End Sub
    Protected Sub Button11_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button11.Click
        data.ConnectionString = con_14
        data.UpdateCommand = "update 大宗郵件執據_bak set yn=0 where 帳號 ='" & Me.TextBox2.Text & " '"
        data.Update()
        data.DataBind()
        Me.SqlDataSource1.DataBind()
        Me.GridView1.DataBind()

        'For i As Long = 0 To Me.GridView1.Rows.Count - 1
        '    CType(Me.GridView1.Rows(i).FindControl("CheckBox1"), CheckBox).Checked = False
        'Next
    End Sub
    Protected Sub Button12_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button12.Click
        data.ConnectionString = con_14
        Me.Label2.Text = ""
        Me.TextBox3.Text = ""
        Me.TextBox3.Visible = False
        data.DeleteCommand = "DELETE FROM 大宗郵件執據_bak WHERE 帳號='" & Me.TextBox2.Text & "'"
        data.Delete()
        data.DataBind()
        Dim _掛號類別 As String = ""
        data.SelectCommand = "SELECT 掛號類別 FROM  大宗郵件執據_郵寄種類 where 序號 ='" & 1 & " '"
        data.DataBind()
        data_dv = data.Select(New DataSourceSelectArguments)
        _掛號類別 = Trim(data_dv(0)(0).ToString())
        data.ConnectionString = con_wf2
        Dim _收費小組 As Long = 0
        data.SelectCommand = "SELECT int_data2 FROM CREATE_DOC INNER JOIN  COM_DATA ON CREATE_DOC.CREATE_MAN = COM_DATA.ID INNER JOIN  USERS ON COM_DATA.STR_DATA2 = USERS.EMP_NO INNER JOIN  DEPT ON USERS.DEPT_ID = DEPT.ID  where (int_data2=5788 or int_data2=11286) and FULL_NO='" & Me.TextBox1.Text & "'"
        data.DataBind()
        data_dv = data.Select(New DataSourceSelectArguments)
        If data_dv.Count > 0 Then
            _收費小組 = 1
        Else
            _收費小組 = 0
        End If
        data.SelectCommand = "SELECT id,SUBJECT FROM  CREATE_DOC where FULL_NO ='" & Me.TextBox1.Text & " '"
        data.DataBind()
        data_dv = data.Select(New DataSourceSelectArguments)
        Dim _id As Long = 0
        Dim _收件人 As String = ""
        Dim _地址 As String = ""
        Dim _郵遞區號 As String = ""
        Dim _yn As Long = 1
        If data_dv.Count > 0 Then
            Me.TextBox3.Visible = True
            Me.TextBox3.Text = "主旨：" + Trim(data_dv(0)(1).ToString())
            _id = data_dv(0)(0).ToString()
            data.SelectCommand = "SELECT ACC_NAME,ADDR,ADDR_CODE FROM  ACCEPTER where DELI_WAY=3  and DOC_ID ='" & _id & " 'ORDER BY ID"
            data.DataBind()
            data_dv = data.Select(New DataSourceSelectArguments)
            data.ConnectionString = con_14
            Dim _郵寄種類 As Long = 1
            For i As Long = 0 To data_dv.Count - 1
                _收件人 = data_dv(i)(0).ToString()
                _地址 = data_dv(i)(1).ToString()
                _郵遞區號 = data_dv(i)(2).ToString()
                data.InsertCommand = "INSERT INTO 大宗郵件執據_bak(帳號,收件人,地址,文號,郵寄種類,郵遞區號,yn,掛號類別,收費小組,附件) VALUES ('" & Me.TextBox2.Text & "',NULLIF(N'" & _收件人 & "',''),NULLIF(N'" & _地址 & "',''),NULLIF(N'" & Trim(Me.TextBox1.Text) & "',''),'" & _郵寄種類 & "',NULLIF('" & _郵遞區號 & "',''),'" & _yn & "',NULLIF('" & _掛號類別 & "',''),'" & _收費小組 & "','0')"
                data.Insert()
                data.DataBind()
            Next
        End If
        Me.SqlDataSource1.DataBind()
        Me.GridView1.DataBind()
        Me.Button10.Visible = True
        Me.Button11.Visible = True
    End Sub
    Protected Sub Button13_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button13.Click
        data.ConnectionString = con_14
        Dim _郵寄種類 As Long
        Dim _序號 As Long
        Dim _id As Long
        data.SelectCommand = "SELECT 序號 FROM 大宗郵件執據_郵寄種類 ORDER BY 序號"
        data.DataBind()
        data_dv = data.Select(New DataSourceSelectArguments)
        For i As Long = 0 To data_dv.Count - 1
            _郵寄種類 = data_dv(i)(0).ToString()
            _序號 = 0
            data.SelectCommand = "SELECT id FROM  大宗郵件執據 where 年='" & Me.DropDownList1.SelectedValue & " ' and 月='" & Me.DropDownList2.SelectedValue & " ' and 日='" & Me.DropDownList3.SelectedValue & "' and 批號='" & Me.DropDownList7.SelectedValue & "' and 郵寄種類='" & _郵寄種類 & "' ORDER BY 序號 "
            data.DataBind()
            data_dv1 = data.Select(New DataSourceSelectArguments)
            For j As Long = 0 To data_dv1.Count - 1
                _id = data_dv1(j)(0).ToString()
                _序號 = _序號 + 1
                data.UpdateCommand = "update 大宗郵件執據 set 序號='" & _序號 & "' where id='" & _id & "'"
                data.Update()
                data.DataBind()
            Next
        Next
        Me.SqlDataSource3.DataBind()
        Me.GridView2.DataBind()
    End Sub
    Protected Sub Button14_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button14.Click
        Me.TextBox23.Text = ""
        Me.SqlDataSource3.DataBind()
        Me.GridView2.DataBind()
    End Sub
    Protected Sub Button15_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button15.Click
        Me.SqlDataSource3.DataBind()
        Me.GridView2.DataBind()
    End Sub
End Class

