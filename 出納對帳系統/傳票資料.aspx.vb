Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel.XlDirection
Imports Microsoft.VisualBasic.Logging
Imports System.IO
Imports System.Drawing
Imports System.Diagnostics
Imports System.Data.Sql
Imports System.Data.SqlTypes
Imports System.Data.SqlClient
Imports System.Configuration
Imports System.Text.RegularExpressions
Partial Class 傳票資料
    Inherits System.Web.UI.Page
    Dim con_14 As String = System.Web.Configuration.WebConfigurationManager.ConnectionStrings("ApplicationServices").ConnectionString
    Dim data As New SqlDataSource
    Dim data_dv As Data.DataView
    Dim CheckBoxLock As Boolean
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        data.ConnectionString = con_14
        Me.Label1.Text = ""
        Me.Label2.Text = ""
        Me.GridView1.PageSize = Me.PageSize.Text
        If Not Page.IsPostBack Then
            DropDownList1_SelectedIndexChanged(sender, e)
            
            Me.Calendar1.Text = DateTime.Now.ToString("yyyy/MM/dd")
            Calendar1_OnTextChanged(sender, e)
            Me.Calendar2.Text = DateTime.Now.AddDays(1).ToString("yyyy/MM/dd")
            Me.Calendar3.Text = DateTime.Now.ToString("yyyy/MM/dd")
            
            Me.TextBox1.Text = (DateTime.Now.Year - 1911).ToString()
            Me.GridView1.PageIndex = Int32.MaxValue
            
            Update(sender, e)
        End If
        CheckBoxLock = Not Me.GridView1.columns(25).Visible
    End Sub
    Protected Sub DropDownList1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)'初始有用到
        '所有
            'Me.TextBox2.Text = ""
            'Me.TextBox3.Text = ""
            'Me.TextBox4.Text = ""
            'Me.TextBox5.Text = ""
            'Me.TextBox6.Text = ""
            'Me.TextBox7.Text = ""
            
            '全選
            Me.Button2.Visible = False
            Me.Button4.Visible = False
            Me.GridView1.Columns(1).Visible = False
            '摘要說明
            'Me.GridView1.Columns(3).Visible = False
            Me.GridView1.columns(4).Visible = False
        'Case "全"
            Me.GridView1.Columns(2).Visible = False
        'Case "土銀405全"
        'Case "土銀405匯款"
            Me.Panel2.Visible = False
            Me.Panel4.Visible = False
            Me.Panel5.Visible = False
            Me.GridView1.columns(13).Visible = True
            Me.GridView1.columns(15).Visible = False
            Me.GridView1.columns(16).Visible = False
            Me.GridView1.columns(17).Visible = False
            Me.GridView1.columns(18).Visible = False
            Me.GridView1.columns(19).Visible = False
            Me.GridView1.columns(20).Visible = False
            Me.GridView1.columns(21).Visible = False
            Me.GridView1.columns(22).Visible = False
            Me.GridView1.columns(23).Visible = False
            Me.GridView1.columns(24).Visible = False
            Me.GridView1.columns(25).Visible = False
            Me.GridView1.columns(26).Visible = False
            Me.GridView1.columns(27).Visible = False
            Me.GridView1.columns(28).Visible = False
            Me.GridView1.columns(29).Visible = False
        'Case "土銀405支票"
            Me.Panel3.Visible = False
            'Me.Panel5.Visible = False
            Me.Button6.Text = "下載"
            Me.GridView1.columns(10).HeaderText = "登錄序號"
            Me.GridView1.columns(12).HeaderText = "預付日期"
            Me.GridView1.columns(15).HeaderText = "分匯金額"
            Me.GridView1.columns(16).HeaderText = "分匯"
            Me.GridView1.columns(25).HeaderText = "下載"
            Me.GridView1.columns(11).Visible = True
            Me.GridView1.columns(9).Visible = False
            'Me.GridView1.columns(13).Visible = True
            'Me.GridView1.columns(21).Visible = False
            'Me.GridView1.columns(25).Visible = False
            'Me.GridView1.columns(27).Visible = False
            'Me.GridView1.columns(29).Visible = False
        'Case "土銀405收入"
            Me.GridView1.columns(12).Visible = True
            Me.GridView1.columns(14).Visible = True
        'Case "中國信託409全"
        'Case "中國信託409收入"
        'Case "中國信託409支出"
            'Me.Panel3.Visible = False
            'Me.Panel5.Visible = False
            'Me.Button6.Text = "下載"
            'Me.GridView1.columns(10).HeaderText = "登錄序號"
            'Me.GridView1.columns(12).HeaderText = "預付日期"
            'Me.GridView1.columns(15).HeaderText = "分匯金額"
            'Me.GridView1.columns(16).HeaderText = "分匯"
            'Me.GridView1.columns(25).HeaderText = "下載"
            'Me.GridView1.columns(11).Visible = True
            'Me.GridView1.columns(13).Visible = True
            'Me.GridView1.columns(21).Visible = False
            'Me.GridView1.columns(25).Visible = False
            'Me.GridView1.columns(27).Visible = False
            'Me.GridView1.columns(29).Visible = False
        
        Select Case Me.DropDownList1.SelectedValue
            Case "全"
            Case "土銀405全"
            Case "土銀405匯款"
                Me.Panel2.Visible = True
                Me.Panel4.Visible = True
                Me.Panel5.Visible = True
                Me.GridView1.columns(13).Visible = False
                Me.GridView1.columns(15).Visible = True
                Me.GridView1.columns(16).Visible = True
                Me.GridView1.columns(17).Visible = True
                '收款人名稱
                'Me.GridView1.columns(18).Visible = True
                Me.GridView1.columns(19).Visible = True
                Me.GridView1.columns(20).Visible = True
                Me.GridView1.columns(21).Visible = True
                Me.GridView1.columns(22).Visible = True
                Me.GridView1.columns(23).Visible = True
                Me.GridView1.columns(24).Visible = True
                Me.GridView1.columns(25).Visible = True
                Me.GridView1.columns(26).Visible = True
                Me.GridView1.columns(27).Visible = True
                Me.GridView1.columns(28).Visible = True
                Me.GridView1.columns(29).Visible = True
                Me.Button6.Enabled = False
                Me.Button7.Enabled = False
            Case "土銀405支票"
                Me.Panel3.Visible = True
                Me.Panel5.Visible = True
                Me.Button6.Text = "填入"
                Me.GridView1.columns(10).HeaderText = "支票編號"
                Me.GridView1.columns(12).HeaderText = "支票日期"
                Me.GridView1.columns(15).HeaderText = "分票金額"
                Me.GridView1.columns(16).HeaderText = "分票"
                Me.GridView1.columns(25).HeaderText = "填入"
                Me.GridView1.columns(9).Visible = True
                Me.GridView1.columns(11).Visible = False
                Me.GridView1.columns(13).Visible = False
                Me.GridView1.columns(15).Visible = False
                Me.GridView1.columns(16).Visible = True
                Me.GridView1.columns(25).Visible = True
                Me.GridView1.columns(27).Visible = True
                Me.GridView1.columns(29).Visible = True
                Me.Button6.Enabled = False
                Me.Button7.Enabled = False
            Case "土銀405收入"
                Me.GridView1.columns(12).Visible = False
                Me.GridView1.columns(14).Visible = False
            Case "中國信託409全"
            Case "中國信託409收入"
            Case "中國信託409支出"
                Me.Panel3.Visible = True
                Me.Panel5.Visible = True
                Me.Button6.Text = "填入"
                Me.GridView1.columns(10).HeaderText = "支票編號"
                Me.GridView1.columns(12).HeaderText = "支票日期"
                Me.GridView1.columns(15).HeaderText = "分票金額"
                Me.GridView1.columns(16).HeaderText = "分票"
                Me.GridView1.columns(25).HeaderText = "填入"
                Me.GridView1.columns(11).Visible = False
                Me.GridView1.columns(13).Visible = False
                Me.GridView1.columns(15).Visible = False
                Me.GridView1.columns(16).Visible = True
                Me.GridView1.columns(25).Visible = True
                Me.GridView1.columns(27).Visible = True
                Me.GridView1.columns(29).Visible = True
                Me.Button6.Enabled = False
                Me.Button7.Enabled = False
        End Select
        
        Me.GridView1.PageIndex = Int32.MaxValue
    End Sub
    
    Protected Sub GridView1_DataBound(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.DataBound
        If (Me.DropDownList1.SelectedValue="土銀405匯款")
            Dim 總金額2 As Long = 0
            For j = 0 To Me.GridView1.Rows.Count - 1
                If CType(Me.GridView1.Rows(j).FindControl("CheckBox1"), Checkbox).Checked = True 'AND (Checked)
                    總金額2 = 總金額2 + CLng(CType(Me.GridView1.Rows(j).FindControl("TextBox11"), TextBox).Text)
                End If
            Next
            For j = 0 To Me.GridView1.Rows.Count - 1
                    If CType(Me.GridView1.Rows(j).FindControl("CheckBox1"), Checkbox).Checked = True 'AND (Checked)
                        If(總金額2>0)
                            CType(Me.GridView1.Rows(j).FindControl("TextBox20"), TextBox).Text=總金額2
                        If (總金額2>50000000)
                            CType(Me.GridView1.Rows(j).FindControl("TextBox20"), TextBox).ForeColor=Color.Red
                            End If
                        Else
                            CType(Me.GridView1.Rows(j).FindControl("TextBox20"), TextBox).Text=""
                        End If
                    End If
                Next
        End If
    End Sub
    Protected Sub DropDownList2_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not Me.DropDownList2.SelectedValue.Contains("選擇")
            Update(sender, e)
            Me.TextBox2.Text = ""
            Me.TextBox3.Text = ""
            Me.TextBox4.Text = ""
            Me.TextBox5.Text = ""
            Me.名稱.Text = ""
            Me.登錄序號s.Text = ""
            Me.TextBox6.Text = ""
            Me.TextBox7.Text = ""
            Me.GridView1.AllowPaging = False
            Me.DropDownList1.Enabled = False
            Me.Panel2.Enabled = False
            Me.Panel3.Enabled = False
            Me.Panel4.Enabled = False
            Me.Panel7.Enabled = False
            Me.Button1.Enabled = False
            Me.Button2.Enabled = False
            Me.Button3.Enabled = False
            Me.Button4.Enabled = False
            Me.Button5.Enabled = False
            Select Case Me.DropDownList1.SelectedValue
                Case "土銀405匯款"
                    Me.Button6.Enabled = True
            End Select
            Me.Button7.Enabled = True
            Me.GridView1.DataBind()
            For i = 0 To Me.GridView1.Rows.Count - 1
                CType(Me.GridView1.Rows(i).FindControl("Label2"), Label).Enabled = False
                CType(Me.GridView1.Rows(i).FindControl("TextBox1"), TextBox).Enabled = False
                CType(Me.GridView1.Rows(i).FindControl("摘要說明"), TextBox).Enabled = False
                CType(Me.GridView1.Rows(i).FindControl("TextBox2"), TextBox).Enabled = False
                CType(Me.GridView1.Rows(i).FindControl("TextBox3"), TextBox).Enabled = False
                CType(Me.GridView1.Rows(i).FindControl("TextBox4"), TextBox).Enabled = False
                CType(Me.GridView1.Rows(i).FindControl("TextBox5"), TextBox).Enabled = False
                CType(Me.GridView1.Rows(i).FindControl("TextBox6"), TextBox).Enabled = False
                CType(Me.GridView1.Rows(i).FindControl("轉匯款"), Button).Enabled = False
                CType(Me.GridView1.Rows(i).FindControl("TextBox11"), TextBox).Enabled = False
                CType(Me.GridView1.Rows(i).FindControl("分匯金額"), TextBox).Enabled = False
                CType(Me.GridView1.Rows(i).FindControl("分匯"), Button).Enabled = False
                CType(Me.GridView1.Rows(i).FindControl("TextBox12"), TextBox).Enabled = False
                CType(Me.GridView1.Rows(i).FindControl("TextBox13"), TextBox).Enabled = False
                CType(Me.GridView1.Rows(i).FindControl("TextBox14"), TextBox).Enabled = False
                CType(Me.GridView1.Rows(i).FindControl("TextBox15"), TextBox).Enabled = False
                CType(Me.GridView1.Rows(i).FindControl("TextBox16"), TextBox).Enabled = False
                CType(Me.GridView1.Rows(i).FindControl("TextBox17"), TextBox).Enabled = False
                CType(Me.GridView1.Rows(i).FindControl("TextBox18"), TextBox).Enabled = False
                CType(Me.GridView1.Rows(i).FindControl("收款人EMAIL"), TextBox).Enabled = False
                CType(Me.GridView1.Rows(i).FindControl("CheckBox1"), CheckBox).Enabled = False
                CType(Me.GridView1.Rows(i).FindControl("TextBox21"), TextBox).Enabled = False
                CType(Me.GridView1.Rows(i).FindControl("清除"), Button).Enabled = False
                CType(Me.GridView1.Rows(i).FindControl("刪除"), Button).Enabled = False
            Next
        Else If Me.DropDownList2.SelectedValue.Contains("選擇")
            Me.GridView1.AllowPaging = True
            Me.DropDownList1.Enabled = True
            Me.Panel2.Enabled = True
            Me.Panel3.Enabled = True
            Me.Panel4.Enabled = True
            Me.Panel7.Enabled = True
            Me.Button1.Enabled = True
            Me.Button2.Enabled = True
            Me.Button3.Enabled = True
            Me.Button4.Enabled = True
            Me.Button5.Enabled = True
            Me.Button6.Enabled = False
            Me.Button7.Enabled = False
            Me.DropDownList2.Enabled = True
            Me.GridView1.DataBind()
            Me.GridView1.PageIndex = Int32.MaxValue
            For i = 0 To Me.GridView1.Rows.Count - 1
                CType(Me.GridView1.Rows(i).FindControl("Label2"), Label).Enabled = True
                CType(Me.GridView1.Rows(i).FindControl("TextBox1"), TextBox).Enabled = True
                CType(Me.GridView1.Rows(i).FindControl("摘要說明"), TextBox).Enabled = True
                CType(Me.GridView1.Rows(i).FindControl("TextBox2"), TextBox).Enabled = True
                CType(Me.GridView1.Rows(i).FindControl("TextBox3"), TextBox).Enabled = True
                CType(Me.GridView1.Rows(i).FindControl("TextBox4"), TextBox).Enabled = True
                CType(Me.GridView1.Rows(i).FindControl("TextBox5"), TextBox).Enabled = True
                CType(Me.GridView1.Rows(i).FindControl("TextBox6"), TextBox).Enabled = True
                CType(Me.GridView1.Rows(i).FindControl("轉匯款"), Button).Enabled = True
                CType(Me.GridView1.Rows(i).FindControl("TextBox11"), TextBox).Enabled = True
                CType(Me.GridView1.Rows(i).FindControl("分匯金額"), TextBox).Enabled = True
                CType(Me.GridView1.Rows(i).FindControl("分匯"), Button).Enabled = True
                CType(Me.GridView1.Rows(i).FindControl("TextBox12"), TextBox).Enabled = True
                CType(Me.GridView1.Rows(i).FindControl("TextBox13"), TextBox).Enabled = True
                CType(Me.GridView1.Rows(i).FindControl("TextBox14"), TextBox).Enabled = True
                CType(Me.GridView1.Rows(i).FindControl("TextBox15"), TextBox).Enabled = True
                CType(Me.GridView1.Rows(i).FindControl("TextBox16"), TextBox).Enabled = True
                CType(Me.GridView1.Rows(i).FindControl("TextBox17"), TextBox).Enabled = True
                CType(Me.GridView1.Rows(i).FindControl("TextBox18"), TextBox).Enabled = True
                CType(Me.GridView1.Rows(i).FindControl("收款人EMAIL"), TextBox).Enabled = True
                CType(Me.GridView1.Rows(i).FindControl("CheckBox1"), CheckBox).Enabled = True
                CType(Me.GridView1.Rows(i).FindControl("TextBox21"), TextBox).Enabled = True
                CType(Me.GridView1.Rows(i).FindControl("清除"), Button).Enabled = True
                CType(Me.GridView1.Rows(i).FindControl("刪除"), Button).Enabled = True
            Next
        End If
    End Sub
    Protected Sub Calendar1_OnTextChanged(ByVal sender As Object, ByVal e As System.EventArgs)'初始有用到
        Dim Calendar1 As String = Me.Calendar1.Text.Replace("/", "")
        data.SelectCommand = "SELECT ISNULL((SELECT MAX(CAST(TXT檔名 AS bigint))+1 FROM 傳票資料 WHERE LEFT(TXT檔名, 8) = '" & Calendar1 & "'), '"& Calendar1 & "001" &"')"
        data_dv = data.Select(New DataSourceSelectArguments)
        Me.TXT檔名.Text = data_dv(0)(0).ToString()
        data.SelectCommand = "SELECT ISNULL((SELECT MAX(CAST(登錄序號 AS bigint))+1 FROM 傳票資料 WHERE LEFT(登錄序號, 8) = '" & Calendar1 & "'), '"& Calendar1 & "0001" &"')"
        data_dv = data.Select(New DataSourceSelectArguments)
        Me.登錄序號.Text = data_dv(0)(0).ToString()
    End Sub
    
    Protected Sub CheckBox1_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        If CheckBoxLock = True
            Exit Sub
        End If
        
        Select Case Me.DropDownList1.SelectedValue
            Case "土銀405匯款"
                Dim i As Long = sender.NamingContainer.RowIndex
                Dim id As String = CType(Me.GridView1.Rows(i).FindControl("Label2"), Label).Text
                Dim Checked As Boolean = CType(Me.GridView1.Rows(i).FindControl("CheckBox1"), CheckBox).Checked
                
                Dim 總金額2 As Long = 0
                For j = 0 To Me.GridView1.Rows.Count - 1
                    If CType(Me.GridView1.Rows(j).FindControl("CheckBox1"), Checkbox).Checked = True 'AND (Checked)
                        總金額2 = 總金額2 + CLng(CType(Me.GridView1.Rows(j).FindControl("TextBox11"), TextBox).Text)
                    End If
                Next
                'TextBox20
                For j = 0 To Me.GridView1.Rows.Count - 1
                    If CType(Me.GridView1.Rows(j).FindControl("CheckBox1"), Checkbox).Checked = True 'AND (Checked)
                        If(總金額2>0)
                            CType(Me.GridView1.Rows(j).FindControl("TextBox20"), TextBox).Text=總金額2
                            If (總金額2>50000000)
                                CType(Me.GridView1.Rows(j).FindControl("TextBox20"), TextBox).ForeColor=Color.Red
                            End If
                        Else
                            CType(Me.GridView1.Rows(j).FindControl("TextBox20"), TextBox).Text=""
                        End If
                    ElseIf j=i AND Not(Checked) 
                        CType(Me.GridView1.Rows(i).FindControl("TextBox20"), TextBox).Text=""
                    End If
                Next
                Dim Calendar2 As String = Me.Calendar2.Text '民國
                Calendar2 = totaiwancalendar(Calendar2).Replace("/", "")
                Calendar2 = If(Checked, Calendar2, "")
                CType(Me.GridView1.Rows(i).FindControl("TextBox9"), TextBox).Text = Calendar2
                data.UpdateCommand = "UPDATE 傳票資料 SET 下載 = CAST('" & Checked & "' AS bit), 預付日期 = '" & Calendar2 & "' WHERE id = '" & id & "'"
                data.Update()
            Case "土銀405支票", "中國信託409支出"
                Dim i As Long = sender.NamingContainer.RowIndex
                Dim id As String = CType(Me.GridView1.Rows(i).FindControl("Label2"), Label).Text
                Dim Checked As Boolean = CType(Me.GridView1.Rows(i).FindControl("CheckBox1"), CheckBox).Checked
                
                data.UpdateCommand = "UPDATE 傳票資料 SET 下載 = CAST('" & Checked & "' AS bit) WHERE id = '" & id & "'"
                data.Update()
        End Select
    End Sub
    Protected Sub SelectAll(ByVal sender As Object, ByVal e As System.EventArgs)
        If CheckBoxLock = True Or Not Me.GridView1.Rows.Count > 0
            Exit Sub
        End IF
        Dim Checked As Boolean = False
        For i = 0 To Me.GridView1.Rows.Count - 1
            Checked = Checked Or Not CType(Me.GridView1.Rows(0).FindControl("CheckBox1"), CheckBox).Checked
        Next
        For i = 0 To Me.GridView1.Rows.Count - 1
            CType(Me.GridView1.Rows(i).FindControl("CheckBox1"), CheckBox).Checked = Checked
            Dim id As String = CType(Me.GridView1.Rows(i).FindControl("Label2"), Label).Text
            data.UpdateCommand = "UPDATE 傳票資料 SET 下載 = CAST('" & Checked & "' AS bit) WHERE id = '" & id & "'"
            data.Update()
        Next
    End Sub
    Protected Sub Insert(ByVal sender As Object, ByVal e As System.EventArgs)
        Me.SqlDataSource1.Insert()
        Me.DropDownList1.SelectedIndex = Me.DropDownList1.Items.IndexOf(Me.DropDownList1.Items.FindByValue("全部"))
        Me.GridView1.PageIndex = Int32.MaxValue
    End Sub
    Protected Sub Update(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim Result As Long = 1
        Select Case Me.DropDownList1.SelectedValue
            Case Else
                For i = 0 To Me.GridView1.Rows.Count - 1
                    Try
                        Dim Label2 As String = CType(Me.GridView1.Rows(i).FindControl("Label2"), Label).Text
                        Dim TextBox1 As String = CType(Me.GridView1.Rows(i).FindControl("TextBox1"), TextBox).Text
                        Dim 摘要說明 As String = CType(Me.GridView1.Rows(i).FindControl("摘要說明"), TextBox).Text
                        Dim TextBox2 As String = CType(Me.GridView1.Rows(i).FindControl("TextBox2"), TextBox).Text
                        Dim TextBox3 As String = CType(Me.GridView1.Rows(i).FindControl("TextBox3"), TextBox).Text
                        Dim TextBox4 As String = CType(Me.GridView1.Rows(i).FindControl("TextBox4"), TextBox).Text
                        Dim TextBox5 As String = CType(Me.GridView1.Rows(i).FindControl("TextBox5"), TextBox).Text
                        Dim TextBox6 As String = CType(Me.GridView1.Rows(i).FindControl("TextBox6"), TextBox).Text
                        Dim TextBox7 As String = CType(Me.GridView1.Rows(i).FindControl("TextBox7"), TextBox).Text
                        Dim TextBox8 As String = CType(Me.GridView1.Rows(i).FindControl("TextBox8"), TextBox).Text
                        Dim TextBox9 As String = CType(Me.GridView1.Rows(i).FindControl("TextBox9"), TextBox).Text
                        Dim TextBox10 As String = CType(Me.GridView1.Rows(i).FindControl("TextBox10"), TextBox).Text
                        Dim TextBox11 As String = CType(Me.GridView1.Rows(i).FindControl("TextBox11"), TextBox).Text
                        Dim TextBox12 As String = CType(Me.GridView1.Rows(i).FindControl("TextBox12"), TextBox).Text
                        Dim TextBox13 As String = CType(Me.GridView1.Rows(i).FindControl("TextBox13"), TextBox).Text
                        Dim TextBox14 As String = CType(Me.GridView1.Rows(i).FindControl("TextBox14"), TextBox).Text
                        Dim TextBox15 As String = CType(Me.GridView1.Rows(i).FindControl("TextBox15"), TextBox).Text
                        Dim TextBox16 As String = CType(Me.GridView1.Rows(i).FindControl("TextBox16"), TextBox).Text
                        Dim TextBox17 As String = CType(Me.GridView1.Rows(i).FindControl("TextBox17"), TextBox).Text
                        Dim TextBox18 As String = CType(Me.GridView1.Rows(i).FindControl("TextBox18"), TextBox).Text
                        Dim 收款人EMAIL As String = CType(Me.GridView1.Rows(i).FindControl("收款人EMAIL"), TextBox).Text
                        Dim TextBox19 As String = CType(Me.GridView1.Rows(i).FindControl("TextBox19"), TextBox).Text
                        Dim TextBox20 As String = CType(Me.GridView1.Rows(i).FindControl("TextBox20"), TextBox).Text
                        Dim TextBox21 As String = CType(Me.GridView1.Rows(i).FindControl("TextBox21"), TextBox).Text
                        
                        TextBox10 = TextBox10.Replace(",", "")
                        TextBox11 = TextBox11.Replace(",", "")
                        TextBox20 = TextBox20.Replace(",", "")
                        
                        TextBox1 = TextBox1.Replace("'", "")
                        摘要說明 = 摘要說明.Replace("'", "")
                        TextBox2 = TextBox2.Replace("'", "")
                        TextBox3 = TextBox3.Replace("'", "")
                        TextBox4 = TextBox4.Replace("'", "")
                        TextBox5 = TextBox5.Replace("'", "")
                        TextBox6 = TextBox6.Replace("'", "")
                        TextBox7 = TextBox7.Replace("'", "")
                        TextBox8 = TextBox8.Replace("'", "")
                        TextBox9 = TextBox9.Replace("'", "")
                        TextBox10 = TextBox10.Replace("'", "")
                        TextBox11 = TextBox11.Replace("'", "")
                        TextBox12 = TextBox12.Replace("'", "")
                        TextBox13 = TextBox13.Replace("'", "")
                        TextBox14 = TextBox14.Replace("'", "")
                        TextBox15 = TextBox15.Replace("'", "")
                        TextBox16 = TextBox16.Replace("'", "")
                        TextBox17 = TextBox17.Replace("'", "")
                        TextBox18 = TextBox18.Replace("'", "")
                        收款人EMAIL = 收款人EMAIL.Replace("'", "")
                        TextBox19 = TextBox19.Replace("'", "")
                        TextBox20 = TextBox20.Replace("'", "")
                        TextBox21 = TextBox21.Replace("'", "")
                        
                        TextBox1 = "N'" & TextBox1 & "'"
                        摘要說明 = "N'" & 摘要說明 & "'"
                        TextBox2 = "N'" & TextBox2 & "'"
                        TextBox3 = "N'" & TextBox3 & "'"
                        TextBox4 = "N'" & TextBox4 & "'"
                        TextBox5 = "N'" & TextBox5 & "'"
                        TextBox6 = "N'" & TextBox6 & "'"
                        TextBox7 = "N'" & TextBox7 & "'"
                        TextBox8 = "N'" & TextBox8 & "'"
                        TextBox9 = "N'" & TextBox9 & "'"
                        TextBox10 = "N'" & TextBox10 & "'"
                        TextBox11 = "N'" & TextBox11 & "'"
                        TextBox12 = "N'" & TextBox12 & "'"
                        TextBox13 = "N'" & TextBox13 & "'"
                        TextBox14 = "N'" & TextBox14 & "'"
                        TextBox15 = "N'" & TextBox15 & "'"
                        TextBox16 = "N'" & TextBox16 & "'"
                        TextBox17 = "N'" & TextBox17 & "'"
                        TextBox18 = "N'" & TextBox18 & "'"
                        收款人EMAIL = "N'" & 收款人EMAIL & "'"
                        TextBox19 = "N'" & TextBox19 & "'"
                        TextBox20 = "N'" & TextBox20 & "'"
                        TextBox21 = "N'" & TextBox21 & "'"
                        
                        TextBox1 = If(Me.GridView1.Columns(2).Visible, TextBox1, "NULL")
                        摘要說明 = If(Me.GridView1.columns(3).Visible, 摘要說明, "NULL")
                        TextBox2 = If(Me.GridView1.columns(4).Visible, TextBox2, "NULL")
                        TextBox3 = If(Me.GridView1.columns(5).Visible, TextBox3, "NULL")
                        TextBox4 = If(Me.GridView1.columns(6).Visible, TextBox4, "NULL")
                        TextBox5 = If(Me.GridView1.columns(7).Visible, TextBox5, "NULL")
                        TextBox6 = If(Me.GridView1.columns(8).Visible, TextBox6, "NULL")
                        TextBox7 = If(Me.GridView1.columns(10).Visible, TextBox7, "NULL")
                        TextBox8 = If(Me.GridView1.columns(11).Visible, TextBox8, "NULL")
                        TextBox9 = If(Me.GridView1.columns(12).Visible, TextBox9, "NULL")
                        TextBox10 = If(Me.GridView1.columns(13).Visible, TextBox10, "NULL")
                        TextBox11 = If(Me.GridView1.columns(14).Visible, TextBox11, "NULL")
                        TextBox12 = If(Me.GridView1.columns(17).Visible, TextBox12, "NULL")
                        TextBox13 = If(Me.GridView1.columns(18).Visible, TextBox13, "NULL")
                        TextBox14 = If(Me.GridView1.columns(19).Visible, TextBox14, "NULL")
                        TextBox15 = If(Me.GridView1.columns(20).Visible, TextBox15, "NULL")
                        TextBox16 = If(Me.GridView1.columns(21).Visible, TextBox16, "NULL")
                        TextBox17 = If(Me.GridView1.columns(22).Visible, TextBox17, "NULL")
                        TextBox18 = If(Me.GridView1.columns(23).Visible, TextBox18, "NULL")
                        收款人EMAIL = If(Me.GridView1.columns(24).Visible, 收款人EMAIL, "NULL")
                        TextBox19 = If(Me.GridView1.columns(26).Visible, TextBox19, "NULL")
                        TextBox20 = If(Me.GridView1.columns(27).Visible, TextBox20, "NULL")
                        TextBox21 = If(Me.GridView1.columns(28).Visible, TextBox21, "NULL")
                        
                        If Me.DropDownList1.SelectedValue = "土銀405匯款"
                            data.UpdateCommand = "UPDATE 傳票資料 SET " & _
                            "傳票送出納檔名 = NULLIF(ISNULL(" & TextBox1 & ", 傳票送出納檔名),''), " & _
                            "摘要說明 = NULLIF(ISNULL(" & 摘要說明 & ", 摘要說明),''), " & _
                            "年 = NULLIF(ISNULL(" & TextBox2 & ", 年),''), " & _
                            "開票日期 = NULLIF(ISNULL(" & TextBox3 & ", 開票日期),''), " & _
                            "傳票號碼 = NULLIF(ISNULL(" & TextBox4 & ", 傳票號碼),''), " & _
                            "之 = NULLIF(ISNULL(" & TextBox5 & ", 之),''), " & _
                            "名稱 = NULLIF(ISNULL(" & TextBox6 & ", 名稱),''), " & _
                            "登錄序號 = NULLIF(ISNULL(" & TextBox7 & ", 登錄序號),''), " & _
                            "登錄日期 = NULLIF(ISNULL(" & TextBox8 & ", 登錄日期),''), " & _
                            "預付日期 = NULLIF(ISNULL(" & TextBox9 & ", 預付日期),''), " & _
                            "收入金額 = NULLIF(ISNULL(" & TextBox10 & ", 收入金額),''), " & _
                            "支出金額 = NULLIF(ISNULL(" & TextBox11 & ", 支出金額),''), " & _
                            "收款人代碼 = NULLIF(ISNULL(" & TextBox12 & ", 收款人代碼),''), " & _
                            "收款人名稱 = NULLIF(ISNULL(" & TextBox13 & ", 收款人名稱),''), " & _
                            "匯入銀行名稱 = NULLIF(ISNULL(" & TextBox14 & ", 匯入銀行名稱),''), " & _
                            "匯入銀行代碼 = NULLIF(ISNULL(" & TextBox15 & ", 匯入銀行代碼),''), " & _
                            "匯入帳號 = ISNULL(NULLIF(ISNULL(" & TextBox16 & ", 匯入帳號),''),'-'), " & _
                            "收款人匯款戶名 = NULLIF(ISNULL(" & TextBox17 & ", 收款人匯款戶名),''), " & _
                            "收款人統編 = NULLIF(ISNULL(" & TextBox18 & ", 收款人統編),''), " & _
                            "收款人EMAIL = NULLIF(ISNULL(" & 收款人EMAIL & ", 收款人EMAIL),''), " & _
                            "TXT檔名 = NULLIF(ISNULL(" & TextBox19 & ", TXT檔名),''), " & _
                            "總金額 = NULLIF(ISNULL(" & TextBox20 & ", 總金額),''), " & _
                            "順序 = NULLIF(ISNULL(" & TextBox21 & ", 順序),'') " & _
                            "WHERE id = '" & Label2 & "'"
                            data.Update()
                        Else
                            data.UpdateCommand = "UPDATE 傳票資料 SET " & _
                            "傳票送出納檔名 = NULLIF(ISNULL(" & TextBox1 & ", 傳票送出納檔名),''), " & _
                            "摘要說明 = NULLIF(ISNULL(" & 摘要說明 & ", 摘要說明),''), " & _
                            "年 = NULLIF(ISNULL(" & TextBox2 & ", 年),''), " & _
                            "開票日期 = NULLIF(ISNULL(" & TextBox3 & ", 開票日期),''), " & _
                            "傳票號碼 = NULLIF(ISNULL(" & TextBox4 & ", 傳票號碼),''), " & _
                            "之 = NULLIF(ISNULL(" & TextBox5 & ", 之),''), " & _
                            "名稱 = NULLIF(ISNULL(" & TextBox6 & ", 名稱),''), " & _
                            "登錄序號 = NULLIF(ISNULL(" & TextBox7 & ", 登錄序號),''), " & _
                            "登錄日期 = NULLIF(ISNULL(" & TextBox8 & ", 登錄日期),''), " & _
                            "預付日期 = NULLIF(ISNULL(" & TextBox9 & ", 預付日期),''), " & _
                            "收入金額 = NULLIF(ISNULL(" & TextBox10 & ", 收入金額),''), " & _
                            "支出金額 = NULLIF(ISNULL(" & TextBox11 & ", 支出金額),''), " & _
                            "收款人代碼 = NULLIF(ISNULL(" & TextBox12 & ", 收款人代碼),''), " & _
                            "收款人名稱 = NULLIF(ISNULL(" & TextBox13 & ", 收款人名稱),''), " & _
                            "匯入銀行名稱 = NULLIF(ISNULL(" & TextBox14 & ", 匯入銀行名稱),''), " & _
                            "匯入銀行代碼 = NULLIF(ISNULL(" & TextBox15 & ", 匯入銀行代碼),''), " & _
                            "匯入帳號 = NULLIF(ISNULL(" & TextBox16 & ", 匯入帳號),''), " & _
                            "收款人匯款戶名 = NULLIF(ISNULL(" & TextBox17 & ", 收款人匯款戶名),''), " & _
                            "收款人統編 = NULLIF(ISNULL(" & TextBox18 & ", 收款人統編),''), " & _
                            "收款人EMAIL = NULLIF(ISNULL(" & 收款人EMAIL & ", 收款人EMAIL),''), " & _
                            "TXT檔名 = NULLIF(ISNULL(" & TextBox19 & ", TXT檔名),''), " & _
                            "總金額 = NULLIF(ISNULL(" & TextBox20 & ", 總金額),''), " & _
                            "順序 = NULLIF(ISNULL(" & TextBox21 & ", 順序),'') " & _
                            "WHERE id = '" & Label2 & "'"
                            data.Update()
                        End If
                    Catch
                        Result = 2
                    End Try
                Next
                
                '重算土銀405匯款總金額
                Try
                    data.UpdateCommand = _
                    "WITH CTE AS (" & _
                        "SELECT TABLE1.id, TABLE1.TXT檔名, 總金額, 計算結果 FROM 傳票資料 TABLE1 " & _
                        "INNER JOIN " & _
                        "(SELECT TXT檔名, SUM(支出金額) AS 計算結果 FROM 傳票資料 WHERE TXT檔名 != '' AND TXT檔名 IS NOT NULL GROUP BY TXT檔名) AS TABLE2 " & _
                        "ON TABLE1.TXT檔名 = TABLE2.TXT檔名" & _
                    ") UPDATE CTE SET 總金額 = 計算結果"
                    data.Update()
                Catch
                    Result = 2
                End Try
                '重算土銀405支票總金額
                Try
                    data.UpdateCommand = _
                    "WITH CTE AS (" & _
                        "SELECT TABLE1.id, TABLE1.登錄序號, 總金額, 計算結果 FROM 傳票資料 TABLE1 " & _
                        "INNER JOIN " & _
                        "(SELECT 登錄序號, SUM(支出金額) AS 計算結果 FROM 傳票資料 WHERE LEN(登錄序號) > 0 AND (TXT檔名 = '' OR TXT檔名 IS NULL) GROUP BY 登錄序號) AS TABLE2 " & _
                        "ON TABLE1.登錄序號 = TABLE2.登錄序號" & _
                    ") UPDATE CTE SET 總金額 = 計算結果"
                    data.Update()
                Catch
                    Result = 2
                End Try
                
                '自動帶入匯款資料
                data.UpdateCommand = _
                "UPDATE 傳票資料 SET " & _
                "傳票資料.名稱 = 收款人.收款人匯款戶名, " & _
                "傳票資料.收款人代碼 = 收款人.收款人代碼, " & _
                "傳票資料.匯入銀行代碼 = 收款人.匯入銀行代碼, " & _
                "傳票資料.匯入帳號 = 收款人.匯入帳號, " & _
                "傳票資料.收款人匯款戶名 = 收款人.收款人匯款戶名, " & _
                "傳票資料.收款人統編 = 收款人.收款人統編, " & _
                "傳票資料.收款人EMAIL = 收款人.收款人EMAIL " & _
                "FROM 傳票資料 INNER JOIN 收款人 " & _
                "ON (收款人.匯入帳號 LIKE N'%'+傳票資料.匯入帳號 OR 傳票資料.匯入帳號='-') " & _
                "AND (收款人.收款人匯款戶名 = 傳票資料.收款人匯款戶名) " & _
                "WHERE (LEN(傳票資料.匯入帳號) > 0) " & _
                "AND (LEN(傳票資料.登錄序號) = 0 OR 傳票資料.登錄序號 IS NULL OR 傳票資料.匯入帳號='-')"
                data.Update()
                
                '清除殘缺不全的匯款資料
                data.UpdateCommand = _
                "UPDATE 傳票資料 SET " & _
                "傳票資料.收款人代碼 = NULL, " & _
                "傳票資料.銀行名稱 = NULL, " & _
                "傳票資料.匯入銀行代碼 = NULL, " & _
                "傳票資料.匯入帳號 = NULL, " & _
                "傳票資料.收款人匯款戶名 = NULL, " & _
                "傳票資料.收款人統編 = NULL, " & _
                "傳票資料.收款人EMAIL = NULL " & _
                "FROM 傳票資料 " & _
                "WHERE (傳票資料.匯入帳號 IS NULL OR 傳票資料.匯入帳號 = '') " & _
                data.Update()
                
                '清除已經下載的下載
                data.UpdateCommand = _
                "UPDATE 傳票資料 SET " & _
                "下載 = NULL " & _
                "FROM 傳票資料 " & _
                "WHERE (LEN(傳票資料.登錄序號) > 0) "
                data.Update()
                
                If Result = 1
                    Me.GridView1.DataBind()
                End If
        End Select
        
        Try
            If sender.ID = "Button3"
                If Result = 1
                    Me.Label1.Text = "成功"
                    Me.Label2.Text = ""
                Else If Result = 2
                    Me.Label1.Text = ""
                    Me.Label2.Text = "失敗"
                End If
            End If
        Catch
        End Try
    End Sub
    Protected Sub Preview(ByVal sender As Object, ByVal e As System.EventArgs)
        If Me.Button5.Text = "預覽"
            If Me.DropDownList1.SelectedValue = "土銀405匯款"
                If Me.Calendar1.Text = "" Or Me.Calendar2.Text = "" Or Me.登錄序號.Text = "" Or Me.TXT檔名.Text = ""
                    ScriptManager.RegisterStartupScript(Me, Page.GetType, "Script1", "alert('請輸入登錄日期、預付日期、登錄序號、TXT檔名。');", True)
                    Exit Sub
                End If
                For j = 0 To Me.GridView1.Rows.Count - 1
                    CType(Me.GridView1.Rows(j).FindControl("TextBox20"), TextBox).Text=""'重設總金額
                Next   
            Else If Me.DropDownList1.SelectedValue = "土銀405支票" Or Me.DropDownList1.SelectedValue = "中國信託409支出"
                If Me.Calendar3.Text = "" Or Me.支票編號.Text = ""
                    ScriptManager.RegisterStartupScript(Me, Page.GetType, "Script1", "alert('請輸入支票日期、支票編號。');", True)
                    Exit Sub
                End If
            End If
            Update(sender, e)
            Label3.Visible=False
            Me.GridView1.AllowPaging = False
            Me.DropDownList1.Enabled = False
            Me.Panel2.Enabled = False
            Me.Panel3.Enabled = False
            Me.Panel4.Enabled = False
            Me.Panel7.Enabled = False
            Me.Button1.Enabled = False
            Me.Button2.Enabled = False
            Me.Button3.Enabled = False
            Me.Button4.Enabled = False
            Me.Button5.Text = "取消"
            Me.Button6.Enabled = True
            Me.DropDownList2.Enabled = False
            Me.DropDownList2.SelectedIndex = Me.DropDownList2.Items.IndexOf(Me.DropDownList2.Items.FindByValue("選擇TXT檔名"))
            Me.GridView1.DataBind()
            For i = 0 To Me.GridView1.Rows.Count - 1
                CType(Me.GridView1.Rows(i).FindControl("Label2"), Label).Enabled = False
                CType(Me.GridView1.Rows(i).FindControl("TextBox1"), TextBox).Enabled = False
                CType(Me.GridView1.Rows(i).FindControl("摘要說明"), TextBox).Enabled = False
                CType(Me.GridView1.Rows(i).FindControl("TextBox2"), TextBox).Enabled = False
                CType(Me.GridView1.Rows(i).FindControl("TextBox3"), TextBox).Enabled = False
                CType(Me.GridView1.Rows(i).FindControl("TextBox4"), TextBox).Enabled = False
                CType(Me.GridView1.Rows(i).FindControl("TextBox5"), TextBox).Enabled = False
                CType(Me.GridView1.Rows(i).FindControl("TextBox6"), TextBox).Enabled = False
                CType(Me.GridView1.Rows(i).FindControl("轉匯款"), Button).Enabled = False
                CType(Me.GridView1.Rows(i).FindControl("TextBox11"), TextBox).Enabled = False
                CType(Me.GridView1.Rows(i).FindControl("分匯金額"), TextBox).Enabled = False
                CType(Me.GridView1.Rows(i).FindControl("分匯"), Button).Enabled = False
                CType(Me.GridView1.Rows(i).FindControl("TextBox12"), TextBox).Enabled = False
                CType(Me.GridView1.Rows(i).FindControl("TextBox13"), TextBox).Enabled = False
                CType(Me.GridView1.Rows(i).FindControl("TextBox14"), TextBox).Enabled = False
                CType(Me.GridView1.Rows(i).FindControl("TextBox15"), TextBox).Enabled = False
                CType(Me.GridView1.Rows(i).FindControl("TextBox16"), TextBox).Enabled = False
                CType(Me.GridView1.Rows(i).FindControl("TextBox17"), TextBox).Enabled = False
                CType(Me.GridView1.Rows(i).FindControl("TextBox18"), TextBox).Enabled = False
                CType(Me.GridView1.Rows(i).FindControl("收款人EMAIL"), TextBox).Enabled = False
                CType(Me.GridView1.Rows(i).FindControl("CheckBox1"), CheckBox).Enabled = False
                CType(Me.GridView1.Rows(i).FindControl("TextBox21"), TextBox).Enabled = False
                CType(Me.GridView1.Rows(i).FindControl("清除"), Button).Enabled = False
                CType(Me.GridView1.Rows(i).FindControl("刪除"), Button).Enabled = False
            Next
            If Me.DropDownList1.SelectedValue = "土銀405匯款"
                Dim Calendar1 As String = Me.Calendar1.Text '民國
                Dim Calendar2 As String = Me.Calendar2.Text '民國
                Calendar1 = totaiwancalendar(Calendar1).Replace("/", "")
                Calendar2 = totaiwancalendar(Calendar2).Replace("/", "")
                Dim TXT檔名 As String = Me.TXT檔名.Text
                Dim 登錄序號 As Long = Me.登錄序號.Text
                Dim 總金額 As Long = 0
                For i = 0 To Me.GridView1.Rows.Count - 1
                    If CType(Me.GridView1.Rows(i).FindControl("CheckBox1"), Checkbox).Checked = True
                        總金額 = 總金額 + CLng(CType(Me.GridView1.Rows(i).FindControl("TextBox11"), TextBox).Text)
                    End If
                Next
                IF 總金額>50000000
                    Me.Button6.Enabled = False
                End If
                For i = 0 To Me.GridView1.Rows.Count - 1
                    If CType(Me.GridView1.Rows(i).FindControl("CheckBox1"), Checkbox).Checked = True
                        CType(Me.GridView1.Rows(i).FindControl("TextBox7"), TextBox).Text = 登錄序號
                        CType(Me.GridView1.Rows(i).FindControl("TextBox8"), TextBox).Text = Calendar1
                        If CType(Me.GridView1.Rows(i).FindControl("TextBox9"), TextBox).Text = ""
                            CType(Me.GridView1.Rows(i).FindControl("TextBox9"), TextBox).Text = Calendar2
                        End If
                        CType(Me.GridView1.Rows(i).FindControl("TextBox19"), TextBox).Text = TXT檔名
                        CType(Me.GridView1.Rows(i).FindControl("TextBox20"), TextBox).Text = 總金額.ToString("N0")
                        登錄序號 = 登錄序號 + 1
                    End If
                Next
            Else If Me.DropDownList1.SelectedValue = "土銀405支票" Or Me.DropDownList1.SelectedValue = "中國信託409支出"
                Dim Calendar3 As String = Me.Calendar3.Text '民國
                Calendar3 = totaiwancalendar(Calendar3).Replace("/", "")
                Dim 支票編號 As String = Me.支票編號.Text
                Dim 總金額 As Long = 0
                For i = 0 To Me.GridView1.Rows.Count - 1
                    If CType(Me.GridView1.Rows(i).FindControl("CheckBox1"), Checkbox).Checked = True
                        總金額 = 總金額 + CLng(CType(Me.GridView1.Rows(i).FindControl("TextBox11"), TextBox).Text)
                    End If
                Next
                For i = 0 To Me.GridView1.Rows.Count - 1
                    If CType(Me.GridView1.Rows(i).FindControl("CheckBox1"), Checkbox).Checked = True
                        CType(Me.GridView1.Rows(i).FindControl("TextBox7"), TextBox).Text = 支票編號
                        CType(Me.GridView1.Rows(i).FindControl("TextBox9"), TextBox).Text = Calendar3
                        CType(Me.GridView1.Rows(i).FindControl("TextBox20"), TextBox).Text = 總金額.ToString("N0")
                    End If
                Next
            End If
        Else If Me.Button5.Text = "取消"
            Label3.Visible=True
            Me.GridView1.AllowPaging = True
            Me.DropDownList1.Enabled = True
            Me.Panel2.Enabled = True
            Me.Panel3.Enabled = True
            Me.Panel4.Enabled = True
            Me.Panel7.Enabled = True
            Me.Button1.Enabled = True
            Me.Button2.Enabled = True
            Me.Button3.Enabled = True
            Me.Button4.Enabled = True
            Me.Button5.Text = "預覽"
            Me.Button6.Enabled = False
            Me.DropDownList2.Enabled = True
            Me.GridView1.DataBind()
            Me.GridView1.PageIndex = Int32.MaxValue
            For i = 0 To Me.GridView1.Rows.Count - 1
                CType(Me.GridView1.Rows(i).FindControl("Label2"), Label).Enabled = True
                CType(Me.GridView1.Rows(i).FindControl("TextBox1"), TextBox).Enabled = True
                CType(Me.GridView1.Rows(i).FindControl("摘要說明"), TextBox).Enabled = True
                CType(Me.GridView1.Rows(i).FindControl("TextBox2"), TextBox).Enabled = True
                CType(Me.GridView1.Rows(i).FindControl("TextBox3"), TextBox).Enabled = True
                CType(Me.GridView1.Rows(i).FindControl("TextBox4"), TextBox).Enabled = True
                CType(Me.GridView1.Rows(i).FindControl("TextBox5"), TextBox).Enabled = True
                CType(Me.GridView1.Rows(i).FindControl("TextBox6"), TextBox).Enabled = True
                CType(Me.GridView1.Rows(i).FindControl("轉匯款"), Button).Enabled = True
                CType(Me.GridView1.Rows(i).FindControl("TextBox11"), TextBox).Enabled = True
                CType(Me.GridView1.Rows(i).FindControl("分匯金額"), TextBox).Enabled = True
                CType(Me.GridView1.Rows(i).FindControl("分匯"), Button).Enabled = True
                CType(Me.GridView1.Rows(i).FindControl("TextBox12"), TextBox).Enabled = True
                CType(Me.GridView1.Rows(i).FindControl("TextBox13"), TextBox).Enabled = True
                CType(Me.GridView1.Rows(i).FindControl("TextBox14"), TextBox).Enabled = True
                CType(Me.GridView1.Rows(i).FindControl("TextBox15"), TextBox).Enabled = True
                CType(Me.GridView1.Rows(i).FindControl("TextBox16"), TextBox).Enabled = True
                CType(Me.GridView1.Rows(i).FindControl("TextBox17"), TextBox).Enabled = True
                CType(Me.GridView1.Rows(i).FindControl("TextBox18"), TextBox).Enabled = True
                CType(Me.GridView1.Rows(i).FindControl("收款人EMAIL"), TextBox).Enabled = True
                CType(Me.GridView1.Rows(i).FindControl("CheckBox1"), CheckBox).Enabled = True
                CType(Me.GridView1.Rows(i).FindControl("TextBox21"), TextBox).Enabled = True
                CType(Me.GridView1.Rows(i).FindControl("清除"), Button).Enabled = True
                CType(Me.GridView1.Rows(i).FindControl("刪除"), Button).Enabled = True
            Next
        End If
    End Sub
    Protected Sub Download(ByVal sender As Object, ByVal e As System.EventArgs)
        Select Case Me.DropDownList1.SelectedValue
            Case "土銀405匯款"
                Dim MyGUID As String = Guid.NewGuid().ToString("N")
                Dim MyTXT As String = MapPath(".\Excel\Temp\") & MyGUID & ".txt"
                File.Copy(MapPath(".\Excel\付款檔.txt"), MyTXT)
                
                Dim content As String = ""
                For i = 0 To Me.GridView1.Rows.Count - 1
                    Dim Label2 As String = CType(Me.GridView1.Rows(i).FindControl("Label2"), Label).Text
                    Dim TextBox1 As String = CType(Me.GridView1.Rows(i).FindControl("TextBox1"), TextBox).Text
                    Dim 摘要說明 As String = CType(Me.GridView1.Rows(i).FindControl("摘要說明"), TextBox).Text
                    Dim TextBox2 As String = CType(Me.GridView1.Rows(i).FindControl("TextBox2"), TextBox).Text
                    Dim TextBox3 As String = CType(Me.GridView1.Rows(i).FindControl("TextBox3"), TextBox).Text
                    Dim TextBox4 As String = CType(Me.GridView1.Rows(i).FindControl("TextBox4"), TextBox).Text
                    Dim TextBox5 As String = CType(Me.GridView1.Rows(i).FindControl("TextBox5"), TextBox).Text
                    Dim TextBox6 As String = CType(Me.GridView1.Rows(i).FindControl("TextBox6"), TextBox).Text
                    Dim TextBox7 As String = CType(Me.GridView1.Rows(i).FindControl("TextBox7"), TextBox).Text
                    Dim TextBox8 As String = CType(Me.GridView1.Rows(i).FindControl("TextBox8"), TextBox).Text
                    Dim TextBox9 As String = CType(Me.GridView1.Rows(i).FindControl("TextBox9"), TextBox).Text
                    Dim TextBox10 As String = CType(Me.GridView1.Rows(i).FindControl("TextBox10"), TextBox).Text
                    Dim TextBox11 As String = CType(Me.GridView1.Rows(i).FindControl("TextBox11"), TextBox).Text
                    Dim TextBox12 As String = CType(Me.GridView1.Rows(i).FindControl("TextBox12"), TextBox).Text
                    Dim TextBox13 As String = CType(Me.GridView1.Rows(i).FindControl("TextBox13"), TextBox).Text
                    Dim TextBox14 As String = CType(Me.GridView1.Rows(i).FindControl("TextBox14"), TextBox).Text
                    Dim TextBox15 As String = CType(Me.GridView1.Rows(i).FindControl("TextBox15"), TextBox).Text
                    Dim TextBox16 As String = CType(Me.GridView1.Rows(i).FindControl("TextBox16"), TextBox).Text
                    Dim TextBox17 As String = CType(Me.GridView1.Rows(i).FindControl("TextBox17"), TextBox).Text
                    Dim TextBox18 As String = CType(Me.GridView1.Rows(i).FindControl("TextBox18"), TextBox).Text
                    Dim 收款人EMAIL As String = CType(Me.GridView1.Rows(i).FindControl("收款人EMAIL"), TextBox).Text
                    Dim TextBox19 As String = CType(Me.GridView1.Rows(i).FindControl("TextBox19"), TextBox).Text
                    Dim TextBox20 As String = CType(Me.GridView1.Rows(i).FindControl("TextBox20"), TextBox).Text
                    
                    '新的一行
                    content = content & Environment.NewLine
                    '登錄序號
                    content = content & TextBox7.PadRight(13)
                    '登錄日期
                    content = content & TextBox8.PadRight(13)
                    '預付日期
                    content = content & TextBox9.PadRight(13)
                    '資料來源
                    content = content & ("人工新增").PadRight(13)
                    '帳款金額
                    content = content & TextBox11.PadLeft(13)
                    '分隔
                    content = content & " "
                    '付款行代號
                    content = content & ("0050773").PadRight(13)
                    '付款帳號
                    content = content & ("077056000014").PadRight(17)
                    '收款人代碼
                    content = content & TextBox12.PadRight(13)
                    '收款人名稱
                    content = content & TextBox17.PadRight(35, "　").Replace("　", "  ")
                    '分隔
                    content = content & " "
                    '收款行代號
                    content = content & TextBox15.PadRight(13)
                    '收款人帳號
                    content = content & TextBox16.PadRight(17)
                    '收款人戶名
                    content = content & TextBox17.PadRight(35, "　").Replace("　", "  ")
                    '分隔
                    content = content & " "
                    '收款人統編
                    content = content & TextBox18.PadRight(13)
                    '其他
                    content = content & 收款人EMAIL.PadRight(214)
                Next
                Using sw As New StreamWriter(MyTXT, True, Encoding.UTF8)
                    sw.WriteLine(content)
                End Using
                Dim TXT檔名 As String = If(Me.GridView1.Rows.Count > 0, CType(Me.GridView1.Rows(0).FindControl("TextBox19"), TextBox).Text, "空白付款檔")
                ScriptManager.RegisterStartupScript(Me, Page.GetType, "Script1", "window.open('下載.aspx?file=" & MyGUID & ".txt&downloadfilename=" & TXT檔名 & ".txt" & "','_blank');", True)
                
                If Me.Button5.Text = "取消"
                    For i = 0 To Me.GridView1.Rows.Count - 1
                        '同步現金備查簿
                        Dim 年 As String = Me.TextBox1.Text
                        Dim 傳票號碼 As String = CType(Me.GridView1.Rows(i).FindControl("TextBox4"), TextBox).Text
                        Dim 登錄序號 As String = CType(Me.GridView1.Rows(i).FindControl("TextBox7"), TextBox).Text
                        Dim 預付日期 As String = CType(Me.GridView1.Rows(i).FindControl("TextBox9"), TextBox).Text
                        Dim 付款日 As String = If(預付日期.Length = 7, 預付日期.Substring(0, 3) & "." & 預付日期.Substring(3, 2) & "." & 預付日期.Substring(5, 2), "")
                        'TODO:使用CROSS APPLY取代Nested REPLACE()、排序支票編號(用換行符號分割字串，經排序後合併字串)
                        data.UpdateCommand = _
                            "UPDATE 現金備查簿 SET " & _
                            "支票編號 = " & _
                            "ISNULL(" & _
                            "NULLIF(" & _
                            "REPLACE(" & _
                            "REPLACE(" & _
                            "REPLACE(支票編號, CHAR(13) + CHAR(10) + N'" & 登錄序號 & "', '')" & _
                            ", N'" & 登錄序號 & "' + CHAR(13) + CHAR(10), '')" & _
                            ", N'" & 登錄序號 & "', '')" & _
                            ", '')" & _
                            " + CHAR(13) + CHAR(10), '')" & _
                            " + N'" & 登錄序號 & "', " & _
                            "付款日 = " & _
                            "ISNULL(" & _
                            "NULLIF(" & _
                            "REPLACE(" & _
                            "REPLACE(" & _
                            "REPLACE(付款日, CHAR(13) + CHAR(10) + N'" & 付款日 & "', '')" & _
                            ", N'" & 付款日 & "' + CHAR(13) + CHAR(10), '')" & _
                            ", N'" & 付款日 & "', '')" & _
                            ", '')" & _
                            " + CHAR(13) + CHAR(10), '')" & _
                            " + N'" & 付款日 & "' " & _
                            "WHERE 年 = N'" & 年 & "' AND 傳票號碼 = N'" & 傳票號碼 & "'"
                        data.Update()
                    Next
                    Update(sender, e)
                    Calendar1_OnTextChanged(sender, e)
                    
                    '清除已經下載的下載
                    data.UpdateCommand = _
                    "UPDATE 傳票資料 SET " & _
                    "下載 = NULL " & _
                    "FROM 傳票資料 " & _
                    "WHERE (LEN(傳票資料.登錄序號) > 0) "
                    data.Update()
                    
                    Preview(sender, e)
                    Me.DropDownList2.DataBind()
                    Me.DropDownList2.SelectedIndex = Me.DropDownList2.Items.IndexOf(Me.DropDownList2.Items.FindByValue(TXT檔名))
                    DropDownList2_SelectedIndexChanged(sender, e)
                End If
            Case "土銀405支票", "中國信託409支出"
                If Me.Button5.Text = "取消"
                    For i = 0 To Me.GridView1.Rows.Count - 1
                        '同步現金備查簿
                        Dim 年 As String = Me.TextBox1.Text
                        Dim 傳票號碼 As String = CType(Me.GridView1.Rows(i).FindControl("TextBox4"), TextBox).Text
                        Dim 支票編號 As String = Me.支票編號.Text
                        Dim 預付日期 As String = CType(Me.GridView1.Rows(i).FindControl("TextBox9"), TextBox).Text
                        Dim 付款日 As String = If(預付日期.Length = 7, 預付日期.Substring(0, 3) & "." & 預付日期.Substring(3, 2) & "." & 預付日期.Substring(5, 2), "")
                        'TODO:使用CROSS APPLY取代Nested REPLACE()、排序支票編號(用換行符號分割字串，經排序後合併字串)
                        data.UpdateCommand = _
                            "UPDATE 現金備查簿 SET " & _
                            "支票編號 = " & _
                            "N'" & 支票編號 & "' + " & _
                            "ISNULL(CHAR(13) + CHAR(10) + " & _
                            "NULLIF(" & _
                            "REPLACE(" & _
                            "REPLACE(" & _
                            "REPLACE(支票編號, CHAR(13) + CHAR(10) + N'" & 支票編號 & "', '')" & _
                            ", N'" & 支票編號 & "' + CHAR(13) + CHAR(10), '')" & _
                            ", N'" & 支票編號 & "', '')" & _
                            ", '')" & _
                            ", ''), " & _
                            "付款日 = " & _
                            "N'" & 付款日 & "' + " & _
                            "ISNULL(CHAR(13) + CHAR(10) + " & _
                            "NULLIF(" & _
                            "REPLACE(" & _
                            "REPLACE(" & _
                            "REPLACE(付款日, CHAR(13) + CHAR(10) + N'" & 付款日 & "', '')" & _
                            ", N'" & 付款日 & "' + CHAR(13) + CHAR(10), '')" & _
                            ", N'" & 付款日 & "', '')" & _
                            ", '')" & _
                            ", '')" & _
                            "WHERE 年 = N'" & 年 & "' AND 傳票號碼 = N'" & 傳票號碼 & "'"
                        data.Update()
                    Next
                    
                    Update(sender, e)
                    
                    '清除已經下載的下載
                    data.UpdateCommand = _
                    "UPDATE 傳票資料 SET " & _
                    "下載 = NULL " & _
                    "FROM 傳票資料 " & _
                    "WHERE (LEN(傳票資料.登錄序號) > 0) "
                    data.Update()
                    
                    Preview(sender, e)
                    Me.DropDownList2.DataBind()
                    Me.DropDownList2.SelectedIndex = Me.DropDownList2.Items.IndexOf(Me.DropDownList2.Items.FindByValue(Me.支票編號.Text))
                    DropDownList2_SelectedIndexChanged(sender, e)
                End If
        End Select
    End Sub
    '刪除TXT檔
    Protected Sub Delete(ByVal sender As Object, ByVal e As System.EventArgs)
        For i = 0 To Me.GridView1.Rows.Count - 1
            '同步現金備查簿
            Dim 年 As String = Me.TextBox1.Text
            Dim 傳票號碼 As String = CType(Me.GridView1.Rows(i).FindControl("TextBox4"), TextBox).Text
            Dim 舊登錄序號 As String = CType(Me.GridView1.Rows(i).FindControl("TextBox7"), TextBox).Text
            Dim 預付日期 As String = CType(Me.GridView1.Rows(i).FindControl("TextBox9"), TextBox).Text
            Dim 付款日 As String = If(預付日期.Length = 7, 預付日期.Substring(0, 3) & "." & 預付日期.Substring(3, 2) & "." & 預付日期.Substring(5, 2), "")
            data.UpdateCommand = _
            "IF NOT N'" & 舊登錄序號 & "' = '' AND (SELECT COUNT(*) FROM 傳票資料 WHERE 年 = N'" & 年 & "' AND 傳票號碼 = N'" & 傳票號碼 & "' AND 登錄序號 = N'" & 舊登錄序號 & "') = 1 " & _
                "BEGIN " & _
                    "UPDATE 現金備查簿 SET " & _
                    "支票編號 = " & _
                    "REPLACE(" & _
                    "REPLACE(" & _
                    "REPLACE(支票編號, CHAR(13) + CHAR(10) + N'" & 舊登錄序號 & "', '')" & _
                    ", N'" & 舊登錄序號 & "' + CHAR(13) + CHAR(10), '')" & _
                    ", N'" & 舊登錄序號 & "', '') " & _
                    "WHERE 年 = N'" & 年 & "' AND 傳票號碼 = N'" & 傳票號碼 & "' " & _
                "END"
            data.Update()
            data.UpdateCommand = _
            "IF NOT N'" & 舊登錄序號 & "' = '' AND (SELECT COUNT(*) FROM 傳票資料 WHERE 年 = N'" & 年 & "' AND 傳票號碼 = N'" & 傳票號碼 & "' AND 預付日期 = REPLACE(N'" & 付款日 & "', '.', '')) = 1 " & _
                "BEGIN " & _
                    "UPDATE 現金備查簿 SET " & _
                    "付款日 = " & _
                    "REPLACE(" & _
                    "REPLACE(" & _
                    "REPLACE(付款日, CHAR(13) + CHAR(10) + N'" & 付款日 & "', '')" & _
                    ", N'" & 付款日 & "' + CHAR(13) + CHAR(10), '')" & _
                    ", N'" & 付款日 & "', '') " & _
                    "WHERE 年 = N'" & 年 & "' AND 傳票號碼 = N'" & 傳票號碼 & "' " & _
                "END"
            data.Update()
            
            CType(Me.GridView1.Rows(i).FindControl("TextBox7"), TextBox).Text = ""
            CType(Me.GridView1.Rows(i).FindControl("TextBox8"), TextBox).Text = ""
            CType(Me.GridView1.Rows(i).FindControl("TextBox9"), TextBox).Text = ""
            CType(Me.GridView1.Rows(i).FindControl("TextBox19"), TextBox).Text = ""
            CType(Me.GridView1.Rows(i).FindControl("TextBox20"), TextBox).Text = ""
        Next
        Update(sender, e)
        Me.DropDownList2.DataBind()
        DropDownList2_SelectedIndexChanged(sender, e)
    End Sub
    Protected Sub GridView1_RowCommand(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles GridView1.RowCommand
        Select Case e.CommandName
            Case "Clean", "CustomDelete"
                Dim i As Long = e.CommandSource.NamingContainer.RowIndex
                '同步現金備查簿
                Dim 年 As String = Me.TextBox1.Text
                Dim 傳票號碼 As String = CType(Me.GridView1.Rows(i).FindControl("TextBox4"), TextBox).Text
                Dim 舊登錄序號 As String = CType(Me.GridView1.Rows(i).FindControl("TextBox7"), TextBox).Text
                data.UpdateCommand = _
                "IF NOT N'" & 舊登錄序號 & "' = '' AND (SELECT COUNT(*) FROM 傳票資料 WHERE 年 = N'" & 年 & "' AND 傳票號碼 = N'" & 傳票號碼 & "' AND 登錄序號 = N'" & 舊登錄序號 & "') = 1 " & _
                    "BEGIN " & _
                        "UPDATE 現金備查簿 SET " & _
                        "支票編號 = " & _
                        "REPLACE(" & _
                        "REPLACE(" & _
                        "REPLACE(支票編號, CHAR(13) + CHAR(10) + N'" & 舊登錄序號 & "', '')" & _
                        ", N'" & 舊登錄序號 & "' + CHAR(13) + CHAR(10), '')" & _
                        ", N'" & 舊登錄序號 & "', '') " & _
                        "WHERE 年 = N'" & 年 & "' AND 傳票號碼 = N'" & 傳票號碼 & "' " & _
                    "END"
                data.Update()
                Dim 預付日期 As String = CType(Me.GridView1.Rows(i).FindControl("TextBox9"), TextBox).Text
                Dim 付款日 As String = If(預付日期.Length = 7, 預付日期.Substring(0, 3) & "." & 預付日期.Substring(3, 2) & "." & 預付日期.Substring(5, 2), "")
                data.UpdateCommand = _
                "IF NOT N'" & 付款日 & "' = '' AND (SELECT COUNT(*) FROM 傳票資料 WHERE 年 = N'" & 年 & "' AND 傳票號碼 = N'" & 傳票號碼 & "' AND 預付日期 = N'" & 預付日期 & "') = 1 " & _
                    "BEGIN " & _
                        "UPDATE 現金備查簿 SET " & _
                        "付款日 = " & _
                        "REPLACE(" & _
                        "REPLACE(" & _
                        "REPLACE(付款日, CHAR(13) + CHAR(10) + N'" & 付款日 & "', '')" & _
                        ", N'" & 付款日 & "' + CHAR(13) + CHAR(10), '')" & _
                        ", N'" & 付款日 & "', '') " & _
                        "WHERE 年 = N'" & 年 & "' AND 傳票號碼 = N'" & 傳票號碼 & "' " & _
                    "END"
                data.Update()
                Select Case e.CommandName
                    Case "Clean"
                        Select Case Me.DropDownList1.SelectedValue
                            Case "土銀405匯款"
                                CType(Me.GridView1.Rows(i).FindControl("TextBox7"), TextBox).Text = ""
                                CType(Me.GridView1.Rows(i).FindControl("TextBox8"), TextBox).Text = ""
                                CType(Me.GridView1.Rows(i).FindControl("TextBox9"), TextBox).Text = ""
                                CType(Me.GridView1.Rows(i).FindControl("TextBox19"), TextBox).Text = ""
                                CType(Me.GridView1.Rows(i).FindControl("TextBox20"), TextBox).Text = ""
                            Case "土銀405支票", "中國信託409支出"
                                CType(Me.GridView1.Rows(i).FindControl("TextBox7"), TextBox).Text = ""
                                CType(Me.GridView1.Rows(i).FindControl("TextBox9"), TextBox).Text = ""
                                CType(Me.GridView1.Rows(i).FindControl("TextBox20"), TextBox).Text = ""
                        End Select
                        Update(sender, e)
                    Case "CustomDelete"
                        Me.GridView1.DeleteRow(i)
                        Update(sender, e)
                End Select
            Case "Separate"
                Dim i As Long = e.CommandSource.NamingContainer.RowIndex
                Dim id As String = CType(Me.GridView1.Rows(i).FindControl("Label2"), Label).Text
                Dim 支出金額 As String = CType(Me.GridView1.Rows(i).FindControl("TextBox11"), TextBox).Text
                'Dim 分匯金額 As String = CType(Me.GridView1.Rows(i).FindControl("分匯金額"), TextBox).Text
                Dim 分匯金額 As String = "0"
                支出金額 = 支出金額.Replace(",","")
                '分匯金額 = 分匯金額.Replace(",","")
                'If 分匯金額 = ""
                '    ScriptManager.RegisterStartupScript(Me, Page.GetType, "s1", "setTimeout(function(){alert('請輸入分匯金額');}, 50);", True)
                '    Exit Sub
                'End If
                CType(Me.GridView1.Rows(i).FindControl("TextBox11"), TextBox).Text = (CLng("0" & 支出金額) - CLng("0" & 分匯金額)).ToString()
                Select Case Me.DropDownList1.SelectedValue
                    Case "土銀405匯款"
                        data.InsertCommand = _
                            "INSERT INTO 傳票資料(傳票送出納檔名, 摘要說明, 年, 開票日期, 傳票號碼, " & _
                            "之, 名稱, 登錄序號, 登錄日期, 預付日期, 收入金額, 支出金額, 收款人代碼, " & _
                            "收款人名稱, 匯入銀行名稱, 匯入銀行代碼, 匯入帳號, 收款人匯款戶名, 收款人統編, " & _
                            "收款人EMAIL, TXT檔名, 總金額, 順序) " & _
                            "SELECT 傳票送出納檔名, 摘要說明, 年, 開票日期, 傳票號碼, " & _
                            "之, NULL, NULL, NULL, NULL, NULL, NULLIF('" & 分匯金額 & "', '') , NULL, " & _
                            "NULL, '中華郵政股份有限公司', NULL, '-', NULL, NULL, " & _
                            "NULL, NULL, NULL, NULL " & _
                            "FROM 傳票資料 WHERE id = N'" & id & "'"
                    Case "土銀405支票", "中國信託409支出"
                        data.InsertCommand = _
                            "INSERT INTO 傳票資料(傳票送出納檔名, 摘要說明, 年, 開票日期, 傳票號碼, " & _
                            "之, 名稱, 登錄序號, 登錄日期, 預付日期, 收入金額, 支出金額, 收款人代碼, " & _
                            "收款人名稱, 匯入銀行名稱, 匯入銀行代碼, 匯入帳號, 收款人匯款戶名, 收款人統編, " & _
                            "收款人EMAIL, TXT檔名, 總金額, 順序) " & _
                            "SELECT 傳票送出納檔名, 摘要說明, 年, 開票日期, 傳票號碼, " & _
                            "之, NULL, NULL, NULL, NULL, NULL, NULLIF('" & 分匯金額 & "', '') , NULL, " & _
                            "NULL, NULL, NULL, NULL, NULL, NULL, " & _
                            "NULL, NULL, NULL, NULL " & _
                            "FROM 傳票資料 WHERE id = N'" & id & "'"
                End Select
                data.Insert()
                Update(sender, e)
            Case "Change"
                Update(sender, e)
                Dim i As Long = e.CommandSource.NamingContainer.RowIndex
                Dim id As String = CType(Me.GridView1.Rows(i).FindControl("Label2"), Label).Text
                data.UpdateCommand = _
                "UPDATE 傳票資料 SET " & _
                "傳票資料.摘要說明 = '網路匯款', " & _
                "傳票資料.匯入銀行名稱 = '中華郵政股份有限公司', " & _
                "傳票資料.匯入帳號 = '-' " & _
                "WHERE 傳票資料.id = '" & id & "'"
                data.Update()
                Me.GridView1.DataBind()
        End Select
    End Sub
    <System.Web.Script.Services.ScriptMethod(), System.Web.Services.WebMethod()>
    Public Shared Function GetMyList(ByVal prefixText As String, ByVal count As Integer)
        Dim con_14 As String = System.Web.Configuration.WebConfigurationManager.ConnectionStrings("ApplicationServices").ConnectionString
        Dim data As New SqlDataSource
        Dim data_dv As Data.DataView
        Dim MyList As New List(Of String)
        data.ConnectionString = con_14
        'data.SelectCommand = "SELECT TOP " & count & " 收款人匯款戶名 FROM 收款人 WHERE 收款人匯款戶名 LIKE N'" & prefixText & "%' ORDER BY 收款人匯款戶名"
        data.SelectCommand = _
            "WITH CTE AS " & _
            "(SELECT DISTINCT TOP " & count & " 收款人匯款戶名 FROM 收款人 WHERE 收款人匯款戶名 LIKE N'%" & prefixText & "%') " & _
            "SELECT * FROM CTE " & _
            "ORDER BY " & _
            "CASE WHEN (收款人匯款戶名 LIKE N'" & prefixText & "%') THEN 0 ELSE 1 END, " & _
            "收款人匯款戶名"
        data_dv = data.Select(New DataSourceSelectArguments)
        For i As Long = 0 To data_dv.Count() - 1
            MyList.Add(data_dv(i)(0).ToString())
        Next
        Return MyList
    End Function
End Class