Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel.XlDirection
Imports Microsoft.Office.Interop.Excel.XlPageBreak
Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.VisualBasic.Logging
Imports System.IO
Imports System.Math
Imports System.Diagnostics
Imports System.Data.Sql
Imports System.Data.SqlTypes
Imports System.Data.SqlClient
Imports System.Configuration
Imports System.Text.RegularExpressions
Imports System.Web.UI.WebControls
Imports System.Drawing
Partial Class 維修紀錄作業
    Inherits System.Web.UI.Page
    Dim con_14 As String = System.Web.Configuration.WebConfigurationManager.ConnectionStrings("ApplicationServices").ConnectionString
    Dim data As New SqlDataSource
    Dim data_dv As Data.DataView
    Dim 編輯權限 As Boolean
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        data.ConnectionString = con_14
        If Not Page.IsPostBack Then
            Me.GridView1.PageIndex = Int32.MaxValue
        Else
            Me.Label1.Text = ""
            Me.Label2.Text = ""
        End If
        data.ConnectionString = con_14
        If not(Session("水_id") is Nothing)
            id.text=Session("水_id")
            品項.text=Session("水_品項")
        End If
        If Session("水_編輯權限")=true
            編輯權限=true
        End If
        PermissionOn()
    End Sub
    Protected Sub Insert(ByVal sender As Object, ByVal e As System.EventArgs)
        'Dim p餘額 As Integer
        'Update(sender, e)
        Dim id As String = me.id.Text
        If id<>""
            'Dim RowIndex As Long = e.CommandSource.NamingContainer.RowIndex
            Dim 餘額 As String
            Dim insert1 as string
            data.InsertCommand = _
            "INSERT INTO 維修紀錄表 " & _
            "( id_水電) " & _
            "VALUES " & _
            "(" & me.id.Text & ")"
            data.Insert()
            Me.GridView1.PageIndex = Int32.MaxValue
            Me.GridView1.DataBind()
        Else
            Label2.text="id為空值"
        End If
    End Sub
    Protected Sub Update(ByVal sender As Object, ByVal e As System.EventArgs)
        'TODO:轉換成整數會失敗
        For i = 0 To Me.GridView1.Rows.Count - 1
            Dim id As String = CType(Me.GridView1.Rows(i).FindControl("id"), TextBox).Text
            Dim 維修日期 As String = nothing
            If CType(Me.GridView1.Rows(i).FindControl("維修日期"), TextBox).text<>""
                維修日期 = CType(Me.GridView1.Rows(i).FindControl("維修日期"), TextBox).text
                維修日期 = taiwancalendarto(維修日期)
            End If
            Dim 金額 As String = Mid(CType(Me.GridView1.Rows(i).FindControl("金額"), TextBox).Text,4)
            金額=金額.Replace(",", "")
            Dim 備註 As String = CType(Me.GridView1.Rows(i).FindControl("備註"), TextBox).Text
            Dim Update1 as string = "UPDATE 維修紀錄表 SET " & _
            "維修日期 = IIF(ISDATE(TRIM(N'" & 維修日期 & "'))=1,TRIM(N'" & 維修日期 & "'),NULL), " & _
            "金額 = REPLACE(ISNULL(NULLIF('" & 金額 & "', ''),'0'), ',', ''), " & _
            "備註 = NULLIF(N'" & 備註 & "', '') " & _
            "WHERE id = '" & id & "'"
            data.UpdateCommand = Update1
            data.Update()
        Next
        Me.GridView1.DataBind()
    End Sub
    Protected Sub test(ByVal sender As Object, ByVal e As System.EventArgs)
    End Sub
    Protected Sub GridView1_RowCommand(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles GridView1.RowCommand
        If e.CommandName = "上傳資料"'上傳資料按紐，不能傳TXT檔，不能有_的字號出現
            '狀況:新增、新增前有檔案、新增檔案與原檔案相同、別的ID_水電(新增、新增前有檔案、新增檔案與原檔案相同)
            Dim i As Long = e.CommandSource.NamingContainer.RowIndex
            Dim id As String = CType(Me.GridView1.Rows(i).FindControl("id"), TextBox).Text
            Dim id_水電 As String = CType(Me.GridView1.Rows(i).FindControl("id_水電"), TextBox).Text
            If Not Me.FileUpload1.HasFile
                Label2.text="如要上傳資料，請先按'選擇檔案'"
                Exit Sub
            End If
            For Each PostedFile As HttpPostedFile In Me.FileUpload1.PostedFiles
                Dim MyGUID As String = Guid.NewGuid().ToString("N")
                Dim Myfiles As String = MapPath(".\data\Temp\") & MyGUID
                PostedFile.SaveAs(Myfiles)
                Try
                    File.Copy(Myfiles, MapPath(".\data\簽案資料\") & PostedFile.FileName, False)
                Catch
                End Try
                '無用程式碼保留至5/23
                '建立該資料列所有資料
                ' data.ConnectionString = con_14
                ' data.SelectCommand = "SELECT * FROM 維修紀錄表 Where id =" & id
                ' data_dv = data.Select(New DataSourceSelectArguments)
                '將原有檔案資料載入
                Dim 簽案資料 As String = CType(Me.GridView1.Rows(i).FindControl("簽案資料"), HyperLink).Text
                ' Dim data_dv2 As Data.DataView
                ' '被取代檔案有資料，則取得該資料表有多少筆相同參考的資料
                ' If 簽案資料<>""
                '     data.SelectCommand = "SELECT count(簽案資料) As 資料數 FROM 維修紀錄表 Where 簽案資料 = N'" & 簽案資料 & "' Group by 簽案資料"
                '     data_dv2 = data.Select(New DataSourceSelectArguments)
                ' End If
                ' '判斷等等資料是否有更新，無更新，會有資料
                ' Dim data_dv3 As Data.DataView
                ' data.SelectCommand = "SELECT * FROM 維修紀錄表 WHERE 簽案資料 = N'" & PostedFile.FileName & "'And id_水電=" & id_水電
                ' data_dv3 = data.Select(New DataSourceSelectArguments)
                ' '更新條件，該ID_水電資料中不可有相同檔案
                data.UpdateCommand = _
                    "IF NOT EXISTS(SELECT * FROM 維修紀錄表 WHERE 簽案資料 = N'" & PostedFile.FileName & "'And id_水電=" & id_水電 & ") "  & _
                    "BEGIN " & _
                    "UPDATE 維修紀錄表 SET " & _
                    "簽案資料 = NULLIF(N'" & PostedFile.FileName & "', '')" & _
                    "WHERE id = '" & id & "'" & _
                    "END"
                data.Update()
                System.IO.File.Delete(Myfiles)
                CheckFileEmpty(簽案資料)
                '如何判斷資料未更新時，不去刪除原有檔案?
                '有更新時，原資料只剩一筆，且為被取代之資料(簽案資料<>"")，即刪除
                '未更新時(data_dv3.count<1)，別動!
                '原有東西，但不為被取代資料，且有更新?EX:有1水電、用EXE取代PNG，水電狀況?不影響
                ' If (簽案資料<>"" And data_dv3.count<1)
                '     '當原本有資料剩1，且資料有更新，再做動作
                '     If data_dv2.count>0
                '         If data_dv2(0)("資料數").ToString()=1'除非只剩一筆，否則不刪除文件檔案，因為有別的資料參照
                '             For i2 As Long = 0 To data_dv.Count() - 1
                '                 System.IO.File.Delete(MapPath(".\data\簽案資料\") & 簽案資料)
                '             Next
                '         End If
                '     End If
                ' End If
                '5/6直接在結尾判定被取待資料是否沒有任何參照 
                ' data.ConnectionString = con_14
                ' data.SelectCommand = "SELECT * FROM 維修紀錄表 Where 簽案資料 =N'" & 簽案資料 &"'"
                ' data_dv = data.Select(New DataSourceSelectArguments)
                ' If data_dv.count<1
                '     System.IO.File.Delete(MapPath(".\data\簽案資料\") & 簽案資料)
                ' End If
                '無用程式碼保留至5/23
            Next
            Me.GridView1.DataBind()
        ElseIf e.CommandName = "上傳照片"'上傳資料按紐
            Dim i As Long = e.CommandSource.NamingContainer.RowIndex
            Dim id As String = CType(Me.GridView1.Rows(i).FindControl("id"), TextBox).Text
            If Not Me.FileUpload1.HasFile
                Label2.text="如要上傳資料，請先按'選擇檔案'"
                Exit Sub
            End If
            Dim regex as new Regex(".jpg$|.png$|.jpeg$")
            If regex.IsMatch(Me.FileUpload1.FileName)
                For Each PostedFile As HttpPostedFile In Me.FileUpload1.PostedFiles
                    Dim fs As System.IO.Stream = FileUpload1.PostedFile.InputStream
                    Dim base64String As String = FileImageToBase64(fs)
                    data.UpdateCommand = "UPDATE 維修紀錄表 SET 照片 = NULLIF(N'" & base64String & "', '')" & _
                    "WHERE id=" & id & ""
                    data.Update()
                Next
            Else
                Label2.text="只能上傳圖片檔"
            End If
            Me.GridView1.DataBind()
        ElseIf e.CommandName = "照片圖"
            Dim i As Long = e.CommandSource.NamingContainer.RowIndex
            IF CType(Me.GridView1.Rows(i).FindControl("照片"), ImageButton).ImageUrl<>nothing
                Session("水_照片")=CType(Me.GridView1.Rows(i).FindControl("照片"), ImageButton).ImageUrl
                Response.Redirect("照片.aspx")
            End IF
        ElseIf e.CommandName = "刪除"
            Update(sender, e)
            Dim i As Long = e.CommandSource.NamingContainer.RowIndex
            Dim id As String = CType(Me.GridView1.Rows(i).FindControl("id"), TextBox).Text
            Dim 簽案資料 As String = CType(Me.GridView1.Rows(i).FindControl("簽案資料"), HyperLink).Text
            '無用程式碼保留至5/23
            ' data.ConnectionString = con_14
            ' data.SelectCommand = "SELECT * FROM 維修紀錄表 Where id =" & id
            ' data_dv = data.Select(New DataSourceSelectArguments)
            ' Dim data_dv2 As Data.DataView
            ' data.SelectCommand = "SELECT count(簽案資料) As 資料數 FROM 維修紀錄表 Where 簽案資料 = '" & 簽案資料 & "' Group by 簽案資料"
            ' data_dv2 = data.Select(New DataSourceSelectArguments)
            ' If (data_dv2.count>0)
            '     If data_dv2(0)("資料數").ToString()=1'除非只剩一筆，否則不刪除文件檔案，因為有別的資料參照
            '         For i2 As Long = 0 To data_dv.Count() - 1
            '             System.IO.File.Delete(MapPath(".\data\簽案資料\") & 簽案資料)
            '         Next
            '     End If
            ' End If
            data.DeleteCommand = "DELETE FROM 維修紀錄表 WHERE id=" & id
            data.Delete()
            CheckFileEmpty(簽案資料)
            ' data.ConnectionString = con_14
            ' data.SelectCommand = "SELECT * FROM 維修紀錄表 Where 簽案資料 =N'" & 簽案資料 &"'"
            ' data_dv = data.Select(New DataSourceSelectArguments)
            ' If data_dv.count<1
            '     System.IO.File.Delete(MapPath(".\data\簽案資料\") & 簽案資料)
            ' End If
            '無用程式碼保留至5/23
            Me.GridView1.DataBind()
        End If
    End Sub
    Protected Sub GridView1_DataBound(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.DataBound
        PermissionOn()
    End Sub
    Protected Sub PermissionOn()
        If 編輯權限=true
            Me.GridView1.columns(6).Visible = true
            Me.GridView1.columns(10).Visible = true
            新增.Visible = true
            存檔.Visible = true
            For i = 0 To Me.GridView1.Rows.Count - 1
                CType(Me.GridView1.Rows(i).FindControl("維修日期"), TextBox).Enabled=true
                Me.FileUpload1.Visible=true
                CType(Me.GridView1.Rows(i).FindControl("上傳資料"), Button).Visible=true
                CType(Me.GridView1.Rows(i).FindControl("上傳照片"), Button).Visible=true
                CType(Me.GridView1.Rows(i).FindControl("金額"), TextBox).Enabled=true
                CType(Me.GridView1.Rows(i).FindControl("備註"), TextBox).Enabled=true
                CType(Me.GridView1.Rows(i).FindControl("刪除"), Button).Visible=true
            Next
        End If
    End Sub
    Public Function FileImageToBase64(ByVal fs As System.IO.Stream) As String'將圖檔轉Base64，以儲存至資料庫
        Try
            If (fs.Length>0)
                Dim br As New System.IO.BinaryReader(fs)
                Dim bytes As Byte() = br.ReadBytes(CType(fs.Length, Integer))
                Dim base64String As String = Convert.ToBase64String(bytes, 0, bytes.Length)
                base64String="data:image/png;base64," & base64String
                Return base64String
            Else
                Return Nothing
            End If
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function ImageToBase64(ByVal image As Drawing.Image, ByVal format As System.Drawing.Imaging.ImageFormat) As String
        Using ms As New MemoryStream()
            image.Save(ms, format)
            Dim imageBytes As Byte() = ms.ToArray()' Convert byte[] to Base64 String
            Dim base64String As String = Convert.ToBase64String(imageBytes)
            base64String="data:image/png;base64," & base64String
            Return base64String
        End Using
    End Function
    '如果該檔案無資料參照，刪除
    Protected Sub CheckFileEmpty(ByVal File As String)'無資料時，被拒絕存取
        '查詢是否有資料
        If File<>""
            data.ConnectionString = con_14
            data.SelectCommand = "SELECT * FROM 維修紀錄表 Where 簽案資料 =N'" & File &"'"
            data_dv = data.Select(New DataSourceSelectArguments)
            If data_dv.count<1
                System.IO.File.Delete(MapPath(".\data\簽案資料\") & File)
            End If
        End If
    End Sub
End Class