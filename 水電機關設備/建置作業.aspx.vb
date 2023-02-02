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
Imports System.Drawing.Imaging
Imports System.Data.OleDb
Imports System.Drawing.Drawing2D
Partial Class 建置作業
    Inherits System.Web.UI.Page
    Dim con_14 As String = System.Web.Configuration.WebConfigurationManager.ConnectionStrings("ApplicationServices").ConnectionString
    Dim data As New SqlDataSource
    Dim data_dv As Data.DataView
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        data.ConnectionString = con_14
        If Not Page.IsPostBack Then
            Me.GridView1.PageIndex = Int32.MaxValue
            If Session("水_Uid")="3855"
                測試.Visible=True
            End If
        Else
            Me.Label1.Text = ""
            Me.Label2.Text = ""
        End If
    End Sub
    Protected Sub Insert(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim RowIndex As Integer=15
        For i = 1 To 15
            Dim 品項 As String = "19"
            Dim 餘額 As String
            Dim insert1 as string
            data.InsertCommand = _
            "INSERT INTO 水電機關設備資料表 " & _
            "( _頁, _列, Id_品項) " & _
            "VALUES " & _
            "(" & (Me.GridView1.PageCount + 1).ToString() & ", " & i & ", '" & 品項 & "')"
            data.Insert()
        Next
        Me.GridView1.PageIndex = Int32.MaxValue
        Me.GridView1.DataBind()
    End Sub
    Protected Sub Update(ByVal sender As Object, ByVal e As System.EventArgs)
        'TODO:轉換成整數會失敗
        For i = 0 To Me.GridView1.Rows.Count - 1
            Dim id As String = CType(Me.GridView1.Rows(i).FindControl("id"), TextBox).Text
            Dim 品項 As String = CType(Me.GridView1.Rows(i).FindControl("品項"),DropDownList).Text
            Dim 建置日期 As String = nothing
            If CType(Me.GridView1.Rows(i).FindControl("建置日期"), TextBox).text<>""
                建置日期 = CType(Me.GridView1.Rows(i).FindControl("建置日期"), TextBox).text
                建置日期 = taiwancalendarto(建置日期)
            End If
            Dim 型號 As String = CType(Me.GridView1.Rows(i).FindControl("型號"), TextBox).Text
            Dim 存置地點 As String = CType(Me.GridView1.Rows(i).FindControl("存置地點"), TextBox).Text
            Dim 維護單位 As String = CType(Me.GridView1.Rows(i).FindControl("維護單位"), TextBox).Text
            Dim 備註 As String = CType(Me.GridView1.Rows(i).FindControl("備註"), TextBox).Text
            Dim Update1 as string ="UPDATE 水電機關設備資料表 SET " & _
            "Id_品項 = NULLIF(N'" & 品項 & "', ''), " & _
            "建置日期 = IIF(ISDATE(TRIM(N'" & 建置日期 & "'))=1,TRIM(N'" & 建置日期 & "'),NULL), " & _
            "型號 = NULLIF(N'" & 型號 & "', ''), " & _
            "存置地點 = NULLIF(N'" & 存置地點 & "', ''), " & _
            "維護單位 = NULLIF(N'" & 維護單位 & "', ''), " & _
            "備註 = NULLIF(N'" & 備註 & "', '') " & _
            "WHERE id = '" & id & "'"
            data.UpdateCommand = Update1
            data.Update()
            '將設備編號產生並排序
            data.UpdateCommand = "WITH CTE AS (Select *,Row_Number() OVER(Partition by Id_品項 order by ID ) AS '序號' From 水電機關設備資料表)" & _
            "UPDATE CTE SET 設備編號 = 序號 Where Id_品項 IS NOT NULL"
            data.Update()
        Next
        Me.GridView1.DataBind()
    End Sub
    Protected Sub Delete(ByVal sender As Object, ByVal e As System.EventArgs)
        'Update(sender, e)
        Me.GridView1.PageIndex = Int32.MaxValue
        Me.GridView1.DataBind()
        Dim delete1 as string=""
        For i = 0 To Me.GridView1.Rows.Count - 1
            Dim id As String = CType(Me.GridView1.Rows(i).FindControl("id"), TextBox).Text
            delete1="DELETE FROM 水電機關設備資料表 " & _
            "WHERE id = '" & id & "'"
            data.deleteCommand =delete1
            data.delete()
        Next
        Me.GridView1.DataBind()
    End Sub
    Protected Sub Test(ByVal sender As Object, ByVal e As System.EventArgs)'改呈現圖長1920,1080 縮圖成長100，高100以下 20220525
        ' data.ConnectionString = con_14
        ' data.SelectCommAnd = "SELECT id,原始照片 FROM 水電機關設備資料表 Where 原始照片 Is Not Null"'全輸出，不輸出無編號
        ' data_dv = data.Select(New DataSourceSelectArguments)
        ' Dim i As Int32 = 0
        ' For i = 0 To data_dv.Count - 1
        '     ' Dim id As String = CType(Me.GridView1.Rows(i).FindControl("id"), TextBox).Text
        '     ' data.SelectCommAnd = "SELECT id,原始照片 FROM 水電機關設備資料表 Where 原始照片 Is Not Null And id=" & id'全輸出，不輸出無編號
        '     ' data_dv = data.Select(New DataSourceSelectArguments)
        '         Dim id As String = data_dv(i)("Id").ToString()
        '         Dim 照片S As String = data_dv(i)("原始照片").ToString()
        '         Dim b64 As string=ImageCompression(照片S,"B")
        '        'Dim b64s As string=ImageCompression(照片S,"S")
        '         ' Label1.Text = Label1.Text & "id = " & id & "原:" & 照片.Width & "," & 照片.Height & "壓縮:" & widthB & "," & HeightB & "縮圖" & widthS & "," & HeightS & "<BR>"
        '     ' data.UpdateCommand = "UPDATE 水電機關設備資料表 SET 照片縮圖 = NULLIF(N'" & b64s & "', ''),壓縮照片 = NULLIF(N'" & b64 & "', '')" & _
        '     '     "WHERE id=" & id & ""
        '     data.UpdateCommand = "UPDATE 水電機關設備資料表 SET 壓縮照片 = NULLIF(N'" & b64 & "', '')" & _
        '          "WHERE id=" & id & ""
        '     data.Update()
        ' Next
        ' Me.GridView1.DataBind()
    End Sub
    Protected Sub GridView1_RowCommand(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles GridView1.RowCommand
        If e.CommandName = "照片圖"'點擊後放大、遇到困難，相對虛擬路徑，遊覽器傳送的字串太小，問要放大還是速度快，目前為速度快
            '原始照片如更改名稱會動到資料庫，故不修正
            Dim i As Long = e.CommandSource.NamingContainer.RowIndex
            Dim id As String = CType(Me.GridView1.Rows(i).FindControl("id"), TextBox).Text
            data.ConnectionString = con_14
            data.SelectCommAnd = "SELECT id,壓縮照片 FROM 水電機關設備資料表 Where id="& id'全輸出，不輸出無編號
            data_dv = data.Select(New DataSourceSelectArguments)
            Session("水_照片")=data_dv(0)("壓縮照片").ToString
            Response.Redirect("照片.aspx")
        ElseIf e.CommandName = "上傳資料"'將照片轉至資料庫儲存
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
                    Dim b64 As String = ImageCompression(base64String,"B")
                    Dim b64s As string=ImageCompression(base64String,"S")
                    If base64String=b64
                        base64String=""
                    End If
                        data.UpdateCommand = "UPDATE 水電機關設備資料表 SET 原始照片 = NULLIF(N'" & base64String & "', '')," & _
                        "壓縮照片 = NULLIF(N'" & b64 & "', '')," & _
                        "照片縮圖 = NULLIF(N'" & b64s & "', '')" & _
                        "WHERE id =" & id & ""
                        data.Update()
                Next
            Else
                Label2.text="只能上傳圖片檔"
            End If
            Me.GridView1.DataBind()
        ElseIf e.CommandName = "維護紀錄"
            Dim i As Long = e.CommandSource.NamingContainer.RowIndex
            Dim id As String = CType(Me.GridView1.Rows(i).FindControl("id"), TextBox).Text
            Dim 品項 As String = CType(Me.GridView1.Rows(i).FindControl("品項"), DropDownList).Text
            Session("水_id")=id
            Session("水_品項")=品項
            Session("水_編輯權限")=true
            Response.Redirect("維修紀錄作業.aspx")
        ElseIf e.CommandName = "刪除"
            '連同維護紀錄作業的相關資料一併刪除
            Update(sender, e)
            Dim i As Long = e.CommandSource.NamingContainer.RowIndex
            Dim id As String = CType(Me.GridView1.Rows(i).FindControl("id"), TextBox).Text
            data.UpdateCommand = "Update 水電機關設備資料表 Set ID_品項='19',建置日期=NULL,型號=NULL,存置地點=NULL,維護單位=NULL,備註=NULL,原始照片=NULL,壓縮照片=NULL,照片縮圖=NULL WHERE id=" & id
            data.Update()
            '
            data.ConnectionString = con_14
            data.SelectCommand = "SELECT 簽案資料 FROM 維修紀錄表 Where id_水電=" & id
            data_dv = data.Select(New DataSourceSelectArguments)
            Dim data_dv2 As Data.DataView
            data.DeleteCommand = "DELETE FROM 維修紀錄表 WHERE id_水電=" & id
            data.Delete()
            for j As Long = 0 To data_dv.Count() - 1
                data.SelectCommand = "SELECT * FROM 維修紀錄表 Where 簽案資料 =N'" & data_dv(j)("簽案資料").ToString() &"'"
                data_dv2 = data.Select(New DataSourceSelectArguments)
                If data_dv2.count<1
                    System.IO.File.Delete(MapPath(".\data\簽案資料\") & data_dv(j)("簽案資料").ToString())
                End If
            Next
            Me.GridView1.DataBind()
            Update(sender,e)
        End If
    End Sub
    public Function ThumbnailCallback() As Boolean
        return true
    End Function
    Public Function ImageCompression(ByVal base64String As String , ByVal Size1 As String) As String
        Try
            If (base64String.Length>0)
                Dim 照片 As Drawing.Image=BASE64_TO_IMG(base64String)
                Dim iScale As int32 = 3
                '//取得圖片大小
                Dim widthB As int32 = 照片.Width
                Dim HeightB As int32 = 照片.Height
                Dim widthSize As int32
                Dim HeightSize As int32 
                Dim 比例 As Double 
                DIm 功能 As Boolean = False
                If Size1="B"
                    widthSize=1280
                    HeightSize=720
                ElseIf Size1="S"
                    widthSize=200
                    HeightSize=200
                End if
                ' While widthB>widthSize OR HeightB>HeightSize
                '     widthB/= iScale
                '     HeightB/= iScale
                ' End While
                If widthB>widthSize
                    比例 = widthB/widthSize
                    widthB=widthSize
                    HeightB=HeightB/比例
                    功能 = True
                End If
                If HeightB>HeightSize
                    比例 = HeightB/HeightSize
                    HeightB=HeightSize
                    widthB=widthB/比例
                    功能 = True
                End If
                Dim b64 As string
                If 功能 = True
                    Dim size As Drawing.Size = new Size(widthB , HeightB)
                    '//新建一個bmp圖片
                    Dim bitmap As Drawing.Image = new System.Drawing.Bitmap(size.Width,size.Height)
                    '//新建一個畫板
                    Dim g As Drawing.Graphics = System.Drawing.Graphics.FromImage(bitmap)
                    '//設定高質量插值法
                    g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.High
                    '//設定高質量,低速度呈現平滑程度
                    g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality
                    '//清空一下畫布
                    g.Clear(Color.White)
                    '//在指定位置畫圖
                    g.DrawImage(照片, new System.Drawing.Rectangle(0, 0, bitmap.Width, bitmap.Height),new System.Drawing.Rectangle(0, 0, 照片.Width,照片.Height),System.Drawing.GraphicsUnit.Pixel)
                    If Size1="B"
                        b64=ImageToBase64(bitmap,System.Drawing.Imaging.ImageFormat.Png)
                    ElseIf Size1="S"
                        '//取得原影象的普通縮圖
                        Dim img As Drawing.Image
                        Dim myCallback As Drawing.Image.GetThumbnailImageAbort = new Drawing.Image.GetThumbnailImageAbort(AddressOf ThumbnailCallback)
                        img = 照片.GetThumbnailImage(widthB, HeightB, myCallback, IntPtr.Zero)
                        b64=ImageToBase64(img,System.Drawing.Imaging.ImageFormat.Png)
                    End if
                Else 
                    b64=base64String
                END if
                Return b64
            Else
                Return Nothing
            End If
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
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
    Function BASE64_TO_IMG(ByVal BASE_STR As String) As Drawing.Image'下載資料轉圖片
        Try
            Dim IMG As Drawing.Image
            If BASE_STR <> Nothing 'And Strings.Right(BASE_STR, 1) = "=" 
                Dim BYT As Byte() = Convert.FromBase64String(BASE_STR.Remove(0,22))
                Dim MS As New IO.MemoryStream(BYT)
                IMG = Drawing.Image.FromStream(MS)
                Return IMG
            Else
                Return Nothing
            End If 
        Catch ex As Exception
            Return Nothing
        End Try
    End Function 
End Class