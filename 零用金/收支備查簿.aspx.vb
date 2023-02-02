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
Partial Class 收支備查簿
    Inherits System.Web.UI.Page
    Dim con_14 As String = System.Web.Configuration.WebConfigurationManager.ConnectionStrings("ApplicationServices").ConnectionString
    Dim data As New SqlDataSource
    Dim data_dv As Data.DataView
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) HAndles Me.Load'每次讀取頁面
        data.ConnectionString = con_14
        If Not Page.IsPostBack Then
            Me.年.Text = (DateTime.Now.Year - 1911).ToString()
            Me.GridView1.PageIndex = Int32.MaxValue
            data.SelectCommand = "select Distinct 科目 from 科目表 Where 科目<>'' OR 科目 IS NOT NULL  order by 科目"
            data_dv = data.Select(New DataSourceSelectArguments)
            Dim 科目 As DropDownList = Me.科目
            科目.Items.Clear()
            科目.Items.Add("")
            科目.Items(0).Value = ""
            For j = 0 To data_dv.Count - 1
                Dim 科目名稱 As String = data_dv(j)(0)
                科目.Items.Add(科目名稱)
                科目.Items(j+1).Value = 科目名稱
            Next
            If Session("Uid")="3855"
                測試.Visible=True
            End If
        Else
            Me.Label1.Text = ""
            Me.Label2.Text = ""
            Me.Label3.Text = "" 
        End If 
    End Sub
    Protected Sub 種類_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)'當種類變成A才顯示收入，更換種類會將最大值設為最後一頁
        If (Me._種類.SelectedValue="A") Then
            Me.GridView1.columns(6).Visible = True
            Me.GridView1.columns(15).Visible = True
            Me.GridView1.columns(17).Visible = True
        Else
            Me.GridView1.columns(6).Visible = False
            Me.GridView1.columns(15).Visible = False
            Me.GridView1.columns(17).Visible = False
        End If
        Me.GridView1.PageIndex = Int32.MaxValue
    End Sub
    Protected Sub 頁1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        If Me.頁2.text="" 
            Me.頁2.text= Me.頁1.text
        End If 
    End Sub
    Protected Sub 頁2_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        If Me.頁1.text="" 
            Me.頁1.text= Me.頁2.text
        End If 
    End Sub
    Protected Sub 號數1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        If Me.號數2.text="" 
            Me.號數2.text= Me.號數1.text
        End If 
    End Sub
    Protected Sub 號數2_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        If Me.號數1.text="" 
            Me.號數1.text= Me.號數2.text
        End If 
    End Sub
    Protected Sub SelectedIndexChanged_尾頁(ByVal sender As Object, ByVal e As System.EventArgs)
        Me.GridView1.PageIndex = Int32.MaxValue
    End Sub
    Protected Sub Insert(ByVal sender As Object, ByVal e As System.EventArgs)'新增頁面，收入、支出預設值寫在資料庫
        initialization()
        Update(sender, e)'先初始化在存檔
        '沒用 Page_Load(sender,e)
            Dim 種類 As String = Me._種類.Text
            For i = 1 To 15
                Dim 年 As String = Me.年.Text
                Dim 摘要 As String = ""
                Dim 餘額 As String
                If i = 1 AND Me.GridView1.PageCount<>0
                    摘要 = "N'承上頁'"
                ElseIf i = 15
                    摘要 = "N'接下頁'"
                Else
                    摘要 = "NULL"
                End If 
                Dim insert1 As string
                data.InsertCommAnd = _
                    "INSERT INTO 收支備查簿 " & _
                    "(年, 取號, _種類, _頁, _列, 摘要) " & _
                    "VALUES " & _
                    "(" & 年 & ", 0, N'" & 種類 & "', " & (Me.GridView1.PageCount + 1).ToString() & ", " & i & ", " & 摘要 & ")"'現階段的總數+1
                data.Insert()
            Next
            If (種類="A") Then
                    data.UpdateCommAnd = _
                    "WITH CTE AS " & _
                    "(SELECT *, " & _
                        "(SELECT TOP 1 (CASE WHEN ISNULL(收入,0) = 0 And ISNULL(支出,0) = 0  THEN 餘額 ELSE 0 END) FROM 收支備查簿 WHERE _種類 = '" & 種類 & "' ORDER BY id) " & _
                        "+ " & _
                        "(SUM(" & _
	                    "(CASE WHEN ISNULL(摘要,'')<>'本月小計' And ISNULL(摘要,'')<>'累計至本月' THEN ISNULL(收入,0)ELSE 0 END )" & _
                        "-" & _ 
	                    "(CASE WHEN ISNULL(摘要,'')<>'本月小計' And ISNULL(摘要,'')<>'累計至本月' THEN ISNULL(支出,0)ELSE 0 END )" & _
	                    ") OVER (ORDER BY id " & _
                        "ROWS BETWEEN UNBOUNDED PRECEDING And CURRENT ROW))" & _
                        "AS RunningTotal " & _
                    "FROM 收支備查簿 WHERE _種類 = '" & 種類 & "') " & _
                    "UPDATE CTE SET 餘額 = RunningTotal"
                    data.Update()
            End If
            Me.GridView1.PageIndex = Int32.MaxValue
            
        Label1.Text="新增一頁成功"
    End Sub
    Protected Sub Update(ByVal sender As Object, ByVal e As System.EventArgs)'修改未鎖定資料步驟
        'TODO:轉換成整數會失敗
        Update_f()
        Me.GridView1.DataBind()
    End Sub
    Protected Sub Test(ByVal sender As Object, ByVal e As System.EventArgs)'測試用，無實質作用
        ' For i = 0 To Me.GridView1.Rows.Count - 1
        '     If CType(Me.GridView1.Rows(i).FindControl("刪除選取"), CheckBox).Checked=True And i<>0 AND i <> Me.GridView1.Rows.Count - 1
        '         Dim s頁 As int32 = Me.GridView1.PageIndex+1
        '         Dim s列 As int32 = CType(Me.GridView1.Rows(i).FindControl("_列"), TextBox).Text
        '         data.ConnectionString = con_14
        '         data.SelectCommAnd = "SELECT * FROM 收支備查簿 WHERE _種類='A' AND ((_頁='" & s頁 & "' AND _列>='" & s列 & "') OR _頁>'" & s頁 & "') AND _列<>'1' AND _列<>'15'"
        '         data_dv = data.Select(New DataSourceSelectArguments)
        '         For j=0 to data_dv.count-2
        '             Dim 頁 As int32 = data_dv(j)("_頁").ToString()
        '             Dim 列 As int32 = data_dv(j)("_列").ToString()
        '             Dim id As int32 = data_dv(j)("id").ToString()
        '             Dim n頁 As int32 = data_dv(j)("_頁").ToString()
        '             Dim n列 As int32 = (data_dv(j)("_列").ToString())+1
        '             If n列>14
        '                 n列=n列-13
        '                 n頁=n頁+1
        '             End If
        '             label1.Text=label1.Text & 頁 & ":" & 列 & " n " & n頁 & ":" & n列 & "<BR>"
        '             If j=0
        '                 label1.Text=label1.Text & 頁 & ":" & 列 & " 刪除<BR>"
        '             END If
        '         Next
        '     End If
        ' Next
        '在第38頁第3行後面插入2行(完成)、修正日誌
        ' data.ConnectionString = con_14
        ' data.SelectCommAnd = "SELECT * FROM 收支備查簿 WHERE _種類='A' AND ((_頁>'37' AND _列>'3') OR _頁>'38') AND _列<>'1' AND _列<>'15'"
        ' data_dv = data.Select(New DataSourceSelectArguments)
        ' For i=0 to data_dv.count-1
        '     Dim 頁 As int32 = data_dv(data_dv.count-1-i)("_頁").ToString()
        '     Dim 列 As int32 = data_dv(data_dv.count-1-i)("_列").ToString()
        '     Dim id As int32 = data_dv(data_dv.count-1-i)("id").ToString()
        '     Dim n頁 As int32 = data_dv(data_dv.count-1-i)("_頁").ToString()
        '     Dim n列 As int32 = (data_dv(data_dv.count-1-i)("_列").ToString()-2)
        '     If n列<2
        '         n列=n列+13
        '         n頁=n頁-1
        '     End If
        '     If 頁=38 AND 列 < 6
        '         n列=n列+11
        '         n頁=40
        '     END If
        '     Dim data_dv2 As Data.DataView
        '     data.ConnectionString = con_14
        '     data.SelectCommAnd = "SELECT id As 原id FROM 收支備查簿 WHERE _種類='A' AND _列='" & n列 & "' AND _頁='" & n頁 & "'"
        '     data_dv2 = data.Select(New DataSourceSelectArguments)
        '     Dim nid As int32 = data_dv2(0)("原id").ToString()
        '     Dim 預支日期 As String = data_dv(data_dv.count-1-i)("預支日期").ToString()
        '     If 預支日期 <>""
        '         預支日期 = 預支日期.substring(0,10)
        '     End If
        '     Dim 歸還日期 As String = data_dv(data_dv.count-1-i)("歸還日期").ToString()
        '     If 歸還日期 <>""
        '         歸還日期 = 歸還日期.substring(0,10)
        '     End If
            'label1.Text=label1.Text & ":" & CType(Me.GridView1.Rows(0).FindControl("_列"), TextBox).Text
            ' data.UpdateCommAnd = "UPDATE 日誌 SET " & _
            '     "id = '" & id & "' " & _
            '     "WHERE id = '" & nid & "'"
            ' data.Update()
            ' If i>data_dv.count-1
                ' label1.Text=label1.Text & "刪除" & n列 & ":" & n頁 & "<BR>"
                ' data.UpdateCommAnd = "UPDATE 收支備查簿 SET " & _
                '     "單位別 = NULL,承辦人 = NULL,月 = NULL,日 = NULL,科目 = NULL,科目2 = NULL,摘要 = NULL,姓名 = NULL,商號 = NULL,經手人 = NULL,種類 = NULL,號數 = NULL," & _
                '     "收入 = 0,支出 = 0,備註 = NULL,取號 = '0',送交主計室日期 = NULL,回覆 = 'False',鎖定 = 'False',過審 = 'False',送出 = 'False',預支日期 = NULL,歸還日期 = NULL " & _
                '     "WHERE _列 = '" & n列 &"' AND _頁 = " & n頁 & " AND _種類='A'"
                ' data.Update()
            ' Else
                ' label1.Text=label1.Text & 頁 & "頁:" & 列 & "列;取代" & n頁 & "頁:" & n列 & "列<BR>"
                ' Dim Update1 As string ="UPDATE 收支備查簿 SET " & _
                ' "單位別 = NULLIF(N'" & data_dv(i)("單位別").ToString() & "', ''), " & _
                ' "承辦人 = NULLIF(N'" & data_dv(i)("承辦人").ToString() & "', ''), " & _
                ' "月 = NULLIF(N'" & data_dv(i)("月").ToString() & "', ''), " & _
                ' "日 = NULLIF(N'" & data_dv(i)("日").ToString() & "', ''), " & _
                ' "科目 = NULLIF(N'" & data_dv(i)("科目").ToString() & "', ''), " & _
                ' "科目2 = NULLIF(N'" & data_dv(i)("科目2").ToString() & "', ''), " & _
                ' "摘要 = NULLIF(N'" & data_dv(i)("摘要").ToString() & "', ''), " & _
                ' "姓名 = NULLIF(N'" & data_dv(i)("姓名").ToString() & "', ''), " & _
                ' "商號 = NULLIF(N'" & data_dv(i)("商號").ToString() & "', ''), " & _
                ' "經手人 = NULLIF(N'" & data_dv(i)("經手人").ToString() & "', ''), " & _
                ' "種類 = NULLIF(N'" & data_dv(i)("種類").ToString() & "', ''), " & _
                ' "號數 = NULLIF(N'" & data_dv(i)("號數").ToString() & "', ''), " & _
                ' "收入 = NULLIF('" & data_dv(i)("收入").ToString() & "', ''), " & _
                ' "支出 = NULLIF('" & data_dv(i)("支出").ToString() & "', ''), " & _
                ' "備註 = NULLIF(N'" & data_dv(i)("備註").ToString() & "', ''), " & _
                ' "取號 = NULLIF(N'" & data_dv(i)("取號").ToString() & "', ''), " & _
                ' "送交主計室日期 = NULLIF(N'" & data_dv(i)("送交主計室日期").ToString() & "', ''), " & _
                ' "回覆 = NULLIF(N'" & data_dv(i)("回覆").ToString() & "', ''), " & _
                ' "鎖定 = NULLIF(N'" & data_dv(i)("鎖定").ToString() & "', ''), " & _
                ' "過審 = NULLIF(N'" & data_dv(i)("過審").ToString() & "', ''), " & _
                ' "送出 = NULLIF(N'" & data_dv(i)("送出").ToString() & "', ''), " & _
                ' "預支日期 = NULLIF('" & 預支日期 & "', '')," & _
                ' "歸還日期 = NULLIF('" & 歸還日期 & "', '')," & _
                ' "主計室簽核 = NULLIF(N'" & data_dv(i)("主計室簽核").ToString() & "', ''), " & _
                ' "駁回原因 = NULLIF(N'" & data_dv(i)("駁回原因").ToString() & "', '') " & _
                ' "WHERE _列 = '" & n列 &"' AND _頁 = " & n頁 & " AND _種類='A'"
                ' data.UpdateCommAnd = Update1
                ' data.Update()
            ' End If
        ' Next
        ' Update(sender,e)
    End Sub
    Protected Sub Download(ByVal sender As Object, ByVal e As System.EventArgs)'下載 A、B、XZ全部通用
        Dim MyGUID As String = Guid.NewGuid().ToString("N")
        Dim MyExcel As String = MapPath(".\Excel\Temp\") & MyGUID & ".xlsx"
        Dim _種類 As String = Me._種類.Text
        System.IO.File.Copy(MapPath(".\Excel\收支備查簿.xlsx"), MyExcel)
        '語法封存
            'Select Case _種類 
            '    Case "B"
            '    Case "XZ"
            '    Case Else
            'End Select
        Dim xlApp As New Excel.ApplicationClass()
        xlApp.DisplayAlerts = False
        xlApp.ScreenUpdating = False
        xlApp.EnableEvents = False
        Dim xlWorkBook As Excel.Workbook = xlApp.Workbooks.Open(MyExcel, 0, False, 5, "", "", True, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", False, False, 0, True, 1, 0)
        Dim xlWorkSheet As Excel.Worksheet
        xlWorkSheet = CType(xlWorkBook.Sheets("零用金用表"), Excel.Worksheet)
        xlWorkSheet.Activate()
        Dim 年 As String = Me.年.Text
        data.ConnectionString = con_14
        data.SelectCommAnd = "SELECT * FROM 收支備查簿 WHERE  ((''='" &  Me.頁1.Text & "' OR ''='" &  Me.頁2.Text & "') "& _
            "OR ( _頁 BETWEEN '" &  Me.頁1.Text & "' AND '" &  Me.頁2.Text & "')) "& _
            "AND ((''='" &  Me.號數1.Text & "' OR ''='" &  Me.號數2.Text & "') "& _
            "OR ( 號數 BETWEEN "& _
            "SUBSTRING('" &  Me.號數1.Text & "', PATINDEX('%[^0]%', '" &  Me.號數1.Text & "'), 3) AND "& _
            "SUBSTRING('" &  Me.號數2.Text & "', PATINDEX('%[^0]%', '" &  Me.號數2.Text & "'), 3))) " & _
            "AND 摘要<>'接下頁' AND 摘要<>'承上頁' AND _種類='" &  _種類 & "'order by _頁,_列"'全輸出，但摘要沒東西不輸出，有月日及科目但沒摘要也會輸出?，加入頁數、號數區間查詢
            'SUBSTRING(int,int)起始於指定的字元位置，並且具有指定的長度
            'PATINDEX(收尋值,子字串)傳回指定之運算式中的模式，在所有有效文字和字元資料類型中第一次出現的起始位置
            '%^[0]%收尋開頭為0的位置
        Dim D_Height As int32
        Dim D_Width As int32
        Dim Data_Height_H As int32
        Dim Data_Height_L As int32
        Dim F_Height As int32
        Dim F_Width As int32
        D_Height=26'資料範圍列
        D_Width=12'資料範圍行
        Data_Height_H=19'資料最高
        Data_Height_L=5'資料最底
        data_dv = data.Select(New DataSourceSelectArguments)
        If data_dv.Count = 0
            xlWorkBook.Close()
            xlApp.Quit()
            ReleaseObject(xlWorkSheet)
            ReleaseObject(xlWorkBook)
            ReleaseObject(xlApp)
            System.IO.File.Delete(MyExcel)
            Exit Sub
        End If
        Dim 總頁數 As Int32 = 0
        Dim Data_Row As Int32
        Data_Row=13'平均資料列
        IF data_dv.Count>15
            If ((data_dv.Count-1) Mod Data_Row)>1  '35-1=34/13=2...8
                總頁數=((data_dv.Count-1)\Data_Row)+1'"/"為整數除法，會四捨五入，"\"為除法，捨棄餘數
            ELSE
                總頁數=(data_dv.Count-1)\Data_Row
            END If
        Else
            總頁數=1
        END If
        For i = 2 To 總頁數'制定範圍並複製，範圍為頁數
            xlWorkSheet.Range(xlWorkSheet.Cells(D_Height * i - (D_Height-1), 1), xlWorkSheet.Cells(D_Height * i, D_Width)).Value(11) = xlWorkSheet.Range(xlWorkSheet.Cells(1, 1), xlWorkSheet.Cells(D_Height, D_Width)).Value(11)'(27,1) (52,12) = (1, 1) (26, 12)
            xlWorkSheet.Range(xlWorkSheet.Cells(D_Height * i - Data_Height_H, 1), xlWorkSheet.Cells(D_Height * i - Data_Height_L, D_Width)).RowHeight = 33'(33,1) (47,12) 高度為33 7~21
            xlWorkSheet.Rows(D_Height * i - (D_Height-1)).PageBreak = xlPageBreakManual'列27從開始載入
        Next
        Dim arr As Object = xlWorkSheet.Range(xlWorkSheet.Cells(1, 1), xlWorkSheet.Cells(總頁數 * D_Height, D_Width)).Value'(1,1) (網頁頁數26,12)
        Dim arr2 As Object = xlWorkSheet.Range(xlWorkSheet.Cells(1, 1), xlWorkSheet.Cells(總頁數 * D_Height, D_Width))
        '已經做出第一頁，中間
        For i = 0 To data_dv.Count - 1
            Dim 月 As String = data_dv(i)("月").ToString()
            Dim 日 As String = data_dv(i)("日").ToString()
            Dim 科目 As String = data_dv(i)("科目").ToString()
            If data_dv(i)("科目2").ToString()<>""
                科目 = 科目 & Chr(10) & data_dv(i)("科目2").ToString()
            End If
            Dim 摘要 As String = Trim(data_dv(i)("摘要").ToString())
            摘要=摘要.Replace(vbCrLf,"")
            摘要=Trim(摘要)
            Dim 姓名 As String = data_dv(i)("姓名").ToString()
            Dim 商號 As String = Trim(data_dv(i)("商號").ToString())
            Dim 經手人 As string = data_dv(i)("經手人").ToString()
            Dim 種類 As String = data_dv(i)("種類").ToString()
            Dim 號數 As String = data_dv(i)("號數").ToString()
            Dim 收入 As String = data_dv(i)("收入").ToString()
            收入 = If (收入 = "", "", CLng(收入).ToString("n0"))
            Dim 支出 As String = data_dv(i)("支出").ToString()
            支出 = If (支出 = "", "", CLng(支出).ToString("n0"))
            Dim 餘額 As String = data_dv(i)("餘額").ToString() 
            Select Case _種類
                Case "B"
                    餘額=0
                Case "XZ"
                    餘額=0
                Case Else
            End Select
            餘額 = If (餘額 = "", "", CLng(餘額).ToString("n0"))
            號數 = If (號數 = "", "", CLng(號數).ToString("000"))
            '以下為合併程式
            'i為第幾個資料、j為輸出的格子
            Dim j As Long 
            Dim 第幾頁 As Int32
            If data_dv.Count>15'判定第一頁能塞滿嗎?
                If ((i-1) \ 13)<總頁數 '14-2,27-2、3,
                    第幾頁=((i-1)\13)+1'"/"為整數除法，會四捨五入，"\"為除法，捨棄餘數
                ELSE
                    第幾頁=(i-1)\13
                END If
            Else
                第幾頁=1
            End If
            If data_dv.Count>15
                If i<14'第一頁
                    Data_Row=14
                    j = (i \ Data_Row ) + (i Mod Data_Row) + 7 '從第一列輸出
                ElseIf 總頁數=第幾頁'14-2
                    Data_Row=14
                    j = D_Height * ((i-1) \ 13 ) + ((i-1) Mod 13) + 8 '前一列為13
                    If (i\13)=總頁數 And ((i-1) Mod 13)=0
                        j=(D_Height * (((i-1) \ 13 )-1) + 13) + 8
                    End If
                Else
                    Data_Row=13
                    j = D_Height * ((i-1) \ Data_Row ) + ((i-1) Mod Data_Row) + 8'從第二列輸出
                End If 
            Else
                Data_Row=15
                j = (i Mod Data_Row) + 7
            End If 
            'i=0、j=7 j為列位置
            
            Dim 年度 As String
            If i=0
                年度 = "　　　　　　　　　　　　　　　　　　　　　　中華民國　　"& 年 &"　　年度　　　　　　　　　　　　　　　　　　　第1頁"
            Else
                年度 = "　　　　　　　　　　　　　　　　　　　　　　中華民國　　"& 年 &"　　年度　　　　　　　　　　　　　　　　　　　第"& 第幾頁 &"頁"
            End If
            If i=0'第一筆輸出年度每頁
                If 總頁數=第幾頁'最後一頁
                    arr(j+14, 4)= ""
                End If
                arr(J-3,1) = 年度
            ElseIf ((i-1) Mod 13)=0 And i>13 And i<>data_dv.count-1'第筆輸出年度每頁
                arr(J-4,1) = 年度'i=14、j=30
                arr(J-1,12) = "=" & arr2(J-14,12).Address'取上一頁最後的餘額
                If 總頁數=第幾頁'最後一頁
                    arr(j+13, 4)= ""
                End If
            END If
            
            arr(j, 1) = 月
            arr(j, 2) = 日
            arr(j, 3) = 科目
            arr(j, 4) = 摘要
            arr2(j, 4).HorizontalAlignment=-4131 '靠左對齊
            If(姓名 <> "")
                Dim string1 As string = MapPath(".\image\")& i &"姓名.png"'建造圖片 完成
                Dim 姓名di As Drawing.image
                姓名di = BASE64_TO_IMG(姓名)
                姓名di.save(string1) 
                姓名di.Dispose()
                xlWorkSheet.Shapes.AddPicture(string1,False,True,arr2(j,5).Left+1,arr2(j,5).Top+1,arr2(j,5).Width-1,arr2(j,5).Height-1)'可以插入圖片,AddPicture (FileName、 LinkToFile、 SaveWithDocument、 Left、 Top、 Width、 Height)
                System.IO.File.Delete(string1)
            End If 
            If(經手人 <> "")
                Dim string1 As string = MapPath(".\image\")& i &"經手人.png"'建造圖片 完成
                Dim 經手人di As Drawing.image
                經手人di = BASE64_TO_IMG(經手人)
                經手人di.save(string1)
                經手人di.Dispose()
                xlWorkSheet.Shapes.AddPicture(string1,False,True,arr2(j,7).Left+1,arr2(j,7).Top+1,arr2(j,7).Width-1,arr2(j,7).Height-1)'插入圖片
                System.IO.File.Delete(string1)
            End If 
            If(姓名 = "") And (經手人 = "")
                arr2(j,1).Rows.AutoFit'自動調整高度
            End If
            arr(j, 6) = 商號
            arr(j, 8) = 種類
            arr(j, 9) = 號數
            arr(j, 10) = 收入
            arr(j, 11) = 支出
            If i=0 '第一筆
                arr(j, 12) = 餘額
            ElseIf 摘要="本月小計" OR 摘要="累計至本月"
                arr(j, 12) = "=" & arr2(j-1, 12).Address
            Else
                arr(j, 12) = "=" & arr2(j-1, 12).Address & "+" & arr2(j, 10).Address & "-" & arr2(j, 11).Address
            END If
            ' arr(j, 12) = 餘額
            If (i Mod Data_Row)=(Data_Row-1) '最後一列
            END If
            If (i Mod Data_Row)=(Data_Row-1) And i<>data_dv.Count - 1'最後一列、最後一筆
            END If
        Next
        xlWorkSheet.Range(xlWorkSheet.Cells(1, 1), xlWorkSheet.Cells(總頁數 * D_Height, D_Width)).Value = arr
       ' xlWorkSheet.Range(xlWorkSheet.Cells(1, 1), xlWorkSheet.Cells(總頁數 * D_Height, D_Width)).Rows.AutoFit'ˋ動態調整高度，但圖片會受影響
        xlWorkBook.Save()
        xlWorkBook.Close()
        xlApp.Quit()
        ReleaseObject(xlWorkSheet)
        ReleaseObject(xlWorkBook)
        ReleaseObject(xlApp)
        Response.Clear()
        Response.ClearHeaders()
        Response.Buffer = True
        Response.ContentType = "application/octet-stream"
        Dim downloadfilename = "收支備查簿.xlsx"
        Response.AddHeader("Content-Disposition", "attachment;FileName=" + Uri.EscapeDataString(downloadfilename))
        Response.WriteFile(MyExcel)
        System.Web.HttpContext.Current.ApplicationInstance.CompleteRequest()
        Response.Flush()
        System.IO.File.Delete(MyExcel)
        Response.End()
    End Sub
    Protected Sub Delete(ByVal sender As Object, ByVal e As System.EventArgs)'刪除末頁資料步驟
        initialization()
        Update(sender, e)
        Me.GridView1.PageIndex = Int32.MaxValue
        Me.GridView1.DataBind()
        Dim delete1 As string=""
        For i = 0 To Me.GridView1.Rows.Count - 1
            Dim id As String = CType(Me.GridView1.Rows(i).FindControl("id"), TextBox).Text
            delete1="DELETE FROM 收支備查簿 " & _
            "WHERE id = '" & id & "'"
            data.ConnectionString = con_14
            Dim sql As string = "SELECT * FROM 收支備查簿 where id=" & id
            data.SelectCommAnd = sql
            data_dv = data.Select(New DataSourceSelectArguments)
            Dim 年 As String = Me.年.Text
            Dim _種類 As String = Me._種類.Text
            Dim 單位別 As String = CType(Me.GridView1.Rows(i).FindControl("單位別"),DropDownList).Text
            Dim 承辦人 As String = CType(Me.GridView1.Rows(i).FindControl("承辦人"), DropDownList).Text
            Dim 月 As String = CType(Me.GridView1.Rows(i).FindControl("月"), DropDownList).Text
            Dim 日 As String = CType(Me.GridView1.Rows(i).FindControl("日"), DropDownList).Text
            Dim 科目 As String = CType(Me.GridView1.Rows(i).FindControl("科目"), DropDownList).Text
            Dim 科目2 As String = CType(Me.GridView1.Rows(i).FindControl("科目2"), DropDownList).Text
            Dim 摘要 As String = CType(Me.GridView1.Rows(i).FindControl("摘要"), TextBox).Text
            Dim 姓名 As String = CType(Me.GridView1.Rows(i).FindControl("姓名"), ImageButton).ImageUrl
            Dim 商號 As String = CType(Me.GridView1.Rows(i).FindControl("商號"), TextBox).Text
            Dim 經手人 As String = CType(Me.GridView1.Rows(i).FindControl("經手人"), ImageButton).ImageUrl
            Dim 種類 As String = CType(Me.GridView1.Rows(i).FindControl("種類"), TextBox).Text
            Dim 號數 As String = CType(Me.GridView1.Rows(i).FindControl("號數"), TextBox).Text
            Dim 收入 As String = CType(Me.GridView1.Rows(i).FindControl("收入"), TextBox).Text
            Dim 支出 As String = CType(Me.GridView1.Rows(i).FindControl("支出"), TextBox).Text
            收入=收入.Replace(",", "").Replace("N", "").Replace("T", "").Replace("$", "")
            支出=支出.Replace(",", "").Replace("N", "").Replace("T", "").Replace("$", "")
            Dim 餘額 As String = CType(Me.GridView1.Rows(i).FindControl("餘額"), TextBox).Text
            Dim 備註 As String = CType(Me.GridView1.Rows(i).FindControl("備註"), TextBox).Text
            If data_dv(0)("送出").ToString()="True"'送出後能刪除，但須將刪除資料記錄在日誌中
                    Dim date1 As string = DateTime.now.tostring()
                    Dim date2 As string = DateTime.now.tostring("yyyy-MM-dd HH:mm:ss")
                    '新增刪除資料表
                    data.insertCommAnd = _
                        "INSERT INTO 刪除資料表 " & _
                        "(id_收, _種類, 單位別, 承辦人, 年, 月, 日, 科目, 科目2, 摘要, 姓名, 商號, 經手人, 種類, 號數, 收入, 支出, 備註 ,date ) " & _
                        "VALUES " & _
                        "(" & id & "," & _
                        "NULLIF(N'" & _種類  & "', ''), " & _
                        "NULLIF(N'" & 單位別 & "', ''), " & _
                        "NULLIF(N'" & 承辦人 & "', ''), " & _
                        "NULLIF(N'" & 年     & "', ''), " & _
                        "NULLIF(N'" & 月     & "', ''), " & _
                        "NULLIF(N'" & 日     & "', ''), " & _
                        "NULLIF(N'" & 科目   & "', ''), " & _
                        "NULLIF(N'" & 科目2  & "', ''), " & _
                        "NULLIF(N'" & 摘要   & "', ''), " & _
                        "NULLIF(N'" & 姓名   & "', ''), " & _
                        "NULLIF(N'" & 商號   & "', ''), " & _
                        "NULLIF(N'" & 經手人 & "', ''), " & _
                        "NULLIF(N'" & 種類   & "', ''), " & _
                        "NULLIF('"  & 號數   & "', ''), " & _
                        "NULLIF('"  & 收入   & "', ''), " & _
                        "NULLIF('"  & 支出   & "', ''), " & _
                        "NULLIF(N'" & 備註   & "', ''), " & _
                        "NULLIF(N'" & date1 & "', ''))"
                    data.insert()
                    '新增日誌
                    sql = "SELECT * FROM [刪除資料表] where id_收=" & id & "And date = '" & date1 & "'"
                    data.SelectCommAnd = sql
                    data_dv = data.Select(New DataSourceSelectArguments)
                    Dim id_2 As string =data_dv(0)("id").ToString()
                    data.insertCommAnd = _
                        "INSERT INTO 日誌 " & _
                        "(id, 動作, 命令,日期,日期2) " & _
                        "VALUES " & _
                        "(N'" & id & "', N'刪除', N'刪除資料id=" & id_2 & "' , N'" & date1 & "' , '" & date2 & "')"
                    data.insert()
            Else
            End If 
            data.DeleteCommAnd =delete1
            data.Delete()
        Next
        Me.GridView1.DataBind()
        Label1.Text="刪除末頁成功"
    End Sub
    Protected Sub 修改已過審_Click(ByVal sender As Object, ByVal e As System.EventArgs)'修改已過審
        Dim 年 As String = Me.年.Text
        Dim _種類 As String = Me._種類.Text
        Dim id_號 As String
        Dim 單位別_號 As String
        Dim 承辦人_號 As String
        Dim 號數_號 As String
        Dim 種類_號 As String
        Dim 備註_號 As String
        Dim 作用 As boolean =False
        For i = 0 To Me.GridView1.Rows.Count - 1
            If CType(Me.GridView1.Rows(i).FindControl("刪除選取"), CheckBox).Checked=True
                作用=true
                Dim id As String = CType(Me.GridView1.Rows(i).FindControl("id"), TextBox).Text
                Dim 單位別 As String = CType(Me.GridView1.Rows(i).FindControl("單位別"),DropDownList).Text
                Dim 承辦人 As String = CType(Me.GridView1.Rows(i).FindControl("承辦人"), DropDownList).Text
                Dim 月 As String = CType(Me.GridView1.Rows(i).FindControl("月"), DropDownList).Text
                Dim 日 As String = CType(Me.GridView1.Rows(i).FindControl("日"), DropDownList).Text
                Dim 科目 As String = CType(Me.GridView1.Rows(i).FindControl("科目"), DropDownList).Text
                Dim 科目2 As String = CType(Me.GridView1.Rows(i).FindControl("科目2"), DropDownList).Text
                Dim 原本科目 As String =科目
                If 科目2<>""
                    原本科目 = (科目 & ";" & 科目2)
                End If
                Dim 摘要 As String = CType(Me.GridView1.Rows(i).FindControl("摘要"), TextBox).Text
                Dim 姓名 As String = CType(Me.GridView1.Rows(i).FindControl("姓名"), ImageButton).ImageUrl
                Dim 商號 As String = CType(Me.GridView1.Rows(i).FindControl("商號"), TextBox).Text
                Dim 經手人 As String = CType(Me.GridView1.Rows(i).FindControl("經手人"), ImageButton).ImageUrl
                Dim 種類 As String = CType(Me.GridView1.Rows(i).FindControl("種類"), TextBox).Text
                Dim 號數 As String = CType(Me.GridView1.Rows(i).FindControl("號數"), TextBox).Text
                Dim 收入 As String = CType(Me.GridView1.Rows(i).FindControl("收入"), TextBox).Text
                收入=收入.Replace(",", "").Replace("N", "").Replace("T", "").Replace("$", "")
                Dim 支出 As String = CType(Me.GridView1.Rows(i).FindControl("支出"), TextBox).Text
                支出=支出.Replace(",", "").Replace("N", "").Replace("T", "").Replace("$", "")
                Dim 餘額 As String = CType(Me.GridView1.Rows(i).FindControl("餘額"), TextBox).Text
                Dim 備註 As String = CType(Me.GridView1.Rows(i).FindControl("備註"), TextBox).Text
                Dim 預支日期 As String = nothing
                If CType(Me.GridView1.Rows(i).FindControl("預支日期"), TextBox).text<>""
                    預支日期 = CType(Me.GridView1.Rows(i).FindControl("預支日期"), TextBox).text
                    預支日期 = taiwancalendarto(預支日期)
                End If
                Dim 歸還日期 As String = nothing
                If CType(Me.GridView1.Rows(i).FindControl("歸還日期"), TextBox).text<>""
                    歸還日期 = CType(Me.GridView1.Rows(i).FindControl("歸還日期"), TextBox).text
                    歸還日期 = taiwancalendarto(歸還日期)
                End If 
                Me.GridView1.Rows(i).FindControl("姓名").Visible = True
                Me.GridView1.Rows(i).FindControl("經手人").Visible = True
                Dim 審核狀態 As String=CType(Me.GridView1.Rows(i).FindControl("審核狀態"), Label).Text
                If 審核狀態="通過" OR 審核狀態= "送交主計室" OR 審核狀態= "已送審" '新增修改已過審紀錄
                    DIM sql As string = "SELECT * FROM [收支備查簿] where id=" & id
                    data.SelectCommAnd = sql
                    data_dv = data.Select(New DataSourceSelectArguments)
                    Dim 單位別_前 As String = data_dv(0)("單位別").ToString()
                    Dim 承辦人_前 As String = data_dv(0)("承辦人").ToString()
                    Dim 月_前 As String = data_dv(0)("月").ToString()
                    Dim 日_前 As String = data_dv(0)("日").ToString()
                    Dim 科目_前 As String = data_dv(0)("科目").ToString()
                    If data_dv(0)("科目2").ToString()<>""
                        科目_前 = data_dv(0)("科目").ToString() & ";" & data_dv(0)("科目2").ToString()
                    End If
                    Dim 摘要_前 As String = data_dv(0)("摘要").ToString()
                    Dim 姓名_前 As String = data_dv(0)("姓名").ToString()
                    Dim 商號_前 As String = data_dv(0)("商號").ToString()
                    Dim 經手人_前 As string = data_dv(0)("經手人").ToString()
                    Dim 種類_前 As String = data_dv(0)("種類").ToString()
                    Dim 號數_前 As String = data_dv(0)("號數").ToString()
                    Dim 收入_前 As String = data_dv(0)("收入").ToString()
                    Dim 支出_前 As String = data_dv(0)("支出").ToString()
                    Dim 備註_前 As String = data_dv(0)("備註").ToString()
                    '判斷資料是否有修改
                    If 號數<>""
                        號數=CType(號數,Int32)
                        號數=CType(號數,string)
                    End If 
                    'B XZ 沒有收入，會到導致""<>"0" ，先將收入及支出空值得收入改成預設值0
                    If 收入=""
                        收入="0"
                    End If
                    If 支出=""
                        支出="0"
                    End If  
                    If 單位別<>單位別_前 or 承辦人<>承辦人_前 or 月<>月_前 or 日<>日_前 or 原本科目<>科目_前 or 摘要<>摘要_前 or 姓名<>姓名_前 or 商號<>商號_前 or 經手人<>經手人_前 or 種類<>種類_前 or 號數<>號數_前 or (收入<>收入_前 And 種類="A") or 支出<>支出_前 or 備註<>備註_前
                        '程式正式執行
                        '新增修改資料
                        Dim date1 As string = DateTime.now.tostring()
                        Dim date2 As string = DateTime.now.tostring("yyyy-MM-dd HH:mm:ss")
                        Dim insert1 As string = _
                            "INSERT INTO 修改資料 " & _
                            "(id_收," & _
                            "單位別,單位別_改," & _
                            "承辦人,承辦人_改," & _
                            "月    ,月_改    ," & _
                            "日    ,日_改    ," & _
                            "科目  ,科目_改  ," & _
                            "摘要  ,摘要_改  ," & _
                            "姓名  ,姓名_改  ," & _
                            "商號  ,商號_改  ," & _
                            "經手人,經手人_改," & _
                            "種類  ,種類_改  ," & _
                            "號數  ,號數_改  ," & _
                            "收入  ,收入_改  ," & _
                            "支出  ,支出_改  ," & _
                            "備註  ,備註_改  ," & _
                            "date) " & _
                            "VALUES " & _
                            "(" & id & "," & _
                            "NULLIF(N'" & 單位別_前 & "', ''),NULLIF(N'" & 單位別   & "', '')," & _
                            "NULLIF(N'" & 承辦人_前   & "', ''),NULLIF(N'" & 承辦人   & "', '')," & _
                            "NULLIF('"  & 月_前     & "', ''),NULLIF('"  & 月       & "', '')," & _
                            "NULLIF('"  & 日_前     & "', ''),NULLIF('"  & 日       & "', '')," & _
                            "NULLIF(N'" & 科目_前   & "', ''),NULLIF(N'" & 原本科目 & "', '')," & _
                            "NULLIF(N'" & 摘要_前   & "', ''),NULLIF(N'" & 摘要     & "', '')," & _
                            "NULLIF(N'" & 姓名_前   & "', ''),NULLIF(N'" & 姓名     & "', '')," & _
                            "NULLIF(N'" & 商號_前   & "', ''),NULLIF(N'" & 商號     & "', '')," & _
                            "NULLIF(N'" & 經手人_前 & "', ''),NULLIF(N'" & 經手人   & "', '')," & _
                            "NULLIF(N'" & 種類_前   & "', ''),NULLIF(N'" & 種類     & "', '')," & _
                            "NULLIF('"  & 號數_前   & "', ''),NULLIF('"  & 號數     & "', '')," & _
                            "NULLIF('"  & 收入_前   & "', ''),NULLIF('"  & 收入     & "', '')," & _
                            "NULLIF('"  & 支出_前   & "', ''),NULLIF('"  & 支出     & "', '')," & _
                            "NULLIF(N'" & 備註_前   & "', ''),NULLIF(N'" & 備註     & "', '')," & _
                            "NULLIF(N'" & date1     & "', ''))"
                        data.insertCommAnd = insert1
                        data.insert()
                        '查詢修改資料表的ID，並把他給日誌
                        sql = "SELECT * FROM [修改資料] where id_收=" & id & "And date = '" & date1 & "'"
                        data.SelectCommAnd = sql
                        data_dv = data.Select(New DataSourceSelectArguments)
                        Dim id_2 As string =data_dv(0)("id").ToString()
                        data.insertCommAnd = _
                        "INSERT INTO 日誌 " & _
                        "(id, 動作, 命令,日期,日期2) " & _
                        "VALUES " & _
                        "(N'" & id & "', N'修改', N'修改資料id="& id_2 &"', N'" & date1 & "', '" & date2 & "')"
                        data.insert()
                    End If 
                End If 
                Dim Update1 As string ="UPDATE 收支備查簿 SET " & _
                    "單位別 = NULLIF(N'" & 單位別 & "', ''), " & _
                    "承辦人 = NULLIF(N'" & 承辦人 & "', ''), " & _
                    "月 = NULLIF(N'" & 月 & "', ''), " & _
                    "日 = NULLIF(N'" & 日 & "', ''), " & _
                    "科目 = NULLIF(N'" & 科目 & "', ''), " & _
                    "科目2 = NULLIF(N'" & 科目2 & "', ''), " & _
                    "摘要 = NULLIF(N'" & 摘要 & "', ''), " & _
                    "姓名 = NULLIF(N'" & 姓名 & "', ''), " & _
                    "商號 = NULLIF(N'" & 商號 & "', ''), " & _
                    "經手人 = NULLIF(N'" & 經手人 & "', ''), " & _
                    "種類 = NULLIF(N'" & 種類 & "', ''), " & _
                    "號數 = NULLIF(N'" & 號數 & "', ''), " & _
                    "收入 = REPLACE(ISNULL(NULLIF('" & 收入 & "', ''),'0'), ',', ''), " & _
                    "支出 = REPLACE(ISNULL(NULLIF('" & 支出 & "', ''),'0'), ',', ''), " & _
                    "備註 = NULLIF(N'" & 備註 & "', ''), " & _
                    "預支日期 = (CASE WHEN ISDATE(NULLIF(N'" & 預支日期 & "', ''))=1 Then NULLIF(N'" & 預支日期 & "', '') End )," & _
                    "歸還日期 = (CASE WHEN ISDATE(NULLIF(N'" & 歸還日期 & "', ''))=1 Then NULLIF(N'" & 歸還日期 & "', '') End )" & _
                    "WHERE id = '" & id & "' And 鎖定 = 'True' "'修改鎖定
                data.UpdateCommAnd = Update1
                data.Update()
            End If 
        Next
        If 作用=False
            Label3.Text="請先選取已送審之資料"
        Else
            Label3.Text=""
        End If
        Update(sender,e)
        Label1.Text="過審資料已修改成功"
    End Sub
    Protected Sub 收回_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Update_f()
        Dim 作用 As boolean =False
        For i = 0 To Me.GridView1.Rows.Count - 1
            If CType(Me.GridView1.Rows(i).FindControl("審核"), CheckBox).Checked=True
                作用=true
                Dim id As String = CType(Me.GridView1.Rows(i).FindControl("id"), TextBox).Text
                Dim 審核狀態 As String=CType(Me.GridView1.Rows(i).FindControl("審核狀態"), Label).Text
                Dim Update1 As string ="UPDATE 收支備查簿 SET " & _
                    "鎖定 = 'False',過審 = 'False',回覆 = 'False',主計室簽核=NULL,送交主計室日期 = NULL,駁回原因 = N'拿回' " & _
                    "WHERE id = '" & id & "'"'修改鎖定
                data.UpdateCommAnd = Update1
                data.Update()
                data.insertCommand = _
                    "INSERT INTO 日誌 " & _
                    "(id, 動作,日期,日期2) " & _
                    "VALUES " & _
                    "(N'" & id & "', N'拿回', N'" & DateTime.now.tostring() & "' , '" & DateTime.now.tostring("yyyy-MM-dd HH:mm:ss") & "')"
                data.insert()
                Label1.Text="資料已收回"
            End If
        Next
        If 作用=False
            Label3.Text="請先選取欲拿回號數之資料"
        Else
            Label3.Text=""
            Me.GridView1.DataBind()
        End If
    End Sub
    Protected Sub 插入_Click(ByVal sender As Object, ByVal e As System.EventArgs)'例外狀況，摘要<>'承上頁' AND 摘要<>'接下頁'
        Update_f()
        Dim 作用 As boolean =False
        Dim ind as int32 = 0
        Dim s頁 As int32 = Me.GridView1.PageIndex+1
        Dim s列 As int32 = 0
        Dim _種類 As String = Me._種類.Text
        For i = 0 To Me.GridView1.Rows.Count - 1
            If CType(Me.GridView1.Rows(i).FindControl("刪除選取"), CheckBox).Checked=True AND i<>0 AND i <> Me.GridView1.Rows.Count - 1
                ind=ind+1
                作用=true
                s列 = CType(Me.GridView1.Rows(i).FindControl("_列"), TextBox).Text
            End If
        Next
        If 作用=False
            Label3.Text="請先選取要插入之資料"
        ElseIf ind=1 AND s列<>0
            data.ConnectionString = con_14
            data.SelectCommAnd = "SELECT * FROM 收支備查簿 WHERE _種類='" & _種類 & "' AND ((_頁='" & s頁 & "' AND _列>='" & s列 & "') OR _頁>'" & s頁 & "') AND ((摘要<>'承上頁' AND 摘要<>'接下頁') OR 摘要 IS NULL) Order by _頁,_列"
            data_dv = data.Select(New DataSourceSelectArguments)
            Dim data_dv2 As Data.DataView
            data.ConnectionString = con_14
            data.SelectCommAnd = "SELECT 主id,id FROM 日誌 WHERE id in (SELECT id FROM 收支備查簿 WHERE _種類='" & _種類 & "' AND ((_頁='" & s頁 & "' AND _列>='" & s列 & "') OR _頁>'" & s頁 & "') AND 摘要<>'承上頁' AND 摘要<>'接下頁')"
            data_dv2 = data.Select(New DataSourceSelectArguments)
            For j=0 to data_dv.count-2
                Dim 頁 As int32 = data_dv(j)("_頁").ToString()
                Dim 列 As int32 = data_dv(j)("_列").ToString()
                ' If j<> data_dv.count-1
                '     Dim Next列 As String = data_dv(j+1)("_列").ToString()
                ' End If
                Dim id As int32 = data_dv(j)("id").ToString()
                ' If j<> data_dv.count-1
                '     Dim Next頁 As String = data_dv(j+1)("_頁").ToString()
                ' End If
                ' Dim n頁 As int32 = Next頁
                Dim n頁 As int32 = data_dv(j+1)("_頁").ToString()
                ' Dim n列 As int32 = Next列
                Dim n列 As int32 = data_dv(j+1)("_列").ToString()
                ' If n列>14
                '     n列=n列-13
                '     n頁=n頁+1
                ' End If
                ' Dim data_dv2 As Data.DataView
                ' data.ConnectionString = con_14
                ' data.SelectCommAnd = "SELECT 主id FROM 日誌 WHERE id = '" & id & "'"
                ' data_dv2 = data.Select(New DataSourceSelectArguments)
                Dim data_dv3 As Data.DataView '10/11日誌更新有問題，如用1466取代1465，之後1467取代1466會把未取代和以取代一起取代掉
                data.ConnectionString = con_14
                data.SelectCommAnd = "SELECT id As 新id FROM 收支備查簿 WHERE _種類='" & _種類 & "' AND _列='" & n列 & "' AND _頁='" & n頁 & "'"
                data_dv3 = data.Select(New DataSourceSelectArguments)
                Dim nid As int32 = nothing
                If data_dv3.count>0
                    nid =data_dv3(0)("新id").ToString()
                End If
                Dim 預支日期 As String = data_dv(j)("預支日期").ToString()
                If 預支日期 <>""
                    預支日期 = 預支日期.substring(0,10)
                End If
                Dim 歸還日期 As String = data_dv(j)("歸還日期").ToString()
                If 歸還日期 <>""
                    歸還日期 = 歸還日期.substring(0,10)
                End If
                ' label1.Text=label1.Text & "新id " & id & " "  & 頁 & "頁:" & 列 & "列;"
                ' label1.Text=label1.Text & "原id " & nid & " "  & n頁 & "頁:" & n列 & "列<BR>"
                
                if j=0
                    data.UpdateCommAnd = "UPDATE 收支備查簿 SET " & _
                       "單位別 = NULL,承辦人 = NULL,月 = NULL,日 = NULL,科目 = NULL,科目2 = NULL,摘要 = NULL,姓名 = NULL,商號 = NULL,經手人 = NULL,種類 = NULL,號數 = NULL," & _
                       "收入 = 0,支出 = 0,備註 = NULL,取號 = '0',送交主計室日期 = NULL,回覆 = 'False',鎖定 = 'False',過審 = 'False',送出 = 'False',預支日期 = NULL,歸還日期 = NULL,主計室簽核=NULL,駁回原因 = NULL " & _
                       "WHERE id = '" & id & "'"
                    data.Update()
                End if
                ' label1.Text=label1.Text & 頁 & "頁:" & 列 & "列;取代" & n頁 & "頁:" & n列 & "列<BR>"
                Dim Update1 As string ="UPDATE 收支備查簿 SET " & _
                "單位別 = NULLIF(N'" & data_dv(j)("單位別").ToString() & "', ''), " & _
                "承辦人 = NULLIF(N'" & data_dv(j)("承辦人").ToString() & "', ''), " & _
                "月 = NULLIF(N'" & data_dv(j)("月").ToString() & "', ''), " & _
                "日 = NULLIF(N'" & data_dv(j)("日").ToString() & "', ''), " & _
                "科目 = NULLIF(N'" & data_dv(j)("科目").ToString() & "', ''), " & _
                "科目2 = NULLIF(N'" & data_dv(j)("科目2").ToString() & "', ''), " & _
                "摘要 = NULLIF(N'" & data_dv(j)("摘要").ToString() & "', ''), " & _
                "姓名 = NULLIF(N'" & data_dv(j)("姓名").ToString() & "', ''), " & _
                "商號 = NULLIF(N'" & data_dv(j)("商號").ToString() & "', ''), " & _
                "經手人 = NULLIF(N'" & data_dv(j)("經手人").ToString() & "', ''), " & _
                "種類 = NULLIF(N'" & data_dv(j)("種類").ToString() & "', ''), " & _
                "號數 = NULLIF(N'" & data_dv(j)("號數").ToString() & "', ''), " & _
                "收入 = NULLIF('" & data_dv(j)("收入").ToString() & "', ''), " & _
                "支出 = NULLIF('" & data_dv(j)("支出").ToString() & "', ''), " & _
                "備註 = NULLIF(N'" & data_dv(j)("備註").ToString() & "', ''), " & _
                "取號 = NULLIF(N'" & data_dv(j)("取號").ToString() & "', ''), " & _
                "送交主計室日期 = NULLIF(N'" & data_dv(j)("送交主計室日期").ToString() & "', ''), " & _
                "回覆 = NULLIF(N'" & data_dv(j)("回覆").ToString() & "', ''), " & _
                "鎖定 = NULLIF(N'" & data_dv(j)("鎖定").ToString() & "', ''), " & _
                "過審 = NULLIF(N'" & data_dv(j)("過審").ToString() & "', ''), " & _
                "送出 = NULLIF(N'" & data_dv(j)("送出").ToString() & "', ''), " & _
                "預支日期 = NULLIF('" & 預支日期 & "', '')," & _
                "歸還日期 = NULLIF('" & 歸還日期 & "', '')," & _
                "主計室簽核 = NULLIF(N'" & data_dv(j)("主計室簽核").ToString() & "', ''), " & _
                "駁回原因 = NULLIF(N'" & data_dv(j)("駁回原因").ToString() & "', '') " & _
                "WHERE _列 = '" & n列 &"' AND _頁 = " & n頁 & " AND _種類='" & _種類 & "'"
                data.UpdateCommAnd = Update1
                data.Update()
                For k=0 to data_dv2.count-1 
                    If id=data_dv2(k)("id").ToString()
                        data.UpdateCommAnd = "UPDATE 日誌 SET " & _
                            "id = '" & nid & "' " & _
                            "WHERE 主id = '" & data_dv2(k)("主id").ToString() & "'"
                        data.Update()
                    End If
                Next
            Next
             '重算餘額，只做A，不能動到本月小計、累計至本月
            If (_種類="A") Then
                data.UpdateCommAnd = _
                "WITH CTE AS " & _
                "(SELECT *, " & _
                    "(SELECT TOP 1 (CASE WHEN ISNULL(收入,0) = 0 And ISNULL(支出,0) = 0  THEN 餘額 ELSE 0 END) FROM 收支備查簿 WHERE _種類 = '" & _種類 & "' ORDER BY id) " & _
                    "+ " & _
                    "(SUM(" & _
	                "(CASE WHEN ISNULL(摘要,'')<>'本月小計' And ISNULL(摘要,'')<>'累計至本月' THEN ISNULL(收入,0)ELSE 0 END )" & _
                    "-" & _ 
	                "(CASE WHEN ISNULL(摘要,'')<>'本月小計' And ISNULL(摘要,'')<>'累計至本月' THEN ISNULL(支出,0)ELSE 0 END )" & _
	                ") OVER (ORDER BY id " & _
                    "ROWS BETWEEN UNBOUNDED PRECEDING And CURRENT ROW))" & _
                    "AS RunningTotal " & _
                "FROM 收支備查簿 WHERE _種類 = '" & _種類 & "') " & _
                "UPDATE CTE SET 餘額 = RunningTotal"
                data.Update()
            End If
            Label3.Text=""
            Me.GridView1.DataBind()
        Else
            Label3.Text="只能選一筆資料"
        End If
    End Sub
    Protected Sub 交換_Click(ByVal sender As Object, ByVal e As System.EventArgs)'交換
        Dim ind as int32 = 0
        Dim id1 As String
        Dim id2 As String
        Dim 單位別 As String
        Dim 承辦人 As String
        Dim 月 As String
        Dim 日 As String
        Dim 科目 As String
        Dim 科目2 As String
        Dim 摘要 As String
        Dim 姓名 As String
        Dim 商號 As String
        Dim 經手人 As String
        Dim 種類 As String
        Dim 號數 As String
        Dim 收入 As String
        Dim 支出 As String
        Dim 備註 As String
        Dim 取號 As String
        Dim 送交主計室日期 As String
        Dim 回覆 As String
        Dim 鎖定 As String
        Dim 過審 As String
        Dim 送出 As String
        Dim 預支日期 As String
        Dim 歸還日期 As String
        Dim 主計室簽核 As String
        Dim 駁回原因 As String
        Dim data_dv2 As Data.DataView
        For i = 0 To Me.GridView1.Rows.Count - 1
            If CType(Me.GridView1.Rows(i).FindControl("刪除選取"), CheckBox).Checked=True
                ind=ind+1
                If ind=1
                    id1=CType(Me.GridView1.Rows(i).FindControl("id"), TextBox).Text
                    data.ConnectionString = con_14
                    data.SelectCommAnd = "SELECT * FROM 收支備查簿 WHERE  id=" &id1
                    data_dv = data.Select(New DataSourceSelectArguments)
                    單位別 = data_dv(0)("單位別").ToString()
                    承辦人 = data_dv(0)("承辦人").ToString()
                    月 = data_dv(0)("月").ToString()
                    日 = data_dv(0)("日").ToString()
                    科目 = data_dv(0)("科目").ToString()
                    科目2 = data_dv(0)("科目2").ToString()
                    摘要 = data_dv(0)("摘要").ToString()
                    姓名 = data_dv(0)("姓名").ToString()
                    商號 = data_dv(0)("商號").ToString()
                    經手人 = data_dv(0)("經手人").ToString()
                    種類 = data_dv(0)("種類").ToString()
                    號數 = data_dv(0)("號數").ToString()
                    收入 = data_dv(0)("收入").ToString()
                    支出 = data_dv(0)("支出").ToString()
                    備註 = data_dv(0)("備註").ToString()
                    取號 = data_dv(0)("取號").ToString()
                    送交主計室日期 = data_dv(0)("送交主計室日期").ToString()
                    回覆 = data_dv(0)("回覆").ToString()
                    鎖定 = data_dv(0)("鎖定").ToString()
                    過審 = data_dv(0)("過審").ToString()
                    送出 = data_dv(0)("送出").ToString()
                    預支日期 = data_dv(0)("預支日期").ToString()
                    歸還日期 = data_dv(0)("歸還日期").ToString()
                    主計室簽核 = data_dv(0)("主計室簽核").ToString()
                    駁回原因 = data_dv(0)("駁回原因").ToString()
                End If
                If ind=2
                    id2=CType(Me.GridView1.Rows(i).FindControl("id"), TextBox).Text
                    data.ConnectionString = con_14
                    data.SelectCommAnd = "SELECT * FROM 收支備查簿 WHERE  id=" &id2
                    data_dv2 = data.Select(New DataSourceSelectArguments)
                End If
            End If
        Next
        If ind=0
            Label3.Text="請先選取需要交換的兩筆資料，在按下[交換]按鈕"
        ElseIf ind=1
            Label3.Text="請選取兩筆資料"
        ElseIf ind=2
            Dim Update1 As string ="UPDATE 收支備查簿 SET " & _
            "單位別 = NULLIF(N'" & data_dv2(0)("單位別").ToString() & "', ''), " & _
            "承辦人 = NULLIF(N'" & data_dv2(0)("承辦人").ToString() & "', ''), " & _
            "月 = NULLIF(N'" & data_dv2(0)("月").ToString() & "', ''), " & _
            "日 = NULLIF(N'" & data_dv2(0)("日").ToString() & "', ''), " & _
            "科目 = NULLIF(N'" & data_dv2(0)("科目").ToString() & "', ''), " & _
            "科目2 = NULLIF(N'" & data_dv2(0)("科目2").ToString() & "', ''), " & _
            "摘要 = NULLIF(N'" & data_dv2(0)("摘要").ToString() & "', ''), " & _
            "姓名 = NULLIF(N'" & data_dv2(0)("姓名").ToString() & "', ''), " & _
            "商號 = NULLIF(N'" & data_dv2(0)("商號").ToString() & "', ''), " & _
            "經手人 = NULLIF(N'" & data_dv2(0)("經手人").ToString() & "', ''), " & _
            "種類 = NULLIF(N'" & data_dv2(0)("種類").ToString() & "', ''), " & _
            "號數 = NULLIF(N'" & data_dv2(0)("號數").ToString() & "', ''), " & _
            "收入 = NULLIF('" & data_dv2(0)("收入").ToString() & "', ''), " & _
            "支出 = NULLIF('" & data_dv2(0)("支出").ToString() & "', ''), " & _
            "備註 = NULLIF(N'" & data_dv2(0)("備註").ToString() & "', ''), " & _
            "取號 = NULLIF(N'" & data_dv2(0)("取號").ToString() & "', ''), " & _
            "送交主計室日期 = NULLIF(N'" & data_dv2(0)("送交主計室日期").ToString() & "', ''), " & _
            "回覆 = NULLIF(N'" & data_dv2(0)("回覆").ToString() & "', ''), " & _
            "鎖定 = NULLIF(N'" & data_dv2(0)("鎖定").ToString() & "', ''), " & _
            "過審 = NULLIF(N'" & data_dv2(0)("過審").ToString() & "', ''), " & _
            "送出 = NULLIF(N'" & data_dv2(0)("送出").ToString() & "', ''), " & _
            "預支日期 = NULLIF(N'" & data_dv2(0)("預支日期").ToString() & "', '')," & _
            "歸還日期 = NULLIF(N'" & data_dv2(0)("歸還日期").ToString() & "', '')," & _
            "主計室簽核 = NULLIF(N'" & data_dv2(0)("主計室簽核").ToString() & "', ''), " & _
            "駁回原因 = NULLIF(N'" & data_dv2(0)("駁回原因").ToString() & "', '') " & _
            "WHERE id = '" & id1 &"'"
            data.UpdateCommAnd = Update1
            data.Update()
            '交換日誌、修改資料
            Update1 ="UPDATE 日誌 SET " & _
            "id=(CASE WHEN id='" & id1 & "' THEN N'" & id2 & "'ELSE N'" & id1 & "' END) " & _
            "WHERE id = '" & id1 & "' OR id = '" & id2 & "'"
            data.UpdateCommAnd = Update1
            data.Update()
            Update1 ="UPDATE 修改資料 SET " & _
            "id_收=(CASE WHEN id_收='" & id1 & "' THEN N'" & id2 & "'ELSE N'" & id1 & "' END) " & _
            "WHERE id_收 = '" & id1 & "' OR id_收 = '" & id2 & "'"
            data.UpdateCommAnd = Update1
            data.Update()
            If data_dv2(0)("送出").ToString()="True"
                Dim sql As string
                Dim date1 As string = DateTime.now.tostring()
                Dim date2 As string = DateTime.now.tostring("yyyy-MM-dd HH:mm:ss")
                Dim insert1 As string = _
                    "INSERT INTO 修改資料 " & _
                    "(id_收," & _
                    "單位別,單位別_改," & _
                    "承辦人,承辦人_改," & _
                    "月    ,月_改    ," & _
                    "日    ,日_改    ," & _
                    "科目  ,科目_改  ," & _
                    "摘要  ,摘要_改  ," & _
                    "姓名  ,姓名_改  ," & _
                    "商號  ,商號_改  ," & _
                    "經手人,經手人_改," & _
                    "種類  ,種類_改  ," & _
                    "號數  ,號數_改  ," & _
                    "收入  ,收入_改  ," & _
                    "支出  ,支出_改  ," & _
                    "備註  ,備註_改  ," & _
                    "date) " & _
                    "VALUES " & _
                    "(" & id1 & "," & _
                    "NULLIF(N'" & 單位別 & "', ''),              NULLIF(N'" & data_dv2(0)("單位別").ToString() & "', ''),"                                       & _
                    "NULLIF(N'" & 承辦人 & "', ''),              NULLIF(N'" & data_dv2(0)("承辦人").ToString() & "', ''),"                                       & _
                    "NULLIF('"  & 月     & "', ''),              NULLIF('"  & data_dv2(0)("月").ToString()     & "', ''),"                                       & _
                    "NULLIF('"  & 日     & "', ''),              NULLIF('"  & data_dv2(0)("日").ToString()     & "', ''),"                                       & _
                    "NULLIF(N'" & 科目   & ";" & 科目2 & "', ';'),NULLIF(N'" & data_dv2(0)("科目").ToString() & ";" & data_dv2(0)("科目2").ToString() & "', ';')," & _
                    "NULLIF(N'" & 摘要   & "', ''),              NULLIF(N'" & data_dv2(0)("摘要").ToString()    & "', ''),"                                      & _
                    "NULLIF(N'" & 姓名   & "', ''),              NULLIF(N'" & data_dv2(0)("姓名").ToString()    & "', ''),"                                      & _
                    "NULLIF(N'" & 商號   & "', ''),              NULLIF(N'" & data_dv2(0)("商號").ToString()    & "', ''),"                                      & _
                    "NULLIF(N'" & 經手人 & "', ''),              NULLIF(N'" & data_dv2(0)("經手人").ToString()  & "', ''),"                                      & _
                    "NULLIF(N'" & 種類   & "', ''),              NULLIF(N'" & data_dv2(0)("種類").ToString()    & "', ''),"                                      & _
                    "NULLIF('"  & 號數   & "', ''),              NULLIF('"  & data_dv2(0)("號數").ToString()    & "', ''),"                                      & _
                    "NULLIF('"  & 收入   & "', ''),              NULLIF('"  & data_dv2(0)("收入").ToString()    & "', ''),"                                      & _
                    "NULLIF('"  & 支出   & "', ''),              NULLIF('"  & data_dv2(0)("支出").ToString()    & "', ''),"                                      & _
                    "NULLIF(N'" & 備註   & "', ''),              NULLIF(N'" & data_dv2(0)("備註").ToString()                                                     & _
                    "', ''),NULLIF(N'" & date1 &  "', ''))"
                    data.insertCommAnd = insert1
                    data.insert()
                    '查詢修改資料表的ID，並把他給日誌
                    sql = "SELECT * FROM [修改資料] where id_收=" & id1 & "And date = '" & date1 & "'"
                    data.SelectCommAnd = sql
                    data_dv = data.Select(New DataSourceSelectArguments)
                    Dim id_2 As string =data_dv(0)("id").ToString()
                    data.insertCommAnd = _
                    "INSERT INTO 日誌 " & _
                    "(id, 動作, 命令,日期,日期2) " & _
                    "VALUES " & _
                    "(N'" & id1 & "', N'修改_交換', N'修改資料id="& id_2 &"', N'" & date1 & "', '" & date2 & "')"
                    data.insert()
            End If
            Dim Update2 As string ="UPDATE 收支備查簿 SET " & _
            "單位別 = NULLIF(N'" & 單位別 & "', ''), " & _
            "承辦人 = NULLIF(N'" & 承辦人 & "', ''), " & _
            "月 = NULLIF(N'" & 月 & "', ''), " & _
            "日 = NULLIF(N'" & 日 & "', ''), " & _
            "科目 = NULLIF(N'" & 科目 & "', ''), " & _
            "科目2 = NULLIF(N'" & 科目2 & "', ''), " & _
            "摘要 = NULLIF(N'" & 摘要 & "', ''), " & _
            "姓名 = NULLIF(N'" & 姓名 & "', ''), " & _
            "商號 = NULLIF(N'" & 商號 & "', ''), " & _
            "經手人 = NULLIF(N'" & 經手人 & "', ''), " & _
            "種類 = NULLIF(N'" & 種類 & "', ''), " & _
            "號數 = NULLIF(N'" & 號數 & "', ''), " & _
            "收入 = NULLIF('" & 收入 & "', ''), " & _
            "支出 = NULLIF('" & 支出 & "', ''), " & _
            "備註 = NULLIF(N'" & 備註 & "', ''), " & _
            "取號 = NULLIF(N'" & 取號 & "', ''), " & _
            "送交主計室日期 = NULLIF(N'" & 送交主計室日期 & "', ''), " & _
            "回覆 = NULLIF(N'" & 回覆 & "', ''), " & _
            "鎖定 = NULLIF(N'" & 鎖定 & "', ''), " & _
            "過審 = NULLIF(N'" & 過審 & "', ''), " & _
            "送出 = NULLIF(N'" & 送出 & "', ''), " & _
            "預支日期 = NULLIF(N'" & 預支日期 & "', '')," & _
            "歸還日期 = NULLIF(N'" & 歸還日期 & "', '')," & _
            "主計室簽核 = NULLIF(N'" & 主計室簽核 & "', ''), " & _
            "駁回原因 = NULLIF(N'" & 駁回原因 & "', '') " & _
            "WHERE id = '" & id2 & "'"
            data.UpdateCommAnd = Update2
            data.Update()
            If 送出="True"
                Dim sql As string
                Dim date1 As string = DateTime.now.tostring()
                Dim date2 As string = DateTime.now.tostring("yyyy-MM-dd HH:mm:ss")
                Dim insert1 As string = _
                    "INSERT INTO 修改資料 " & _
                    "(id_收,單位別,單位別_改,承辦人,承辦人_改,月,月_改,日,日_改,科目,科目_改,摘要,摘要_改,姓名,姓名_改,商號,商號_改,經手人,經手人_改,種類,種類_改,號數,號數_改,收入,收入_改,支出,支出_改,備註,備註_改,date) " & _
                    "VALUES " & _
                    "(" & id2 & ",NULLIF(N'" & _
                    data_dv2(0)("單位別").ToString() & "',''),NULLIF(N'" & 單位別 & "', ''),NULLIF(N'" & _
                    data_dv2(0)("承辦人").ToString()  & "',''),NULLIF(N'" & 承辦人  & "', ''), NULLIF('" & _
                    data_dv2(0)("月").ToString()  & "',''), NULLIF('" & 月  & "',''), NULLIF('" & _
                    data_dv2(0)("日").ToString()  & "',''), NULLIF('" & 日  & "',''),NULLIF(N'" & _
                    data_dv2(0)("科目").ToString() & ";" & data_dv2(0)("科目2").ToString()  & "',';'),NULLIF(N'" & 科目 & ";" & 科目2  & "', ';'),NULLIF(N'" & _
                    data_dv2(0)("摘要").ToString()  & "',''),NULLIF(N'" & 摘要  & "', ''),NULLIF(N'" & _
                    data_dv2(0)("姓名").ToString()  & "',''),NULLIF(N'" & 姓名  & "', ''),NULLIF(N'" & _ 
                    data_dv2(0)("商號").ToString()  & "',''),NULLIF(N'" & 商號  & "', ''),NULLIF(N'" & _
                    data_dv2(0)("經手人").ToString() & "',''),NULLIF(N'" & 經手人 & "', ''),NULLIF(N'" & _
                    data_dv2(0)("種類").ToString() & "',''),NULLIF(N'" & 種類  & "', ''), NULLIF('" & _
                    data_dv2(0)("號數").ToString() & "',''), NULLIF('" & 號數  & "',''), NULLIF('" & _
                    data_dv2(0)("收入").ToString() & "',''), NULLIF('" & 收入 & "',''), NULLIF('" & _
                    data_dv2(0)("支出").ToString()  & "',''), NULLIF('" & 支出  & "',''),NULLIF(N'" & _
                    data_dv2(0)("備註").ToString()  & "',''),NULLIF(N'" & 備註  & "', ''),NULLIF(N'" & date1 &  "',''))"
                    data.insertCommAnd = insert1
                    data.insert()
                    '查詢修改資料表的ID，並把他給日誌
                    sql = "SELECT * FROM [修改資料] where id_收=" & id2 & "And date = '" & date1 & "'"
                    data.SelectCommAnd = sql
                    data_dv = data.Select(New DataSourceSelectArguments)
                    Dim id_2 As string =data_dv(0)("id").ToString()
                    data.insertCommAnd = _
                    "INSERT INTO 日誌 " & _
                    "(id, 動作, 命令,日期,日期2) " & _
                    "VALUES " & _
                    "(N'" & id2 & "', N'修改_交換', N'修改資料id="& id_2 &"', N'" & date1 & "', '" & date2 & "')"
                    data.insert()
            End If
        ElseIf ind>2
            Label3.Text="請勿選擇三筆以上的資料"
        End If
        Me.GridView1.DataBind()
        Update(sender,e)
        Label1.Text="交換成功"
    End Sub
    Protected Sub 取號_Click(ByVal sender As Object, ByVal e As System.EventArgs)'取號
        '先使用存檔,則無法偵測checkbox
        Update_f()
        Dim 年 As String = Me.年.Text
        Dim _種類 As String = Me._種類.Text
        Dim 單位別 As String 
        Dim 承辦人 As String
        Dim 號數 As String
        Dim 種類 As String 
        Dim 備註 As String 
        Dim Id_Sum As String = ""
        Dim b As Byte = 0 '是否有選取方塊
        Dim h As Byte = 0 '判斷所有位址任一不為空
        For i=0 to Me.GridView1.Rows.Count - 1
            Dim id As String = CType(Me.GridView1.Rows(i).FindControl("id"), TextBox).Text
            號數 = CType(Me.GridView1.Rows(i).FindControl("號數"), TextBox).Text
            If CType(Me.GridView1.Rows(i).FindControl("取號勾選"), CheckBox).Checked=True '取號被勾選
                If 號數=Nothing And CType(Me.GridView1.Rows(i).FindControl("取號勾選"), CheckBox).Checked=true And h=0 '號數無內容
                    b=1'要取號
                    種類=_種類
                    If Id_Sum=""
                        Id_Sum=id
                    Else
                        Id_Sum=Id_Sum & " OR id="& id
                    End If 
                ElseIf 號數<>Nothing And CType(Me.GridView1.Rows(i).FindControl("取號勾選"), CheckBox).Checked=true '號數有內容
                    h=1
                    For k = 0 to Me.GridView1.Rows.Count - 1
                        If CType(Me.GridView1.Rows(k).FindControl("取號勾選"), CheckBox).Checked=true '任一當前有號數之值取代全部有勾選號數
                            Dim id2 As String = CType(Me.GridView1.Rows(k).FindControl("id"), TextBox).Text
                            DIM sql As string = "SELECT * FROM [收支備查簿] where _種類='" & me._種類.text & "'And 號數=" & 號數 & " And 取號='True'"
                            data.SelectCommAnd = sql
                            data_dv = data.Select(New DataSourceSelectArguments)
                            If data_dv.count <> 0
                                For i2 = 0 to data_dv.count-1
                                    單位別 = data_dv(i2)("單位別").ToString()
                                    承辦人 = data_dv(i2)("承辦人").ToString()
                                    號數 = data_dv(i2)("號數").ToString()
                                    種類 = data_dv(i2)("種類").ToString()
                                    備註 = data_dv(i2)("備註").ToString()
                                    data.UpdateCommAnd = "UPDATE 收支備查簿 SET " & _
                                        "單位別 = NULLIF(N'" & 單位別 & "', ''), 承辦人 = NULLIF(N'" & 承辦人 & "', ''), 號數 = NULLIF(N'" & 號數 & "', ''), 種類 = NULLIF(N'" & 種類 & "', ''), 備註 = NULLIF(N'" & 備註 & "', '') WHERE 鎖定 = 'False' And id=" & id2
                                    data.Update()
                                    Me.GridView1.DataBind()
                                Next
                            End If 
                        End If 
                    Next
                End If 
            End If 
        Next
        If b=1 And h=0 '全部勾選號數之值都無內容
            data.UpdateCommAnd = "UPDATE 收支備查簿 SET " & _
                "種類 = NULLIF(N'" & 種類 & "', ''), 號數 = (SELECT ISNULL(MAX(號數) + 1, 1) FROM 收支備查簿 WHERE 年 = " & 年 & " And _種類 = '" & 種類 & "')" & _
                "WHERE 鎖定 = 'False' And id=" & Id_Sum
            data.Update()
            Me.GridView1.DataBind()
        End If 
        If b=0
            Label3.Text="請先勾取[取號列]中想取號的資料"
        Else
            Label3.Text=""
        End If 
    End Sub
    Protected Sub 新增科目_Click(ByVal sender As Object, ByVal e As System.EventArgs)'新增科目
        Dim 作用 As boolean = False
        For i=0 to Me.GridView1.Rows.Count - 1
            If CType(Me.GridView1.Rows(i).FindControl("刪除選取"), CheckBox).Checked=True
                CType(Me.GridView1.Rows(i).FindControl("科目2"), DropDownList).Visible=True
                作用=True
            End If 
        Next
        If 作用=False
            Label3.Text="請先選取欲新增科目之資料"
        Else
            Label3.Text=""
            'Me.GridView1.DataBind() 不可更新
            Label1.Text="新增科目成功"
        End If
    End Sub
    Protected Sub SendToDirector(ByVal sender As Object, ByVal e As System.EventArgs)'送交主任
        Update_f()
        Dim 作用 As boolean = False
        For i=0 to Me.GridView1.Rows.Count - 1
            Dim id As String = CType(Me.GridView1.Rows(i).FindControl("id"), TextBox).Text
            Dim 號數 As String = CType(Me.GridView1.Rows(i).FindControl("號數"), TextBox).Text
            Dim 前號數 As String 
            Dim 動作 As String 
            If CType(Me.GridView1.Rows(i).FindControl("審核"), CheckBox).Checked=True And 號數<>前號數
                作用=True
                data.ConnectionString = con_14
                data.SelectCommAnd = "SELECT * FROM 收支備查簿 " & _
                    "Where 過審 = 'False' AND 取號='False' And 鎖定 = 'False' AND 號數 = '" & 號數 & "' And _種類 = '" & Me._種類.Text & "'"
                data_dv = data.Select(New DataSourceSelectArguments)
                If data_dv.count <> 0
                    For j=0 to data_dv.count-1
                        id=data_dv(j)("id").ToString()
                        If data_dv(j)("送出").ToString()="False"
                            動作="新增"
                        Else
                            動作="送審"
                        End If 
                        '可能出現情況，送審後，補資料(判斷是否有鎖定)
                        '更新日誌
                        Dim data_dv2 As Data.DataView
                        data.SelectCommAnd = "Select id From 收支備查簿 " & _
                            "Where 過審 = 'False' AND 取號='False' And 鎖定 = 'False' AND 號數 = '" & 號數 & "' And _種類 = '" & Me._種類.Text & "'"
                        data_dv2 = data.Select(New DataSourceSelectArguments)
                        For k=0 to data_dv2.count-1
                            data.insertCommAnd = _
                                "INSERT INTO 日誌 " & _
                                "(id, 動作,日期,日期2) " & _
                                "VALUES " & _
                                "(N'" & data_dv2(k)("id").ToString() & "', N'" & 動作 & "' , N'" & DateTime.now.tostring() & "' , '" & DateTime.now.tostring("yyyy-MM-dd HH:mm:ss") & "')"
                            data.insert()
                        Next
                        '送交資料
                        data.UpdateCommAnd = "UPDATE 收支備查簿 SET " & _
                            "鎖定 = 'True',送出 = 'True', 主計室簽核 = NULL  " & _
                            "Where 過審 = 'False' AND 取號='False' And 鎖定 = 'False' AND 號數 = '" & 號數 & "' And _種類 = '" & Me._種類.Text & "'"
                        data.Update()
                    Next
                    前號數=號數
                End If 
            End If
        Next
        If 作用=False
            Label3.Text="請先選取號數再送審資料"
        Else
            Label1.Text="成功送審"
            Me.GridView1.DataBind()
        End If
    End Sub
    Protected Sub ReSetRow(ByVal sender As Object, ByVal e As System.EventArgs)'先存檔，因直接下達SQL，所以要先重新載入GridView頁面
        update_f()
        Dim 作用 As boolean = False
        For i=0 to Me.GridView1.Rows.Count - 1
            Dim id As String = CType(Me.GridView1.Rows(i).FindControl("id"), TextBox).Text
            Dim 動作 As String 
            If CType(Me.GridView1.Rows(i).FindControl("刪除選取"), CheckBox).Checked=True
                作用=True 
                data.ConnectionString = con_14
                Dim sql As string = "SELECT * FROM 收支備查簿 where id=" & id
                data.SelectCommAnd = sql
                data_dv = data.Select(New DataSourceSelectArguments)
                If data_dv(0)("鎖定").ToString()="False"
                    '---
                    Dim 年 As String = Me.年.Text
                    Dim _種類 As String = Me._種類.Text
                    Dim 單位別 As String = CType(Me.GridView1.Rows(i).FindControl("單位別"),DropDownList).Text
                    Dim 承辦人 As String = CType(Me.GridView1.Rows(i).FindControl("承辦人"), DropDownList).Text
                    Dim 月 As String = CType(Me.GridView1.Rows(i).FindControl("月"), DropDownList).Text
                    Dim 日 As String = CType(Me.GridView1.Rows(i).FindControl("日"), DropDownList).Text
                    Dim 科目 As String = CType(Me.GridView1.Rows(i).FindControl("科目"), DropDownList).Text
                    Dim 科目2 As String = CType(Me.GridView1.Rows(i).FindControl("科目2"), DropDownList).Text
                    Dim 摘要 As String = CType(Me.GridView1.Rows(i).FindControl("摘要"), TextBox).Text
                    Dim 姓名 As String = CType(Me.GridView1.Rows(i).FindControl("姓名"), ImageButton).ImageUrl
                    Dim 商號 As String = CType(Me.GridView1.Rows(i).FindControl("商號"), TextBox).Text
                    Dim 經手人 As String = CType(Me.GridView1.Rows(i).FindControl("經手人"), ImageButton).ImageUrl
                    Dim 種類 As String = CType(Me.GridView1.Rows(i).FindControl("種類"), TextBox).Text
                    Dim 號數 As String = CType(Me.GridView1.Rows(i).FindControl("號數"), TextBox).Text
                    Dim 收入 As String = CType(Me.GridView1.Rows(i).FindControl("收入"), TextBox).Text
                    收入=收入.Replace(",", "").Replace("N", "").Replace("T", "").Replace("$", "")
                    Dim 支出 As String = CType(Me.GridView1.Rows(i).FindControl("支出"), TextBox).Text
                    支出=支出.Replace(",", "").Replace("N", "").Replace("T", "").Replace("$", "")
                    Dim 備註 As String = CType(Me.GridView1.Rows(i).FindControl("備註"), TextBox).Text
                    If data_dv(0)("送出").ToString()="True"
                        Dim date1 as string = DateTime.now.tostring()
                        Dim date2 as string = DateTime.now.tostring("yyyy-MM-dd HH:mm:ss")
                        '新增刪除資料表
                        data.insertCommand = _
                        "INSERT INTO 刪除資料表 " & _
                        "(id_收, _種類, 單位別, 承辦人,年 ," & _
                        " 月, 日, 科目, 科目2, 摘要, 姓名, 商號," & _
                        " 經手人, 種類, 號數, 收入, 支出, 備註 ,date ) " & _
                        "VALUES " & _
                        "(" & id & "," & _
                        "NULLIF(N'" & _種類  & "','')," & _
                        "NULLIF(N'" & 單位別 & "','')," & _
                        "NULLIF(N'" & 承辦人 & "','')," & _
                        "NULLIF('"  & 年     & "','')," & _
                        "NULLIF('"  & 月     & "','')," & _
                        "NULLIF('"  & 日     & "','')," & _
                        "NULLIF(N'" & 科目   & "','')," & _
                        "NULLIF(N'" & 科目2  & "','')," & _
                        "NULLIF(N'" & 摘要   & "','')," & _
                        "NULLIF(N'" & 姓名   & "','')," & _
                        "NULLIF(N'" & 商號   & "','')," & _
                        "NULLIF(N'" & 經手人 & "','')," & _
                        "NULLIF(N'" & 種類   & "','')," & _
                        "NULLIF('"  & 號數   & "','')," & _
                        "NULLIF('"  & 收入   & "','')," & _
                        "NULLIF('"  & 支出   & "','')," & _
                        "NULLIF(N'" & 備註   & "','')," & _
                        "NULLIF(N'" & date1  & "',''))"
                        data.insert()
                        '新增日誌
                        sql = "SELECT * FROM [刪除資料表] where id_收=" & id & "and date = '" & date1 & "'"
                        data.SelectCommand = sql
                        data_dv = data.Select(New DataSourceSelectArguments)
                        Dim id_2 as string =data_dv(0)("id").ToString()                    
                        data.insertCommand = _
                        "INSERT INTO 日誌 " & _
                        "(id, 動作, 命令,日期,日期2) " & _
                        "VALUES " & _
                        "(N'" & id & "', N'刪除', N'刪除資料id=" & id_2 & "' , N'" & date1 & "' , '" & date2 & "')"
                        data.insert()
                    End If
                    '---
                    '重置收支備查簿
                    data.UpdateCommAnd = "UPDATE 收支備查簿 SET " & _
                        "單位別 = NULL,承辦人 = NULL,月 = NULL,日 = NULL,科目 = NULL,科目2 = NULL,摘要 = NULL,姓名 = NULL,商號 = NULL,經手人 = NULL,種類 = NULL,號數 = NULL," & _
                        "收入 = 0,支出 = 0,備註 = NULL,取號 = '0',送交主計室日期 = NULL,回覆 = 'False',鎖定 = 'False',過審 = 'False',送出 = 'False',預支日期 = NULL,歸還日期 = NULL,主計室簽核=NULL,駁回原因 = NULL  " & _
                        "WHERE 鎖定 = 'False' And id = " & id
                    data.Update()
                Else
                    Label3.Text="無法刪除已送審之資料"
                End If
            End If 
        Next
        If 作用=False
            Label3.Text="請先選取欲刪除之資料"
        Else
            Me.GridView1.DataBind()'先更新至刪除之畫面
            Update(sender,e)'重新計算
            Label1.Text="已刪除該列"
        End If
    End Sub
    Protected Sub DeleteRow(ByVal sender As Object, ByVal e As System.EventArgs)'先存檔，因直接下達SQL，所以要先重新載入GridView頁面
        update_f()
        Dim 作用 As boolean = False
        Dim ind as int32 = 0
        Dim s頁 As int32 = Me.GridView1.PageIndex+1
        Dim s列 As int32 = 0
        Dim _種類 As String = Me._種類.Text
        For i = 0 To Me.GridView1.Rows.Count - 1
            If CType(Me.GridView1.Rows(i).FindControl("刪除選取"), CheckBox).Checked=True AND i<>0 AND i <> Me.GridView1.Rows.Count - 1
                ind=ind+1
                作用=true
                s列 = CType(Me.GridView1.Rows(i).FindControl("_列"), TextBox).Text
            End If
        Next
        If 作用=False
            Label3.Text="請先選取要刪除之資料"
        ElseIf ind=1 AND s列<>0
            data.ConnectionString = con_14
            data.SelectCommAnd = "SELECT * FROM 收支備查簿 WHERE _種類='" & _種類 & "' AND ((_頁='" & s頁 & "' AND _列>='" & s列 & "') OR _頁>'" & s頁 & "') AND ((摘要<>'承上頁' AND 摘要<>'接下頁') OR 摘要 IS NULL) Order by _頁 Desc,_列 Desc"
            data_dv = data.Select(New DataSourceSelectArguments)
            Dim data_dv2 As Data.DataView
            data.ConnectionString = con_14
            data.SelectCommAnd = "SELECT 主id,id FROM 日誌 WHERE id in (SELECT id FROM 收支備查簿 WHERE _種類='" & _種類 & "' AND ((_頁='" & s頁 & "' AND _列>='" & s列 & "') OR _頁>'" & s頁 & "') AND 摘要<>'承上頁' AND 摘要<>'接下頁')"
            data_dv2 = data.Select(New DataSourceSelectArguments)
            For j=0 to data_dv.count-1
                Dim 頁 As int32 = data_dv(j)("_頁").ToString()
                Dim 列 As int32 = data_dv(j)("_列").ToString()
                ' If j<> data_dv.count-1
                '     Dim Next列 As String = data_dv(j+1)("_列").ToString()
                ' End If
                Dim id As int32 = data_dv(j)("id").ToString()
                ' If j<> data_dv.count-1
                '     Dim Next頁 As String = data_dv(j+1)("_頁").ToString()
                ' End If
                ' Dim n頁 As int32 = Next頁
                Dim p頁 As int32
                Dim p列 As int32
                If j<>data_dv.count-1 
                    p頁 = data_dv(j+1)("_頁").ToString()
                    ' Dim n列 As int32 = Next列
                    p列 = data_dv(j+1)("_列").ToString()
                End If
                '---
                If j=data_dv.count-1 AND data_dv(j)("送出").ToString()="True"
                    Dim 年 As String = Me.年.Text
                    Dim 單位別 As String = data_dv(j)("單位別").ToString()
                    Dim 承辦人 As String = data_dv(j)("承辦人").ToString()
                    Dim 月 As String = data_dv(j)("月").ToString()
                    Dim 日 As String = data_dv(j)("日").ToString()
                    Dim 科目 As String = data_dv(j)("科目").ToString()
                    Dim 科目2 As String = data_dv(j)("科目2").ToString()
                    Dim 摘要 As String = data_dv(j)("摘要").ToString()
                    Dim 姓名 As String = data_dv(j)("姓名").ToString()
                    Dim 商號 As String = data_dv(j)("商號").ToString()
                    Dim 經手人 As String = data_dv(j)("經手人").ToString()
                    Dim 種類 As String = data_dv(j)("種類").ToString()
                    Dim 號數 As String = data_dv(j)("號數").ToString()
                    Dim 收入 As String = data_dv(j)("收入").ToString()
                    Dim 支出 As String = data_dv(j)("支出").ToString()
                    Dim 備註 As String = data_dv(j)("備註").ToString()
                    Dim date1 as string = DateTime.now.tostring()
                    Dim date2 as string = DateTime.now.tostring("yyyy-MM-dd HH:mm:ss")
                    '新增刪除資料表
                    data.insertCommand = _
                    "INSERT INTO 刪除資料表 " & _
                    "(id_收, _種類, 單位別, 承辦人,年 ," & _
                    " 月, 日, 科目, 科目2, 摘要, 姓名, 商號," & _
                    " 經手人, 種類, 號數, 收入, 支出, 備註 ,date ) " & _
                    "VALUES " & _
                    "(" & id & "," & _
                    "NULLIF(N'" & _種類  & "','')," & _
                    "NULLIF(N'" & 單位別 & "','')," & _
                    "NULLIF(N'" & 承辦人 & "','')," & _
                    "NULLIF('"  & 年     & "','')," & _
                    "NULLIF('"  & 月     & "','')," & _
                    "NULLIF('"  & 日     & "','')," & _
                    "NULLIF(N'" & 科目   & "','')," & _
                    "NULLIF(N'" & 科目2  & "','')," & _
                    "NULLIF(N'" & 摘要   & "','')," & _
                    "NULLIF(N'" & 姓名   & "','')," & _
                    "NULLIF(N'" & 商號   & "','')," & _
                    "NULLIF(N'" & 經手人 & "','')," & _
                    "NULLIF(N'" & 種類   & "','')," & _
                    "NULLIF('"  & 號數   & "','')," & _
                    "NULLIF('"  & 收入   & "','')," & _
                    "NULLIF('"  & 支出   & "','')," & _
                    "NULLIF(N'" & 備註   & "','')," & _
                    "NULLIF(N'" & date1  & "',''))"
                    data.insert()
                    '新增日誌
                    Dim data_dv4 As Data.DataView
                    data.ConnectionString = con_14
                    data.SelectCommand = "SELECT * FROM [刪除資料表] where id_收=" & id & "and date = '" & date1 & "'"
                    data_dv4 = data.Select(New DataSourceSelectArguments)
                    Dim id_2 as string =data_dv4(0)("id").ToString()                    
                    data.insertCommand = _
                    "INSERT INTO 日誌 " & _
                    "(id, 動作, 命令,日期,日期2) " & _
                    "VALUES " & _
                    "(N'" & id & "', N'刪除', N'刪除資料id=" & id_2 & "' , N'" & date1 & "' , '" & date2 & "')"
                    data.insert()
                End If
                '---
                Dim data_dv3 As Data.DataView 
                data.ConnectionString = con_14
                data.SelectCommAnd = "SELECT id As 新id FROM 收支備查簿 WHERE _種類='" & _種類 & "' AND _列='" & p列 & "' AND _頁='" & p頁 & "'"
                data_dv3 = data.Select(New DataSourceSelectArguments)
                Dim nid As int32 = nothing
                If data_dv3.count>0
                    nid =data_dv3(0)("新id").ToString()
                End If
                Dim 預支日期 As String = data_dv(j)("預支日期").ToString()
                If 預支日期 <>""
                    預支日期 = 預支日期.substring(0,10)
                End If
                Dim 歸還日期 As String = data_dv(j)("歸還日期").ToString()
                If 歸還日期 <>""
                    歸還日期 = 歸還日期.substring(0,10)
                End If
                If j<>data_dv.count-1 
                Dim Update1 As string ="UPDATE 收支備查簿 SET " & _
                    "單位別 = NULLIF(N'" & data_dv(j)("單位別").ToString() & "', ''), " & _
                    "承辦人 = NULLIF(N'" & data_dv(j)("承辦人").ToString() & "', ''), " & _
                    "月 = NULLIF(N'" & data_dv(j)("月").ToString() & "', ''), " & _
                    "日 = NULLIF(N'" & data_dv(j)("日").ToString() & "', ''), " & _
                    "科目 = NULLIF(N'" & data_dv(j)("科目").ToString() & "', ''), " & _
                    "科目2 = NULLIF(N'" & data_dv(j)("科目2").ToString() & "', ''), " & _
                    "摘要 = NULLIF(N'" & data_dv(j)("摘要").ToString() & "', ''), " & _
                    "姓名 = NULLIF(N'" & data_dv(j)("姓名").ToString() & "', ''), " & _
                    "商號 = NULLIF(N'" & data_dv(j)("商號").ToString() & "', ''), " & _
                    "經手人 = NULLIF(N'" & data_dv(j)("經手人").ToString() & "', ''), " & _
                    "種類 = NULLIF(N'" & data_dv(j)("種類").ToString() & "', ''), " & _
                    "號數 = NULLIF(N'" & data_dv(j)("號數").ToString() & "', ''), " & _
                    "收入 = NULLIF('" & data_dv(j)("收入").ToString() & "', ''), " & _
                    "支出 = NULLIF('" & data_dv(j)("支出").ToString() & "', ''), " & _
                    "備註 = NULLIF(N'" & data_dv(j)("備註").ToString() & "', ''), " & _
                    "取號 = NULLIF(N'" & data_dv(j)("取號").ToString() & "', ''), " & _
                    "送交主計室日期 = NULLIF(N'" & data_dv(j)("送交主計室日期").ToString() & "', ''), " & _
                    "回覆 = NULLIF(N'" & data_dv(j)("回覆").ToString() & "', ''), " & _
                    "鎖定 = NULLIF(N'" & data_dv(j)("鎖定").ToString() & "', ''), " & _
                    "過審 = NULLIF(N'" & data_dv(j)("過審").ToString() & "', ''), " & _
                    "送出 = NULLIF(N'" & data_dv(j)("送出").ToString() & "', ''), " & _
                    "預支日期 = NULLIF('" & 預支日期 & "', '')," & _
                    "歸還日期 = NULLIF('" & 歸還日期 & "', '')," & _
                    "主計室簽核 = NULLIF(N'" & data_dv(j)("主計室簽核").ToString() & "', ''), " & _
                    "駁回原因 = NULLIF(N'" & data_dv(j)("駁回原因").ToString() & "', '') " & _
                    "WHERE _列 = '" & p列 &"' AND _頁 = " & p頁 & " AND _種類='" & _種類 & "'"
                    data.UpdateCommAnd = Update1
                    data.Update()
                End If
                For k=0 to data_dv2.count-1 
                    If id=data_dv2(k)("id").ToString()
                        data.UpdateCommAnd = "UPDATE 日誌 SET " & _
                            "id = '" & nid & "' " & _
                            "WHERE 主id = '" & data_dv2(k)("主id").ToString() & "'"
                        data.Update()
                    End If
                Next
            Next
             '重算餘額，只做A，不能動到本月小計、累計至本月
            If (_種類="A") Then
                data.UpdateCommAnd = _
                "WITH CTE AS " & _
                "(SELECT *, " & _
                    "(SELECT TOP 1 (CASE WHEN ISNULL(收入,0) = 0 And ISNULL(支出,0) = 0  THEN 餘額 ELSE 0 END) FROM 收支備查簿 WHERE _種類 = '" & _種類 & "' ORDER BY id) " & _
                    "+ " & _
                    "(SUM(" & _
	                "(CASE WHEN ISNULL(摘要,'')<>'本月小計' And ISNULL(摘要,'')<>'累計至本月' THEN ISNULL(收入,0)ELSE 0 END )" & _
                    "-" & _ 
	                "(CASE WHEN ISNULL(摘要,'')<>'本月小計' And ISNULL(摘要,'')<>'累計至本月' THEN ISNULL(支出,0)ELSE 0 END )" & _
	                ") OVER (ORDER BY id " & _
                    "ROWS BETWEEN UNBOUNDED PRECEDING And CURRENT ROW))" & _
                    "AS RunningTotal " & _
                "FROM 收支備查簿 WHERE _種類 = '" & _種類 & "') " & _
                "UPDATE CTE SET 餘額 = RunningTotal"
                data.Update()
            End If
            Label3.Text=""
            Me.GridView1.DataBind()
        Else
            Label3.Text="只能選一筆資料"
        End If
    End Sub
    Protected Sub 審核_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)'審核選取方塊勾選出，對於月日有參考價值?
        Dim checkbox As CheckBox = sender '下面三段為取審核_CheckedChanged在GridView的位置
        Dim row As GridViewRow = checkbox.NamingContainer
        Dim index As Integer = row.RowIndex
        Dim i As Long = index
        Dim 號數 As String = CType(Me.GridView1.Rows(i).FindControl("號數"), TextBox).Text
        If 號數<>""
            For j = 0 to Me.GridView1.Rows.Count - 1
                If CType(Me.GridView1.Rows(j).FindControl("號數"), TextBox).Text=號數
                    CType(Me.GridView1.Rows(j).FindControl("審核"), CheckBox).Checked=CType(Me.GridView1.Rows(i).FindControl("審核"), CheckBox).Checked
                End If
            Next
        End If 
    End Sub
    Protected Sub 月_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)'選月，日會更新，但月為空白時，日不會初始化
        Dim 月 As DropDownList = sender
        Dim row As GridViewRow = 月.NamingContainer
        Dim index As Integer = row.RowIndex
        Dim i As Long = index
        月 = CType(Me.GridView1.Rows(i).FindControl("月"), DropDownList)
        Dim 日 As DropDownList = CType(Me.GridView1.Rows(i).FindControl("日"), DropDownList)
        GetDay(月,日)
        If CType(Me.GridView1.Rows(i).FindControl("月"), DropDownList).Text=""
            Exit Sub
        End if
        Dim 月1 As Integer = 月.Text
        Dim 月2 As Integer = 0
        If i>0
            For j=i-1 to 0 Step -1
                If CType(Me.GridView1.Rows(j).FindControl("月"), DropDownList).Text<>""
                    月2 = CType(Me.GridView1.Rows(j).FindControl("月"), DropDownList).Text
                End If
                If 月1<月2
                    Label3.Text="請確認月、日是否正確"
                End If
            Next
        End If
    End Sub
    Protected Sub 日_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)'選月，日會更新，但月為空白時，日不會初始化
        Dim 日 As DropDownList = sender
        Dim row As GridViewRow = 日.NamingContainer
        Dim index As Integer = row.RowIndex
        Dim i As Long = index
        If CType(Me.GridView1.Rows(i).FindControl("日"), DropDownList).Text="" OR CType(Me.GridView1.Rows(i).FindControl("月"), DropDownList).Text=""
            Exit Sub
        End if
        Dim 日1 As Integer = CType(Me.GridView1.Rows(i).FindControl("日"), DropDownList).Text
        Dim 月1 As Integer = CType(Me.GridView1.Rows(i).FindControl("月"), DropDownList).Text
        Dim 日2 As Integer = 0
        Dim 月2 As Integer = 0
        If i>0
            For j=i-1 to 0 Step -1
                If CType(Me.GridView1.Rows(j).FindControl("月"), DropDownList).Text<>""
                    月2 = CType(Me.GridView1.Rows(j).FindControl("月"), DropDownList).Text
                End If
                If CType(Me.GridView1.Rows(j).FindControl("日"), DropDownList).Text<>""
                    日2 = CType(Me.GridView1.Rows(j).FindControl("日"), DropDownList).Text
                End If
                If 月1<月2 OR (月1=月2 AND 日1<日2)
                    Label3.Text="請確認月、日是否正確"
                End If
            Next
        End If
    End Sub
    Protected Sub GridView1_RowCommAnd(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewCommAndEventArgs) HAndles GridView1.RowCommAnd
        If e.CommAndName = "簽名圖"'簽名按紐
            Update(sender, e)
            Dim i As Long = e.CommAndSource.NamingContainer.RowIndex
            Dim id As String = CType(Me.GridView1.Rows(i).FindControl("id"), TextBox).Text
            Dim dataURL As string = Request.Form("dataURL")
            Dim 年 As String = Me.年.Text
            Dim 種類 As String = Me._種類.Text
            data.UpdateCommAnd = "UPDATE 收支備查簿 SET 姓名 = NULLIF(N'" & dataURL & "', '')" & _ 
            "FROM 收支備查簿 WHERE 年 = " & 年 & " And _種類 = '" & 種類 & "' And id=" & id & "And 過審='False'"
            data.Update()
            Me.GridView1.DataBind()
            Label1.Text="新增簽名檔儲存功能執行成功，如無改變，請檢查資料狀態。"
        ElseIf e.CommAndName = "經手人圖"
            Update(sender, e)
            Dim i As Long = e.CommAndSource.NamingContainer.RowIndex
            Dim id As String = CType(Me.GridView1.Rows(i).FindControl("id"), TextBox).Text
            Dim dataURL As string = Request.Form("dataURL")
            Dim 年 As String = Me.年.Text
            Dim 種類 As String = Me._種類.Text
            data.UpdateCommAnd = "UPDATE 收支備查簿 SET 經手人 = NULLIF(N'" & dataURL & "', '')" & _ 
            "FROM 收支備查簿 WHERE 年 = " & 年 & " And _種類 = '" & 種類 & "' And id=" & id & "And 過審 = 'False'"
            data.Update()
            Me.GridView1.DataBind()
            Label1.Text="新增經手人簽名檔儲存功能執行成功，如無改變，請檢查資料狀態。"
        ElseIf e.CommAndName = "本月小計"'改進:修改B、XZ不做收入，如剛好在第一頁做會有問題
            Update(sender, e)
            Dim 作用 As boolean = False
            Dim i As Long = e.CommAndSource.NamingContainer.RowIndex
            Dim id As String = CType(Me.GridView1.Rows(i).FindControl("id"), TextBox).Text
            Dim 摘要 As String = CType(Me.GridView1.Rows(i).FindControl("摘要"), TextBox).Text
            Dim 月 As string = nothing
            Dim 日 As String = nothing
            Dim j As Int32
            If i>=0 
                j=i'不為該頁第一筆資料
                    While j>0 AND CType(Me.GridView1.Rows(j).FindControl("月"), DropDownList).Text=""'找最後一筆有資料
                        j=j-1
                    End While
                    If CType(Me.GridView1.Rows(j).FindControl("月"), DropDownList).Text<>""
                        月=CType(Me.GridView1.Rows(j).FindControl("月"), DropDownList).Text
                        日=CType(Me.GridView1.Rows(j).FindControl("日"), DropDownList).Text
                    End If 
            End If 
            IF 月=""
                月=(DateTime.Now.month-1).ToString()
            END IF
            Dim 年 As String = Me.年.Text
            Dim 種類 As String = Me._種類.Text
            'Label1.Text=月 + "月" + 日 + "日"
            If 摘要="" OR 摘要="本月小計"
                CType(Me.GridView1.Rows(i).FindControl("月"), DropDownList).Text=月
                'SQL 加上ISUNLL 31 會變成 * 可能是因為ISNULL會將後變數類型轉成前變數類型，日期改用CASE WHEN
                data.UpdateCommAnd = "UPDATE 收支備查簿 SET "& _
                "月=ISNULL(NULLIF('" & 月 & "' ,''),(CASE WHEN(DAY(Getdate())<16) THEN Month(Dateadd(Month,-1,Getdate())) ELSE Month(Getdate()) END)), "  & _
	            "日=(CASE WHEN NULLIF(TRIM('" & 日 & "'),'') IS NULL "  & _
	            "THEN DAY(EOMONTH(STR(" & 年 & "+1911)+'/'+STR(ISNULL(NULLIF( '" & 月 & "' ,''),(CASE WHEN(DAY(Getdate())<16) THEN Month(Dateadd(Month,-1,Getdate())) ELSE Month(Getdate()) END)))+'/01')) "  & _
                "Else TRIM('" & 日 & "') END) "  & _
                ",收入 = NULLIF( " & _
                "(SELECT sum(收入) FROM 收支備查簿 WHERE " & _
                "id >=(SELECT TOP 1 id FROM 收支備查簿 WHERE 年 = " & 年 & " And _種類 = '" & 種類 & "' And 月=ISNULL(NULLIF('" & 月 & "' ,''),(CASE WHEN(DAY(Getdate())<16) THEN Month(Dateadd(Month,-1,Getdate())) ELSE Month(Getdate()) END)) And 摘要 <> '本月小計' And 摘要 <> '累計至本月' ORDER BY id) And " & _
                "id <=(SELECT TOP 1 id FROM 收支備查簿 WHERE 年 = " & 年 & " And _種類 = '" & 種類 & "' And 月=ISNULL(NULLIF('" & 月 & "' ,''),(CASE WHEN(DAY(Getdate())<16) THEN Month(Dateadd(Month,-1,Getdate())) ELSE Month(Getdate()) END)) And 摘要 <> '本月小計' And 摘要 <> '累計至本月' ORDER BY id Desc) And " & _
                "年 = " & 年 & " And _種類 = '" & 種類 & "' And 摘要 <> '本月小計' And 摘要 <> '累計至本月') " & _
                ", '')" & _ 
                ",支出 = NULLIF( " & _ 
                "(SELECT sum(支出) FROM 收支備查簿 WHERE " & _
                "id >=(SELECT TOP 1 id FROM 收支備查簿 WHERE 年 = " & 年 & " And _種類 = '" & 種類 & "' And 月=ISNULL(NULLIF('" & 月 & "' ,''),(CASE WHEN(DAY(Getdate())<16) THEN Month(Dateadd(Month,-1,Getdate())) ELSE Month(Getdate()) END)) And 摘要 <> '本月小計' And 摘要 <> '累計至本月' ORDER BY id) And " & _
                "id <=(SELECT TOP 1 id FROM 收支備查簿 WHERE 年 = " & 年 & " And _種類 = '" & 種類 & "' And 月=ISNULL(NULLIF('" & 月 & "' ,''),(CASE WHEN(DAY(Getdate())<16) THEN Month(Dateadd(Month,-1,Getdate())) ELSE Month(Getdate()) END)) And 摘要 <> '本月小計' And 摘要 <> '累計至本月' ORDER BY id Desc) And " & _
                "年 = " & 年 & " And _種類 = '" & 種類 & "' And 摘要 <> '本月小計' And 摘要 <> '累計至本月') " & _
                ", '') " & _ 
                ",摘要= N'本月小計' " & _
                "FROM 收支備查簿 WHERE 年 = '" & 年 & "' And _種類 = '" & 種類 & "' And id='" & id & "' And 鎖定 = 'False' "
                data.Update()
                Me.GridView1.DataBind()
                Label1.Text="本月小計成功"
            Else
                Label3.Text="摘要有內容，請刪除摘要內容"
            End If 
        ElseIf e.CommAndName = "累計至本月"'改進:修改B、XZ不做收入
            Update(sender, e)
            Dim i As Long = e.CommAndSource.NamingContainer.RowIndex
            Dim id As String = CType(Me.GridView1.Rows(i).FindControl("id"), TextBox).Text
            Dim 摘要 As String = CType(Me.GridView1.Rows(i).FindControl("摘要"), TextBox).Text
            Dim 月 As string =""
            Dim 日 As string =""
            Dim j As Int32
            If i>=0 
                j=i'不為該頁第一筆資料
                    While  j>0 AND CType(Me.GridView1.Rows(j).FindControl("月"), DropDownList).Text=""'找最後一筆有資料
                        j=j-1
                    End While
                    If CType(Me.GridView1.Rows(j).FindControl("月"), DropDownList).Text<>""
                        月=CType(Me.GridView1.Rows(j).FindControl("月"), DropDownList).Text
                    End If 
            End If 
            IF 月=""
                月=(DateTime.Now.month-1).ToString()
                日=DateTime.DaysInMonth((CLng(Me.年.text) + 1911), CLng(月))
            END IF
            If i>=0 
                j=i'不為該頁第一筆資料
                    While  j>0 AND CType(Me.GridView1.Rows(j).FindControl("日"), DropDownList).Text=""'找最後一筆有資料
                        j=j-1
                    End While
                    If CType(Me.GridView1.Rows(j).FindControl("日"), DropDownList).Text<>""
                        日=CType(Me.GridView1.Rows(j).FindControl("日"), DropDownList).Text
                    End If 
            End If 
            Dim 年 As String = Me.年.Text
            Dim 種類 As String = Me._種類.Text
            If 摘要="" OR 摘要="累計至本月"
                data.UpdateCommAnd = "UPDATE 收支備查簿 SET "& _
                "月 = ISNULL(NULLIF('" & 月 & "', ''),Month(Dateadd(Month,-1,Getdate())))"  & _
	            ",日=(CASE WHEN NULLIF(TRIM('" & 日 & "'),'') IS NULL "  & _
	            "THEN DAY(EOMONTH(Dateadd(Month,-1,Getdate()))) "  & _
                "Else TRIM('" & 日 & "') END) "  & _
                ",收入 = NULLIF(" & _
                "(SELECT sum(收入) FROM 收支備查簿 WHERE 年 = " & 年 & " And _種類 = '" & 種類 & "' And id <=" & id & " And 摘要 <> '本月小計' And 摘要 <> '累計至本月')" & _
                ", '')" & _ 
                ",支出 = NULLIF(" & _ 
                "(SELECT sum(支出) FROM 收支備查簿 WHERE 年 = " & 年 & " And _種類 = '" & 種類 & "' And id <=" & id & " And 摘要 <> '本月小計' And 摘要 <> '累計至本月')"  & _
                ", '')" & _ 
                ",摘要= N'累計至本月'" & _
                "FROM 收支備查簿 WHERE 年 = " & 年 & " And _種類 = '" & 種類 & "' And id=" & id & " And 鎖定 = 'False'"
                data.Update()
                Me.GridView1.DataBind()
                Label1.Text="累計至本月成功"
            Else
                Label3.Text="摘要有內容，請刪除摘要內容"
            End If 
        End If 
    End Sub
    Protected Sub GridView1_DataBound(ByVal sender As Object, ByVal e As System.EventArgs) HAndles GridView1.DataBound
        GenerateDropdownlist()
    End Sub
    Protected Sub Update_f()'存檔方法，PS:將原有的一鍵計算刪除，只留重作餘額
        Dim 年 As String = Me.年.Text
        Dim _種類 As String = Me._種類.Text
        Dim id_號 As String
        Dim 單位別_號 As String
        Dim 承辦人_號 As String
        Dim 號數_號 As String
        Dim 種類_號 As String
        Dim 備註_號 As String
        For i = 0 To Me.GridView1.Rows.Count - 1
            Dim id As String = CType(Me.GridView1.Rows(i).FindControl("id"), TextBox).Text
            Dim 單位別 As String = CType(Me.GridView1.Rows(i).FindControl("單位別"),DropDownList).Text
            Dim 承辦人 As String = CType(Me.GridView1.Rows(i).FindControl("承辦人"), DropDownList).Text
            Dim 月 As String = CType(Me.GridView1.Rows(i).FindControl("月"), DropDownList).Text
            Dim 日 As String = CType(Me.GridView1.Rows(i).FindControl("日"), DropDownList).Text
            Dim 科目 As String = CType(Me.GridView1.Rows(i).FindControl("科目"), DropDownList).Text
            Dim 科目2 As String = CType(Me.GridView1.Rows(i).FindControl("科目2"), DropDownList).Text
            Dim 原本科目 As String =科目
            If 科目2<>""
                原本科目 = (科目 & ";" & 科目2)
            End If
            Dim 摘要 As String = CType(Me.GridView1.Rows(i).FindControl("摘要"), TextBox).Text
            Dim 姓名 As String = CType(Me.GridView1.Rows(i).FindControl("姓名"), ImageButton).ImageUrl
            'Dim 姓名text As String = CType(Me.GridView1.Rows(i).FindControl("姓名text"), TextBox).Text
            Dim 商號 As String = CType(Me.GridView1.Rows(i).FindControl("商號"), TextBox).Text
            Dim 經手人 As String = CType(Me.GridView1.Rows(i).FindControl("經手人"), ImageButton).ImageUrl
            'Dim 經手人text As String = CType(Me.GridView1.Rows(i).FindControl("經手人text"), TextBox).Text
            Dim 種類 As String = CType(Me.GridView1.Rows(i).FindControl("種類"), TextBox).Text
            Dim 號數 As String = CType(Me.GridView1.Rows(i).FindControl("號數"), TextBox).Text
            Dim 收入 As String = CType(Me.GridView1.Rows(i).FindControl("收入"), TextBox).Text
            收入=收入.Replace(",", "").Replace("N", "").Replace("T", "").Replace("$", "")
            Dim 支出 As String = CType(Me.GridView1.Rows(i).FindControl("支出"), TextBox).Text
            支出=支出.Replace(",", "").Replace("N", "").Replace("T", "").Replace("$", "")
            Dim 餘額 As String = CType(Me.GridView1.Rows(i).FindControl("餘額"), TextBox).Text
            Dim 備註 As String = CType(Me.GridView1.Rows(i).FindControl("備註"), TextBox).Text
            'Dim 預支日期 As String = CType(Me.GridView1.Rows(i).FindControl("預支日期"), TextBox).Text
            Dim 預支日期 As String = nothing
            If CType(Me.GridView1.Rows(i).FindControl("預支日期"), TextBox).text<>""
                預支日期 = CType(Me.GridView1.Rows(i).FindControl("預支日期"), TextBox).text
                預支日期 = taiwancalendarto(預支日期)
            End If
            Dim 歸還日期 As String = nothing
            If CType(Me.GridView1.Rows(i).FindControl("歸還日期"), TextBox).text<>""
                歸還日期 = CType(Me.GridView1.Rows(i).FindControl("歸還日期"), TextBox).text
                歸還日期 = taiwancalendarto(歸還日期)
            End If
            Me.GridView1.Rows(i).FindControl("姓名").Visible = True
            Me.GridView1.Rows(i).FindControl("經手人").Visible = True
            IF _種類<>種類 AND 種類<>"" '當種類與_種類不符合，且種類不為空(例:_種類=A，種類=B、XZ)
                Label3.Text="種類不相符，請確定種類頁面使否相符"
            ELSE
                If CType(Me.GridView1.Rows(i).FindControl("審核狀態"), Label).Text="駁回"'駁回新增修改紀錄，駁回為 鎖定False,送出true，要新增至修改資料
                DIM sql As string = "SELECT * FROM [收支備查簿] where id=" & id
                data.SelectCommAnd = sql
                data_dv = data.Select(New DataSourceSelectArguments)
                Dim 單位別_前 As String = data_dv(0)("單位別").ToString()
                Dim 承辦人_前 As String = data_dv(0)("承辦人").ToString()
                Dim 月_前 As String = data_dv(0)("月").ToString()
                Dim 日_前 As String = data_dv(0)("日").ToString()
                Dim 科目_前 As String = data_dv(0)("科目").ToString()
                If data_dv(0)("科目2").ToString()<>""
                    科目_前 = data_dv(0)("科目").ToString() & ";" & data_dv(0)("科目2").ToString()
                End If
                Dim 摘要_前 As String = data_dv(0)("摘要").ToString()
                Dim 姓名_前 As String = data_dv(0)("姓名").ToString()
                Dim 商號_前 As String = data_dv(0)("商號").ToString()
                Dim 經手人_前 As string = data_dv(0)("經手人").ToString()
                Dim 種類_前 As String = data_dv(0)("種類").ToString()
                Dim 號數_前 As String = data_dv(0)("號數").ToString()
                Dim 收入_前 As String = data_dv(0)("收入").ToString()
                Dim 支出_前 As String = data_dv(0)("支出").ToString()
                Dim 備註_前 As String = data_dv(0)("備註").ToString()
                '判斷資料是否有修改
                If 號數<>""
                    號數=CType(號數,Int32)
                    號數=CType(號數,string)
                End If 
                'B XZ 沒有收入，會到導致""<>"0" ，先將收入及支出空值得收入改成預設值0
                If 收入=""
                    收入="0"
                End If
                If 支出=""
                    支出="0"
                End If  
                If 單位別<>單位別_前 or 承辦人<>承辦人_前 or 月<>月_前 or 日<>日_前 or 原本科目<>科目_前 or 摘要<>摘要_前 or 姓名<>姓名_前 or 商號<>商號_前 or 經手人<>經手人_前 or 種類<>種類_前 or 號數<>號數_前 or (收入<>收入_前 And 種類="A" ) or 支出<>支出_前 or 備註<>備註_前
                    '程式正式執行
                    '新增修改資料
                    Dim date1 As string = DateTime.now.tostring()
                    Dim date2 As string = DateTime.now.tostring("yyyy-MM-dd HH:mm:ss")
                    Dim insert1 As string = _
                    "INSERT INTO 修改資料 " & _
                    "(id_收,單位別,單位別_改,承辦人,承辦人_改,月,月_改,日,日_改,科目,科目_改,摘要,摘要_改,姓名,姓名_改,商號,商號_改,經手人,經手人_改,種類,種類_改,號數,號數_改,收入,收入_改,支出,支出_改,備註,備註_改,date) " & _
                    "VALUES " & _
                    "(" & id & ",NULLIF(N'" & 單位別_前 & "',''),NULLIF(N'" & 單位別 & "',''),NULLIF(N'" & _
                    承辦人_前 & "',''),NULLIF(N'" & 承辦人 & "',''),NULLIF('" & _
                    月_前 & "',''),NULLIF('" & 月 & "',''),NULLIF('" & _
                    日_前 & "',''),NULLIF('" & 日 & "',''),NULLIF(N'" & _
                    科目_前 & "',''),NULLIF(N'" & 原本科目 & "',''),NULLIF(N'" & _
                    摘要_前 & "',''),NULLIF(N'" & 摘要 & "',''),NULLIF(N'" & _
                    姓名_前 & "',''),NULLIF(N'" & 姓名 & "',''),NULLIF(N'" & _ 
                    商號_前 & "',''),NULLIF(N'" & 商號 & "',''),NULLIF(N'" & _
                    經手人_前 & "',''),NULLIF(N'" & 經手人 & "',''),NULLIF(N'" & _
                    種類_前 & "',''),NULLIF(N'" & 種類 & "',''),NULLIF('" & _
                    號數_前 & "',''),NULLIF('" & 號數 & "',''),NULLIF('" & _
                    收入_前 & "',''),NULLIF('" & 收入 & "',''),NULLIF('" & _
                    支出_前 & "',''),NULLIF('" & 支出 & "',''),NULLIF(N'" & _
                    備註_前 & "',''),NULLIF(N'" & 備註 & "',''),NULLIF(N'" & date1 &  "',''))"
                    data.insertCommAnd = insert1
                    data.insert()
                    '查詢修改資料表的ID，並把他給日誌
                    sql = "SELECT * FROM [修改資料] where id_收=" & id & "And date = '" & date1 & "'"
                    data.SelectCommAnd = sql
                    data_dv = data.Select(New DataSourceSelectArguments)
                    Dim id_2 As string =data_dv(0)("id").ToString()
                    data.insertCommAnd = _
                    "INSERT INTO 日誌 " & _
                    "(id, 動作, 命令,日期,日期2) " & _
                    "VALUES " & _
                    "(N'" & id & "', N'修改', N'修改資料id="& id_2 &"', N'" & date1 & "', '" & date2 & "')"
                    data.insert()
                End If 
                End If 
                Dim Update1 As string ="UPDATE 收支備查簿 SET " & _
                    "單位別 = NULLIF(N'" & 單位別 & "', ''), " & _
                    "承辦人 = NULLIF(N'" & 承辦人 & "', ''), " & _
                    "月 = NULLIF(N'" & 月 & "', ''), " & _
                    "日 = NULLIF(N'" & 日 & "', ''), " & _
                    "科目 = NULLIF(N'" & 科目 & "', ''), " & _
                    "科目2 = NULLIF(N'" & 科目2 & "', ''), " & _
                    "摘要 = REPLACE(REPLACE(NULLIF(N'" & 摘要 & "', ''),' ',''),CHAR(13)+CHAR(10),''), " & _
                    "姓名 = NULLIF(N'" & 姓名 & "', ''), " & _
                    "商號 = NULLIF(N'" & 商號 & "', ''), " & _
                    "經手人 = NULLIF(N'" & 經手人 & "', ''), " & _
                    "種類 = NULLIF(N'" & 種類 & "', ''), " & _
                    "號數 = NULLIF(N'" & 號數 & "', ''), " & _
                    "收入 = REPLACE(ISNULL(NULLIF('" & 收入 & "', ''),'0'), ',', ''), " & _
                    "支出 = REPLACE(ISNULL(NULLIF('" & 支出 & "', ''),'0'), ',', ''), " & _
                    "備註 = NULLIF(N'" & 備註 & "', ''), " & _
                    "預支日期 = (CASE WHEN ISDATE(NULLIF(N'" & 預支日期 & "', ''))=1 Then NULLIF(N'" & 預支日期 & "', '') End )," & _
                    "歸還日期 = (CASE WHEN ISDATE(NULLIF(N'" & 歸還日期 & "', ''))=1 Then NULLIF(N'" & 歸還日期 & "', '') End )" & _
                    "WHERE id = '" & id & "' And 鎖定 = 'False' "'鎖定則不修改
                data.UpdateCommAnd = Update1
                data.Update()
                '重算餘額，只做A，不能動到本月小計、累計至本月
                If (_種類="A") Then
                    data.UpdateCommAnd = _
                    "WITH CTE AS " & _
                    "(SELECT *, " & _
                        "(SELECT TOP 1 (CASE WHEN ISNULL(收入,0) = 0 And ISNULL(支出,0) = 0  THEN 餘額 ELSE 0 END) FROM 收支備查簿 WHERE _種類 = '" & _種類 & "' ORDER BY id) " & _
                        "+ " & _
                        "(SUM(" & _
	                    "(CASE WHEN ISNULL(摘要,'')<>'本月小計' And ISNULL(摘要,'')<>'累計至本月' THEN ISNULL(收入,0)ELSE 0 END )" & _
                        "-" & _ 
	                    "(CASE WHEN ISNULL(摘要,'')<>'本月小計' And ISNULL(摘要,'')<>'累計至本月' THEN ISNULL(支出,0)ELSE 0 END )" & _
	                    ") OVER (ORDER BY id " & _
                        "ROWS BETWEEN UNBOUNDED PRECEDING And CURRENT ROW))" & _
                        "AS RunningTotal " & _
                    "FROM 收支備查簿 WHERE _種類 = '" & _種類 & "') " & _
                    "UPDATE CTE SET 餘額 = RunningTotal"
                    data.Update()
                End If
            End If
        Next
        Label1.Text="存檔成功"
    End Sub
    Public Sub GenerateDropdownlist()'動態輸入科目名稱
        data.SelectCommand = "select Distinct 科目 from 科目表 Where 科目<>'' OR 科目 IS NOT NULL  order by 科目"
        data_dv = data.Select(New DataSourceSelectArguments)
        For i = 0 To Me.GridView1.Rows.Count - 1
            Dim id As String = CType(Me.GridView1.Rows(i).FindControl("id"), TextBox).Text
            data.SelectCommand = "select * from 收支備查簿 Where id='"& id &"'"
            Dim data_dv2 As Data.DataView = data.Select(New DataSourceSelectArguments)
            Dim 科目1 As DropDownList = CType(Me.GridView1.Rows(i).FindControl("科目"), DropDownList)
            Dim 科目2 As DropDownList = CType(Me.GridView1.Rows(i).FindControl("科目2"), DropDownList)
            Dim 科目S1 As String = data_dv2(0)("科目").ToString()
            Dim 科目S2 As String = data_dv2(0)("科目2").ToString()
            科目1.Items.Clear()
            科目2.Items.Clear()
            科目1.Items.Add("")
            科目1.Items(0).Value = ""
            科目2.Items.Add("")
            科目2.Items(0).Value = ""
            For j = 0 To data_dv.Count - 1
                Dim 科目名稱 As String = data_dv(j)(0)
                科目1.Items.Add(科目名稱)
                科目1.Items(j+1).Value = 科目名稱
                科目2.Items.Add(科目名稱)
                科目2.Items(j+1).Value = 科目名稱
            Next
            科目1.SelectedIndex=科目1.Items.IndexOf(科目1.Items.FindByValue(科目S1))
            科目2.SelectedIndex=科目2.Items.IndexOf(科目2.Items.FindByValue(科目S2))
        Next
    End Sub
    Public Sub initialization()'初始化
        頁1.Text=""
        頁2.Text=""
        單位別.Text=""
        承辦人.Text=""
        月1.Text=""
        日1.Text=""
        月2.Text=""
        日2.Text=""
        科目.Text=""
        摘要.Text=""
        商號.Text=""
        號數1.Text=""
        號數2.Text=""
    End Sub
    Protected Sub Clear_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Response.Redirect(Request.Url.ToString())
    End Sub
    Protected Sub 年_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles 年.TextChanged
        月1_SelectedIndexChanged(sender,e)
        月2_SelectedIndexChanged(sender,e)
    End Sub
    Protected Sub 月1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles 月1.SelectedIndexChanged
       GetDay(月1,日1)
    End Sub
    Protected Sub 月2_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles 月2.SelectedIndexChanged
       GetDay(月2,日2)
    End sub
    'Shared之前，需要下行程式碼
    <System.Web.Script.Services.ScriptMethod(), System.Web.Services.WebMethod()>
    Public Shared Function GetMyList(ByVal prefixText As String, ByVal count As Integer)'常用文字
        Dim con_14 As String = System.Web.Configuration.WebConfigurationManager.ConnectionStrings("ApplicationServices").ConnectionString
        Dim data As New SqlDataSource
        Dim data_dv As Data.DataView
        Dim MyList As New List(Of String)
        data.ConnectionString = con_14
        data.SelectCommAnd = "SELECT TOP " & count & " 常用文字 FROM 常用清單 WHERE 常用文字 LIKE '%" & prefixText & "%' ORDER BY CASE WHEN 常用文字 IS NULL THEN 1 ELSE 0 END, 常用文字"
        data_dv = data.Select(New DataSourceSelectArguments)
        For i As Long = 0 To data_dv.Count() - 1
            MyList.Add(data_dv(i)(0).ToString())
        Next
        Return MyList
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
    Public Sub GetDay(ByVal month As Object,ByVal day As Object)'以月取日，收尋，日可不留白
        If month.SelectedValue<>"" And Me.年.text<>""
            Dim currentdate = day.SelectedValue
            day.Items.Clear()
            day.Items.Add("")
            day.Items(0).Value = ""
            For i = 1 To DateTime.DaysInMonth((CLng(Me.年.text) + 1911), CLng(month.SelectedValue))
                day.Items.Add((i).ToString("0"))
                day.Items(i).Value = (i).ToString("0")
            Next
            If day.Items.IndexOf(day.Items.FindByValue(currentdate)) = -1
                day.SelectedIndex = day.Items.Count - 1
            Else
                day.SelectedIndex = day.Items.IndexOf(day.Items.FindByValue(currentdate))
            End If
        End If
    End Sub
End Class