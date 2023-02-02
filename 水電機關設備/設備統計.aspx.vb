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
Partial Class 設備統計
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
        Me.狀態.text=""
        Update(sender, e)
        For i = 1 To 6
            Dim insert1 as string
            data.InsertCommand = _
            "INSERT INTO 設備 " & _
            "( _頁, _列,年限,保管人) " & _
            "VALUES " & _
            "('" & (Me.GridView1.PageCount + 1).ToString() & "', '" & i & "','2年','林芳瑜')"
            data.Insert()
        Next
        Me.GridView1.PageIndex = Int32.MaxValue
        Me.GridView1.DataBind()
        ' Me.SqlDataSource1.Insert()
        ' Me.GridView1.PageIndex = Int32.MaxValue
    End Sub
    Protected Sub Update(ByVal sender As Object, ByVal e As System.EventArgs)
        'TODO:轉換成整數會失敗
        For i = 0 To Me.GridView1.Rows.Count - 1
            Dim id As String = CType(Me.GridView1.Rows(i).FindControl("id"), TextBox).Text
            Dim 項 As String = CType(Me.GridView1.Rows(i).FindControl("項"),Label).Text
            Dim 財產編號 As String = CType(Me.GridView1.Rows(i).FindControl("財產編號"),TextBox).Text
            Dim 財產名稱 As String = CType(Me.GridView1.Rows(i).FindControl("財產名稱"),TextBox).Text
            Dim 財產別名 As String = CType(Me.GridView1.Rows(i).FindControl("財產別名"),TextBox).Text
            Dim 分群 As String = CType(Me.GridView1.Rows(i).FindControl("分群"), DropDownList).Text
            Dim 廠牌 As String = CType(Me.GridView1.Rows(i).FindControl("廠牌"),TextBox).Text
            Dim 型號 As String = CType(Me.GridView1.Rows(i).FindControl("型號"), TextBox).Text
            型號=型號.Replace("'", "''")
            Dim 購置日期 As String = nothing
            If CType(Me.GridView1.Rows(i).FindControl("購置日期"), TextBox).text<>""
                購置日期 = CType(Me.GridView1.Rows(i).FindControl("購置日期"), TextBox).text
                購置日期 = taiwancalendarto(購置日期)
            End If
            Dim 單位 As String = CType(Me.GridView1.Rows(i).FindControl("單位"), TextBox).Text
            Dim 數量 As String = CType(Me.GridView1.Rows(i).FindControl("數量"), TextBox).Text
            數量=數量.Replace(",", "")
            Dim 年限 As String = CType(Me.GridView1.Rows(i).FindControl("年限"), TextBox).Text
            Dim 年數 As String = CType(Me.GridView1.Rows(i).FindControl("年數"), TextBox).Text
            Dim 保管人 As String = CType(Me.GridView1.Rows(i).FindControl("保管人"), TextBox).Text
            Dim 存置地點 As String = CType(Me.GridView1.Rows(i).FindControl("存置地點"), TextBox).Text
            Dim Update1 as string ="UPDATE 設備 SET " & _
            "項 = NULLIF(N'" & 項 & "', ''), " & _
            "財產編號 = NULLIF(N'" & 財產編號 & "', ''), " & _
            "財產名稱 = NULLIF(N'" & 財產名稱 & "', ''), " & _
            "財產別名 = NULLIF(N'" & 財產別名 & "', ''), " & _
            "分群 = NULLIF(N'" & 分群 & "', ''), " & _
            "廠牌 = NULLIF(N'" & 廠牌 & "', ''), " & _
            "型號 = NULLIF(N'" & 型號 & "', ''), " & _
            "購置日期 = IIF(ISDATE(TRIM(N'" & 購置日期 & "'))=1,TRIM(N'" & 購置日期 & "'),NULL), " & _
            "單位 = NULLIF(N'" & 單位 & "', ''), " & _
            "數量 = NULLIF(N'" & 數量 & "', ''), " & _
            "年限 = NULLIF(N'" & 年限 & "', ''), " & _
            "年數 = NULLIF(N'" & 年數 & "', ''), " & _
            "保管人 = NULLIF(N'" & 保管人 & "', ''), " & _
            "存置地點 = NULLIF(N'" & 存置地點 & "', '') " & _
            "WHERE id = '" & id & "'"
            data.UpdateCommand = Update1
            data.Update()
            Update1 = "UPDATE 設備統計分群 SET " & _
            "總計 =  (Select Sum(數量) From 設備 Where 分群 = NULLIF(N'" & 分群 & "', '') Group by 分群)" & _
            "WHERE 分群 = NULLIF(N'" & 分群 & "', '')"
            data.UpdateCommand = Update1
            data.Update()
            If CType(Me.GridView1.Rows(i).FindControl("項"), Label).text=""
                 Dim Update2 as string ="WITH CTE AS (Select *,Row_Number() OVER(Partition by 分群 order by ID ) AS '序號' From 設備)" & _
                "UPDATE CTE SET 項 = 序號 Where 分群 IS NOT NULL"
                data.UpdateCommand = Update2
                data.Update()
            End If
        Next
        Me.GridView1.DataBind()
    End Sub
    Protected Sub Delete(ByVal sender As Object, ByVal e As System.EventArgs)
        Me.狀態.text=""
        Update(sender, e)
        Me.GridView1.PageIndex = Int32.MaxValue
        Me.GridView1.DataBind()
        Dim delete1 as string=""
        For i = 0 To Me.GridView1.Rows.Count - 1
            Dim id As String = CType(Me.GridView1.Rows(i).FindControl("id"), TextBox).Text
            delete1="DELETE FROM 設備 " & _
            "WHERE id = '" & id & "'"
            data.deleteCommand =delete1
            data.delete()
        Next
        Me.GridView1.DataBind()
    End Sub
    Protected Sub 年數偵測_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        data.ConnectionString = con_14
        data.SelectCommAnd = "SELECT id,購置日期,年數 FROM 設備 "
        data_dv = data.Select(New DataSourceSelectArguments)
        Dim id As String
        Dim 購置日期 As String
        Dim 年數 As String=""
        Dim 年 AS Int32=0
        Dim 月 As Int32=0
        Dim 滿 As Int32=0 '剛好滿一個月
        Dim 相差月 As Int32=0
        For i = 0 To data_dv.count-1
            年 = 0
            月 = 0
            年數=""
            相差月=0
            id = data_dv(i)("id").ToString()
            購置日期 = data_dv(i)("購置日期").ToString()
            If IsDate(購置日期)
                If Day(購置日期)>Day(DateTime.Now)
                    滿=1
                Else
                    滿=0
                End If
                相差月=(DateDiff("m",購置日期,DateTime.Now)-滿)
                If 相差月\12<>0
                    年數=相差月\12 & "年"
                End If
                If 相差月-((相差月\12)*12)<>0
                    年數=年數 & 相差月-((相差月\12)*12) & "月"
                ElseIf 相差月\12=0'狀況為0年0月
                    年數="不到1月"
                End If
                If InStr(data_dv(i)("年數").ToString(), "年")<>0 AND InStr(data_dv(i)("年數").ToString(), "月")<>0
                    月 = Mid(data_dv(i)("年數").ToString(),InStr(data_dv(i)("年數").ToString(), "年")+1,InStr(data_dv(i)("年數").ToString(), "月")-InStr(data_dv(i)("年數").ToString(), "年")-1)
                    年 = LEFT(data_dv(i)("年數").ToString(),InStr(data_dv(i)("年數").ToString(), "年")-1)
                ElseIF InStr(data_dv(i)("年數").ToString(), "年")=0 AND InStr(data_dv(i)("年數").ToString(), "月")<>0 AND data_dv(i)("年數").ToString()<>"不到1月"
                    月 = LEFT(data_dv(i)("年數").ToString(),InStr(data_dv(i)("年數").ToString(), "月")-1)
                ElseIF InStr(data_dv(i)("年數").ToString(), "月")=0 AND InStr(data_dv(i)("年數").ToString(), "年")<>0
                    年 = LEFT(data_dv(i)("年數").ToString(),InStr(data_dv(i)("年數").ToString(), "年")-1)
                End If
                If (相差月<=(年*12+月+1) AND 相差月<>(年*12+月)) OR data_dv(i)("年數").ToString()=""'相差1月，但相同字串不執行
                data.UpdateCommand = "UPDATE 設備 SET " & _
                    "年數 = NULLIF(N'" & 年數 & "', '') " & _
                    "WHERE id = '" & id & "'"
                data.Update()
                'Me.Label1.Text=Me.Label1.Text & id & ":TURE<BR>"
                End If
            End If
        Next
        Me.GridView1.DataBind()
    End Sub
    Protected Sub Test(ByVal sender As Object, ByVal e As System.EventArgs)
        狀態.text=""
    End Sub
    Protected Sub Download(ByVal sender As Object, ByVal e As System.EventArgs)
        年數偵測_Click(sender,e)
        '8/3 每張頁首、頁尾調整
        Dim MyGUID As String = Guid.NewGuid().ToString("N")
        Dim MyExcel As String = MapPath(".\Excel\Temp\") & MyGUID & ".xlsx"
        System.IO.File.Copy(MapPath(".\Excel\水電設備統計表.xlsx"), MyExcel)
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
        xlWorkSheet = CType(xlWorkBook.Sheets("統計表"), Excel.Worksheet)
        xlWorkSheet.Activate()
        data.ConnectionString = con_14
        Dim 狀態 As string = Me.狀態.Text
        data.SelectCommAnd = "SELECT * FROM 設備 Where 分群 Is Not NULL AND ('" & 狀態 & "'='' OR ('" & 狀態 & "'='過期' AND (REPLACE(年限,'年','')*12)<=DateDiff(m,購置日期,GETDATE())) OR ('" & 狀態 & "'='未過期' AND NOT((REPLACE(年限,'年','')*12)<=DateDiff(m,購置日期,GETDATE())))) ORDER BY 分群,項"
            '全輸出，之後可能加入頁數、號數區間查詢
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
        D_Width=10'資料範圍行
        Data_Height_H=16'資料最高
        Data_Height_L=0'資料最底
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
        Dim data_dv2 As Data.DataView'算各自分群
        data.ConnectionString = con_14
        data.SelectCommAnd = "SELECT 分群,Sum(總計) AS 各自總計 FROM 設備統計分群 Where 分群 IN(SELECT distinct 分群 FROM 設備 Where 分群 Is Not NULL AND ('" & 狀態 & "'=''OR ('" & 狀態 & "'='過期' AND (REPLACE(年限,'年','')*12)<=DateDiff(m,購置日期,GETDATE())) OR ('" & 狀態 & "'='未過期' AND NOT((REPLACE(年限,'年','')*12)<=DateDiff(m,購置日期,GETDATE()))))) Group By 分群"
        data_dv2 = data.Select(New DataSourceSelectArguments)
        Dim 總頁數 As Int32 = 0
        Dim Data_Row As Int32 =10
        ' Data_Row=10'平均資料列
        Dim DataCount As Int32=data_dv.Count+data_dv2.Count'9+1
        If (DataCount MOD 9)>0'8+2,9+1+2
            DataCount = DataCount+((DataCount\9)+1)'9+1=10
        ELSE
            DataCount = DataCount+(DataCount\9)'9=9
        END If
        '總資料數
        IF DataCount>10
            If ((DataCount) Mod 10)>0  '8
                總頁數=((DataCount)\10)+1'"/"為整數除法，會四捨五入，"\"為除法，捨棄餘數
            ELSE
                總頁數=(DataCount)\10
            END If
        Else
            總頁數=1
        END If
        For i = 2 To 總頁數'制定範圍並複製，範圍為頁數
            xlWorkSheet.Range(xlWorkSheet.Cells(D_Height * i - (D_Height-1), 1), xlWorkSheet.Cells(D_Height * i, D_Width)).Value(11) = xlWorkSheet.Range(xlWorkSheet.Cells(1, 1), xlWorkSheet.Cells(D_Height, D_Width)).Value(11)'(27,1) (52,10) = (1, 1) (26, 10)
            xlWorkSheet.Range(xlWorkSheet.Cells(D_Height * i - Data_Height_H, 1), xlWorkSheet.Cells(D_Height * i - Data_Height_L, D_Width)).RowHeight = 33'(33,1) (47,10) 高度為33 7~21
            xlWorkSheet.Rows(D_Height * i - (D_Height-1)).PageBreak = xlPageBreakManual'列23從開始載入
        Next
        Dim arr As Object = xlWorkSheet.Range(xlWorkSheet.Cells(1, 1), xlWorkSheet.Cells(總頁數 * D_Height, D_Width)).Value'(1,1) (網頁頁數25,10)
        Dim arr2 As Object = xlWorkSheet.Range(xlWorkSheet.Cells(1, 1), xlWorkSheet.Cells(總頁數 * D_Height, D_Width))
        '已經做出第一頁，中間
        Dim i1 As Long = 0 '主要資料
        Dim i2 As Long = 0 '小計資料
        Dim 數量位址 As String = ""
        Dim 小計位置 As String = ""
        Dim 總計位置 As String = ""
        Dim 影印日期 As String = ToTaiwanCalendar((DateTime.Now).ToString("yyyy/MM/dd"))
        Dim 登帳截止日期 As String = Me.登帳截止日期.Text
        For i = 0 To DataCount-1'DataCount=10
            Dim 項 As String 
            Dim 財產編號 As String 
            Dim 財產名稱 As String
            Dim 財產別名 As String
            Dim 分群 As String
            Dim 廠牌 As String
            Dim 型號 As String
            Dim 購置日期 As string
            Dim 單位 As String
            Dim 數量 As String
            Dim 年限 As String
            Dim 年數 As String 
            Dim 保管人 As String
            Dim 存置地點 As String
            If i1 < data_dv.count
                項 = data_dv(i1)("項").ToString()
                財產編號 = data_dv(i1)("財產編號").ToString()
                財產名稱 = Trim(data_dv(i1)("財產名稱").ToString())
                財產名稱 = 財產名稱.Replace(vbCrLf,"")
                財產名稱 = Trim(財產名稱)
                財產別名 = Trim(data_dv(i1)("財產別名").ToString())
                財產別名 = 財產別名.Replace(vbCrLf,"")
                財產別名 = Trim(財產別名)
                分群 = data_dv(i1)("分群").ToString()
                廠牌 = data_dv(i1)("廠牌").ToString()
                型號 = data_dv(i1)("型號").ToString()
                購置日期 = data_dv(i1)("購置日期").ToString()
                購置日期 = ToTaiwanCalendar(購置日期)
                單位 = data_dv(i1)("單位").ToString()
                數量 = data_dv(i1)("數量").ToString()
                數量 = If (數量 = "", "", CLng(數量).ToString("n0"))
                年限 = Trim(data_dv(i1)("年限").ToString())
                年限 = 年限.Replace(vbCrLf,"")
                年限 = Trim(年限)
                年數 = data_dv(i1)("年數").ToString()
                保管人 = data_dv(i1)("保管人").ToString()
                存置地點 = data_dv(i1)("存置地點").ToString()
            END If
            'i為第幾個資料、j為輸出的格子
            Dim j As Long '輸出位置
            Dim 第幾頁 As Int32
            'Dim I_dataCount AS Int32 =i1+i2
            If DataCount>Data_Row'判定第一頁能塞滿嗎?
                If ((i) \ Data_Row)<總頁數 '14-2,27-2、3,
                    第幾頁=((i)\Data_Row)+1'"/"為整數除法，會四捨五入，"\"為除法，捨棄餘數
                ELSE
                    第幾頁=(i)\Data_Row
                END If
            Else
                第幾頁=1
            End If
            If DataCount>10
                    j = ((i) \ Data_Row )*26 + ((i) Mod Data_Row)*2 + 7 '從第一列輸出
            Else
                j = (i Mod Data_Row)*2 + 7
            End If 
            ' i=0、j=7 j為列位置
            Dim 頁數 As String
            Dim 總頁 As String
            If (i=0) '每頁開頭
                頁數 ="第1頁"
                總頁 ="共"& 總頁數 &"頁"
            Else
                頁數 ="第"& 第幾頁 &"頁"
                總頁 ="共"& 總頁數 &"頁"
            End If
            If ((i) Mod Data_Row)=0 '第一筆輸出年度每頁
                arr(J-3,6)=登帳截止日期
                arr(J-4,9)=影印日期
                arr(J-3,8) = 頁數
                arr(J-3,10) = 總頁
            END If
            If i1 < Data_dv.Count AND (i Mod Data_Row<9)
                IF data_dv(i1)("分群").ToString()=data_dv2(i2)("分群").ToString()
                    arr(j, 1) = 項
                    'arr2(j, 1).Borders.Weight=2 '邊框線
                    arr(j, 2) = 財產編號
                    arr(j, 3) = 財產名稱 & Chr(10) & 財產別名
                    arr(j, 4) = 廠牌
                    arr(j+1, 4) = 型號
                    arr(j, 5) = 購置日期
                    arr(j, 6) = 單位
                    arr(j+1, 6) = 數量
                    If 數量位址=""
                        數量位址="=" & arr2(j+1, 6).Address
                    Else
                        數量位址=數量位址 & "+" & arr2(j+1, 6).Address
                    End if
                    arr(j, 7) = 年限
                    arr(j+1, 7) = 年數
                    '計算年數
                    Dim 年 AS Int32=0
                    Dim 月 As Int32=0
                    If InStr(data_dv(i1)("年數").ToString(), "年")<>0 AND InStr(data_dv(i1)("年數").ToString(), "月")<>0
                        月 = Mid(data_dv(i1)("年數").ToString(),InStr(data_dv(i1)("年數").ToString(), "年")+1,InStr(data_dv(i1)("年數").ToString(), "月")-InStr(data_dv(i1)("年數").ToString(), "年")-1)
                        年 = LEFT(data_dv(i1)("年數").ToString(),InStr(data_dv(i1)("年數").ToString(), "年")-1)
                    ElseIF InStr(data_dv(i1)("年數").ToString(), "年")=0 AND InStr(data_dv(i1)("年數").ToString(), "月")<>0 AND data_dv(i1)("年數").ToString()<>"不到1月"
                        月 = LEFT(data_dv(i1)("年數").ToString(),InStr(data_dv(i1)("年數").ToString(), "月")-1)
                    ElseIF InStr(data_dv(i1)("年數").ToString(), "月")=0 AND InStr(data_dv(i1)("年數").ToString(), "年")<>0
                        年 = LEFT(data_dv(i1)("年數").ToString(),InStr(data_dv(i1)("年數").ToString(), "年")-1)
                    End If
                    Dim 年限I As Int32=0
                    年限I=LEFT(data_dv(i1)("年限").ToString(),InStr(data_dv(i1)("年限").ToString(), "年")-1)'載入年限
                    月=月+年*12'年數換算單位為"月"
                    '比較年限(換算單位為月，年*12)與年數(單位為月)
                    If (年限I*12)<=月 AND Me.過期資料.Checked = True
                        arr2(j+1, 7).Interior.ColorIndex = 3 '背景色，紅色
                    End If
                    arr(j, 8) = 保管人
                    arr(j, 9) = 存置地點
                    i1=i1+1
                Else
                    Dim 總計項 As Int32=data_dv(i1-1)("項").ToString()+1
                    arr(j, 1) = 總計項 & " " & data_dv2(i2)("分群").ToString() & "小計"
                    arr2(j, 1).HorizontalAlignment = -4152'靠右對齊
                    xlWorkSheet.Range(arr2(j, 1),arr2(j+1, 5)).Merge
                    arr(j, 6) = 數量位址
                    arr2(j, 6).VerticalAlignment = -4108'垂直置中對齊
                    xlWorkSheet.Range(arr2(j, 6),arr2(j+1, 6)).Merge
                    xlWorkSheet.Range(arr2(j, 7),arr2(j+1, 10)).Merge
                    數量位址=""
                        If 總計位置=""
                            總計位置 = "=" & arr2(j, 6).Address
                        Else
                            總計位置=總計位置 & "+" & arr2(j, 6).Address
                        End if
                    i2=i2+1
                End if
            ElseIf i2 < Data_dv2.Count AND (i Mod Data_Row<9)
                Dim 總計項 As Int32 =data_dv(i1-1)("項").ToString()+1
                arr(j, 1) = 總計項 & " " & data_dv2(i2)("分群").ToString() & "小計"
                arr2(j, 1).HorizontalAlignment = -4152'靠右對齊
                'Range.Merge 合併儲存格
                xlWorkSheet.Range(arr2(j, 1),arr2(j+1, 5)).Merge
                arr(j, 6) = 數量位址
                arr2(j, 6).VerticalAlignment = -4108'垂直置中對齊
                xlWorkSheet.Range(arr2(j, 6),arr2(j+1, 6)).Merge
                xlWorkSheet.Range(arr2(j, 7),arr2(j+1, 10)).Merge
                If 總計位置=""
                    總計位置 = "=" & arr2(j, 6).Address
                Else
                    總計位置=總計位置 & "+" & arr2(j, 6).Address
                End if
                i2=i2+1
            End If
            IF(i Mod Data_Row=9) OR (i=DataCount-1)
                arr(j, 1) = "總計"
                arr2(j, 1).HorizontalAlignment = -4152'靠右對齊
                xlWorkSheet.Range(arr2(j, 1),arr2(j+1, 5)).Merge
                arr(j, 6) = 總計位置
                arr2(j, 6).VerticalAlignment = -4108'垂直置中對齊
                xlWorkSheet.Range(arr2(j, 6),arr2(j+1, 6)).Merge
                xlWorkSheet.Range(arr2(j, 7),arr2(j+1, 10)).Merge
            End If
        Next
        xlWorkSheet.Range(xlWorkSheet.Cells(1, 1), xlWorkSheet.Cells(總頁數 * D_Height, D_Width)).Value = arr
        xlWorkSheet.Range(xlWorkSheet.Cells(1, 1), xlWorkSheet.Cells(總頁數 * D_Height, D_Width)).Rows.AutoFit'ˋ動態調整高度，但圖片會受影響
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
        Dim downloadfilename = "水電設備統計表.xlsx"
        Response.AddHeader("Content-Disposition", "attachment;FileName=" + Uri.EscapeDataString(downloadfilename))
        Response.WriteFile(MyExcel)
        System.Web.HttpContext.Current.ApplicationInstance.CompleteRequest()
        Response.Flush()
        System.IO.File.Delete(MyExcel)
        Response.End()
    End Sub
    Protected Sub GridView1_DataBound(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.DataBound
        data.SelectCommand = "select Distinct 分群 from 設備統計分群 Where 分群<>'' OR 分群 IS NOT NULL  order by 分群"
        data_dv = data.Select(New DataSourceSelectArguments)
        For i = 0 To Me.GridView1.Rows.Count - 1
            Dim id As String = CType(Me.GridView1.Rows(i).FindControl("id"), TextBox).Text
            data.SelectCommand = "select * from 設備 Where id='"& id &"'"
            Dim data_dv2 As Data.DataView = data.Select(New DataSourceSelectArguments)
            Dim 分群1 As DropDownList = CType(Me.GridView1.Rows(i).FindControl("分群"), DropDownList)
            Dim 分群S1 As String = data_dv2(0)("分群").ToString()
            分群1.Items.Clear()
            分群1.Items.Add("")
            分群1.Items(0).Value = ""
            For j = 0 To data_dv.Count - 1
                Dim 分群名稱 As String = data_dv(j)(0)
                分群1.Items.Add(分群名稱)
                分群1.Items(j+1).Value = 分群名稱
            Next
            分群1.SelectedIndex=分群1.Items.IndexOf(分群1.Items.FindByValue(分群S1))
        Next
    End Sub
    Protected Sub GridView1_RowCommand(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles GridView1.RowCommand
        If e.CommandName = "刪除"
            '連同維護紀錄作業的相關資料一併刪除
            Update(sender, e)
            Dim i As Long = e.CommandSource.NamingContainer.RowIndex
            Dim id As String = CType(Me.GridView1.Rows(i).FindControl("id"), TextBox).Text
            data.UpdateCommand = "Update 設備 Set " & _
            "項 = NULL," & _
            "財產編號 = NULL," & _
            "財產名稱 = NULL," & _
            "財產別名 = NULL," & _
            "分群 = NULL," & _
            "廠牌 = NULL," & _
            "型號 = NULL," & _
            "購置日期 = NULL," & _
            "單位 = NULL," & _
            "數量 = NULL," & _
            "年限 = '2年'," & _
            "年數 = NULL," & _
            "保管人 = NULL," & _
            "存置地點 = NULL " & _
            "WHERE id='" & id & "'"
            data.Update()
            Me.GridView1.DataBind()
        End If
    End Sub 
End Class