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
Partial Class 核銷支出明細備查簿
    Inherits System.Web.UI.Page
    Dim con_14 As String = System.Web.Configuration.WebConfigurationManager.ConnectionStrings("ApplicationServices").ConnectionString
    Dim data As New SqlDataSource
    Dim data_dv As Data.DataView
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        data.ConnectionString = con_14
        If Not Page.IsPostBack Then
            Me.年.Text = (DateTime.Now.Year - 1911).ToString()
            月1_SelectedIndexChanged(sender,e)
            日1.SelectedValue="1"
            月2_SelectedIndexChanged(sender,e)
            日2.SelectedValue="31"
            Me.GridView1.PageIndex = 0
        Else
            Me.Label1.Text = ""
            Me.Label2.Text = ""
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
    Protected Sub Download(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim MyGUID As String = Guid.NewGuid().ToString("N")
        Dim MyExcel As String = MapPath(".\Excel\Temp\") & MyGUID & ".xls"
        Select Case _種類.Text
            Case "A"
                System.IO.File.Copy(MapPath(".\Excel\零用金報銷-秘A.xls"), MyExcel)
            Case "B"
                System.IO.File.Copy(MapPath(".\Excel\報銷單-秘B.xls"), MyExcel)
            Case "XZ"
                System.IO.File.Copy(MapPath(".\Excel\暫付轉正-秘XZ.xlsx"), MyExcel)
        End Select 
        Dim xlApp As New Excel.ApplicationClass()
        xlApp.DisplayAlerts = False
        xlApp.ScreenUpdating = false
        xlApp.EnableEvents = false
        Dim xlWorkBook As Excel.Workbook = xlApp.Workbooks.Open(MyExcel, 0, False, 5, "", "", True, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", False, False, 0, True, 1, 0)
        Dim xlWorkSheet As Excel.Worksheet
        Select Case _種類.Text
            Case "A"
                xlWorkSheet = CType(xlWorkBook.Sheets(" "), Excel.Worksheet)
            Case "B"
                xlWorkSheet = CType(xlWorkBook.Sheets("電子採購"), Excel.Worksheet)
            Case "XZ"
                xlWorkSheet = CType(xlWorkBook.Sheets("暫付轉正-111年"), Excel.Worksheet)
        End Select 
        xlWorkSheet.Activate()
        Dim 年 As String = Me.年.Text
        data.ConnectionString = con_14
        Dim 是否勾選 As Boolean = false
        Dim id_array As String = "(NULL"
        For i = 0 to Me.GridView1.Rows.Count - 1
            Dim id As String = CType(Me.GridView1.Rows(i).FindControl("id"), Label).Text
            If CType(Me.GridView1.Rows(i).FindControl("勾選下載"), CheckBox).Checked=True'需勾選以選取要下載的資料
                id_array = id_array & ", " &  id
                是否勾選=true
            End If 
        Next
        Dim 號數1 As String = Me.號數1.Text
        Dim 號數2 As String = Me.號數2.Text
        Dim data_dv2 As Data.DataView'計算號數多少筆
        Dim data_dv3 As Data.DataView'計算有借金的支出-收入 只有A要做
        If 是否勾選=false
            data.SelectCommand = "SELECT id,月,日,摘要,支出,號數,商號 FROM 收支備查簿 WHERE _種類 = '" & Me._種類.text & "' And 號數 Is Not NULL And 支出>0 " & _
            "AND ((''=TRIM('" & 號數1 & "') OR ''=TRIM('" & 號數2 & "'))" & _
            "OR ( 號數 BETWEEN " & _
            "SUBSTRING(TRIM('" & 號數1 & "'), PATINDEX('%[^0]%', TRIM('" & 號數1 & "')), 3) AND " & _
            "SUBSTRING(TRIM('" & 號數2 & "'), PATINDEX('%[^0]%', TRIM('" & 號數2 & "')), 3)))" & _
            "order by 號數,_頁, _列"
            data_dv = data.Select(New DataSourceSelectArguments)
            '算有多少號數
            data.ConnectionString = con_14
            data.SelectCommand = "SELECT count(id) As 筆數 FROM 收支備查簿 WHERE _種類 = '" & Me._種類.text & "' And 號數 Is Not NULL And 支出>0 " & _
            "AND ((''=TRIM('" & 號數1 & "') OR ''=TRIM('" & 號數2 & "'))" & _
            "OR ( 號數 BETWEEN " & _
            "SUBSTRING(TRIM('" & 號數1 & "'), PATINDEX('%[^0]%', TRIM('" & 號數1 & "')), 3) AND " & _
            "SUBSTRING(TRIM('" & 號數2 & "'), PATINDEX('%[^0]%', TRIM('" & 號數2 & "')), 3)))" & _
            "group by 號數"
            data_dv2 = data.Select(New DataSourceSelectArguments)
        Else
            id_array = id_array & ")"
            data.SelectCommand = "SELECT id,月,日,摘要,支出,號數,商號 FROM 收支備查簿 WHERE _種類 = '" & Me._種類.text & "' And id in " & id_array & " order by 號數,_頁, _列"
            data_dv = data.Select(New DataSourceSelectArguments)
            '算有多少號數
            data.ConnectionString = con_14
            data.SelectCommand = "SELECT count(id) As 筆數 FROM 收支備查簿 WHERE _種類 = '" & Me._種類.text & "' And id in " & id_array & " group by 號數"
            data_dv2 = data.Select(New DataSourceSelectArguments)
        End If 
        '算借金，怕有問題的狀況(忘記打繳回?)
        '特例:將行政訴訟暫時排除、將"春節連假服務台之班人員餐費"做另外處裡，請芳瑜修正之後再更正此程式
        If _種類.Text="A"
            data.ConnectionString = con_14
            data.SelectCommand = "Select a.id As id, a.號數 As 號數,(a.支出-b.收入)As 支出2 From 收支備查簿 As a,收支備查簿 As b Where " & _
                "a._種類='A' AND " & _
                "(left(REPLACE(REPLACE(a.摘要,' ',''),CHAR(13)+CHAR(10),''), 10) Like substring(REPLACE(REPLACE(b.摘要,' ',''),CHAR(13)+CHAR(10),''), 3, 10) AND " & _
                "b.摘要 Like'_回%' AND " & _
                "(a.號數=b.號數 Or b.號數 Is NULL) AND " & _
                "a.摘要 not Like '行政訴訟%' ) OR " & _
                "(a.id='417' And b.id='499')"
            data_dv3 = data.Select(New DataSourceSelectArguments)
        End If
        dim no As Int32=0
        If data_dv2.count>0
            Dim 單頁筆數 As Int32
            Dim D_Height As int32
            Dim D_Width As int32
            Dim Data_Height_H As int32
            Dim Data_Height_L As int32
            Select Case _種類.Text
            Case "A"
                單頁筆數=7
                D_Height=32
                D_Width=11
                Data_Height_H=25
                Data_Height_L=19
            Case "B"
                單頁筆數=5
                D_Height=30
                D_Width=8
                Data_Height_H=25
                Data_Height_L=19
            Case "XZ"
                單頁筆數=4
                D_Height=22
                D_Width=8
                Data_Height_H=25
                Data_Height_L=19
            End Select 
            For i = 0 to data_dv2.count-1
                If data_dv2(i)("筆數").ToString()>單頁筆數
                    no=no+Math.Floor(data_dv2(i)("筆數").ToString()/(單頁筆數+1))'"/"為整數除法，會四捨五入，"\"為除法，捨棄餘數 PS:注意此段是否有錯誤
                End If 
            Next
            For i = 2 To data_dv2.Count+no '算張數
                xlWorkSheet.Range(xlWorkSheet.Cells(D_Height * i - (D_Height-1), 1), xlWorkSheet.Cells(D_Height * i, D_Width)).Value(11) = xlWorkSheet.Range(xlWorkSheet.Cells(1, 1), xlWorkSheet.Cells(D_Height, D_Width)).Value(11)'(27,1) (52,12) = (1, 1) (26, 12)
                For j = 1 to D_Height-1'RowHeight回傳範圍時，值只有第一列的高度
                    xlWorkSheet.Range(xlWorkSheet.Cells(D_Height * i - j, 1), xlWorkSheet.Cells(D_Height * i - j, 8)).RowHeight = xlWorkSheet.Range(xlWorkSheet.Cells(D_Height-j, 1), xlWorkSheet.Cells(D_Height-j, 8)).RowHeight
                Next
                xlWorkSheet.Rows(D_Height * i - (D_Height-1)).PageBreak = xlPageBreakManual'列27從開始載入
            Next
            Dim arr As Object = xlWorkSheet.Range(xlWorkSheet.Cells(1, 1), xlWorkSheet.Cells((data_dv2.Count+no) * D_Height, D_Width)).Value'範圍
            Dim arr2 As Object = xlWorkSheet.Range(xlWorkSheet.Cells(1, 1), xlWorkSheet.Cells((data_dv2.Count+no) * D_Height, D_Width))'範圍
            'dim 支出sum As Int32 = 0
            Dim 張 As Long =0
            Select Case _種類.Text
            Case "A"
                For i = 0 To data_dv.Count - 1'輸出資料，A、B、XZ都長得不一樣
                    Dim 月 As String =  data_dv(i)("月").ToString()
                    Dim 日 As String =  data_dv(i)("日").ToString()
                    Dim 上月 As String
                    Dim 上日 As String
                    Dim 摘要 As String = data_dv(i)("摘要").ToString()
                    Dim 支出 As Int32 = data_dv(i)("支出").ToString()
                    dim 位置 As Long
                    Dim 年度 As String = 年 &"年"
                    dim 號數i As Int32 = data_dv(i)("號數").ToString()
                    dim 號數 As string = 號數i.ToString("D3")
                    dim 秘字 As string = "秘" & me._種類.text & "字第" & 號數 & "號"
                    dim 秘字2 As string = "秘" & 號數
                    dim 變動 As Boolean = false
                    For i2=0 To data_dv3.Count - 1
                        If data_dv(i)("號數").ToString()=data_dv3(i2)("號數").ToString() AND data_dv(i)("id").ToString()=data_dv3(i2)("id").ToString()'被收尋字串.IndexOf(收尋字串)
                            支出=data_dv3(i2)("支出2").ToString()
                        ENd if
                    Next
                    If i>0 '非第一筆
                        If data_dv(i)("號數").ToString()<>data_dv(i-1)("號數").ToString() or 位置=6'換號數、換頁
                            變動=true
                            張=張+1
                            位置=0
                        Else
                        位置=位置+1
                        End If 
                    Else'第一筆
                        位置 =(i Mod 7)
                    End If 
                    Dim j As Long = 32 * 張 + 位置 + 6'32*張+第幾個+初始位址
                    If i =0 or 變動=true'第一頁
                        arr(J-3,1) = 年度
                        arr(J-3,10) = 秘字
                        arr(J+13,1) = 年度
                        arr(J+13,10) = 秘字
                    END If
                    'If 過審=true
                    If 上月<>月 OR 上日<>日 OR 變動=true
                        arr(j, 1) = 月
                        arr(j, 2) = 日
                        arr(J, 3) = 秘字2
                        arr(j+16, 1) = 月
                        arr(j+16, 2) = 日
                        arr(J+16, 3) = 秘字2
                        上月=月
                        上日=日
                    End If
                    arr2(j,1).Rows.AutoFit
                    arr2(j+16,1).Rows.AutoFit
                    arr(j, 5) = 摘要
                    arr(j, 6) = 支出
                    arr(j+16, 5) = 摘要
                    arr(j+16, 6) = 支出
                    ' 支出sum=支出sum+支出
                    If i<data_dv.Count - 1'最後一筆或換頁，輸出資料，但非最後一筆且不換頁，則不動 ---'8/24，第一筆的總和在第二筆
                        If data_dv(i)("號數").ToString()<>data_dv(i+1)("號數").ToString() or 位置=6'號碼有變動，換頁
                            arr(j+7-位置,6) = "=SUM(" & arr2(j+6-位置,6).Address & ":" & arr2(j-位置,6).Address & ")"
                            arr(j+16+7-位置,6) = "=SUM(" & arr2(j+16+6-位置,6).Address & ":" & arr2(j+16-位置,6).Address & ")"
                        End If 
                    Else'最後一筆
                        arr(j+7-位置,6) = "=SUM(" & arr2(j+6-位置,6).Address & ":" & arr2(j-位置,6).Address & ")"
                        arr(j+16+7-位置,6) = "=SUM(" & arr2(j+16+6-位置,6).Address & ":" & arr2(j+16-位置,6).Address & ")"
                    END If
                Next
            Case "B"
                For i = 0 To data_dv.Count - 1'輸出資料
                    Dim 月 As String =  data_dv(i)("月").ToString()
                    Dim 日 As String =  data_dv(i)("日").ToString()
                    Dim 上月 As String
                    Dim 上日 As String
                    Dim 摘要 As String = data_dv(i)("摘要").ToString()
                    Dim 支出 As Int32 = data_dv(i)("支出").ToString()
                    dim 位置 As Long
                    Dim 年度 As String = 年 &"年"
                    dim 號數i As Int32 = data_dv(i)("號數").ToString()
                    dim 號數 As string = 號數i.ToString("D3")
                    dim 秘字 As string = "秘" & me._種類.text & "字第" & 號數 & "號"
                    dim 秘字2 As string = "秘" & 號數
                    dim 變動 As Boolean = false
                    Dim 商號 As String = data_dv(i)("商號").ToString()
                    If i>0 '非第一筆
                        If data_dv(i)("號數").ToString()<>data_dv(i-1)("號數").ToString() or 位置=4'換號數、換頁
                            變動=true
                            張=張+1
                            位置=0
                        Else
                        位置=位置+1
                        End If 
                    Else'第一筆
                        位置 =(i Mod 5)
                    End If 
                    Dim j As Long = 30 * 張 + 位置 + 5'32*張+第幾個+初始位址
                    If i=0 or 變動=true'第一頁
                        arr(J-3,1) = 年度
                        arr(J-3,8) = 秘字
                        arr(J+12,1) = 年度
                        arr(J+12,8) = 秘字
                        arr(j, 1) = 月
                        arr(j, 2) = 日
                        arr(j, 3) = 商號
                        arr(j, 5) = 摘要
                        arr(j, 6) = 支出
                        arr(j+15, 1) = 月
                        arr(j+15, 2) = 日
                        arr(j+15, 3) = 商號
                        arr(j+15, 5) = 摘要
                        arr(j+15, 6) = 支出
                    END If
                    If 上月<>月 OR 上日<>日 OR 變動=true
                        arr(j, 1) = 月
                        arr(j, 2) = 日
                        arr(j, 3) = 商號
                        arr(j+15, 1) = 月
                        arr(j+15, 2) = 日
                        arr(j+15, 3) = 商號
                        上月=月
                        上日=日
                    End If
                    'If 過審=true
                    arr(j, 5) = 摘要
                    arr(j, 6) = 支出
                    arr(j+15, 5) = 摘要
                    arr(j+15, 6) = 支出
                    arr2(j,3).Rows.AutoFit
                    arr2(j+15,3).Rows.AutoFit
                    If i<data_dv.Count - 1'最後一筆或換頁，輸出資料，但非最後一筆且不換頁，則不動 ---'8/24，第一筆的總和在第二筆
                        If data_dv(i)("號數").ToString()<>data_dv(i+1)("號數").ToString()or 位置=4'號碼有變動，換頁
                            arr(j+7-位置,6) = "=SUM(" & arr2(j+6-位置,6).Address & ":" & arr2(j-位置,6).Address & ")"
                            arr(j+16+6-位置,6) = "=SUM(" & arr2(j+16+5-位置,6).Address & ":" & arr2(j+15-位置,6).Address & ")"
                        End If 
                    Else'最後一筆
                        arr(j+7-位置,6) = "=SUM(" & arr2(j+6-位置,6).Address & ":" & arr2(j-位置,6).Address & ")"
                        arr(j+16+6-位置,6) = "=SUM(" & arr2(j+16+5-位置,6).Address & ":" & arr2(j+15-位置,6).Address & ")"
                    END If
                Next
            Case "XZ"
                For i = 0 To data_dv.Count - 1'輸出資料
                    Dim 月 As String =  data_dv(i)("月").ToString()
                    Dim 日 As String =  data_dv(i)("日").ToString()
                    Dim 上月 As String
                    Dim 上日 As String
                    Dim 摘要 As String = data_dv(i)("摘要").ToString()
                    Dim 支出 As Int32 = data_dv(i)("支出").ToString()
                    dim 位置 As Long 
                    Dim 年度 As String = 年 &"年"
                    dim 號數i As Int32 = data_dv(i)("號數").ToString()
                    dim 號數 As string = 號數i.ToString("D3")
                    dim 秘字 As string = "秘" & me._種類.text & "字第" & 號數 & "號"
                    dim 秘字2 As string = "秘" & 號數
                    dim 變動 As Boolean = false
                    If i>0 '非第一筆
                        If data_dv(i)("號數").ToString()<>data_dv(i-1)("號數").ToString() or 位置=4'換號數、換頁
                            變動=true
                            張=張+1
                            位置=0
                        Else
                        位置=位置+1
                        End If 
                    Else'第一筆
                        位置 =(i Mod 5)
                    End If 
                    Dim j As Long = 22 * 張 + 位置 + 5'32*張+第幾個+初始位址
                    If i=0 or 變動=true'第一頁
                        arr(J-3,1) = 年度
                        arr(J-3,8) = 秘字
                        arr(J+8,1) = 年度
                        arr(J+8,8) = 秘字
                        arr(j, 1) = 月
                        arr(j, 2) = 日
                        arr(j, 5) = 摘要
                        arr(j, 6) = 支出
                        arr(j+11, 1) = 月
                        arr(j+11, 2) = 日
                        arr(j+11, 5) = 摘要
                        arr(j+11, 6) = 支出
                    END If
                    If 上月<>月 OR 上日<>日 OR 變動=true
                        arr(j, 1) = 月
                        arr(j, 2) = 日
                        arr(j+11, 1) = 月
                        arr(j+11, 2) = 日
                        上月=月
                        上日=日
                    End If
                    'If 過審=true
                    arr(j, 5) = 摘要
                    arr(j, 6) = 支出
                    arr(j+11, 5) = 摘要
                    arr(j+11, 6) = 支出
                    arr2(j,5).Rows.AutoFit
                    arr2(j+11,5).Rows.AutoFit
                    If i<data_dv.Count - 1'最後一筆或換頁，輸出資料，但非最後一筆且不換頁，則不動 ---'8/24，第一筆的總和在第二筆
                        If data_dv(i)("號數").ToString()<>data_dv(i+1)("號數").ToString() or 位置=4'號碼有變動，換頁
                            arr(j+4-位置,6) = "=SUM(" & arr2(j+3-位置,6).Address & ":" & arr2(j-位置,6).Address & ")"
                            arr(j+11+4-位置,6) = "=SUM(" & arr2(j+11+3-位置,6).Address & ":" & arr2(j+11-位置,6).Address & ")"
                        End If 
                    Else'最後一筆
                        arr(j+4-位置,6) = "=SUM(" & arr2(j+3-位置,6).Address & ":" & arr2(j-位置,6).Address & ")"
                        arr(j+11+4-位置,6) = "=SUM(" & arr2(j+11+3-位置,6).Address & ":" & arr2(j+11-位置,6).Address & ")"
                    END If
                Next
            End Select 
            xlWorkSheet.Range(xlWorkSheet.Cells(1, 1), xlWorkSheet.Cells((data_dv2.Count+no) * D_Height, D_Width)).Value = arr
            xlWorkBook.Save()
        End If 
        xlWorkBook.Close()
        xlApp.Quit()
        ReleaseObject(xlWorkSheet)
        ReleaseObject(xlWorkBook)
        ReleaseObject(xlApp)
        Response.Clear()
        Response.ClearHeaders()
        Response.Buffer = True
        Response.ContentType = "application/octet-stream"
        Dim downloadfilename As string
        Select Case _種類.Text
            Case "A"
                downloadfilename = "零用金報銷-秘A(" & 年 & "年度).xls"
            Case "B"
                downloadfilename = "-報銷單-注意銀河系-秘B(110年度).xls"
            Case "XZ"
                downloadfilename = "暫付轉正-秘XZ.xlsx"
        End Select 
        Response.AddHeader("Content-Disposition", "attachment;FileName=" + Uri.EscapeDataString(downloadfilename))
        Response.WriteFile(MyExcel)
        System.Web.HttpContext.Current.ApplicationInstance.CompleteRequest()
        Response.Flush()
        System.IO.File.Delete(MyExcel)
        Response.End()
        '以上為合併版，!此段請勿刪除!
        '     If(Me._種類.text="A")
        '         'A
        '         Dim MyGUID As String = Guid.NewGuid().ToString("N")
        '         Dim MyExcel As String = MapPath(".\Excel\Temp\") & MyGUID & ".xls"
        '         System.IO.File.Copy(MapPath(".\Excel\零用金報銷-秘A.xls"), MyExcel)
        '         Dim xlApp As New Excel.ApplicationClass()
        '         xlApp.DisplayAlerts = False
        '         xlApp.ScreenUpdating = false
        '         xlApp.EnableEvents = false
        '         Dim xlWorkBook As Excel.Workbook = xlApp.Workbooks.Open(MyExcel, 0, False, 5, "", "", True, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", False, False, 0, True, 1, 0)
        '         Dim xlWorkSheet As Excel.Worksheet
        '         xlWorkSheet = CType(xlWorkBook.Sheets(" "), Excel.Worksheet)
        '         xlWorkSheet.Activate()
        '         Dim 年 As String = Me.年.Text
        '         data.ConnectionString = con_14
        '         Dim 是否勾選 As Boolean = false
        '         Dim id_array As String = "(NULL"
        '         For i = 0 to Me.GridView1.Rows.Count - 1
        '             Dim id As String = CType(Me.GridView1.Rows(i).FindControl("id"), Label).Text
        '             If CType(Me.GridView1.Rows(i).FindControl("勾選下載"), CheckBox).Checked=True'需勾選以選取要下載的資料
        '                 id_array = id_array & ", " &  id
        '                 是否勾選=true
        '             End If 
        '         Next
        '         Dim data_dv2 As Data.DataView'計算號數多少筆
        '         Dim data_dv3 As Data.DataView'計算有借金的支出-收入
        '         If 是否勾選=false
        '             data.SelectCommand = "SELECT * FROM 收支備查簿 WHERE _種類 = '" & Me._種類.text & "' And 號數 Is Not NULL And 支出>0 order by 號數,_頁, _列"
        '             data_dv = data.Select(New DataSourceSelectArguments)
        '             '算有多少號數
        '             data.ConnectionString = con_14
        '             data.SelectCommand = "SELECT count(*) As 筆數 FROM 收支備查簿 WHERE _種類 = '" & Me._種類.text & "' And 號數 Is Not NULL And 支出>0 group by 號數"
        '             data_dv2 = data.Select(New DataSourceSelectArguments)
        '         Else
        '             id_array = id_array & ")"
        '             data.SelectCommand = "SELECT * FROM 收支備查簿 WHERE _種類 = '" & Me._種類.text & "' And id in " & id_array & " order by 號數,_頁, _列"
        '             data_dv = data.Select(New DataSourceSelectArguments)
        '             '算有多少號數
        '             data.ConnectionString = con_14
        '             data.SelectCommand = "SELECT count(*) As 筆數 FROM 收支備查簿 WHERE _種類 = '" & Me._種類.text & "' And id in " & id_array & " group by 號數"
        '             data_dv2 = data.Select(New DataSourceSelectArguments)
        '         End If 
        '         '算借金，怕有問題的狀況(忘記打繳回?)
        '         '特例:將行政訴訟暫時排除、將"春節連假服務台之班人員餐費"做另外處裡，請芳瑜修正之後再更正此程式
        '         data.ConnectionString = con_14
        '         data.SelectCommand = "Select a.id As id, a.號數 As 號數,(a.支出-b.收入)As 支出2 From 收支備查簿 As a,收支備查簿 As b Where " & _
        '             "a._種類='A' AND " & _
        '             "(left(REPLACE(REPLACE(a.摘要,' ',''),CHAR(13)+CHAR(10),''), 10) Like substring(REPLACE(REPLACE(b.摘要,' ',''),CHAR(13)+CHAR(10),''), 3, 10) AND " & _
        '             "b.摘要 Like'_回%' AND " & _
        '             "(a.號數=b.號數 Or b.號數 Is NULL) AND " & _
        '             "a.摘要 not Like '行政訴訟%' ) OR " & _
        '             "(a.id='417' And b.id='499')"
        '         data_dv3 = data.Select(New DataSourceSelectArguments)
        '         dim no As Int32=0
        '         If data_dv2.count>0
        '             For i = 0 to data_dv2.count-1
        '                 If data_dv2(i)("筆數").ToString()>7
        '                     no=no+Math.Floor(data_dv2(i)("筆數").ToString()/8)'"/"為整數除法，會四捨五入，"\"為除法，捨棄餘數 PS:注意此段是否有錯誤
        '                 End If 
        '             Next
        '             For i = 2 To data_dv2.Count+no '算張數
        '                 xlWorkSheet.Range(xlWorkSheet.Cells(32 * i - 31, 1), xlWorkSheet.Cells(32 * i, 11)).Value(11) = xlWorkSheet.Range(xlWorkSheet.Cells(1, 1), xlWorkSheet.Cells(32, 11)).Value(11)
        '                 xlWorkSheet.Range(xlWorkSheet.Cells(32 * i - 25, 1), xlWorkSheet.Cells(32 * i - 19, 11)).RowHeight = 32
        '                 xlWorkSheet.Rows(32 * i - 31).PageBreak = xlPageBreakManual
        '             Next
        '             Dim arr As Object = xlWorkSheet.Range(xlWorkSheet.Cells(1, 1), xlWorkSheet.Cells((data_dv2.Count+no) * 32, 11)).Value'範圍
        '             Dim arr2 As Object = xlWorkSheet.Range(xlWorkSheet.Cells(1, 1), xlWorkSheet.Cells((data_dv2.Count+no) * 32, 11))'範圍
        '             'dim 支出sum As Int32 = 0
        '             Dim 張 As Long =0
        '             For i = 0 To data_dv.Count - 1'輸出資料
        '                 'Dim 過審 As String = data_dv(i)("過審").ToString()
        '                 Dim 月 As String =  data_dv(i)("月").ToString()
        '                 Dim 日 As String =  data_dv(i)("日").ToString()
        '                 Dim 上月 As String
        '                 Dim 上日 As String
        '                 Dim 摘要 As String = data_dv(i)("摘要").ToString()
        '                 Dim 支出 As Int32 = data_dv(i)("支出").ToString()
        '                 dim 位置 As Long
        '                 Dim 年度 As String = 年 &"年"
        '                 dim 號數i As Int32 = data_dv(i)("號數").ToString()
        '                 dim 號數 As string = 號數i.ToString("D3")
        '                 dim 秘字 As string = "秘" & me._種類.text & "字第" & 號數 & "號"
        '                 dim 秘字2 As string = "秘" & 號數
        '                 dim 變動 As Boolean = false
        '                 For i2=0 To data_dv3.Count - 1
        '                     If data_dv(i)("號數").ToString()=data_dv3(i2)("號數").ToString() AND data_dv(i)("id").ToString()=data_dv3(i2)("id").ToString()'被收尋字串.IndexOf(收尋字串)
        '                         支出=data_dv3(i2)("支出2").ToString()
        '                     ENd if
        '                 Next
        '                 If i>0 '非第一筆
        '                     If data_dv(i)("號數").ToString()<>data_dv(i-1)("號數").ToString() or 位置=6'換號數、換頁
        '                         變動=true
        '                         張=張+1
        '                         位置=0
        '                     Else
        '                     位置=位置+1
        '                     End If 
        '                 Else'第一筆
        '                     位置 =(i Mod 7)
        '                 End If 
        '                 Dim j As Long = 32 * 張 + 位置 + 6'32*張+第幾個+初始位址
        '                 If i =0 or 變動=true'第一頁
        '                     arr(J-3,1) = 年度
        '                     arr(J-3,10) = 秘字
        '                     arr(J+13,1) = 年度
        '                     arr(J+13,10) = 秘字
        '                 END If
        '                 'If 過審=true
        '                 If 上月<>月 OR 上日<>日 OR 變動=true
        '                     arr(j, 1) = 月
        '                     arr(j, 2) = 日
        '                     arr(J, 3) = 秘字2
        '                     arr(j+16, 1) = 月
        '                     arr(j+16, 2) = 日
        '                     arr(J+16, 3) = 秘字2
        '                     上月=月
        '                     上日=日
        '                 End If
        '                 arr(j, 5) = 摘要
        '                 arr(j, 6) = 支出
        '                 arr(j+16, 5) = 摘要
        '                 arr(j+16, 6) = 支出
        '                 If i<data_dv.Count - 1'最後一筆或換頁，輸出資料，但非最後一筆且不換頁，則不動 ---'8/24，第一筆的總和在第二筆
        '                     If data_dv(i)("號數").ToString()<>data_dv(i+1)("號數").ToString() or 位置=6'號碼有變動，換頁
        '                         arr(j+7-位置,6) = "=SUM(" & arr2(j+6-位置,6).Address & ":" & arr2(j-位置,6).Address & ")"
        '                         arr(j+16+7-位置,6) = "=SUM(" & arr2(j+16+6-位置,6).Address & ":" & arr2(j+16-位置,6).Address & ")"
        '                     End If 
        '                 Else'最後一筆
        '                     arr(j+7-位置,6) = "=SUM(" & arr2(j+6-位置,6).Address & ":" & arr2(j-位置,6).Address & ")"
        '                     arr(j+16+7-位置,6) = "=SUM(" & arr2(j+16+6-位置,6).Address & ":" & arr2(j+16-位置,6).Address & ")"
        '                 END If
        '             Next
        '             xlWorkSheet.Range(xlWorkSheet.Cells(1, 1), xlWorkSheet.Cells((data_dv2.Count+no) * 32, 11)).Value = arr
        '             xlWorkBook.Save()
        '         End If 
        '         xlWorkBook.Close()
        '         xlApp.Quit()
        '         ReleaseObject(xlWorkSheet)
        '         ReleaseObject(xlWorkBook)
        '         ReleaseObject(xlApp)
        '         Response.Clear()
        '         Response.ClearHeaders()
        '         Response.Buffer = True
        '         Response.ContentType = "application/octet-stream"
        '         Dim downloadfilename = "零用金報銷-秘A(" & 年 & "年度).xls"
        '         Response.AddHeader("Content-Disposition", "attachment;FileName=" + Uri.EscapeDataString(downloadfilename))
        '         Response.WriteFile(MyExcel)
        '         System.Web.HttpContext.Current.ApplicationInstance.CompleteRequest()
        '         Response.Flush()
        '         System.IO.File.Delete(MyExcel)
        '         Response.End()
        '     Elseif(Me._種類.text="B")
        '         'B
        '         Dim MyGUID As String = Guid.NewGuid().ToString("N")
        '         Dim MyExcel As String = MapPath(".\Excel\Temp\") & MyGUID & ".xls"
        '         System.IO.File.Copy(MapPath(".\Excel\報銷單-秘B.xls"), MyExcel)
        '         Dim xlApp As New Excel.ApplicationClass()
        '         xlApp.DisplayAlerts = False
        '         xlApp.ScreenUpdating = false
        '         xlApp.EnableEvents = false
        '         Dim xlWorkBook As Excel.Workbook = xlApp.Workbooks.Open(MyExcel, 0, False, 5, "", "", True, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", False, False, 0, True, 1, 0)
        '         Dim xlWorkSheet As Excel.Worksheet
        '         xlWorkSheet = CType(xlWorkBook.Sheets("電子採購"), Excel.Worksheet)
        '         xlWorkSheet.Activate()
        '         Dim 年 As String = Me.年.Text
        '         data.ConnectionString = con_14
        '         Dim 是否勾選 As Boolean = false
        '         Dim id_array As String = "(NULL"
        '         For i = 0 to Me.GridView1.Rows.Count - 1
        '             Dim id As String = CType(Me.GridView1.Rows(i).FindControl("id"), Label).Text
        '             If CType(Me.GridView1.Rows(i).FindControl("勾選下載"), CheckBox).Checked=True
        '                 id_array = id_array & ", " &  id
        '                 是否勾選=true
        '             End If 
        '         Next
        '         Dim data_dv2 As Data.DataView'計算號數多少筆
        '         If 是否勾選=false
        '             data.SelectCommand = "SELECT * FROM 收支備查簿 WHERE _種類 = '" & Me._種類.text & "' And 號數 Is Not NULL And 支出>0 order by 號數,_頁, _列"
        '             data_dv = data.Select(New DataSourceSelectArguments)
        '             '算有多少號數
        '             data.ConnectionString = con_14
        '             data.SelectCommand = "SELECT count(*) As 筆數 FROM 收支備查簿 WHERE _種類 = '" & Me._種類.text & "' And 號數 Is Not NULL And 支出>0 group by 號數"
        '             data_dv2 = data.Select(New DataSourceSelectArguments)
        '         Else
        '             id_array = id_array & ")"
        '             data.SelectCommand = "SELECT * FROM 收支備查簿 WHERE _種類 = '" & Me._種類.text & "' And id in " & id_array & " order by 號數,_頁, _列"
        '             data_dv = data.Select(New DataSourceSelectArguments)
        '             '算有多少號數
        '             data.ConnectionString = con_14
        '             data.SelectCommand = "SELECT count(*) As 筆數 FROM 收支備查簿 WHERE _種類 = '" & Me._種類.text & "' And id in " & id_array & " group by 號數"
        '             data_dv2 = data.Select(New DataSourceSelectArguments)
        '         End If 
        '         dim no As Int32=0
        '         If data_dv2.count>0
        '             For i = 0 to data_dv2.count-1
        '                 If data_dv2(i)("筆數").ToString()>5
        '                     no=no+Math.Floor(data_dv2(i)("筆數").ToString()/6)'"/"為整數除法，會四捨五入，"\"為除法，捨棄餘數 PS:注意此段是否有錯誤
        '                 End If 
        '             Next
        '             For i = 2 To data_dv2.Count+no'算張數
        '                 xlWorkSheet.Range(xlWorkSheet.Cells(30 * i - 29, 1), xlWorkSheet.Cells(30 * i, 8)).Value(11) = xlWorkSheet.Range(xlWorkSheet.Cells(1, 1), xlWorkSheet.Cells(30, 8)).Value(11)
        '                 xlWorkSheet.Range(xlWorkSheet.Cells(30 * i - 25, 1), xlWorkSheet.Cells(30 * i - 19, 8)).RowHeight = 30
        '                 xlWorkSheet.Rows(30 * i - 14).PageBreak = xlPageBreakManual
        '             Next
        '             Dim arr As Object = xlWorkSheet.Range(xlWorkSheet.Cells(1, 1), xlWorkSheet.Cells((data_dv2.Count+no) * 30, 8)).Value
        '             Dim arr2 As Object = xlWorkSheet.Range(xlWorkSheet.Cells(1, 1), xlWorkSheet.Cells((data_dv2.Count+no) * 30, 8))'範圍
        '             dim 支出sum As Int32 = 0
        '             Dim 張 As Long =0
        '             For i = 0 To data_dv.Count - 1'輸出資料
        '                 Dim 月 As String =  DateTime.now.tostring("MM")
        '                 Dim 日 As String =  DateTime.now.tostring("dd")
        '                 Dim 摘要 As String = data_dv(i)("摘要").ToString()
        '                 Dim 支出 As Int32 = data_dv(i)("支出").ToString()
        '                 dim 位置 As Long
        '                 Dim 年度 As String = 年 &"年"
        '                 dim 號數i As Int32 = data_dv(i)("號數").ToString()
        '                 dim 號數 As string = 號數i.ToString("D3")
        '                 dim 秘字 As string = "秘" & me._種類.text & "字第" & 號數 & "號"
        '                 dim 秘字2 As string = "秘" & 號數
        '                 dim 變動 As Boolean = false
        '                 Dim 商號 As String = data_dv(i)("商號").ToString()
        '                 If i>0 '非第一筆
        '                     If data_dv(i)("號數").ToString()<>data_dv(i-1)("號數").ToString() or 位置=4'換號數、換頁
        '                         變動=true
        '                         張=張+1
        '                         位置=0
        '                     Else
        '                     位置=位置+1
        '                     End If 
        '                 Else'第一筆
        '                     位置 =(i Mod 5)
        '                 End If 
        '                 Dim j As Long = 30 * 張 + 位置 + 5'32*張+第幾個+初始位址
        '                 If i=0 or 變動=true'第一頁
        '                     arr(J-3,1) = 年度
        '                     arr(J-3,8) = 秘字
        '                     arr(J+12,1) = 年度
        '                     arr(J+12,8) = 秘字
        '                     arr(j, 1) = 月
        '                     arr(j, 2) = 日
        '                     arr(j, 3) = 商號
        '                     arr(j, 5) = 摘要
        '                     arr(j, 6) = 支出
        '                     arr(j+15, 1) = 月
        '                     arr(j+15, 2) = 日
        '                     arr(j+15, 3) = 商號
        '                     arr(j+15, 5) = 摘要
        '                     arr(j+15, 6) = 支出
        '                 END If
        '                 'If 過審=true
        '                 arr(j, 5) = 摘要
        '                 arr(j, 6) = 支出
        '                 arr(j+15, 5) = 摘要
        '                 arr(j+15, 6) = 支出
        '                 If i<data_dv.Count - 1'最後一筆或換頁，輸出資料，但非最後一筆且不換頁，則不動 ---'8/24，第一筆的總和在第二筆
        '                     If data_dv(i)("號數").ToString()<>data_dv(i+1)("號數").ToString()or 位置=4'號碼有變動，換頁
        '                         arr(j+7-位置,6) = "=SUM(" & arr2(j+6-位置,6).Address & ":" & arr2(j-位置,6).Address & ")"
        '                         arr(j+16+6-位置,6) = "=SUM(" & arr2(j+16+5-位置,6).Address & ":" & arr2(j+15-位置,6).Address & ")"
        '                     End If 
        '                 Else'最後一筆
        '                     arr(j+7-位置,6) = "=SUM(" & arr2(j+6-位置,6).Address & ":" & arr2(j-位置,6).Address & ")"
        '                     arr(j+16+6-位置,6) = "=SUM(" & arr2(j+16+5-位置,6).Address & ":" & arr2(j+15-位置,6).Address & ")"
        '                 END If
        '             Next
        '             xlWorkSheet.Range(xlWorkSheet.Cells(1, 1), xlWorkSheet.Cells((data_dv2.Count+no) * 30, 8)).Value = arr
        '             xlWorkBook.Save()
        '         End If 
        '         xlWorkBook.Close()
        '         xlApp.Quit()
        '         ReleaseObject(xlWorkSheet)
        '         ReleaseObject(xlWorkBook)
        '         ReleaseObject(xlApp)
        '         Response.Clear()
        '         Response.ClearHeaders()
        '         Response.Buffer = True
        '         Response.ContentType = "application/octet-stream"
        '         Dim downloadfilename = "-報銷單-注意銀河系-秘B(110年度).xls"
        '         Response.AddHeader("Content-Disposition", "attachment;FileName=" + Uri.EscapeDataString(downloadfilename))
        '         Response.WriteFile(MyExcel)
        '         System.Web.HttpContext.Current.ApplicationInstance.CompleteRequest()
        '         Response.Flush()
        '         System.IO.File.Delete(MyExcel)
        '         Response.End()
        '     Elseif(Me._種類.text="XZ")
        '         'XZ
        '         Dim MyGUID As String = Guid.NewGuid().ToString("N")
        '         Dim MyExcel As String = MapPath(".\Excel\Temp\") & MyGUID & ".xlsx"
        '         System.IO.File.Copy(MapPath(".\Excel\暫付轉正-秘XZ.xlsx"), MyExcel)
        '         Dim xlApp As New Excel.ApplicationClass()
        '         xlApp.DisplayAlerts = False
        '         xlApp.ScreenUpdating = false
        '         xlApp.EnableEvents = false
        '         Dim xlWorkBook As Excel.Workbook = xlApp.Workbooks.Open(MyExcel, 0, False, 5, "", "", True, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", False, False, 0, True, 1, 0)
        '         Dim xlWorkSheet As Excel.Worksheet
        '         xlWorkSheet = CType(xlWorkBook.Sheets("暫付轉正-111年"), Excel.Worksheet)
        '         xlWorkSheet.Activate()
        '         Dim 年 As String = Me.年.Text
        '         data.ConnectionString = con_14
        '         Dim 是否勾選 As Boolean = false
        '         Dim id_array As String = "(NULL"
        '         For i = 0 to Me.GridView1.Rows.Count - 1
        '             Dim id As String = CType(Me.GridView1.Rows(i).FindControl("id"), Label).Text
        '             If CType(Me.GridView1.Rows(i).FindControl("勾選下載"), CheckBox).Checked=True
        '                 id_array = id_array & ", " &  id
        '                 是否勾選=true
        '             End If 
        '         Next
        '         Dim data_dv2 As Data.DataView'計算號數多少筆
        '         If 是否勾選=false
        '             data.SelectCommand = "SELECT * FROM 收支備查簿 WHERE _種類 = '" & Me._種類.text & "' And 號數 Is Not NULL And 支出>0 order by 號數,_頁, _列"
        '             data_dv = data.Select(New DataSourceSelectArguments)
        '             '算有多少號數
        '             data.ConnectionString = con_14
        '             data.SelectCommand = "SELECT count(*) As 筆數 FROM 收支備查簿 WHERE _種類 = '" & Me._種類.text & "' And 號數 Is Not NULL And 支出>0 group by 號數"
        '             data_dv2 = data.Select(New DataSourceSelectArguments)
        '         Else
        '             id_array = id_array & ")"
        '             data.SelectCommand = "SELECT * FROM 收支備查簿 WHERE _種類 = '" & Me._種類.text & "' And id in " & id_array & " order by 號數,_頁, _列"
        '             data_dv = data.Select(New DataSourceSelectArguments)
        '             '算有多少號數
        '             data.ConnectionString = con_14
        '             data.SelectCommand = "SELECT count(*) As 筆數 FROM 收支備查簿 WHERE _種類 = '" & Me._種類.text & "' And id in " & id_array & " group by 號數"
        '             data_dv2 = data.Select(New DataSourceSelectArguments)
        '         End If 
        '         dim no As Int32=0
        '         If data_dv2.count>0
        '             For i = 0 to data_dv2.count-1
        '                 If data_dv2(i)("筆數").ToString()>4
        '                     no=no+Math.Floor(data_dv2(i)("筆數").ToString()/6)'"/"為整數除法，會四捨五入，"\"為除法，捨棄餘數 PS:注意此段是否有錯誤
        '                 End If 
        '             Next
        '             For i = 2 To data_dv2.Count+no'算張數
        '                 xlWorkSheet.Range(xlWorkSheet.Cells(22 * i - 21, 1), xlWorkSheet.Cells(22 * i, 8)).Value(11) = xlWorkSheet.Range(xlWorkSheet.Cells(1, 1), xlWorkSheet.Cells(22, 8)).Value(11)
        '                 '設定高度
        '                 For j = 1 to 21'RowHeight回傳範圍時，值只有第一列的高度
        '                     xlWorkSheet.Range(xlWorkSheet.Cells(22 * i - j, 1), xlWorkSheet.Cells(22 * i - j, 8)).RowHeight = xlWorkSheet.Range(xlWorkSheet.Cells(22-j, 1), xlWorkSheet.Cells(22-j, 8)).RowHeight
        '                 Next
        '                 xlWorkSheet.Rows(22 * i - 21).PageBreak = xlPageBreakManual
        '             Next
        '             Dim arr As Object = xlWorkSheet.Range(xlWorkSheet.Cells(1, 1), xlWorkSheet.Cells((data_dv2.Count+no) * 22, 8)).Value
        '             Dim arr2 As Object = xlWorkSheet.Range(xlWorkSheet.Cells(1, 1), xlWorkSheet.Cells((data_dv2.Count+no) * 22, 8))'範圍
        '             dim 支出sum As Int32 = 0
        '             Dim 張 As Long =0
        '             For i = 0 To data_dv.Count - 1'輸出資料
        '                 Dim 月 As String =  DateTime.now.tostring("MM")
        '                 Dim 日 As String =  DateTime.now.tostring("dd")
        '                 Dim 摘要 As String = data_dv(i)("摘要").ToString()
        '                 Dim 支出 As Int32 = data_dv(i)("支出").ToString()
        '                 dim 位置 As Long 
        '                 Dim 年度 As String = 年 &"年"
        '                 dim 號數i As Int32 = data_dv(i)("號數").ToString()
        '                 dim 號數 As string = 號數i.ToString("D3")
        '                 dim 秘字 As string = "秘" & me._種類.text & "字第" & 號數 & "號"
        '                 dim 秘字2 As string = "秘" & 號數
        '                 dim 變動 As Boolean = false
        '                 If i>0 '非第一筆
        '                     If data_dv(i)("號數").ToString()<>data_dv(i-1)("號數").ToString() or 位置=3'換號數、換頁
        '                         變動=true
        '                         張=張+1
        '                         位置=0
        '                     Else
        '                     位置=位置+1
        '                     End If 
        '                 Else'第一筆
        '                     位置 =(i Mod 4)
        '                 End If 
        '                 Dim j As Long = 22 * 張 + 位置 + 5'32*張+第幾個+初始位址
        '                 If i=0 or 變動=true'第一頁
        '                     arr(J-3,1) = 年度
        '                     arr(J-3,8) = 秘字
        '                     arr(J+8,1) = 年度
        '                     arr(J+8,8) = 秘字
        '                     arr(j, 1) = 月
        '                     arr(j, 2) = 日
        '                     arr(j, 5) = 摘要
        '                     arr(j, 6) = 支出
        '                     arr(j+11, 1) = 月
        '                     arr(j+11, 2) = 日
        '                     arr(j+11, 5) = 摘要
        '                     arr(j+11, 6) = 支出
        '                 END If
        '                 'If 過審=true
        '                 arr(j, 5) = 摘要
        '                 arr(j, 6) = 支出
        '                 arr(j+11, 5) = 摘要
        '                 arr(j+11, 6) = 支出
        '                 支出sum=支出sum+支出
        '                 If i<data_dv.Count - 1'最後一筆或換頁，輸出資料，但非最後一筆且不換頁，則不動 ---'8/24，第一筆的總和在第二筆
        '                     If data_dv(i)("號數").ToString()<>data_dv(i+1)("號數").ToString() or 位置=3'號碼有變動，換頁
        '                         arr(j+4-位置,6) = "=SUM(" & arr2(j+3-位置,6).Address & ":" & arr2(j-位置,6).Address & ")"
        '                         arr(j+9+5-位置,6) = "=SUM(" & arr2(j+9+4-位置,6).Address & ":" & arr2(j+9+2-位置,6).Address & ")"
        '                     End If 
        '                 Else'最後一筆
        '                     arr(j+4-位置,6) = "=SUM(" & arr2(j+3-位置,6).Address & ":" & arr2(j-位置,6).Address & ")"
        '                     arr(j+9+5-位置,6) = "=SUM(" & arr2(j+9+4-位置,6).Address & ":" & arr2(j+5-位置,6).Address & ")"
        '                 END If
        '             Next
        '             xlWorkSheet.Range(xlWorkSheet.Cells(1, 1), xlWorkSheet.Cells((data_dv2.Count+no) * 22, 8)).Value = arr
        '             xlWorkBook.Save()
        '         End If 
        '         xlWorkBook.Close()
        '         xlApp.Quit()
        '         ReleaseObject(xlWorkSheet)
        '         ReleaseObject(xlWorkBook)
        '         ReleaseObject(xlApp)
        '         Response.Clear()
        '         Response.ClearHeaders()
        '         Response.Buffer = True
        '         Response.ContentType = "application/octet-stream"
        '         Dim downloadfilename = "暫付轉正-秘XZ.xlsx"
        '         Response.AddHeader("Content-Disposition", "attachment;FileName=" + Uri.EscapeDataString(downloadfilename))
        '         Response.WriteFile(MyExcel)
        '         System.Web.HttpContext.Current.ApplicationInstance.CompleteRequest()
        '         Response.Flush()
        '         System.IO.File.Delete(MyExcel)
        '         Response.End()
        '     End If 
    End Sub
    Protected Sub test(ByVal sender As Object, ByVal e As System.EventArgs)
    End Sub
    Protected Sub 全選_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
            For i=0 to Me.GridView1.Rows.Count - 1
                    Dim 勾選_CheckBox As CheckBox=CType(Me.GridView1.Rows(i).FindControl("勾選"), CheckBox)
                If 勾選_CheckBox.Enabled=True
                    勾選_CheckBox.Checked=全選.Checked
                End If 
                CType(Me.GridView1.Rows(i).FindControl("勾選下載"), CheckBox).Checked=全選.Checked
            Next
    End Sub
    Protected Sub send(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim 作用 As boolean = False
        Dim 種類 As String = Me._種類.Text
        For i=0 to Me.GridView1.Rows.Count - 1
            Dim id As String = CType(Me.GridView1.Rows(i).FindControl("id"), Label).Text
            Dim 號數 As String = CType(Me.GridView1.Rows(i).FindControl("號數"), Label).Text
            If CType(Me.GridView1.Rows(i).FindControl("勾選"), CheckBox).Checked=True And CType(Me.GridView1.Rows(i).FindControl("回覆R"), RadioButtonList).SelectedIndex=1 And CType(Me.GridView1.Rows(i).FindControl("主計室日期"), TextBox).Text=""
                作用=True
               '更新日誌
                Dim data_dv2 As Data.DataView
                data.SelectCommAnd = "Select id From 收支備查簿 " & _
                "Where 送交主計室日期 IS NULL AND 過審 = 'True'AND 駁回原因 IS NULL AND 支出>0 AND 號數 = '" & 號數 & "'AND _種類 = '" & 種類 & "'"
                data_dv2 = data.Select(New DataSourceSelectArguments)
                For k=0 to data_dv2.count-1
                    data.insertCommAnd = _
                        "INSERT INTO 日誌 " & _
                        "(id, 動作,日期,日期2) " & _
                        "VALUES " & _
                        "(N'" & data_dv2(k)("id").ToString() & "', N'送交主計室', N'" & DateTime.now.tostring() & "' , '" & DateTime.now.tostring("yyyy-MM-dd HH:mm:ss") & "')"
                    data.insert()
                Next
                '送交資料
                data.UpdateCommAnd = "UPDATE 收支備查簿 SET " & _
                    "送交主計室日期 = N'" & datetime.Today() & "'" & _
                    "Where 送交主計室日期 IS NULL AND 過審 = 'True'AND 駁回原因 IS NULL AND 支出>0 AND 號數 = '" & 號數 & "'AND _種類 = '" & 種類 & "'"
                data.Update()
                '--
            End If
        Next
        '作用=1，會出現bug，必須用True
        If 作用=True
            Me.GridView1.DataBind()
            Label1.Text="已送出成功"
        Else
            label2.Text="請先勾選欲送交主計室之資料"
        End If
    End Sub
    Protected Sub 勾選_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        CheckedChanged(sender)
    End Sub
    Protected Sub 勾選下載_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        CheckedChanged(sender)
    End Sub
    '將勾選_CheckedChanged和勾選下載_CheckedChanged合併，目的在於將所有同號數的選取方塊一起作用
    Protected Sub CheckedChanged(ByVal CheckBox As CheckBox)
        '下面三段為取審核_CheckedChanged在GridView的位置
        Dim row As GridViewRow = CheckBox.NamingContainer
        Dim index As Integer = row.RowIndex
        Dim i As Long = index
        Dim 號數 As String = CType(Me.GridView1.Rows(i).FindControl("號數"), Label).Text
        If 號數<>""
                For j = 0 to Me.GridView1.Rows.Count - 1
                    If CType(Me.GridView1.Rows(j).FindControl("號數"), Label).Text=號數
                        CType(Me.GridView1.Rows(j).FindControl(CheckBox.id), CheckBox).Checked=CType(Me.GridView1.Rows(i).FindControl(CheckBox.id), CheckBox).Checked
                    End If 
                Next
        End If 
    End Sub
    Protected Sub 年_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles 年.TextChanged
        If Me.年.Text=""
            Me.年.Text = (DateTime.Now.Year - 1911).ToString()
        End If
        月1_SelectedIndexChanged(sender,e)
        月2_SelectedIndexChanged(sender,e)
    End Sub
    Protected Sub 月1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles 月1.SelectedIndexChanged
        GetDay(月1,日1)
    End Sub
    Protected Sub 月2_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles 月2.SelectedIndexChanged
        GetDay(月2,日2)
    End Sub
    Protected Sub Clear_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Response.Redirect(Request.Url.ToString())
    End Sub
    Protected Sub GetDay(ByVal month As Object,ByVal day As Object)'以月取日，收尋，日可不留白
        If month.SelectedValue<>"" AND Me.年.text<>""
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
End Class