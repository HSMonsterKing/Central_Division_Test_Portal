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
Partial Class 濾心清單
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
        For i = 1 To 15
            Dim insert1 as string
            data.InsertCommand = _
            "INSERT INTO 濾心清單表 " & _
            "( _頁, _列) " & _
            "VALUES " & _
            "('" & (Me.GridView1.PageCount + 1).ToString() & "', '" & i & "')"
            data.Insert()
        Next
        Me.GridView1.PageIndex = Int32.MaxValue
        Me.GridView1.DataBind()
    End Sub
    Protected Sub Update(ByVal sender As Object, ByVal e As System.EventArgs)
        'TODO:轉換成整數會失敗
        For i = 0 To Me.GridView1.Rows.Count - 1
            Dim id As String = CType(Me.GridView1.Rows(i).FindControl("id"), TextBox).Text
            Dim 編號 As String = CType(Me.GridView1.Rows(i).FindControl("編號"),Label).Text
            Dim 項目 As String = CType(Me.GridView1.Rows(i).FindControl("項目"),TextBox).Text
            Dim 品名型號 As String = CType(Me.GridView1.Rows(i).FindControl("品名型號"),TextBox).Text
            Dim 更換週期 As String = CType(Me.GridView1.Rows(i).FindControl("更換週期"),TextBox).Text
            Dim 上次更換 As String = nothing
            If CType(Me.GridView1.Rows(i).FindControl("上次更換"), TextBox).text<>""
                上次更換 = CType(Me.GridView1.Rows(i).FindControl("上次更換"), TextBox).text
                上次更換 = 上次更換 & "/01"
                上次更換 = taiwancalendarto(上次更換)
            End If
            '增加將更換日期自動轉成上次更換
            Dim 更換日期 As String = nothing
            If CType(Me.GridView1.Rows(i).FindControl("更換日期"), TextBox).text<>""
                更換日期 = CType(Me.GridView1.Rows(i).FindControl("更換日期"), TextBox).text
                更換日期 = taiwancalendarto(更換日期)
                If IsDate(上次更換) AND IsDate(更換日期)
                    If Year(上次更換) <> Year(更換日期) AND Month(上次更換) <> Month(更換日期)
                    Dim insert1 as string = _
                        "INSERT INTO 濾心更換_日誌 " & _
                        "(濾心ID, 更換日期) " & _
                        "VALUES " & _
                        "('" & id & "', '" & 更換日期 & "')"
                        data.InsertCommand = insert1
                        data.Insert()
                    End If
                End If
                上次更換 = 更換日期
            End If
            Dim 單價 As String = CType(Me.GridView1.Rows(i).FindControl("單價"), TextBox).Text
            單價=單價.Replace(",", "")
            Dim 數量 As String = CType(Me.GridView1.Rows(i).FindControl("數量"), TextBox).Text
            數量=數量.Replace(",", "")
            Dim 更換地點 As String = CType(Me.GridView1.Rows(i).FindControl("更換地點"), TextBox).Text
            Dim Update1 as string ="UPDATE 濾心清單表 SET " & _
            "編號 = NULLIF(N'" & 編號 & "', ''), " & _
            "項目 = NULLIF(N'" & 項目 & "', ''), " & _
            "品名型號 = NULLIF(N'" & 品名型號 & "', ''), " & _
            "更換週期 = IIF(ISNUMERIC(Replace(TRIM(N'" & 更換週期 & "'),'個月',''))=1,TRIM(N'" & 更換週期 & "'),NULL), " & _
            "上次更換 = IIF(ISDATE(TRIM(N'" & 上次更換 & "'))=1,TRIM(N'" & 上次更換 & "'),NULL), " & _
            "更換日期 = IIF(ISDATE(TRIM(N'" & 更換日期 & "'))=1,TRIM(N'" & 更換日期 & "'),NULL), " & _
            "單價 = NULLIF(N'" & 單價 & "', ''), " & _
            "數量 = NULLIF(N'" & 數量 & "', ''), " & _
            "更換地點 = NULLIF(N'" & 更換地點 & "', '') " & _
            "WHERE id = '" & id & "'"
            data.UpdateCommand = Update1
            data.Update()
            If CType(Me.GridView1.Rows(i).FindControl("編號"), Label).text=""
                 Dim Update2 as string ="WITH CTE AS (Select *,Row_Number() OVER(order by ID ) AS '序號' From 濾心清單表)" & _
                "UPDATE CTE SET 編號 = 序號 Where 項目 IS NOT NULL"
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
            delete1="DELETE FROM 濾心清單表 " & _
            "WHERE id = '" & id & "'"
            data.deleteCommand =delete1
            data.delete()
        Next
        Me.GridView1.DataBind()
    End Sub
    Protected Sub Test(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim 更換週期 As String="2個月"
        Dim 上次更換 As String="2022-09-01"
        label1.text=label1.text & DateDiff("m",CDate(DATEADD("m",CInt(left(更換週期,1)),CDate(上次更換))),DateTime.Now.ToString("yyyy-MM-dd"))
        'label1.text=label1.text+DateDiff("m",CDate(DATEADD("m",CInt(left(更換週期,1)),CDate(上次更換))),DateTime.Now.ToString("yyyy-MM-dd"))
    End Sub
    Protected Sub Download(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim MyGUID As String = Guid.NewGuid().ToString("N")
        Dim MyExcel As String = MapPath(".\Excel\Temp\") & MyGUID & ".xlsx"
        System.IO.File.Copy(MapPath(".\Excel\濾心報價單.xlsx"), MyExcel)
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
        xlWorkSheet = CType(xlWorkBook.Sheets("濾心報價單"), Excel.Worksheet)
        xlWorkSheet.Activate()
        data.ConnectionString = con_14
        Dim 狀態 As string = Me.狀態.Text
        data.SelectCommAnd = "SELECT * FROM 濾心清單表 Where 編號 Is Not NULL AND (REPLACE(更換週期,'個月','')-DateDiff(m,上次更換,GETDATE())) in (select min(REPLACE(更換週期,'個月','')-DateDiff(m,上次更換,GETDATE())) FROM 濾心清單表) ORDER BY 編號"
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
        D_Height=17'資料範圍列
        D_Width=11'資料範圍行
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
        Dim 總頁數 As Int32 = 0
        Dim Data_Row As Int32 =15'平均資料列
        '總資料數
        IF data_dv.Count>Data_Row
            If ((data_dv.Count) Mod Data_Row)>0  '8
                總頁數=((data_dv.Count)\Data_Row)+1'"/"為整數除法，會四捨五入，"\"為除法，捨棄餘數
            ELSE
                總頁數=(data_dv.Count)\Data_Row
            END If
        Else
            總頁數=1
        END If
        For i = 2 To 總頁數'制定範圍並複製，範圍為頁數
            xlWorkSheet.Range(xlWorkSheet.Cells(D_Height * i - (D_Height-1), 1), xlWorkSheet.Cells(D_Height * i, D_Width)).Value(11) = xlWorkSheet.Range(xlWorkSheet.Cells(1, 1), xlWorkSheet.Cells(D_Height, D_Width)).Value(11)'(27,1) (52,10) = (1, 1) (26, 10)
            xlWorkSheet.Range(xlWorkSheet.Cells(D_Height * i - Data_Height_H, 1), xlWorkSheet.Cells(D_Height * i - Data_Height_L, D_Width)).RowHeight = 44'(33,1) (47,10) 高度為33 7~21
            xlWorkSheet.Rows(D_Height * i - (D_Height-1)).PageBreak = xlPageBreakManual'列23從開始載入
        Next
        Dim arr As Object = xlWorkSheet.Range(xlWorkSheet.Cells(1, 1), xlWorkSheet.Cells(總頁數 * D_Height, D_Width)).Value'(1,1) (網頁頁數25,10)
        Dim arr2 As Object = xlWorkSheet.Range(xlWorkSheet.Cells(1, 1), xlWorkSheet.Cells(總頁數 * D_Height, D_Width))
        Dim 總價款位置 As String =""
        For i = 0 To data_dv.Count-1'DataCount=10
            Dim 編號 As String 
            Dim 項目 As String 
            Dim 品名型號 As String
            Dim 更換週期 As String
            Dim 上次更換 As String
            Dim 下次更換 As String
            Dim 需更換 As Boolean = False
            Dim 更換日期 As String 
            Dim 單價 As String
            Dim 數量 As String
            Dim 合計 As String =""
            Dim 更換地點 As String
            Dim 總價款中文 As String
            Dim j As Long '輸出位置
            If i < data_dv.count
                If data_dv.count>Data_Row
                    j = ((i) \ Data_Row )*17 + ((i) Mod Data_Row) + 2 '從第一列輸出
                Else
                    j = (i Mod Data_Row) + 2
                End If 
                編號 = data_dv(i)("編號").ToString()
                項目 = Trim(data_dv(i)("項目").ToString())
                品名型號 = Trim(data_dv(i)("品名型號").ToString())
                品名型號 = 品名型號.Replace(vbCrLf,"")
                品名型號 = Trim(品名型號)
                更換週期 = Trim(data_dv(i)("更換週期").ToString())
                上次更換 = data_dv(i)("上次更換").ToString()
                If IsDate(上次更換) AND NOT(更換週期 IS DBNull.Value)
                    下次更換 =(Year(CDate(DATEADD("m",CInt(Replace(更換週期,"個月","")),CDate(上次更換))))-1911).ToString() & "/" & (Month(CDate(DATEADD("m",CInt(Replace(更換週期,"個月","")),CDate(上次更換))))).ToString()
                    If DateDiff("m",DateTime.Now.ToString("yyyy-MM-dd"),CDate(DATEADD("m",CInt(Replace(更換週期,"個月","")),CDate(上次更換))))<1
                        需更換=True
                    End IF
                Else
                    下次更換 =""
                END If
                If IsDate(上次更換)
                    上次更換 = Year(ToTaiwanCalendar(上次更換)) & "/" & Month(ToTaiwanCalendar(上次更換))
                END If
                更換日期 = data_dv(i)("更換日期").ToString()
                更換日期 = ToTaiwanCalendar(更換日期)
                單價 = data_dv(i)("單價").ToString()
                單價 = If (單價 = "", "", CLng(單價).ToString("n0"))
                數量 = data_dv(i)("數量").ToString()
                數量 = If (數量 = "", "", CLng(數量).ToString("n0"))
                ' If data_dv(i)("單價").ToString()<>"" AND data_dv(i)("數量").ToString()<>""
                '     合計 = data_dv(i)("單價").ToString()*data_dv(i)("數量").ToString()
                '     合計 = If (合計 = "", "", CLng(合計).ToString("n0"))
                ' Else
                '     合計 = ""
                ' END If
                更換地點 = Trim(data_dv(i)("更換地點").ToString())
                更換地點 = 更換地點.Replace(vbCrLf,"")
                更換地點 = Trim(更換地點)
                arr(j, 1) = 編號
                arr(j, 2) = 項目
                arr(j, 3) = 品名型號
                arr(j, 4) = 更換週期
                arr(j, 5) = 上次更換
                arr(j, 6) = 下次更換
                'If 需更換=True AND Me.需更換資料.Checked = True
                '    arr2(j, 6).Interior.ColorIndex = 3 '背景色，紅色
                'End If
                'arr(j, 7) = 更換日期
                arr(j, 8) = 單價
                arr(j, 9) = 數量
                arr(j, 10) = "=" & arr2(j, 8).Address & "*" & arr2(j, 9).Address
                ' If 合計<>""
                    If 總價款位置=""
                        總價款位置="=" & arr2(j, 10).Address
                    Else
                        總價款位置=總價款位置 & "+" & arr2(j, 10).Address
                    End If
                ' End If
                arr(j, 11) = 更換地點
            End If
            If(i Mod Data_Row=14) OR (i=data_dv.count-1)
                j=((i\Data_Row)+1)*17
                總價款中文=總價款位置
                arr(j, 3)="新臺幣:"
                arr(j, 4)=總價款中文
                arr(j, 6)="元整(含稅價)"
                If 總價款位置<>""
                    總價款位置=總價款位置 & "&" & Chr(34)  & "元" & Chr(34)
                End If
                arr(j, 11) = 總價款位置
                '總價款位置=""
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
        Dim downloadfilename = "濾心報價單.xlsx"
        Response.AddHeader("Content-Disposition", "attachment;FileName=" + Uri.EscapeDataString(downloadfilename))
        Response.WriteFile(MyExcel)
        System.Web.HttpContext.Current.ApplicationInstance.CompleteRequest()
        Response.Flush()
        System.IO.File.Delete(MyExcel)
        Response.End()
    End Sub
    Protected Sub Download2(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim MyGUID As String = Guid.NewGuid().ToString("N")
        Dim MyExcel As String = MapPath(".\Excel\Temp\") & MyGUID & ".xlsx"
        System.IO.File.Copy(MapPath(".\Excel\濾心更換單.xlsx"), MyExcel)
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
        xlWorkSheet = CType(xlWorkBook.Sheets("濾心更換單"), Excel.Worksheet)
        xlWorkSheet.Activate()
        data.ConnectionString = con_14
        Dim 狀態 As string = Me.狀態.Text
        data.SelectCommAnd = "SELECT * FROM 濾心清單表 Where 編號 Is Not NULL AND (REPLACE(更換週期,'個月','')-DateDiff(m,上次更換,GETDATE())) in (select min(REPLACE(更換週期,'個月','')-DateDiff(m,上次更換,GETDATE())) FROM 濾心清單表) ORDER BY 編號"
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
        D_Height=16'資料範圍列
        D_Width=12'資料範圍行
        Data_Height_H=15'資料最高
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
        Dim 總頁數 As Int32 = 0
        Dim Data_Row As Int32 =15
        '總資料數
        IF data_dv.Count>Data_Row
            If ((data_dv.Count) Mod Data_Row)>0  '8
                總頁數=((data_dv.Count)\Data_Row)+1'"/"為整數除法，會四捨五入，"\"為除法，捨棄餘數
            ELSE
                總頁數=(data_dv.Count)\Data_Row
            END If
        Else
            總頁數=1
        END If
        For i = 2 To 總頁數'制定範圍並複製，範圍為頁數
            xlWorkSheet.Range(xlWorkSheet.Cells(D_Height * i - (D_Height-1), 1), xlWorkSheet.Cells(D_Height * i, D_Width)).Value(11) = xlWorkSheet.Range(xlWorkSheet.Cells(1, 1), xlWorkSheet.Cells(D_Height, D_Width)).Value(11)'(27,1) (52,10) = (1, 1) (26, 10)
            xlWorkSheet.Range(xlWorkSheet.Cells(D_Height * i - Data_Height_H, 1), xlWorkSheet.Cells(D_Height * i - Data_Height_L, D_Width)).RowHeight = 44'(33,1) (47,10) 高度為33 7~21
            xlWorkSheet.Rows(D_Height * i - (D_Height-1)).PageBreak = xlPageBreakManual'列23從開始載入
        Next
        Dim arr As Object = xlWorkSheet.Range(xlWorkSheet.Cells(1, 1), xlWorkSheet.Cells(總頁數 * D_Height, D_Width)).Value'(1,1) (網頁頁數25,10)
        Dim arr2 As Object = xlWorkSheet.Range(xlWorkSheet.Cells(1, 1), xlWorkSheet.Cells(總頁數 * D_Height, D_Width))
        For i = 0 To data_dv.Count-1'DataCount=10
            Dim 編號 As String 
            Dim 項目 As String 
            Dim 品名型號 As String
            Dim 更換週期 As String
            Dim 上次更換 As String
            Dim 下次更換 As String
            Dim 需更換 As Boolean = False
            'Dim 更換日期 As String 
            Dim 單價 As String
            Dim 數量 As String
            Dim 合計 As String
            Dim 更換地點 As String
            Dim j As Long '輸出位置
            If i < data_dv.count
                If data_dv.count>Data_Row
                    j = ((i) \ Data_Row )*16 + ((i) Mod Data_Row) + 2 '從第一列輸出
                Else
                    j = (i Mod Data_Row) + 2
                End If 
                編號 = (i Mod Data_Row)+1'data_dv(i)("編號").ToString()
                項目 = Trim(data_dv(i)("項目").ToString())
                品名型號 = Trim(data_dv(i)("品名型號").ToString())
                品名型號 = 品名型號.Replace(vbCrLf,"")
                品名型號 = Trim(品名型號)
                更換週期 = Trim(data_dv(i)("更換週期").ToString())
                上次更換 = data_dv(i)("上次更換").ToString()
                If IsDate(上次更換) AND NOT(更換週期 IS DBNull.Value)
                    下次更換 =(Year(CDate(DATEADD("m",CInt(Replace(更換週期,"個月","")),CDate(上次更換))))-1911).ToString() & "/" & (Month(CDate(DATEADD("m",CInt(Replace(更換週期,"個月","")),CDate(上次更換))))).ToString()
                    If DateDiff("m",DateTime.Now.ToString("yyyy-MM-dd"),CDate(DATEADD("m",CInt(Replace(更換週期,"個月","")),CDate(上次更換))))<1
                        需更換=True
                    End IF
                Else
                    下次更換 =""
                END If
                If IsDate(上次更換)
                    上次更換 = Year(ToTaiwanCalendar(上次更換)) & "/" & Month(ToTaiwanCalendar(上次更換))
                END If
                '更換日期 = data_dv(i)("更換日期").ToString()
                '更換日期 = ToTaiwanCalendar(更換日期)
                單價 = data_dv(i)("單價").ToString()
                單價 = If (單價 = "", "", CLng(單價).ToString("n0"))
                數量 = data_dv(i)("數量").ToString()
                數量 = If (數量 = "", "", CLng(數量).ToString("n0"))
                If data_dv(i)("單價").ToString()<>"" AND data_dv(i)("數量").ToString()<>""
                    合計 = 單價*數量
                    合計 = If (合計 = "", "", CLng(合計).ToString("n0"))
                Else
                    合計 = ""
                END If
                更換地點 = Trim(data_dv(i)("更換地點").ToString())
                更換地點 = 更換地點.Replace(vbCrLf,"")
                更換地點 = Trim(更換地點)
                arr(j, 1) = 編號
                arr(j, 2) = 項目
                arr(j, 3) = 品名型號
                arr(j, 4) = 更換週期
                arr(j, 5) = 上次更換
                arr(j, 6) = 下次更換
                ' If 需更換=True AND Me.需更換資料.Checked = True
                '     arr2(j, 6).Interior.ColorIndex = 3 '背景色，紅色
                ' End If
                'arr(j, 7) = 更換日期
                arr(j, 8) = 單價
                arr(j, 9) = 數量
                arr(j, 10) = 合計
                arr(j, 11) = "□符合"& vbCrLf &"□不符合"
                arr(j, 12) = 更換地點
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
        Dim downloadfilename = "濾心更換單.xlsx"
        Response.AddHeader("Content-Disposition", "attachment;FileName=" + Uri.EscapeDataString(downloadfilename))
        Response.WriteFile(MyExcel)
        System.Web.HttpContext.Current.ApplicationInstance.CompleteRequest()
        Response.Flush()
        System.IO.File.Delete(MyExcel)
        Response.End()
    End Sub
    Protected Sub GridView1_RowCommand(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles GridView1.RowCommand
        If e.CommandName = "刪除"
            Update(sender, e)
            Dim i As Long = e.CommandSource.NamingContainer.RowIndex
            Dim id As String = CType(Me.GridView1.Rows(i).FindControl("id"), TextBox).Text
            data.UpdateCommand = "Update 濾心清單表 Set " & _
            "編號 = NULL," & _
            "項目 = NULL," & _
            "品名型號 = NULL," & _
            "更換週期 = NULL," & _
            "上次更換 = NULL," & _
            "更換日期 = NULL," & _
            "單價 = NULL," & _
            "數量 = NULL," & _
            "更換地點 = NULL " & _
            "WHERE id = '" & id & "'"
            data.Update()
            Me.GridView1.DataBind()
        End If
    End Sub 
End Class