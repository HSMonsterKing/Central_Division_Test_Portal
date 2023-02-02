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
Partial Class 水質檢驗
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
        'Me.狀態.text=""
        Update(sender, e)
        For i = 1 To 15
            Dim insert1 as string
            data.InsertCommand = _
            "INSERT INTO 水質檢驗表 " & _
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
            'Dim 編號 As String = CType(Me.GridView1.Rows(i).FindControl("編號"),Label).Text
            Dim 檢驗週期 As String = CType(Me.GridView1.Rows(i).FindControl("檢驗週期"),TextBox).Text
            ' Dim 上次檢驗 As String = nothing
            ' If CType(Me.GridView1.Rows(i).FindControl("上次檢驗"), TextBox).text<>""
            '     上次檢驗 = CType(Me.GridView1.Rows(i).FindControl("上次檢驗"), TextBox).text
            '     上次檢驗 = taiwancalendarto(上次檢驗)
            ' End If
            '增加將檢驗日期自動轉成上次檢驗
            Dim 檢驗日期 As String = nothing
            If CType(Me.GridView1.Rows(i).FindControl("檢驗日期"), TextBox).text<>""
                檢驗日期 = CType(Me.GridView1.Rows(i).FindControl("檢驗日期"), TextBox).text & "-01"
                檢驗日期=檢驗日期.Replace("-","/")
                檢驗日期 = taiwancalendarto(檢驗日期)
                ' 上次檢驗 = 檢驗日期
            End If
            Dim 檢驗地點 As String = CType(Me.GridView1.Rows(i).FindControl("檢驗地點"), TextBox).Text
            Dim 檢驗項目 As String = CType(Me.GridView1.Rows(i).FindControl("檢驗項目"), TextBox).Text
            Dim Update1 as string ="UPDATE 水質檢驗表 SET " & _
            "檢驗週期 = NULLIF(N'" & 檢驗週期 & "', ''), " & _
            "檢驗地點 = NULLIF(N'" & 檢驗地點 & "', ''), " & _
            "檢驗日期 = IIF(ISDATE(TRIM(N'" & 檢驗日期 & "'))=1,TRIM(N'" & 檢驗日期 & "'),NULL), " & _
            "檢驗項目 = NULLIF(N'" & 檢驗項目 & "', '') " & _
            "WHERE id = '" & id & "'"
            '"上次檢驗 = IIF(ISDATE(TRIM(N'" & 上次檢驗 & "'))=1,TRIM(N'" & 上次檢驗 & "'),NULL), " & _
            data.UpdateCommand = Update1
            data.Update()
            ' If CType(Me.GridView1.Rows(i).FindControl("編號"), Label).text=""
            '      Dim Update2 as string ="WITH CTE AS (Select *,Row_Number() OVER(order by ID ) AS '序號' From 水質檢驗表)" & _
            '     "UPDATE CTE SET 編號 = 序號 Where 檢驗地點 IS NOT NULL"
            '     data.UpdateCommand = Update2
            '     data.Update()
            ' End If
        Next
        Me.GridView1.DataBind()
    End Sub
    Protected Sub Delete(ByVal sender As Object, ByVal e As System.EventArgs)
        'Me.狀態.text=""
        Update(sender, e)
        Me.GridView1.PageIndex = Int32.MaxValue
        Me.GridView1.DataBind()
        Dim delete1 as string=""
        For i = 0 To Me.GridView1.Rows.Count - 1
            Dim id As String = CType(Me.GridView1.Rows(i).FindControl("id"), TextBox).Text
            delete1="DELETE FROM 水質檢驗表 " & _
            "WHERE id = '" & id & "'"
            data.deleteCommand =delete1
            data.delete()
        Next
        Me.GridView1.DataBind()
    End Sub
    Protected Sub Test(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim 檢驗日期 As String = nothing
        檢驗日期 = CType(Me.GridView1.Rows(0).FindControl("檢驗日期"), TextBox).text & "-01"
        檢驗日期=檢驗日期.Replace("-","/")
        label1.text=label1.text+檢驗日期
    End Sub
    Protected Sub Download(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim MyGUID As String = Guid.NewGuid().ToString("N")
        Dim MyExcel As String = MapPath(".\Excel\Temp\") & MyGUID & ".xlsx"
        System.IO.File.Copy(MapPath(".\Excel\水質檢驗單.xlsx"), MyExcel)
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
        xlWorkSheet = CType(xlWorkBook.Sheets("水質檢驗單"), Excel.Worksheet)
        xlWorkSheet.Activate()
        data.ConnectionString = con_14
        'Dim 狀態 As string = Me.狀態.Text
        data.SelectCommAnd = "SELECT * FROM 水質檢驗表 Where 檢驗地點 Is Not NULL ORDER BY _頁,_列"
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
        D_Width=7'資料範圍行
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
            Dim 檢驗週期 As String
            ' Dim 上次檢驗 As String
            ' Dim 下次檢驗 As String
            ' Dim 需檢驗 As Boolean = False
            Dim 檢驗日期 As String
            Dim 檢驗地點 As String
            Dim 檢驗項目 As String
            Dim j As Long '輸出位置
            If i < data_dv.count
                If data_dv.count>Data_Row
                    j = ((i) \ Data_Row )*16 + ((i) Mod Data_Row) + 2 '從第一列輸出
                Else
                    j = (i Mod Data_Row) + 2
                End If 
                編號 = data_dv(i)("編號").ToString()
                檢驗週期 = Trim(data_dv(i)("檢驗週期").ToString())
                ' 上次檢驗 = data_dv(i)("上次檢驗").ToString()
                ' If IsDate(上次檢驗) AND NOT(檢驗週期 IS DBNull.Value)
                '     下次檢驗 =(Year(CDate(DATEADD("m",CInt(left(檢驗週期,1)),CDate(上次檢驗))))-1911).ToString() & "/" & (CDate(DATEADD("m",CInt(left(檢驗週期,1)),CDate(上次檢驗)))).ToString("MM/dd")
                '     If DateDiff("m",DateTime.Now.ToString("yyyy-MM-dd"),CDate(DATEADD("m",CInt(left(檢驗週期,1)),CDate(上次檢驗))))<1
                '         需檢驗=True
                '     End IF
                ' Else
                '     下次檢驗 =""
                ' END If
                ' 上次檢驗 = ToTaiwanCalendar(上次檢驗)
                檢驗日期 = data_dv(i)("檢驗日期").ToString()
                If IsDate(檢驗日期)
                    檢驗日期 = Year(ToTaiwanCalendar(檢驗日期)) & "-" & Month(ToTaiwanCalendar(檢驗日期))
                END If
                檢驗地點 = Trim(data_dv(i)("檢驗地點").ToString())
                檢驗地點 = 檢驗地點.Replace(vbCrLf,"")
                檢驗地點 = Trim(檢驗地點)
                檢驗項目 = Trim(data_dv(i)("檢驗項目").ToString())
                檢驗項目 = 檢驗項目.Replace(vbCrLf,"")
                檢驗項目 = Trim(檢驗項目)
                arr(j, 1) = 檢驗週期
                ' arr(j, 3) = 上次檢驗
                ' arr(j, 4) = 下次檢驗
                ' If 需檢驗=True AND Me.需檢驗資料.Checked = True
                '     arr2(j, 4).Interior.ColorIndex = 3 '背景色，紅色
                ' End If
                arr(j, 2) = 檢驗日期
                arr(j, 3) = 檢驗地點
                arr(j, 4) = 檢驗項目
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
        Dim downloadfilename = "水質檢驗單.xlsx"
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
            data.UpdateCommand = "Update 水質檢驗表 Set " & _
            "檢驗週期 = NULL," & _
            "檢驗地點 = NULL, " & _
            "檢驗日期 = NULL," & _
            "檢驗項目 = NULL " & _
            "WHERE id = '" & id & "'"
            '"編號 = NULL," & _
            '"上次檢驗 = NULL," & _
            data.Update()
            Me.GridView1.DataBind()
        End If
    End Sub 
End Class