Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel.XlDirection
Imports Microsoft.Office.Interop.Excel.XlPageBreak
Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.VisualBasic.Logging
Imports System.IO
Imports System.Math
Imports System.Drawing
Imports System.Diagnostics
Imports System.Data.Sql
Imports System.Data.SqlTypes
Imports System.Data.SqlClient
Imports System.Configuration
Imports System.Text.RegularExpressions
Partial Class 例行故障維修
    Inherits System.Web.UI.Page
    Dim con_14 As String = System.Web.Configuration.WebConfigurationManager.ConnectionStrings("ApplicationServices").ConnectionString
    Dim data As New SqlDataSource
    Dim data_dv As Data.DataView
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        data.ConnectionString = con_14
        If Not Page.IsPostBack Then
            Me.GridView1.PageIndex = Int32.MaxValue
        Else
            Me.Label1.Text = ""
            Me.Label2.Text = ""
        End If
    End Sub
    Protected Sub Insert(ByVal sender As Object, ByVal e As System.EventArgs)
        Update(sender, e)
        Dim RowIndex As Integer=15
        For i = 1 To 15
            data.InsertCommand = _
            "INSERT INTO 例行故障維修表 " & _
            "(_頁,_列,編號,維修報修內容,處理情形) " & _
            "VALUES " & _
            "(" & (Me.GridView1.PageCount + 1).ToString() & ", '"& i &"', '"& i &"','','')"
            data.Insert()
        Next
        Me.GridView1.PageIndex = Int32.MaxValue
        Me.GridView1.DataBind()
    End Sub
    Protected Sub Update(ByVal sender As Object, ByVal e As System.EventArgs)
        'TODO:轉換成整數會失敗
        For i = 0 To Me.GridView1.Rows.Count - 1
            Dim id As String = CType(Me.GridView1.Rows(i).FindControl("id"), TextBox).Text
            Dim 編號 As String = CType(Me.GridView1.Rows(i).FindControl("編號"), TextBox).Text
            Dim 故障日期 As String = CType(Me.GridView1.Rows(i).FindControl("故障日期"), TextBox).Text
            If CType(Me.GridView1.Rows(i).FindControl("故障日期"), TextBox).text<>""
                故障日期 = CType(Me.GridView1.Rows(i).FindControl("故障日期"), TextBox).text
                故障日期 = TaiwanCalendarTo(故障日期)
            End If
            Dim 維修報修內容 As String = CType(Me.GridView1.Rows(i).FindControl("維修報修內容"), TextBox).Text
            Dim 處理情形 As String = CType(Me.GridView1.Rows(i).FindControl("處理情形"), TextBox).Text
            data.UpdateCommand = "UPDATE 例行故障維修表 SET " & _
            "編號 = NULLIF(TRIM(N'" & 編號 & "'), ''), " & _
            "故障日期 = IIF(ISDATE(TRIM(N'" & 故障日期 & "'))=1,TRIM(N'" & 故障日期 & "'),NULL), " & _
            "維修報修內容 = TRIM(N'" & 維修報修內容 & "'), " & _
            "處理情形 = TRIM(N'" & 處理情形 & "') " & _
            "WHERE id = '" & id & "'"
            data.Update()
        Next
        Me.GridView1.DataBind()
    End Sub
    Protected Sub Delete(ByVal sender As Object, ByVal e As System.EventArgs)
        Update(sender, e)
        Me.GridView1.PageIndex = Int32.MaxValue'跳至最後一頁
        Me.GridView1.DataBind()
        Dim delete1 as string=""
        For i = 0 To Me.GridView1.Rows.Count - 1
            Dim id As String = CType(Me.GridView1.Rows(i).FindControl("id"), TextBox).Text
            delete1="DELETE FROM 例行故障維修表 " & _
            "WHERE id = '" & id & "'"
            data.deleteCommand =delete1
            data.delete()
        Next
        Me.GridView1.DataBind()
    End Sub
    Protected Sub Download(ByVal sender As Object, ByVal e As System.EventArgs)'下載
        Dim MyGUID As String = Guid.NewGuid().ToString("N")
        Dim MyExcel As String = MapPath(".\Excel\Temp\") & MyGUID & ".xlsx"
        System.IO.File.Copy(MapPath(".\Excel\水電業務報表.xlsx"), MyExcel)
        Dim xlApp As New Excel.ApplicationClass()
        xlApp.DisplayAlerts = False
        xlApp.ScreenUpdating = false
        xlApp.EnableEvents = false
        Dim xlWorkBook As Excel.Workbook = xlApp.Workbooks.Open(MyExcel, 0, False, 5, "", "", True, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", False, False, 0, True, 1, 0)
        Dim xlWorkSheet As Excel.Worksheet
        xlWorkSheet = CType(xlWorkBook.Sheets("水電業務報表"), Excel.Worksheet)
        xlWorkSheet.Activate()
        data.ConnectionString = con_14
        data.SelectCommAnd = "SELECT * FROM 例行故障維修表 WHERE 編號 Is Not Null order by id,編號"'全輸出，不輸出無編號
        data_dv = data.Select(New DataSourceSelectArguments)
        '3/29
        For i = 2 To (data_dv.Count/16)+1'制定範圍並複製，範圍為頁數
            xlWorkSheet.Range(xlWorkSheet.Cells(16 * i - 15 , 1), xlWorkSheet.Cells(16 * i , 7)).Value(11) = xlWorkSheet.Range(xlWorkSheet.Cells(1, 1), xlWorkSheet.Cells(16, 7)).Value(11)'(17,1) (32,12) = (1, 1) (16, 7)
            xlWorkSheet.Range(xlWorkSheet.Cells(16 * i - 14 , 1), xlWorkSheet.Cells(16 * i , 7)).RowHeight = 42'(18,1) (32,12) 高度為26
            xlWorkSheet.Rows(16 * i - 15).PageBreak = xlPageBreakManual'列27從開始載入
        Next
        Dim 印出頁數 As Int32 = 0
        IF ((data_dv.Count-1) Mod 16)>0
            印出頁數=((data_dv.Count-1)/16)+1
        ELSE
            印出頁數=(data_dv.Count-1)/16
        END IF
        Dim arr As Object = xlWorkSheet.Range(xlWorkSheet.Cells(1, 1), xlWorkSheet.Cells(印出頁數 * 16, 7)).Value'(1,1) (網頁頁數16,7)
        Dim arr2 As Object = xlWorkSheet.Range(xlWorkSheet.Cells(1, 1), xlWorkSheet.Cells(印出頁數 * 16, 7))
        Dim 上次餘額 As String
        For i = 0 To data_dv.Count - 1
            Dim 編號 As String = data_dv(i)("編號").ToString()
            Dim 故障日期 As String = data_dv(i)("故障日期").ToString()
            Dim 維修報修內容 As String = data_dv(i)("維修報修內容").ToString()
            Dim 處理情形 As String = Trim(data_dv(i)("處理情形").ToString())
            'i為第幾個資料、j為輸出的格子
            Dim j As Long = 16 * (i \ 15) + (i Mod 15) + 2'i=14、j=34，i=26、j=46
            If (i Mod 15)=0'第一筆輸出年度每頁，先留存
            END If
            arr(j, 1) = 編號
            arr(j, 2) = 故障日期
            arr(j, 3) = 維修報修內容
            arr(j, 4) = 處理情形
            'End If
            If (i Mod 15)=14 '最後一列
            END If
            If (i Mod 15)=14 And i<>data_dv.Count - 1'最後一列、最後一筆
            END If
        Next
        xlWorkSheet.Range(xlWorkSheet.Cells(1, 1), xlWorkSheet.Cells(印出頁數 * 16, 7)).Value = arr
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
        Dim downloadfilename = "水電業務報表.xlsx"
        Response.AddHeader("Content-Disposition", "attachment;FileName=" + Uri.EscapeDataString(downloadfilename))
        Response.WriteFile(MyExcel)
        System.Web.HttpContext.Current.ApplicationInstance.CompleteRequest()
        Response.Flush()
        System.IO.File.Delete(MyExcel)
        Response.End()
    End Sub
    Protected Sub GridView1_RowCommand(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles GridView1.RowCommand
    End Sub
    Protected Sub test(ByVal sender As Object, ByVal e As System.EventArgs)
    End Sub
End Class