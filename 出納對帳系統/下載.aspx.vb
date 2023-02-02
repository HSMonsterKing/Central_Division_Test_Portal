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
Partial Class 下載
    Inherits System.Web.UI.Page
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim MyTXT as String = MapPath(".\Excel\Temp\") & Request.QueryString("file")
        
        Response.Clear()
        Response.ClearHeaders()
        Response.Buffer = True
        Response.ContentType = "application/octet-stream"
        Dim downloadfilename = Request.QueryString("downloadfilename")
        Response.AddHeader("Content-Disposition", "attachment;FileName=" + Uri.EscapeDataString(downloadfilename))
        Response.WriteFile(MyTXT)
        System.Web.HttpContext.Current.ApplicationInstance.CompleteRequest()
        Response.Flush()
        File.Delete(MyTXT)
        Response.End()
    End Sub
End Class